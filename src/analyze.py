"""Rebuild all analytical tables from the original monthly WB exports."""
from pathlib import Path
import argparse
import calendar
import hashlib
import json
import shutil
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
MONTHS = {'Январь':1,'Февраль':2,'Март':3,'Апрель':4,'Май':5,'Июнь':6,'Июль':7,'Август':8,'Сентябрь':9,'Октябрь':10,'Ноябрь':11,'Декабрь':12}
MAP = {'Артикул WB':'sku','Артикул (ЦМ)':'article','Название':'name','Предмет':'category','Удаленный товар':'deleted','Рейтинг карточки':'rating','Переходы в карточку':'visits','Положили в корзину':'carts','Заказали, шт':'orders','Выкупили, шт':'buyouts','Отменили, шт':'cancels','Заказали на сумму, ₽':'ordered_rub','Выкупили на сумму, ₽':'revenue','Средняя цена, ₽':'price','Остатки склад ВБ, шт':'stock_wb','Остатки МП, шт':'stock_mp','Доля карточки в выручке':'source_share_pct','Заказали, шт (предыдущий период)':'prev_orders_export','Выкупили на сумму, ₽ (предыдущий период)':'prev_revenue_export'}

def save(df, name):
    df.to_csv(ROOT/'data'/f'{name}.csv',index=False,encoding='utf-8-sig',float_format='%.10g')

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--source',type=Path,default=ROOT/'data'/'raw')
    args=parser.parse_args()
    files=list(args.source.glob('*.xlsx'))
    if not files:
        raise SystemExit('Нет исходных XLSX. Укажите --source PATH или заполните data/raw.')
    records=[]; manifest=[]; raw={}
    for path in files:
        month,year=path.stem.split()
        period=f'{year}-{MONTHS[month]:02}'
        df=pd.read_excel(path,sheet_name='Товары',dtype={'Артикул WB':str})
        raw[period]=df
        assert set(MAP).issubset(df.columns),path.name
        df=df[list(MAP)].rename(columns=MAP)
        for col in set(MAP.values())-{'sku','article','name','category','deleted'}:
            original=df[col]
            df[col]=pd.to_numeric(original,errors='coerce')
            assert not (original.notna() & df[col].isna()).any(),(path.name,col)
        required=['orders','revenue','visits','stock_wb','stock_mp']
        assert not df[required].isna().any().any(),path.name
        assert (df[required]>=0).all().all(),path.name
        assert df.sku.notna().all() and not df.sku.duplicated().any(),path.name
        df.insert(0,'month',period)
        df['source_file']=path.name
        records.append(df)
        manifest.append({'file':path.name,'period':period,'rows':len(df),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
    facts=pd.concat(records,ignore_index=True).sort_values(['month','sku']).reset_index(drop=True)
    periods=sorted(facts.month.unique())
    assert len(periods)==11
    assert list(pd.period_range(periods[0],periods[-1],freq='M').astype(str))==periods
    latest=periods[-1]; previous=periods[-2]
    facts['stock']=facts.stock_wb+facts.stock_mp
    save(facts,'monthly_facts')
    history=facts.groupby('month').agg(skus=('sku','size'),revenue=('revenue','sum'),orders=('orders','sum'),visits=('visits','sum'),stock=('stock','sum')).reset_index()
    history['conversion']=history.orders/history.visits
    save(history,'monthly_summary')
    cur=facts[facts.month==latest].copy().sort_values(['revenue','sku'],ascending=[False,True]).reset_index(drop=True)
    cur['rank']=np.arange(1,len(cur)+1)
    total=cur.revenue.sum()
    cur['share']=cur.revenue/total
    cur['cum_share']=cur.share.cumsum()
    # Include the item crossing the threshold in the preceding class.
    before=cur.cum_share-cur.share
    cur['abc']=np.where(cur.revenue<=0,'C',np.where(before<.8,'A',np.where(before<.95,'B','C')))
    prev=facts[facts.month==previous].set_index('sku')
    cur['orders_previous']=cur.sku.map(prev.orders)
    cur['revenue_previous']=cur.sku.map(prev.revenue)
    cur['forecast']=np.ceil((cur.orders+cur.orders_previous)/2)
    cur['growth']=np.where(cur.orders_previous>0,cur.orders/cur.orders_previous-1,np.nan)
    cur['conversion']=np.where(cur.visits>0,cur.orders/cur.visits,np.nan)
    cur['turnover_proxy']=np.where(cur.stock>0,cur.orders/cur.stock,np.nan)
    days=calendar.monthrange(int(latest[:4]),int(latest[-2:]))[1]
    cur['coverage_days']=np.where(cur.orders>0,cur.stock/(cur.orders/days),np.nan)
    demand_cut=float(cur.loc[cur.orders>0,'orders'].median())
    cur['demand_class']=np.where(cur.orders>=demand_cut,'X','Z')
    cur['turnover_class']=np.where(cur.stock==0,np.where(cur.orders>0,'Нет остатка','Нет спроса'),np.where(cur.turnover_proxy>=1,'Быстрый','Медленный'))
    cur['xz']=cur.demand_class+' / '+cur.turnover_class
    # 30 days target + 20% buffer are explicit planning assumptions, not measured parameters.
    cur['target_stock']=np.ceil(cur.forecast*1.2)
    cur['stock_gap']=np.maximum(cur.target_stock-cur.stock,0)
    cur['excess_stock']=np.maximum(cur.stock-cur.target_stock,0)
    cur['recommendation']=np.select([cur.forecast.isna(),cur.deleted.eq('Да'),cur.stock_gap.gt(0),cur.excess_stock.gt(0)],['Проверить историю','Проверить удаление','Пополнить','Снизить закупку'],default='Поддерживать')
    # Reliability caveat applies to every row, not only the portfolio total.
    cur['stock_months']=cur.sku.map(facts.groupby('sku').stock.nunique())
    cur['stock_warning']=np.where(cur.stock_months==1,'Остаток неизменен за историю','Проверить дату остатка')
    weighted_conversion=float(cur.orders.sum()/cur.visits.sum())
    cur['low_conversion']=cur.visits.ge(1000)&cur.conversion.lt(weighted_conversion/2)
    cur['marketing_candidate']=cur.rating.ge(9)&cur.share.lt(.001)&cur.visits.ge(1000)&cur.deleted.ne('Да')
    save(cur,'assortment_april')
    save(cur.head(10),'forecast_top10')
    save(cur.head(20),'stock_top20')
    save(cur[cur.low_conversion].sort_values('visits',ascending=False),'low_conversion')
    save(cur[cur.marketing_candidate].sort_values('rating',ascending=False),'marketing_candidates')
    save(cur.groupby('abc',sort=True).agg(skus=('sku','size'),revenue=('revenue','sum'),share=('share','sum')).reset_index(),'abc_summary')
    save(cur.groupby(['demand_class','turnover_class']).agg(skus=('sku','size'),orders=('orders','sum'),stock=('stock','sum'),gap=('stock_gap','sum')).reset_index(),'xz_summary')
    # Rolling-origin tests: select top 10 by revenue of month t-1, never by month t.
    tests=[]
    for i in range(2,len(periods)):
        older=facts[facts.month==periods[i-2]].set_index('sku')
        prior=facts[facts.month==periods[i-1]].sort_values(['revenue','sku'],ascending=[False,True]).head(10).set_index('sku')
        actual=facts[facts.month==periods[i]].set_index('sku')
        for sku,row in prior.iterrows():
            if sku not in older.index or sku not in actual.index:
                tests.append({'month':periods[i],'sku':sku,'status':'Нет полной истории'})
                continue
            prediction=float((older.loc[sku,'orders']+row.orders)/2)
            observed=float(actual.loc[sku,'orders'])
            naive=float(row.orders)
            tests.append({'month':periods[i],'sku':sku,'status':'OK','older_orders':float(older.loc[sku,'orders']),'prior_orders':naive,'actual':observed,'prediction':prediction,'abs_error':abs(observed-prediction),'naive':naive,'naive_error':abs(observed-naive)})
    backtest=pd.DataFrame(tests)
    save(backtest,'backtest_detail')
    valid=backtest[backtest.status=='OK']
    bt=valid.groupby('month').agg(skus=('sku','size'),actual=('actual','sum'),abs_error=('abs_error','sum'),naive_error=('naive_error','sum')).reset_index()
    bt['wape']=bt.abs_error/bt.actual
    bt['naive_wape']=bt.naive_error/bt.actual
    save(bt,'backtest_summary')
    # Repeated previous-period values are audited, never appended as new observations.
    comparisons=[]
    for i in range(1,len(periods)):
        a=facts[facts.month==periods[i]].set_index('sku')
        b=facts[facts.month==periods[i-1]].set_index('sku')
        keys=a.index.intersection(b.index)
        comparisons.append({'month':periods[i],'common_skus':len(keys),'orders_mismatch':int((a.loc[keys,'prev_orders_export']!=b.loc[keys,'orders']).sum()),'revenue_mismatch':int((a.loc[keys,'prev_revenue_export']!=b.loc[keys,'revenue']).sum())})
    save(pd.DataFrame(comparisons),'previous_period_checks')
    # Check suspicious funnel relationships without silently changing the inputs.
    anomalies={'orders_gt_visits':int((facts.orders>facts.visits).sum()),'buyouts_plus_cancels_gt_orders':int((facts.buyouts+facts.cancels>facts.orders).sum()),'zero_visits_positive_orders':int(((facts.visits==0)&(facts.orders>0)).sum()),'negative_revenue':int((facts.revenue<0).sum())}
    stock_changes=facts.groupby('sku').agg(months=('month','size'),unique_stock=('stock','nunique'))
    same=stock_changes[(stock_changes.months==11)&(stock_changes.unique_stock==1)]
    summary={'period_start':periods[0],'period_end':latest,'forecast_month':str(pd.Period(latest,freq='M')+1),'rows':len(facts),'skus':int(facts.sku.nunique()),'latest_skus':len(cur),'revenue':float(total),'previous_revenue':float(prev.revenue.sum()),'revenue_growth':float(total/prev.revenue.sum()-1),'orders':int(cur.orders.sum()),'previous_orders':int(prev.orders.sum()),'conversion':weighted_conversion,'previous_conversion':float(prev.orders.sum()/prev.visits.sum()),'stock':int(cur.stock.sum()),'top10_forecast':int(cur.head(10).forecast.sum()),'top20_gap':int(cur.head(20).stock_gap.sum()),'top20_excess':int(cur.head(20).excess_stock.sum()),'demand_median':demand_cut,'low_conversion_count':int(cur.low_conversion.sum()),'marketing_count':int(cur.marketing_candidate.sum()),'wape':float(valid.abs_error.sum()/valid.actual.sum()),'naive_wape':float(valid.naive_error.sum()/valid.actual.sum()),'latest_wape':float(bt.iloc[-1].wape),'latest_naive_wape':float(bt.iloc[-1].naive_wape),'test_rows':len(valid),'test_missing':int((backtest.status!='OK').sum()),'all_months_skus':int((stock_changes.months==11).sum()),'unchanged_stock_skus':len(same),'anomalies':anomalies,'previous_period_mismatches':comparisons,'abc':cur.groupby('abc').size().to_dict(),'top20':cur.head(20)[['sku','article','revenue','orders','stock','forecast','stock_gap','recommendation']].to_dict('records')}
    (ROOT/'validation'/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    (ROOT/'data'/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    assert abs(cur.share.sum()-1)<1e-10
    assert len(cur.head(10))==10 and len(cur.head(20))==20
    assert cur.head(20).forecast.notna().all()
    assert int(history.orders.sum())==int(facts.orders.sum())
    print(json.dumps({k:v for k,v in summary.items() if k not in ['top20','previous_period_mismatches']},ensure_ascii=False,indent=2))

if __name__=='__main__':
    main()
