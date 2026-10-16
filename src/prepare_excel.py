from pathlib import Path
import json
import pandas as pd

root=Path(__file__).resolve().parents[1]
def read(name):
    return pd.read_csv(root/'data'/f'{name}.csv',dtype={'sku':str})
a=read('assortment_april')
payload={
    'current':a[['sku','article','category','rating','visits','orders','revenue','stock_wb','stock_mp','orders_previous','price','deleted','revenue_previous']].astype(object).where(pd.notna(a),None).values.tolist(),
    'history':read('monthly_summary').to_dict('records'),
    'backtest':read('backtest_detail').to_dict('records'),
    'backtestMonths':read('backtest_summary').month.tolist(),
    'stockWarnings':a.head(20).stock_warning.tolist(),
    'summary':json.loads((root/'validation'/'summary.json').read_text(encoding='utf-8')),
}
(root/'validation'/'excel_input.json').write_text(json.dumps(payload,ensure_ascii=False,allow_nan=False),encoding='utf-8')
