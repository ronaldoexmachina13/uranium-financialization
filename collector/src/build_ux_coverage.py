"""Build auditable UX daily coverage and monthly observations from saved CME PDFs.

Run after download_cme_uranium.py --parse. No network, interpolation, or filling.
Weeks qualify only if every weekday has a product-total observation. This is a
conservative sufficient rule: a holiday without a verified calendar prevents
qualification rather than manufacturing a zero.
"""
from datetime import timedelta
import re
import json
import pandas as pd
import pdfplumber
from common import RAW, PROCESSED, log_retrieval
from download_cme_uranium import parse_s02b, status_rank

START, END = '2021-07-01', '2026-10-05'
daily = []
for path in sorted((RAW/'cme/daily_bulletin').glob('*S02B.pdf')):
    r = parse_s02b(path)
    for lookback in (False, True):
        date = r['trade_date'] - timedelta(days=364) if lookback else r['trade_date']
        daily.append(dict(trade_date=date,
            ux_total_volume=r['s02b_volume_52w_ago'] if lookback else r['s02b_total_volume'],
            ux_open_interest=r['s02b_oi_52w_ago'] if lookback else r['s02b_open_interest'],
            status=r['status'], source_file=path.name,
            provenance='52_week_comparison' if lookback else 'current_product_total',
            bulletin_trade_date=r['trade_date']))
browser_path=RAW/'cme/UX_browser_observed_records__20261006.json'
browser=json.loads(browser_path.read_text(encoding='utf-8'))
for r in browser['records']:
    if 'contract_total_volumes' in r:
        assert sum(r['contract_total_volumes'])==r['ux_total_volume']
        assert sum(r['contract_open_interest'])==r['ux_open_interest']
    daily.append({k:r[k] for k in ['trade_date','ux_total_volume','ux_open_interest','status','provenance']} | {'source_file':browser_path.name})
log_retrieval('UX_browser','cme_browser',browser['source_url'],browser_path,browser_path.read_bytes(),'parsed','DOM transcription, explicitly labelled')
d = pd.DataFrame(daily)
d['trade_date'] = pd.to_datetime(d.trade_date)
d = (d.assign(rank=d.status.map(status_rank)).sort_values(['trade_date','rank'])
       .drop_duplicates('trade_date',keep='last').drop(columns='rank'))
d = d[d.trade_date.between(START,END)]
grid = pd.DataFrame({'trade_date':pd.bdate_range(START,END)})
grid = grid.merge(d,on='trade_date',how='left')
grid['coverage'] = grid.ux_total_volume.notna().map({True:'observed',False:'missing_not_zero'})
d.to_csv(PROCESSED/'daily/uranium_futures_UX_observed_daily__CME.csv',index=False)
grid.to_csv(PROCESSED/'daily/uranium_futures_UX_daily_coverage__CME.csv',index=False)
weekly=[]
for end,g in grid.groupby(grid.trade_date.dt.to_period('W-FRI')):
    week_end=end.end_time.normalize()
    week_start=week_end-pd.Timedelta(days=4)
    complete=(len(g)==5 and g.ux_total_volume.notna().all() and week_start>=pd.Timestamp(START) and week_end<=pd.Timestamp(END))
    weekly.append(dict(week_ending=week_end.date(),candidate_weekdays=len(g),
        observed_days=int(g.ux_total_volume.notna().sum()),complete_coverage=complete,
        V_t=g.ux_total_volume.sum() if complete else None,
        observed_partial_volume=g.ux_total_volume.sum(min_count=1),
        week_end_open_interest=(g.ux_open_interest.iloc[-1] if g.trade_date.iloc[-1]==week_end else None),
        coverage_rule='all_five_weekdays_observed; holidays_unverified',
        includes_preliminary=bool((g.status=='PRELIMINARY').any()),
        includes_unspecified_status=bool((g.status=='UNSPECIFIED_CHART').any())))
w=pd.DataFrame(weekly)
(PROCESSED/'weekly').mkdir(exist_ok=True)
w.to_csv(PROCESSED/'weekly/uranium_futures_UX_volume_weekly_coverage__CME.csv',index=False)

records=[]
for kind,prefix in [('volume','Web_Volume_Report'),('open_interest','Web_OI_Report')]:
    for path in sorted((RAW/'cme').glob(prefix+'_NYMEX_COMEX*.pdf')):
        b=path.read_bytes()
        url='https://www.cmegroup.com/daily_bulletin/monthly_volume/'+prefix+'_NYMEX_COMEX.pdf'
        log_retrieval('UX_monthly','cme_browser',url,path,b,'parsed','Personal local research only')
        with pdfplumber.open(path) as pdf:
            for page_no,page in enumerate(pdf.pages,1):
                txt=page.extract_text() or ''
                match=re.search(r'^URANIUM UX (.+)$',txt,re.M)
                if not match: continue
                months=re.findall(r'\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC) (20\d{2})\b',txt)
                assert len(months)>=3, 'Month headers missing'
                tokens=match[1].split()
                offsets=[0,1,3] if kind=='volume' else [0,1,3]
                for (month,year),off in zip(months[:3],offsets):
                    records.append(dict(month=pd.to_datetime(month+year,format='%b%Y').strftime('%Y-%m'),
                        measure=kind,value=int(tokens[off].replace(',','')),unit='contracts',
                        source_file=path.name,source_page=page_no,source_url=url,
                        preliminary_oi=(kind=='open_interest'),raw_row=match[0]))
                if kind=='volume':
                    for year,off in [('2026',5),('2025',6)]:
                        records.append(dict(month='',measure='year_to_date_volume',value=int(tokens[off].replace(',','')),
                            unit='contracts',period_start=year+'-01-01',period_end=year+'-09-30',
                            source_file=path.name,source_page=page_no,source_url=url,raw_row=match[0]))
                break
(PROCESSED/'monthly').mkdir(exist_ok=True)
m=pd.DataFrame(records)
m.to_csv(PROCESSED/'monthly/uranium_futures_UX_monthly_observations__CME.csv',index=False)
assert len(d)==10
assert w.V_t.notna().sum()==1
assert w.loc[w.week_ending.astype(str)=='2026-10-02','V_t'].iloc[0]==524
assert sorted(m[m.measure=='volume'].value)==[4,150,508]
print(f'Observed daily: {len(d)}; grid weekdays: {len(grid)}; weekly rows: {len(w)}; complete weeks: {w.complete_coverage.sum()}')
print(m[['month','measure','value']].to_string(index=False))
