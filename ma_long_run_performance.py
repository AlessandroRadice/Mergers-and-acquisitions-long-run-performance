# %% [markdown]
# # Does M&A Create Value? A Long-Run Performance Study
#
# **April 2025. A board is weighing a large acquisition. Announcement-day returns say what investors expect a deal to do; this
# notebook measures what actually happened afterwards. It follows every significant acquisition completed by a US listed
# non-financial company from 2012 to 2021 (at least $100m and 10% of the acquirer's market value) for three to five years, and
# asks whether the acquirer's shareholders beat comparable stocks, whether any shortfall survives a risk adjustment, which kinds
# of deal did worse, and whether margins and goodwill tell the same story.**
#
# Everything is rebuilt from public, free sources, as available on 31 March 2025:
#
# | Step | What it does |
# |---|---|
# | 1. Data | SEC XBRL frames (acquisition spending, fundamentals, cover-page shares and float), SEC submissions (8-K items, filing history), DERA SIC codes, DoltHub prices, splits and dividends including delisted stocks, Kenneth French factors and portfolios |
# | 2. Deals | Acquirer-years with $100m+ of acquisitions, dated by the first 8-K Item 2.01 (completion of acquisition) |
# | 3. Prices | Every SEC filer matched to a ticker, and every match checked: price times shares outstanding must agree with the public float |
# | 4. Event study | Buy-and-hold abnormal returns against 25 size and book-to-market portfolios, skewness-adjusted and bootstrapped tests |
# | 5. Calendar time | Monthly portfolios of recent acquirers regressed on the Fama-French five factors plus momentum |
# | 6. Operations | Industry-adjusted operating return on assets (Healy, Palepu and Ruback regression) and goodwill impairments |
# | 7. Cross-section | Sorts and a regression on payment, size, valuation, run-up, serial buying and goodwill; robustness and placebo |
# | 8. Export | **Interactive HTML page** with every deal, and an **Excel model with live formulas** that reproduces the headline numbers |
#
# > Educational project, not investment advice. The final sample has 786 deals.

# %%
import subprocess, sys
_ = subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pandas", "numpy", "scipy", "matplotlib", "requests", "openpyxl", "statsmodels"], check=False, capture_output=True)

# %%
import os, io, re, json, gzip, time, zipfile, warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import requests
import statsmodels.api as sm
import statsmodels.formula.api as smf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", lambda x: f"{x:,.4f}")
pd.set_option("display.width", 200)

# %% [markdown]
# ## 0. Configuration
# `DATA_MODE = "snapshot"` uses the files in `data/` (exactly what the download produced, trimmed to 31 March 2025) and runs
# in a few minutes. `DATA_MODE = "live"` downloads everything again from the SEC, DoltHub and the French data library
# (several hours, mostly the SEC's rate limit of 10 requests a second; put your own contact in `UA`).

# %%
DATA_MODE = "snapshot"            # "snapshot" or "live"
DATA_DIR = "data"                 # source data (json.gz)
R = "work/"                       # intermediate files written by the pipeline
CUTOFF = "2025-03-31"                  # nothing after this date is used
os.makedirs(DATA_DIR, exist_ok=True); os.makedirs(R, exist_ok=True)

def load(name):
    with gzip.open(f"{DATA_DIR}/{name}.json.gz", "rt", encoding="utf-8") as f:
        return json.load(f)

def save(name, obj):
    with gzip.open(f"{DATA_DIR}/{name}.json.gz", "wt", encoding="utf-8") as f:
        json.dump(obj, f, separators=(",", ":"))

# %% [markdown]
# ## 1. Data
# The download functions below are the ones used to build the snapshot. In snapshot mode they are only defined, not run.
# * **Acquisitions**: SEC XBRL frames for six tags, calendar years 2009 to 2024: cash paid for acquisitions (net, gross, including
#   affiliates), stock issued for acquisitions, consideration transferred and goodwill acquired.
# * **Fundamentals**: assets, equity, goodwill, revenue, operating income, net income, operating cash flow, impairments, depreciation;
#   cover-page public float and shares outstanding (dei namespace), quarterly.
# * **Filing history**: SEC submissions for every candidate acquirer: 8-K items 1.01 and 2.01, 10-K and 10-Q dates and the
#   ticker prefix of the primary document.
# * **Prices**: DoltHub `post-no-preference/stocks` month-end closes from January 2011, splits and dividends, including delisted stocks.
# * **Benchmarks**: Kenneth R. French data library: five factors, momentum, 25 size and book-to-market portfolios, NYSE breakpoints.

# %%
UA = {"User-Agent": "Your Name your.email@example.com"}   # the SEC asks for a contact    # put your own contact here
DOLT = "https://www.dolthub.com/api/v1alpha1/post-no-preference/stocks/master"
FRENCH = "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
_last = [0.0]
def sec(url, tries=5):
    for k in range(tries):
        wait = 0.13 - (time.time() - _last[0])
        if wait > 0: time.sleep(wait)
        _last[0] = time.time()
        r = requests.get(url, headers=UA, timeout=60)
        if r.status_code == 200: return r
        if r.status_code in (429, 503): time.sleep(60 * (k + 1)); continue
        if r.status_code == 404: return None
    r.raise_for_status()
def sec_frame(ns, tag, unit, period):
    r = sec(f"https://data.sec.gov/api/xbrl/frames/{ns}/{tag}/{unit}/{period}.json")
    return [] if r is None else r.json()["data"]
ACQ_TAGS = ["PaymentsToAcquireBusinessesNetOfCashAcquired", "PaymentsToAcquireBusinessesGross", "StockIssuedDuringPeriodValueAcquisitions",
            "BusinessCombinationConsiderationTransferred1", "GoodwillAcquiredDuringPeriod", "PaymentsToAcquireBusinessesAndInterestInAffiliates"]
def acquisition_frames(y0=2009, y1=2024):
    rows = []
    for y in range(y0, y1 + 1):
        for t in ACQ_TAGS:
            for d in sec_frame("us-gaap", t, "USD", f"CY{y}"):
                rows.append([t, y, d["cik"], d["entityName"], d.get("start"), d["end"], d["val"], d["accn"]])
    return pd.DataFrame(rows, columns=["tag", "cy", "cik", "name", "start", "end", "val", "accn"])
STOCK_TAGS = ["Assets", "StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "Goodwill"]
FLOW_TAGS = ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet", "OperatingIncomeLoss", "NetIncomeLoss",
             "NetCashProvidedByUsedInOperatingActivities", "GoodwillImpairmentLoss", "DepreciationDepletionAndAmortization"]
def fundamental_frames(y0=2008, y1=2024):
    rows = []
    for y in range(y0, y1 + 1):
        specs = [("us-gaap", t, "USD", f"CY{y}Q4I") for t in STOCK_TAGS] + [("us-gaap", t, "USD", f"CY{y}") for t in FLOW_TAGS]
        specs += [("dei", "EntityPublicFloat", "USD", f"CY{y}Q{q}I") for q in (1, 2, 3, 4)]
        specs += [("dei", "EntityCommonStockSharesOutstanding", "shares", f"CY{y}Q{q}I") for q in (1, 2, 3, 4)]
        for ns, t, u, p in specs:
            for d in sec_frame(ns, t, u, p):
                rows.append([t, p, d["cik"], d["end"], d["val"]])
    for t in ("Assets", "StockholdersEquity", "Goodwill"):
        for d in sec_frame("us-gaap", t, "USD", "CY2025Q1I"): rows.append([t, "CY2025Q1I", d["cik"], d["end"], d["val"]])
    for d in sec_frame("dei", "EntityCommonStockSharesOutstanding", "shares", "CY2025Q1I"): rows.append(["EntityCommonStockSharesOutstanding", "CY2025Q1I", d["cik"], d["end"], d["val"]])
    return pd.DataFrame(rows, columns=["tag", "per", "cik", "end", "val"])
KEEP = {"8-K", "10-K", "10-K405", "10-KT", "10-Q", "25", "25-NSE", "15-12B", "15-12G", "15-15D"}
def submissions(ciks):
    out = {}
    def take(f, arr):
        for i, fm in enumerate(f["form"]):
            if fm not in KEEP: continue
            items = f["items"][i] or ""
            if fm == "8-K" and not re.search(r"1\.01|2\.01", items): continue
            if fm in ("10-K", "10-KT", "10-Q"):
                m = re.match(r"^([a-z]{1,5})[-_]\d{8}", (f["primaryDocument"][i] or "").lower())
                if fm == "10-Q":
                    if m: arr.append(["Q", f["filingDate"][i], m.group(1)])
                else: arr.append(["K", f["filingDate"][i], m.group(1) if m else "", f["accessionNumber"][i]])
            elif fm == "8-K": arr.append(["8", f["filingDate"][i], items])
            else: arr.append([fm, f["filingDate"][i]])
    for c in ciks:
        r = sec(f"https://data.sec.gov/submissions/CIK{int(c):010d}.json")
        if r is None: continue
        j = r.json(); arr = []; take(j["filings"]["recent"], arr)
        for fl in j["filings"].get("files", []):
            r2 = sec("https://data.sec.gov/submissions/" + fl["name"])
            if r2 is not None: take(r2.json(), arr)
        out[str(c)] = [j["name"], j["sic"], j["tickers"], j["exchanges"], arr]
    return out
def sic_codes(quarters=("2009q3", "2010q2", "2011q1", "2011q3", "2012q3", "2013q3", "2014q1", "2014q3", "2015q1", "2016q3", "2017q1", "2018q1",
                        "2019q3", "2020q1", "2021q3", "2022q1", "2023q1", "2024q1")):
    sic = {}
    for q in quarters:
        r = sec(f"https://www.sec.gov/files/dera/data/financial-statement-data-sets/{q}.zip")
        if r is None: continue
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            sub = pd.read_csv(z.open("sub.txt"), sep="\t", usecols=["cik", "sic"], dtype=str)
        for c, s in sub.dropna().values: sic.setdefault(str(int(c)), s)
    return sic
def dolt(q, tries=5):
    for k in range(tries):
        try:
            j = requests.get(DOLT, params={"q": q}, timeout=120).json()
            if "rows" in j: return j["rows"]
        except Exception: pass
        time.sleep(3)
    raise RuntimeError("DoltHub query failed: " + q[:80])
def dolt_symbols():
    out, last = [], ""
    while True:
        rows = dolt(f"SELECT act_symbol,security_name,listing_exchange,is_etf,last_seen FROM symbol WHERE act_symbol > '{last}' ORDER BY act_symbol LIMIT 1000")
        if not rows: break
        out += rows; last = rows[-1]["act_symbol"].replace("'", "''")
        if len(rows) < 1000: break
    return [[r["act_symbol"], r["security_name"], r["listing_exchange"], r["last_seen"]] for r in out if r["is_etf"] != "1"]
def dolt_splits():
    out, off = [], 0
    while True:
        rows = dolt(f"SELECT * FROM split ORDER BY act_symbol, ex_date LIMIT 1000 OFFSET {off}")
        out += rows; off += 1000
        if len(rows) < 1000: break
    return [[r["act_symbol"], r["ex_date"], float(r["to_factor"]), float(r["for_factor"])] for r in out]
def month_ends(first="2011-01", last="2025-03"):
    days = []
    for p in pd.period_range(first, last, freq="M"):
        e = p.to_timestamp(how="end")
        days += [(e - pd.Timedelta(days=k)).strftime("%Y-%m-%d") for k in range(7)]
    got = []
    for i in range(0, len(days), 300):
        got += [r["date"] for r in dolt("SELECT date FROM ohlcv WHERE act_symbol='AAPL' AND date IN (" + ",".join(f"'{d}'" for d in days[i:i + 300]) + ")")]
    s = pd.Series(pd.to_datetime(got)); return sorted(s.groupby(s.dt.to_period("M")).max().dt.strftime("%Y-%m-%d"))
def dolt_prices(symbols, dates):
    dl = ",".join(f"'{d}'" for d in ["2011-01-03"] + list(dates)); out = {}
    for i in range(0, len(symbols), 5):
        b = ",".join("'" + s.replace("'", "''") + "'" for s in symbols[i:i + 5])
        for r in dolt(f"SELECT act_symbol,date,close FROM ohlcv WHERE act_symbol IN ({b}) AND date IN ({dl})"):
            out.setdefault(r["act_symbol"], []).append([r["date"], float(r["close"])])
    return out
def dolt_dividends(symbols):
    out = []
    for i in range(0, len(symbols), 12):
        b = ",".join("'" + s + "'" for s in symbols[i:i + 12])
        out += [[r["act_symbol"], r["ex_date"], float(r["amount"])] for r in
                dolt(f"SELECT act_symbol,ex_date,amount FROM dividend WHERE act_symbol IN ({b}) AND ex_date >= '2010-12-01' AND ex_date <= '2025-03-31'")]
    return out
def french(name):
    r = requests.get(FRENCH + name + ".zip", timeout=120)
    with zipfile.ZipFile(io.BytesIO(r.content)) as z: return z.read(z.namelist()[0]).decode("latin-1")

# %%
FRENCH_FILES = ["F-F_Research_Data_5_Factors_2x3_CSV", "F-F_Momentum_Factor_CSV", "25_Portfolios_5x5_CSV", "ME_Breakpoints_CSV", "BE-ME_Breakpoints_CSV"]

def download_sources():
    """SEC and French data. Prices and dividends follow in section 3, once the candidate tickers are known."""
    af = acquisition_frames(2009, 2024); af = af[af.end <= CUTOFF]
    save("acq_frames", af.values.tolist())
    a = acq_year()
    cands = sorted(set(a[(a.dv >= 1e8) & a.cy.between(2011, 2022)].cik))
    ff = fundamental_frames(); ff = ff[ff.end <= CUTOFF]
    ff = ff[~(ff.per.eq("CY2025Q1I") & ff.tag.ne("EntityCommonStockSharesOutstanding"))]   # balances at 31 March 2025 were filed after the cutoff
    ff["e"] = ff.end.str.replace("-", "").astype(int)
    shares = ff.tag == "EntityCommonStockSharesOutstanding"
    ff["v"] = np.where(shares, ff.val, ff.val / 1000.0)
    c = ff[ff.cik.isin(cands)]
    save("fund_cand", {f"{t}|{p}": [g.cik.tolist(), g.e.tolist(), g.v.tolist()] for (t, p), g in c.groupby(["tag", "per"])})
    o = ff[~ff.cik.isin(cands) & ff.tag.isin(["Assets", "OperatingIncomeLoss"] + REV)]
    save("fund_all", {f"{t}|{p}": [g.cik.tolist(), g.v.tolist()] for (t, p), g in o.groupby(["tag", "per"])})
    subs = submissions(cands)
    for k, v in subs.items(): v[4] = [f for f in v[4] if f[1] <= CUTOFF]
    save("subs", subs)
    save("sic", sic_codes())
    save("dolt_sym", {"sym": dolt_symbols(), "spl": [r for r in dolt_splits() if r[1] <= CUTOFF]})
    save("french", {n: french(n) for n in FRENCH_FILES})

need = ["acq_frames", "fund_cand", "fund_all", "sic", "subs", "dolt_sym", "french"]
if DATA_MODE == "live" or not all(os.path.exists(f"{DATA_DIR}/{n}.json.gz") for n in need):
    print("Downloading from the SEC and the French data library..."); download_sources()
print({n: f"{os.path.getsize(f'{DATA_DIR}/{n}.json.gz') / 1e6:.1f} MB" for n in need})

# %% [markdown]
# ## 2. Benchmarks and fundamentals
# French's CSV files are split into titled blocks; the 25 portfolios come value- and equal-weighted, returns in percent.
# Size breakpoints are monthly NYSE percentiles of market value ($m); book-to-market breakpoints are annual.

# %%
D=load('french')
def sections(txt):
    """split a French csv into titled blocks of monthly rows (YYYYMM)"""
    out={}; title='main'; rows=[]; hdr=None
    for l in txt.split('\n'):
        s=l.strip()
        if not s: continue
        m=re.match(r'^(\d{6}),',s)
        if m:
            rows.append([x.strip() for x in s.split(',')]); continue
        if re.match(r'^\d{4},',s): continue           # annual rows
        if s.startswith(','):
            if rows: out[title]=(hdr,rows); rows=[]
            hdr=[x.strip() for x in s.split(',')][1:]
            continue
        if rows: out[title]=(hdr,rows); rows=[]
        title=s
    if rows: out[title]=(hdr,rows)
    return out
def frame(hdr,rows):
    df=pd.DataFrame([r[1:] for r in rows],index=pd.PeriodIndex([r[0] for r in rows],freq='M'),columns=hdr).astype(float)
    return df
def factors():
    s=sections(D['F-F_Research_Data_5_Factors_2x3_CSV']); k=[k for k in s][0]; f=frame(*s[k])/100
    m=sections(D['F-F_Momentum_Factor_CSV']); km=[k for k in m][0]; mo=frame(*m[km])/100
    f['MOM']=mo.iloc[:,0]; return f
def p25():
    s=sections(D['25_Portfolios_5x5_CSV'])
    vw=[k for k in s if 'Value Weight' in k and 'Monthly' in k][0]; ew=[k for k in s if 'Equal Weight' in k and 'Monthly' in k][0]
    VW=frame(*s[vw]); EW=frame(*s[ew]); VW[VW<=-99.99]=np.nan; EW[EW<=-99.99]=np.nan
    return VW/100, EW/100
def me_bp():
    rows=[]
    for l in D['ME_Breakpoints_CSV'].split('\n'):
        if re.match(r'^\d{6},',l.strip()): rows.append([float(x) for x in l.strip().split(',')])
    df=pd.DataFrame(rows); df.index=pd.PeriodIndex(df[0].astype(int).astype(str),freq='M'); return df.iloc[:,2:]  # 20 cols: 5th..100th pct, $m
def beme_bp():
    rows=[]
    for l in D['BE-ME_Breakpoints_CSV'].split('\n'):
        s=l.strip()
        if re.match(r'^\d{4},',s): rows.append([float(x) for x in s.split(',')])
    df=pd.DataFrame(rows); df.index=df[0].astype(int); return df.iloc[:,3:]

f = factors(); VW25, EW25 = p25()
print(f"Factors {f.index[0]} to {f.index[-1]} · 25 portfolios: {VW25.shape[1]} · last month {VW25.index[-1]}")
f.loc["2020-01":"2020-06"]

# %%
def cand_long():
    C=load('fund_cand'); out=[]
    for k,(ciks,ends,vals) in C.items():
        tag,per=k.split('|')
        out.append(pd.DataFrame({'tag':tag,'per':per,'cik':ciks,'end':pd.to_datetime(pd.Series(ends).astype(str)),'val':vals}))
    df=pd.concat(out,ignore_index=True)
    df['val']=np.where(df.tag=='EntityCommonStockSharesOutstanding',df.val,df.val*1000.0)
    return df
def all_long():
    A=load('fund_all'); out=[]
    for k,(ciks,vals) in A.items():
        tag,per=k.split('|'); out.append(pd.DataFrame({'tag':tag,'per':per,'cik':ciks,'val':np.array(vals,float)*1000}))
    return pd.concat(out,ignore_index=True)
REV=['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax','SalesRevenueNet']
def annual(df):
    """firm-year table keyed on frame calendar year; flows from CYyyyy, stocks from CYyyyyQ4I"""
    d=df.copy(); d['cy']=d.per.str[2:6].astype(int); d['kind']=np.where(d.per.str.contains('Q'),'I','D')
    flows=d[(d.kind=='D')].pivot_table(index=['cik','cy'],columns='tag',values='val',aggfunc='first')
    stocks=d[(d.per.str.endswith('Q4I'))].pivot_table(index=['cik','cy'],columns='tag',values='val',aggfunc='first')
    t=flows.join(stocks,how='outer')
    if set(REV)&set(t.columns):
        t['rev']=np.nan
        for r in REV:
            if r in t: t['rev']=t['rev'].fillna(t[r])
    return t

def shares_float():
    c=cand_long()
    sh=c[c.tag=='EntityCommonStockSharesOutstanding'][['cik','end','val']].drop_duplicates(['cik','end']).rename(columns={'val':'shares','end':'date'})
    fl=c[c.tag=='EntityPublicFloat'][['cik','end','val']].drop_duplicates(['cik','end']).rename(columns={'val':'float','end':'date'})
    fl=fl[fl.date<='2024-07-31']      # a float is reported in the 10-K filed six to nine months after its measurement date: later ones were not public by 31 March 2025
    return sh.sort_values(['cik','date']), fl.sort_values(['cik','date'])

sh, fl = shares_float(); sh.to_pickle(R + "shares.pkl"); fl.to_pickle(R + "float.pkl")
print(f"{len(sh):,} cover-page share counts · {len(fl):,} public-float values")

# %% [markdown]
# ## 3. Deals
# **Deal value** for a company and fiscal year is the larger of cash plus stock paid, consideration transferred and, where no
# payment is tagged, goodwill acquired. **Completion date**: the first 8-K with Item 2.01 filed between the start of the fiscal
# year and ten days after its end. Banks, insurers and other financials (SIC 6000 to 6999) and utilities (4900 to 4999) are excluded.

# %%
def acq_year():
    df=pd.DataFrame(load('acq_frames'),columns=['tag','cy','cik','name','start','end','val','accn'])
    df['end']=pd.to_datetime(df.end); df['start']=pd.to_datetime(df.start)
    p=df.pivot_table(index=['cik','cy'],columns='tag',values='val',aggfunc='first')
    e=df.groupby(['cik','cy']).agg(end=('end','max'),start=('start','min'),name=('name','first'))
    p=p.join(e)
    p['cash']=p['PaymentsToAcquireBusinessesNetOfCashAcquired'].fillna(p['PaymentsToAcquireBusinessesGross']).fillna(p['PaymentsToAcquireBusinessesAndInterestInAffiliates'])
    p['stock']=p['StockIssuedDuringPeriodValueAcquisitions']
    p['paid']=p[['cash','stock']].clip(lower=0).fillna(0).sum(axis=1)
    p['consid']=p['BusinessCombinationConsiderationTransferred1']
    p['gw_acq']=p['GoodwillAcquiredDuringPeriod']
    p['dv_src']=np.where(p['consid'].fillna(0)>p['paid'],'consideration',np.where(p['paid']>0,'cash+stock','goodwill'))
    p['dv']=np.fmax(np.fmax(p['paid'].values, p['consid'].fillna(0).values), np.where(p['paid']>0,0,p['gw_acq'].fillna(0).values))
    p['stock_share']=(p['stock'].fillna(0).clip(lower=0)/p['paid'].replace(0,np.nan)).clip(0,1)
    return p.reset_index()

STOP={'inc','incorporated','corp','corporation','co','company','ltd','limited','plc','llc','lp','l','p','sa','nv','n','v','ag','se','the','holdings','holding','group','common','stock','shares','share','class','a','b','ordinary','new','de','del','md','nv','ny','tx','ca','ma','nj','pa','va','in','oh','fl','il','mi','wa','co','depositary','american','ads','adr','each','representing','par','value','usd','0','01','001','0001','trust','units','unit','of','and','&','inc.'}
def norm(s):
    s=s.lower()
    s=re.sub(r'/[a-z]{2,3}/?$','',s.strip())          # EDGAR state suffix e.g. /DE/
    s=re.sub(r'\s-\s.*$','',s)                          # Nasdaq style " - Common Stock"
    s=s.replace('&',' and ')
    s=re.sub(r'common stock.*$|class [a-z] .*$|ordinary shares.*$|american depositary.*$','',s)
    s=re.sub(r'[^a-z0-9 ]',' ',s)
    toks=[t for t in s.split() if t not in STOP]
    return ' '.join(toks)

SUB = load("subs"); SIC = load("sic")
def base_events():
    a=acq_year()
    a['sic']=pd.to_numeric(a.cik.astype(str).map(lambda c: SUB.get(c,[None,None])[1] or SIC.get(c)),errors='coerce')
    a=a[(a.dv>=1e8)&(a.cy.between(2011,2022))].copy()
    a['fin']=a.sic.between(6000,6999); a['util']=a.sic.between(4900,4999)
    rows=[]
    for r in a.itertuples():
        f=SUB.get(str(r.cik),[None]*5)[4] or []
        lo=r.start if pd.notna(r.start) else r.end-pd.Timedelta(days=365); hi=r.end+pd.Timedelta(days=10)
        c=[pd.Timestamp(x[1]) for x in f if x[0]=='8' and '2.01' in x[2] and lo<=pd.Timestamp(x[1])<=hi]
        rows.append((min(c) if c else pd.NaT, len(c)))
    a['d201']=[x[0] for x in rows]; a['n201']=[x[1] for x in rows]
    return a
def ticker_candidates(ciks):
    sym=load('dolt_sym')['sym']
    byname={}
    for s,nm,ex,last in sym:
        byname.setdefault(norm(nm),[]).append(s)
    out={}
    for c in ciks:
        if str(c) not in SUB: out[c]=[]; continue
        n,sic,t,x,f=SUB[str(c)]
        cand=set(t or [])
        for r in f:
            if r[0] in ('K','Q') and r[2]: cand.add(r[2].upper())
        cand|=set(byname.get(norm(n),[]))
        out[c]=sorted(cand)
    return out

a = base_events()
e = a[(~a.fin) & (~a.util)]
e2 = e[e.d201.notna() & (e.d201 >= "2012-01-01") & (e.d201 <= "2021-12-31")]
tc = ticker_candidates(sorted(set(e.cik)))
json.dump({str(k): v for k, v in tc.items()}, open(R + "ticker_cand.json", "w")); a.to_pickle(R + "acq_years.pkl")
print(f"Acquirer-years with $100m+: {len(a):,} · non-financial: {len(e):,} · with an Item 2.01 in 2012-2021: {len(e2):,} ({e2.cik.nunique():,} firms)")
e2.sort_values("dv", ascending=False)[["name", "cy", "d201", "dv", "cash", "stock", "consid", "stock_share"]].head(10)

# %% [markdown]
# ## 4. Prices and ticker validation
# Tickers are reused and company names change, so every candidate ticker is checked: at every 10-K cover date, the month-end
# price times the shares outstanding must be between 20% and 120% of the public float reported on the same cover (the float
# excludes insiders, so the ratio sits below 1). Firms without a validated ticker are accepted only when the ticker in the
# filings and the DoltHub security name both match the EDGAR name. After a firm stops trading it earns its benchmark.

# %%
if DATA_MODE == "live" or not os.path.exists(f"{DATA_DIR}/px_month.json.gz"):
    syms = sorted({s for v in tc.values() for s in v})
    save("px_month", dolt_prices(syms, month_ends()))
def load_px():
    P=load('px_month')
    rows=[(s,d,c) for s,v in P.items() for d,c in v]
    df=pd.DataFrame(rows,columns=['sym','date','close']); df['date']=pd.to_datetime(df.date)
    df['m']=df.date.dt.to_period('M')
    df.loc[df.date==pd.Timestamp('2011-01-03'),'m']=pd.Period('2010-12','M')
    df=df.sort_values(['sym','date']).drop_duplicates(['sym','m'],keep='last')   # 2011-01-03 and 2011-01-31 -> keep month-end; base kept separately
    return df
def split_factors():
    spl=load('dolt_sym')['spl']
    s=pd.DataFrame(spl,columns=['sym','ex','to','for']); s['ex']=pd.to_datetime(s.ex); s['f']=s['to']/s['for']
    return s[(s.f>0)&np.isfinite(s.f)]
def monthly_returns(px, spl, divs=None):
    """price return (+ dividends if given) per symbol-month, split adjusted"""
    out=[]
    S={k:g for k,g in spl.groupby('sym')}
    D={k:g for k,g in divs.groupby('sym')} if divs is not None else {}
    for sym,g in px.groupby('sym'):
        g=g.sort_values('date'); d=g.date.values; c=g.close.values
        f=np.ones(len(g)); dv=np.zeros(len(g))
        if sym in S:
            for e,fac in S[sym][['ex','f']].values:
                i=np.searchsorted(d,np.datetime64(e),side='left')   # first obs with date>=ex
                if 0<i<len(g): f[i]*=fac
        if sym in D:
            for e,amt in D[sym][['ex','amt']].values:
                i=np.searchsorted(d,np.datetime64(e),side='left')
                if 0<i<len(g): dv[i]+=amt
        r=np.full(len(g),np.nan); r[1:]=(c[1:]*f[1:]+dv[1:])/c[:-1]-1
        # a gap in monthly observations (missing months) -> not a one-month return
        mm=g.m.values; gap=np.ones(len(g),bool); gap[1:]=[(mm[i]-mm[i-1]).n==1 for i in range(1,len(g))]
        r[~gap]=np.nan
        out.append(pd.DataFrame({'sym':sym,'m':g.m.values,'date':g.date.values,'close':c,'ret':r}))
    return pd.concat(out,ignore_index=True)

def build_map():
    tc={int(k):v for k,v in json.load(open(R+'ticker_cand.json')).items()}
    px=load_px(); pxd={s:g.set_index('m').close for s,g in px.groupby('sym')}
    sh=pd.read_pickle(R+'shares.pkl'); fl=pd.read_pickle(R+'float.pkl')
    sh['cik']=sh.cik.astype(int); fl['cik']=fl.cik.astype(int)
    SH={c:g for c,g in sh.groupby('cik')}; FL={c:g for c,g in fl.groupby('cik')}
    rows=[]
    for c,syms in tc.items():
        syms=[s for s in syms if s in pxd]
        if not syms: continue
        # implied market cap checks at every cover date: price(month of cover) * shares vs float (same year)
        g=SH.get(c); f=FL.get(c)
        for s in syms:
            p=pxd[s]
            ok=[];
            if f is not None:
                for d,flt in f[['date','float']].values:
                    d=pd.Timestamp(d); m=d.to_period('M')
                    if m not in p.index or g is None: continue
                    gg=g[(g.date-d).abs()<=pd.Timedelta(days=150)]
                    if not len(gg): continue
                    shs=gg.iloc[(gg.date-d).abs().argmin()].shares
                    mc=p[m]*shs
                    if mc>0: ok.append((d.year, flt/mc))
            for y,rat in ok: rows.append((c,s,y,rat))
    return pd.DataFrame(rows,columns=['cik','sym','year','ratio']), tc, pxd
def firm_symbol_months(v, pxd, events_ciks):
    """firm -> DataFrame(m, sym) using validated symbol-years; carry forward 2y / back 1y"""
    v=v[v.ratio.between(0.2,1.2)].copy(); v['dev']=(v.ratio-0.9).abs()
    best=v.sort_values('dev').drop_duplicates(['cik','year']).set_index(['cik','year']).sym
    out={}
    for c in events_ciks:
        if c not in best.index.get_level_values(0): continue
        b=best.loc[c]; yrs=range(2010,2026); sym_y={}
        for y in yrs:
            if y in b.index: sym_y[y]=b[y]
        for y in yrs:                     # forward fill up to 2 years
            if y not in sym_y:
                for k in (1,2):
                    if y-k in b.index: sym_y[y]=b[y-k]; break
        for y in yrs:                     # back fill 1 year
            if y not in sym_y and y+1 in b.index: sym_y[y]=b[y+1]
        rows=[]
        for y,s in sym_y.items():
            p=pxd.get(s)
            if p is None: continue
            for m in p.index:
                if m.year==y or (y==2010 and m==pd.Period('2010-12','M')): rows.append((m,s))
        if rows: out[c]=pd.DataFrame(rows,columns=['m','sym']).drop_duplicates('m').sort_values('m')
    return out

def fallback_symbols(unmapped, pxd):
    """accept a symbol without the float check only if two independent sources agree:
    (current ticker or filing-name prefix) AND DoltHub security name == EDGAR name"""
    SUB=load('subs'); sym=load('dolt_sym')['sym']
    nm={s:norm(n) for s,n,ex,last in sym}
    out={}
    for c in unmapped:
        if str(c) not in SUB: continue
        n,sic,t,x,f=SUB[str(c)]
        src=set(t or [])|{r[2].upper() for r in f if r[0] in('K','Q') and r[2]}
        ok=[s for s in src if s in pxd and nm.get(s)==norm(n)]
        if len(ok)>=1:
            s=sorted(ok,key=lambda s:-len(pxd[s]))[0]
            out[c]=pd.DataFrame({'m':pxd[s].index,'sym':s})
    return out

v, tc_, pxd = build_map(); v.to_pickle(R + "sym_valid.pkl")
print(f"{len(v):,} ticker-year checks · {v.ratio.between(0.2, 1.2).mean():.1%} within 0.2 to 1.2")

# %% [markdown]
# ## 5. The event panel
# Monthly total returns (split-adjusted, dividends added back). DoltHub dividend amounts are split-adjusted for some symbols
# and not others, so amounts next to large splits are re-scaled to whichever version matches the stock's own yield.

# %%
if DATA_MODE == "live" or not os.path.exists(f"{DATA_DIR}/divs.json.gz"):
    used = sorted({s for s in v.sym.unique()})
    save("divs", dolt_dividends(used))
LAST=pd.Period('2025-03','M')
def clean_divs(df, px, spl):
    """DoltHub dividend amounts are split-adjusted for some symbols and not for others, while prices are unadjusted.
    For a dividend followed by later splits (cumulative price factor C), keep the amount as reported or its
    unadjusted version amt*C (for splits of 3-for-1 or larger, either way), whichever gives a yield closer to the symbol's own yield after its last split."""
    out=[]; S={k:g for k,g in spl.groupby('sym')}; P={k:g for k,g in px.groupby('sym')}; fixed=0
    for sym,g in df.groupby('sym'):
        g=g.sort_values('ex').copy()
        if sym not in P: out.append(g); continue
        p=P[sym]; i=np.searchsorted(p.date.values,g.ex.values)-1
        prev=np.where(i>=0,p.close.values[np.clip(i,0,None)],np.nan)
        C=np.ones(len(g))
        if sym in S:
            for e,f in S[sym][['ex','f']].values:
                if pd.Timestamp(e)<=LAST.to_timestamp(how='end'): C[g.ex.values<np.datetime64(e)]*=f
        y=g.amt.values/prev
        ref=np.nanmedian(np.log(y[(C==1)&(y>0)])) if ((C==1)&(y>0)).any() else np.log(0.01)
        a=g.amt.values.copy()
        for j in np.where(np.abs(np.log(C))>=np.log(2.9))[0]:   # only splits large enough to separate from dividend growth
            if not np.isfinite(y[j]) or y[j]<=0: continue
            if abs(np.log(y[j]*C[j])-ref)<abs(np.log(y[j])-ref): a[j]=a[j]*C[j]; fixed+=1
        g['amt']=a
        g=g[~(a>prev)]   # a distribution larger than the stock price is a data error (e.g. a rights offering booked as a dividend)
        out.append(g)
    print('dividends re-scaled for split adjustment:',fixed)
    return pd.concat(out,ignore_index=True)
def load_divs(px=None, spl=None):
    try:
        d=load('divs')
    except FileNotFoundError:
        return None
    df=pd.DataFrame(d,columns=['sym','ex','amt']); df['ex']=pd.to_datetime(df.ex)
    if px is not None and spl is not None: df=clean_divs(df,px,spl)
    return df
def build():
    px=load_px(); spl=split_factors(); divs=load_divs(px,spl)
    rets=monthly_returns(px,spl,divs)
    RET={s:g.set_index('m') for s,g in rets.groupby('sym')}
    pxd={s:g.set_index('m').close for s,g in px.groupby('sym')}
    v=pd.read_pickle(R+'sym_valid.pkl')
    a=pd.read_pickle(R+'acq_years.pkl')
    ev=a[(~a.fin)&(~a.util)&a.d201.notna()&(a.d201>='2012-01-01')&(a.d201<='2021-12-31')].copy()
    fm=firm_symbol_months(v,pxd,sorted(set(ev.cik)))
    fb=fallback_symbols([c for c in set(ev.cik) if c not in fm],pxd)
    fm.update(fb); ev['validated']=~ev.cik.isin(list(fb))
    fm=fill_gaps(fm,v,pxd)
    # firm monthly return series
    FR={}
    for c,d in fm.items():
        rr=[]; cl=[]
        for m,s in d[['m','sym']].values:
            g=RET[s]
            rr.append(g.ret.get(m,np.nan)); cl.append(g.close.get(m,np.nan))
        FR[c]=pd.DataFrame({'ret':rr,'close':cl,'sym':d.sym.values},index=pd.PeriodIndex(d.m.values,freq='M'))
    FR,log=repair(FR,spl)
    global REPAIRS; REPAIRS={k:len(x) for k,x in log.items()}
    print('return repairs:',REPAIRS)
    return ev, FR
def fill_gaps(fm,v,pxd):
    """Months inside a firm's span with no price for the symbol chosen that year (typically a ticker change): use another
    validated symbol of the same firm that has a price that month, preferring the symbol of the next observed month."""
    good=v[v.ratio.between(0.2,1.2)].groupby('cik').sym.apply(set).to_dict()
    out={}
    for c,d in fm.items():
        d=d.sort_values('m').reset_index(drop=True); have=dict(zip(d.m,d.sym))
        syms=sorted(good.get(c,set())|set(d.sym))
        full=pd.period_range(d.m.min(),d.m.max(),freq='M'); rows=[]
        for m in full:
            if m in have: rows.append((m,have[m])); continue
            nxt=next((have[x] for x in full if x>m and x in have),None); prv=rows[-1][1] if rows else None
            opts=[s for s in [nxt,prv]+syms if s is not None and s in pxd and m in pxd[s].index]
            if opts: rows.append((m,opts[0]))
        out[c]=pd.DataFrame(rows,columns=['m','sym'])
    return out
SPLIT_K=[2,3,4,5,6,8,10,12,15,20,25,30,40,50]
def repair(FR,spl):
    """Three data repairs, all checked against the firm's own cover-page share counts:
    1. ticker switch: the first month under a new symbol has no return inside that symbol; splice it from the old close
       (adjusted for splits recorded on the new symbol that month) when the result is a plausible monthly move;
    2. a recorded split applied one month early or late (two offsetting jumps): move the factor to the right month;
    3. a price jump of about k-for-1 (or 1-for-k) with no split recorded, or a recorded split with no change in shares
       outstanding: correct the return when the share count confirms it."""
    sh=pd.read_pickle(R+'shares.pkl'); sh['cik']=sh.cik.astype(int); SH={c:g.sort_values('date') for c,g in sh.groupby('cik')}
    SP={}
    for sym,g in spl.groupby('sym'):
        for e,f in g[['ex','f']].values: SP.setdefault(sym,{}).setdefault(pd.Timestamp(e).to_period('M'),1.0); SP[sym][pd.Timestamp(e).to_period('M')]*=f
    cand=SPLIT_K+[1/k for k in SPLIT_K]
    near=lambda x: min(cand,key=lambda q:abs(np.log(x/q)))
    big=lambda x: np.isfinite(x) and abs(np.log(x))>np.log(1.8)
    log={'splice':[],'pair':[],'split':[]}
    def share_ratio(c,m):
        g=SH.get(c)
        if g is None: return np.nan
        t0=(m-1).to_timestamp(how='end'); t1=m.to_timestamp(how='end')
        b=g[(g.date<=t0)&(g.date>=t0-pd.Timedelta(days=200))]; a=g[(g.date>=t1-pd.Timedelta(days=20))&(g.date<=t1+pd.Timedelta(days=200))]
        if not len(b) or not len(a): return np.nan
        return a.shares.iloc[0]/b.shares.iloc[-1]
    for c,f in FR.items():
        f=f.copy(); r=f.ret.values.copy(); cl=f.close.values; sy=f.sym.values; ms=list(f.index)
        for i in range(1,len(f)):
            if sy[i]!=sy[i-1] and not np.isfinite(r[i]) and (ms[i]-ms[i-1]).n==1 and cl[i-1]>0 and np.isfinite(cl[i]):
                x=cl[i]*SP.get(sy[i],{}).get(ms[i],1.0)/cl[i-1]
                if abs(np.log(x))<np.log(1.6): r[i]=x-1; log['splice'].append((c,str(ms[i])))
        g=1+r
        for i in range(len(f)-1):
            if big(g[i]) and big(g[i+1]) and abs(np.log(g[i]*g[i+1]))<np.log(1.4):
                q=near(g[i]); recs=[SP.get(sy[j],{}).get(ms[j],1.0) for j in (i,i+1)]
                if not any(abs(np.log(x/q))<np.log(1.15) or abs(np.log(x*q))<np.log(1.15) for x in recs if abs(np.log(x))>1e-9): continue
                if abs(np.log(g[i]/q))<np.log(1.4) and abs(np.log(g[i+1]*q))<np.log(1.4):
                    g[i]/=q; g[i+1]*=q; log['pair'].append((c,str(ms[i])))
        for i in range(1,len(f)):
            if not big(g[i]): continue
            q=near(g[i]); s=share_ratio(c,ms[i])
            if abs(np.log(g[i]/q))>np.log(1.15) or not np.isfinite(s): continue
            rec=SP.get(sy[i],{}).get(ms[i],1.0)
            if abs(np.log(s*q))<np.log(1.25) and abs(np.log(rec))<1e-9:          # price fell ~1/k, shares rose ~k: split missing
                g[i]*=1/q; log['split'].append((c,str(ms[i]),'missing'))
            elif abs(np.log(rec))>np.log(1.5) and abs(np.log(s))<np.log(1.25) and abs(np.log(q/rec))<np.log(1.15):   # split recorded, shares unchanged
                g[i]/=rec; log['split'].append((c,str(ms[i]),'spurious'))
        f['ret']=g-1; FR[c]=f
    return FR,log
def shares_at(sh, cik, date, tol=150):
    g=sh.get(cik)
    if g is None: return np.nan
    dd=(g.date-date).abs(); i=dd.idxmin()
    return g.shares[i] if dd[i]<=pd.Timedelta(days=tol) else np.nan

# %% [markdown]
# ## 6. Event study
# * **Sample**: market value at the month before completion (price times the nearest cover-page share count); deals of at
#   least 10% of it; one deal per acquirer every 36 months.
# * **Benchmark**: one of the 25 value-weighted size and book-to-market portfolios, NYSE breakpoints at completion.
# * **BHAR** = compounded firm return minus compounded benchmark return, months 1 to H. Inference: Johnson skewness-adjusted
#   t-statistic, bootstrap p-value (1,000 resamples), Wilcoxon signed-rank on the median (Lyon, Barber and Tsai, 1999).

# %%
LAST=pd.Period('2025-03','M')
HOR=[12,24,36,60]
def q5(x,cuts):
    return int(np.searchsorted(cuts,x,side='right'))+1 if np.isfinite(x) else np.nan
def equity_table():
    c=cand_long(); e=c[c.tag.isin(['StockholdersEquity','StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest'])].copy()
    e['pri']=(e.tag=='StockholdersEquity').astype(int)
    e=e.sort_values(['cik','end','pri']).drop_duplicates(['cik','end'],keep='last')
    return {int(k):g[['end','val']].reset_index(drop=True) for k,g in e.groupby('cik')}
def events(ev, FR, min_rs=0.10, gap=36):
    sh=pd.read_pickle(R+'shares.pkl'); sh['cik']=sh.cik.astype(int); SH={c:g.reset_index(drop=True) for c,g in sh.groupby('cik')}
    EQ=equity_table(); MEB=me_bp(); BMB=beme_bp()
    rows=[]
    for r in ev.itertuples():
        if r.cik not in FR: continue
        f=FR[r.cik]; t=pd.Period(r.d201,'M'); t0=t-1
        if t0 not in f.index or not np.isfinite(f.close.get(t0,np.nan)): continue
        d0=t0.to_timestamp(how='end')
        g=SH.get(r.cik)
        if g is None: continue
        dd=(g.date-d0).abs(); i=dd.idxmin()
        if dd[i]>pd.Timedelta(days=150): continue
        me=f.close[t0]*g.shares[i]
        if not me>0: continue
        # shares issued around the deal: last cover before t-1 vs first cover 1-9 months after completion
        sb=g[(g.date<=d0)&(g.date>=d0-pd.Timedelta(days=150))]; sa=g[(g.date>=(t+1).to_timestamp())&(g.date<=(t+9).to_timestamp(how='end'))]
        dsh=(sa.shares.iloc[0]-sb.shares.iloc[-1]) if len(sb) and len(sa) else np.nan
        stock_est=dsh*f.close[t0] if np.isfinite(dsh) and dsh>0.03*sb.shares.iloc[-1] else 0.0
        eq=EQ.get(r.cik); be=np.nan
        if eq is not None:
            ok=eq[(eq.end<=d0-pd.Timedelta(days=90))&(eq.end>=d0-pd.Timedelta(days=550))]
            if len(ok): be=ok.val.iloc[-1]
        mc=MEB.loc[t0].values if t0 in MEB.index else MEB.iloc[-1].values
        szq=q5(me/1e6,mc[[3,7,11,15]])
        yr=t.year if t.month>=7 else t.year-1
        bc=BMB.loc[yr].values if yr in BMB.index else BMB.iloc[-1].values
        bm=be/me if np.isfinite(be) else np.nan
        bmq=(1 if bm<=0 else q5(bm,bc[[3,7,11,15]])) if np.isfinite(bm) else np.nan
        dv=r.dv; ss=r.stock_share
        eq_iss=dsh/sb.shares.iloc[-1] if np.isfinite(dsh) and len(sb) and sb.shares.iloc[-1]>0 else np.nan
        rows.append(dict(cik=r.cik,name=r.name,t=t,cy=r.cy,dv=dv,dv_src=r.dv_src,cash=r.cash,stock=r.stock,stock_est=stock_est,eq_iss=eq_iss,stock_share=ss,
                         gw_acq=r.gw_acq,sic=r.sic,me=me,rs=dv/me,be=be,bm=bm,szq=szq,bmq=bmq,n201=r.n201,validated=r.validated))
    E=pd.DataFrame(rows).sort_values(['t','cik'])
    E=E[E.rs>=min_rs]
    keep=[]; last={}
    for r in E.itertuples():
        if r.cik in last and (r.t-last[r.cik]).n<gap: continue
        keep.append(r.Index); last[r.cik]=r.t
    E=E.loc[keep].reset_index(drop=True)
    return E
def bench_series(E):
    VW,EW=p25(); cols=list(VW.columns)
    E=E.copy()
    E['bmq_f']=E.bmq.fillna(3)          # missing book equity -> middle BM quintile (flagged)
    E['port']=[cols[int((s-1)*5+(b-1))] if np.isfinite(s) else None for s,b in zip(E.szq,E.bmq_f)]
    return E,VW,EW
def bhar(E, FR, VW, H):
    out=[]; paths=[]
    for r in E.itertuples():
        t=r.t; f=FR[r.cik].ret; b=VW[r.port]
        end=t+H
        if end>LAST: out.append(np.nan); continue
        idx=pd.period_range(t+1,end,freq='M')
        bb=b.reindex(idx).values
        if np.isnan(bb).any(): out.append(np.nan); continue
        fr=f.reindex(idx).values
        if np.isnan(fr[0]): out.append(np.nan); continue
        alive=~np.isnan(fr)
        fr=np.where(alive,fr,bb)                     # after delisting / gaps: earn the benchmark
        out.append(np.prod(1+fr)-np.prod(1+bb))
    return np.array(out)
def bhar_path(E, FR, VW, H=36):
    P=[]
    for r in E.itertuples():
        t=r.t; idx=pd.period_range(t-12,t+H,freq='M')
        if t+1>LAST: continue
        f=FR[r.cik].ret.reindex(idx).values; b=VW[r.port].reindex(idx).values
        f=np.where(np.isnan(f),b,f)
        # cumulative from event month end (k=0) forwards and backwards
        post=np.cumprod(1+f[13:])-np.cumprod(1+b[13:])
        P.append(np.r_[np.nan*np.ones(13),post])
    return np.array(P)
def skew_t(x):
    x=np.asarray(x); x=x[np.isfinite(x)]; n=len(x); S=x.mean()/x.std(ddof=1); g=((x-x.mean())**3).mean()/x.std(ddof=0)**3
    return np.sqrt(n)*(S+g*S*S/3+g/(6*n))
def boot_p(x, B=2000, seed=7):
    x=np.asarray(x); x=x[np.isfinite(x)]; n=len(x); rng=np.random.default_rng(seed); t0=skew_t(x)
    xc=x-x.mean(); ts=np.array([skew_t(rng.choice(xc,size=max(n//4,20),replace=True)) for _ in range(B)])
    lo=np.mean(ts<=t0); return 2*min(lo,1-lo)
def summ(x):
    x=np.asarray(x); x=x[np.isfinite(x)]
    from scipy import stats
    return dict(n=len(x),mean=x.mean(),median=np.median(x),pos=(x>0).mean(),t=stats.ttest_1samp(x,0).statistic,
                t_skew=skew_t(x),p_boot=boot_p(x),p_wilcoxon=stats.wilcoxon(x).pvalue)

# %% [markdown]
# ## 7. Calendar time, operations and cross-section
# * **Calendar time**: each month, the portfolio of all acquirers within 36 months of completion (equal- or value-weighted),
#   regressed on CAPM, Fama-French 3 and 5 factors plus momentum; Newey-West t-statistics, three lags.
# * **Operations**: operating income over average assets minus the two-digit SIC industry median of all SEC filers; post-deal
#   (years +1 to +3) regressed on pre-deal (year -1), as in Healy, Palepu and Ruback (1992). Goodwill impairments to fiscal 2024.
# * **Cross-section**: sorts and an OLS of the winsorized 36-month BHAR on deal traits, clustered by completion year.
# * **Placebo**: the same acquirers in three-year windows after their deal with no other $100m+ deal, 50 random draws.

# %%
LAST=pd.Period('2025-03','M')
def portfolio(E, FR, H=36, weight='ew', mask=None, min_n=10):
    E=E if mask is None else E[mask]
    months=pd.period_range('2012-02','2025-03',freq='M'); out={}
    rets={}; w={}
    for r in E.itertuples():
        s=FR[r.cik].ret
        for k in range(1,H+1):
            m=r.t+k
            if m>LAST: break
            v=s.get(m,np.nan)
            if np.isfinite(v): rets.setdefault(m,[]).append(v); w.setdefault(m,[]).append(r.me)
    rows=[]
    for m in months:
        x=np.array(rets.get(m,[]))
        if len(x)<min_n: continue
        if weight=='ew': rp=x.mean()
        else: ww=np.array(w[m]); rp=(x*ww).sum()/ww.sum()
        rows.append((m,rp,len(x)))
    return pd.DataFrame(rows,columns=['m','rp','n']).set_index('m')
def alpha(P, model='ff5m'):
    f=factors(); d=P.join(f,how='inner'); y=d.rp-d.RF
    X={'capm':['Mkt-RF'],'ff3':['Mkt-RF','SMB','HML'],'ff5m':['Mkt-RF','SMB','HML','RMW','CMA','MOM']}[model]
    m=sm.OLS(y,sm.add_constant(d[X])).fit(cov_type='HAC',cov_kwds={'maxlags':3})
    return dict(alpha_m=m.params['const'],t=m.tvalues['const'],alpha_y=(1+m.params['const'])**12-1,n_months=len(d),
                avg_firms=P.n.mean(),**{f'b_{k}':m.params[k] for k in X})
def longshort(P1,P2,model='ff5m'):
    d=P1[['rp']].join(P2[['rp']],lsuffix='1',rsuffix='2',how='inner'); d['rp']=d.rp1-d.rp2
    f=factors(); d=d.join(f,how='inner'); X={'ff3':['Mkt-RF','SMB','HML'],'ff5m':['Mkt-RF','SMB','HML','RMW','CMA','MOM']}[model]
    m=sm.OLS(d.rp,sm.add_constant(d[X])).fit(cov_type='HAC',cov_kwds={'maxlags':3})
    return dict(alpha_m=m.params['const'],t=m.tvalues['const'],alpha_y=(1+m.params['const'])**12-1,n_months=len(d))

def firm_year():
    c=annual(cand_long()); c=c.reset_index(); c['cand']=True
    a=all_long(); a['cy']=a.per.str[2:6].astype(int)
    af=a[~a.per.str.contains('Q')].pivot_table(index=['cik','cy'],columns='tag',values='val',aggfunc='first')
    ai=a[a.per.str.endswith('Q4I')].pivot_table(index=['cik','cy'],columns='tag',values='val',aggfunc='first')
    A=af.join(ai,how='outer')
    A['rev']=np.nan
    for r in REV:
        if r in A: A['rev']=A['rev'].fillna(A[r])
    A=A.reset_index(); A['cand']=False
    F=pd.concat([c,A],ignore_index=True)
    sic=load('sic'); SUB=load('subs')
    F['sic']=pd.to_numeric(F.cik.astype(int).astype(str).map(lambda k:(SUB.get(k) or [None,None])[1] or sic.get(k)),errors='coerce')
    F=F.sort_values(['cik','cy'])
    F['assets_l']=F.groupby('cik').Assets.shift(1)
    F.loc[F.groupby('cik').cy.diff()!=1,'assets_l']=np.nan
    F['avg_assets']=F[['Assets','assets_l']].mean(axis=1)
    F['roa']=F.OperatingIncomeLoss/F.avg_assets
    F['sic2']=(F.sic//100)
    ok=(F.Assets>1e7)&F.roa.between(-1,1)
    med=F[ok].groupby(['sic2','cy']).roa.agg(['median','count']).rename(columns={'median':'ind_roa','count':'ind_n'})
    F=F.join(med,on=['sic2','cy'])
    F['roa_adj']=np.where(F.ind_n>=5,F.roa-F.ind_roa,np.nan)
    F['cfo_roa']=F.NetCashProvidedByUsedInOperatingActivities/F.avg_assets if 'NetCashProvidedByUsedInOperatingActivities' in F else np.nan
    return F
def op_perf(E,F):
    G=F.set_index(['cik','cy'])
    rows=[]
    for r in E.itertuples():
        def get(y,col):
            try: v=G.loc[(r.cik,y),col]; return float(v.iloc[0]) if hasattr(v,'iloc') else float(v)
            except KeyError: return np.nan
        pre=get(r.cy-1,'roa_adj'); post=[get(r.cy+k,'roa_adj') for k in (1,2,3)]
        cpre=get(r.cy-1,'cfo_roa'); cpost=[get(r.cy+k,'cfo_roa') for k in (1,2,3)]
        imp=[get(r.cy+k,'GoodwillImpairmentLoss') for k in range(0,6) if r.cy+k<=2024]
        gw=get(r.cy-1,'Goodwill'); gw0=get(r.cy,'Goodwill')
        rows.append(dict(pre_adj=pre,post_adj=np.nanmean(post) if np.isfinite(post).any() else np.nan,post_n=int(np.isfinite(post).sum()),
                         cfo_pre=cpre,cfo_post=np.nanmean(cpost) if np.isfinite(cpost).any() else np.nan,
                         imp_sum=np.nansum(imp),imp_any=bool(np.nansum(imp)>0),imp_years=len(imp),gw_end=gw0))
    return pd.concat([E.reset_index(drop=True),pd.DataFrame(rows)],axis=1)

def runup(E,FR,VW):
    out=[]
    for r in E.itertuples():
        idx=pd.period_range(r.t-12,r.t-1,freq='M'); f=FR[r.cik].ret.reindex(idx).values; b=VW[r.port].reindex(idx).values
        if np.isnan(f).sum()>2 or np.isnan(b).any(): out.append(np.nan); continue
        f=np.where(np.isnan(f),b,f); out.append(np.prod(1+f)-np.prod(1+b))
    return np.array(out)
def serial(E):
    a=pd.read_pickle(R+'acq_years.pkl'); a=a[a.dv>=1e8]
    s=[]
    for r in E.itertuples():
        g=a[(a.cik==r.cik)&(a.cy<r.cy)&(a.cy>=r.cy-5)]; s.append(len(g))
    return np.array(s)
def features(E,FR,VW):
    X=E.copy()
    X['runup']=runup(X,FR,VW); X['n_prior']=serial(X); X['serial']=(X.n_prior>=2).astype(int)
    X['pay']=np.select([X.stock_share.fillna(0)<0.05, X.stock_share>=0.5],['Cash','Stock-heavy'],'Mixed')
    X['eq_issuer']=(X.eq_iss>=0.10).astype(int)
    X['log_rs']=np.log(X.rs.clip(upper=5)); X['log_me']=np.log(X.me/1e6)
    X['gw_int']=(X.gw_acq/X.dv).clip(0,1.5)
    X['glamour']=(X.bmq==1).astype(int)
    X['late']=(X.t.dt.year>=2017).astype(int)
    X['sic1']=(X.sic//1000).astype('Int64').astype(str)
    X['yr']=X.t.dt.year
    return X
def wins(s,p=0.01):
    lo,hi=s.quantile([p,1-p]); return s.clip(lo,hi)
def regress(X,y='bhar36'):
    d=X[X[y].notna()&X.runup.notna()].copy(); d['y']=wins(d[y])
    f='y ~ stock_share + log_rs + log_me + gw_int + serial + runup + glamour + late + C(sic1)'
    d=d.dropna(subset=['stock_share','gw_int'])
    m=smf.ols(f,d).fit(cov_type='cluster',cov_kwds={'groups':d.yr})
    return m

def subgroups(X):
    g={}
    def add(name,mask_dict):
        g[name]={k:summ(X.loc[m,'bhar36']) for k,m in mask_dict.items()}
    add('Payment (disclosed)',{'Cash':X.pay=='Cash','Mixed':X.pay=='Mixed','Stock-heavy':X.pay=='Stock-heavy'})
    add('New shares issued around the deal',{'Under 10%':X.eq_issuer==0,'10% or more':X.eq_issuer==1})
    q=X.rs.quantile([1/3,2/3]).values
    add('Deal size vs acquirer',{'Small (10%% to %.1f%%)'%(q[0]*100):X.rs<q[0],'Medium':(X.rs>=q[0])&(X.rs<q[1]),'Large (%.1f%% or more)'%(q[1]*100):X.rs>=q[1]})
    add('Acquirer valuation',{'Glamour (lowest book-to-market)':X.bmq==1,'Middle':X.bmq.isin([2,3]),'Value (top two quintiles)':X.bmq.isin([4,5])})
    add('Acquisition history',{'Occasional (0-1 prior)':X.serial==0,'Serial (2+ in 5 years)':X.serial==1})
    ru=X.runup.quantile([1/3,2/3]).values
    add('Stock run-up before the deal',{'Low':X.runup<ru[0],'Middle':(X.runup>=ru[0])&(X.runup<ru[1]),'High':X.runup>=ru[1]})
    gi=X.gw_int.quantile([1/3,2/3]).values
    add('Goodwill share of price',{'Low':X.gw_int<gi[0],'Middle':(X.gw_int>=gi[0])&(X.gw_int<gi[1]),'High':X.gw_int>=gi[1]})
    add('Period',{'2012-2016':X.late==0,'2017-2021':X.late==1})
    return g
def robustness(ev,FR,VW,EW,X):
    f=factors(); MK=pd.DataFrame({p:(f['Mkt-RF']+f['RF']) for p in VW.columns})
    rows={}
    rows['Base: size/book-to-market portfolios, value-weighted']=summ(X.bhar36)
    rows['Equal-weighted size/book-to-market portfolios']=summ(X.bhar36_ew)
    rows['Market portfolio']=summ(X.bhar36_mkt)
    for mr,lab in [(0.05,'Deals of 5% or more of acquirer value'),(0.25,'Deals of 25% or more')]:
        E=events(ev,FR,min_rs=mr); E,_,_=bench_series(E); rows[lab]=summ(bhar(E,FR,VW,36))
    E=events(ev,FR,gap=0); E,_,_=bench_series(E); rows['All deals, overlaps included']=summ(bhar(E,FR,VW,36))
    rows['Price matched to float only (drop name-only matches)']=summ(X.loc[X.validated,'bhar36'])
    rows['Deal value from cash-flow and equity tags only']=summ(X.loc[X.dv_src!='goodwill','bhar36'])
    w=wins(X.bhar36.dropna()); rows['Winsorized at 1% and 99%']=summ(w)
    return rows
def delist_stats(X,FR,H=36):
    out=[]
    for r in X.itertuples():
        s=FR[r.cik].ret; last=s.dropna().index.max()
        out.append(last < min(r.t+H, pd.Period('2025-03','M')))
    return np.array(out)
def op_path(X,F):
    G=F.set_index(['cik','cy']); res={}
    for k in range(-2,4):
        v=[]; c=[]
        for r in X.itertuples():
            try: row=G.loc[(r.cik,r.cy+k)]
            except KeyError: v.append(np.nan); c.append(np.nan); continue
            row=row.iloc[0] if isinstance(row,pd.DataFrame) else row
            v.append(row.roa_adj); c.append(row.cfo_roa)
        v=np.array(v,float); c=np.array(c,float)
        res[k]=dict(roa_adj_med=np.nanmedian(v),roa_adj_mean=float(np.nanmean(np.clip(v,*np.nanpercentile(v,[1,99])))) if np.isfinite(v).any() else np.nan,cfo_med=np.nanmedian(c),n=int(np.isfinite(v).sum()))
    return res
def healy(X):
    ok=X.pre_adj.notna()&X.post_adj.notna()&(X.post_n>=2)
    d=X[ok]; y=wins(d.post_adj); x=wins(d.pre_adj)
    m=sm.OLS(y,sm.add_constant(x)).fit(cov_type='HC1')
    return dict(n=int(ok.sum()),a=m.params.iloc[0],t_a=m.tvalues.iloc[0],b=m.params.iloc[1],t_b=m.tvalues.iloc[1],r2=m.rsquared,
                pre_med=d.pre_adj.median(),post_med=d.post_adj.median(),chg_med=(d.post_adj-d.pre_adj).median(),
                cfo_pre_med=d.cfo_pre.median(),cfo_post_med=d.cfo_post.median())
def impairment_curve(X,F):
    G=F.set_index(['cik','cy']).GoodwillImpairmentLoss
    res={}
    for k in range(1,6):
        m=X.cy+k<=2024; hit=[]
        for r in X[m].itertuples():
            s=0.0
            for j in range(0,k+1):
                try: v=G.loc[(r.cik,r.cy+j)]; v=float(v.iloc[0]) if hasattr(v,'iloc') else float(v); s+=0 if np.isnan(v) else v
                except KeyError: pass
            hit.append((s>0, s>=0.10*r.dv))
        h=np.array(hit)
        res[k]=dict(n=len(h),any=h[:,0].mean(),big=h[:,1].mean())
    return res

_C={}
def _ctx():
    if not _C:
        sh=pd.read_pickle(R+'shares.pkl'); sh['cik']=sh.cik.astype(int)
        _C['SH']={c:g.reset_index(drop=True) for c,g in sh.groupby('cik')}
        _C['EQ']=equity_table(); _C['MEB']=me_bp(); _C['BMB']=beme_bp(); _C['cols']=list(p25()[0].columns)
    return _C
def port_at(cik,t,FR):
    """Size and book-to-market portfolio of a firm at month t, same rules as for the deals."""
    C=_ctx(); f=FR[cik]; t0=t-1; d0=t0.to_timestamp(how='end')
    if t0 not in f.index or not np.isfinite(f.close.get(t0,np.nan)): return None
    g=C['SH'].get(cik)
    if g is None: return None
    dd=(g.date-d0).abs(); i=dd.idxmin()
    if dd[i]>pd.Timedelta(days=150): return None
    me=f.close[t0]*g.shares[i]
    if not me>0: return None
    eq=C['EQ'].get(cik); be=np.nan
    if eq is not None:
        ok=eq[(eq.end<=d0-pd.Timedelta(days=90))&(eq.end>=d0-pd.Timedelta(days=550))]
        if len(ok): be=ok.val.iloc[-1]
    MEB,BMB=C['MEB'],C['BMB']
    mc=MEB.loc[t0].values if t0 in MEB.index else MEB.iloc[-1].values
    szq=q5(me/1e6,mc[[3,7,11,15]])
    yr=t.year if t.month>=7 else t.year-1
    bc=BMB.loc[yr].values if yr in BMB.index else BMB.iloc[-1].values
    bm=be/me if np.isfinite(be) else np.nan
    bmq=(1 if bm<=0 else q5(bm,bc[[3,7,11,15]])) if np.isfinite(bm) else 3
    return C['cols'][int((szq-1)*5+(bmq-1))]
def placebo(X,FR,VW,seed=11,side='after'):
    """Same acquirers, start months with no acquisition of $100m+ within three years either side.
    The benchmark portfolio is re-assigned at the placebo date. side='after' keeps only windows starting after the firm's first sample deal."""
    a=pd.read_pickle(R+'acq_years.pkl'); a=a[a.dv>=1e8]; busy={c:set(g.cy) for c,g in a.groupby('cik')}
    rng=np.random.default_rng(seed); rows=[]
    for c,g in X.groupby('cik'):
        s=FR[c].ret.dropna()
        if len(s)<60: continue
        first=g.t.min()
        cands=[m for m in s.index if m+36<=pd.Period('2025-03','M') and m>=pd.Period('2012-01','M')
               and not any(abs(m.year-y)<=3 for y in busy.get(c,())) and (side=='any' or m>first)]
        rng.shuffle(cands)
        for m in cands[:5]:
            pt=port_at(c,m,FR)
            if pt is not None: rows.append(dict(cik=c,t=m,port=pt,pre=m<first)); break
    P=pd.DataFrame(rows); P['bhar36']=bhar(P,FR,VW,36)
    return P
def placebo_draws(X,FR,VW,n=50):
    """Repeat the random draw n times; also compare each firm's own deal BHAR with its quiet-window BHAR."""
    dealm=X.dropna(subset=['bhar36']).groupby('cik').bhar36.mean()
    out=[]
    for sd in range(n):
        P=placebo(X,FR,VW,seed=100+sd,side='after').dropna(subset=['bhar36'])
        sm=summ(P.bhar36)
        d=(dealm.reindex(P.cik).values-P.bhar36.values); d=d[np.isfinite(d)]
        out.append(dict(t_skew=sm['t_skew'],mean=sm['mean'],median=sm['median'],pos=sm['pos'],n=len(P),dmean=d.mean(),dmed=np.median(d),dt=d.mean()/(d.std(ddof=1)/np.sqrt(len(d))),dn=len(d)))
    o=pd.DataFrame(out)
    return dict(draws=n,n=int(o.n.median()),t_skew=o.t_skew.mean(),mean=o['mean'].mean(),median=o['median'].mean(),pos=o['pos'].mean(),
                diff_mean=o.dmean.mean(),diff_median=o.dmed.mean(),diff_t=o.dt.mean(),diff_n=int(o.dn.median()),share_sig=(o.dt<-1.96).mean())

# %% [markdown]
# ## 8. Run the study

# %%
def run():
    ev,FR=build(); pd.to_pickle((ev,FR),R+"panel_tmp.pkl")
    E=events(ev,FR); E,VW,EW=bench_series(E)
    for H in HOR:
        E[f'bhar{H}']=bhar(E,FR,VW,H); E[f'bhar{H}_ew']=bhar(E,FR,EW,H)
    # market benchmark
    f=factors(); MK=pd.DataFrame({p:(f['Mkt-RF']+f['RF']) for p in VW.columns})
    E['bhar36_mkt']=bhar(E,FR,MK,36)
    X=features(E,FR,VW)
    F=firm_year(); X=op_perf(X,F)
    res={}
    res['summ']={H:summ(X[f'bhar{H}']) for H in HOR}
    res['summ_ew']={H:summ(X[f'bhar{H}_ew']) for H in (12,36)}
    res['summ_mkt']=summ(X['bhar36_mkt'])
    res['by_pay']={p:summ(g.bhar36) for p,g in X.groupby('pay')}
    res['path']=bhar_path(X,FR,VW,36)
    ct={}
    for w in ('ew','vw'):
        P=portfolio(X,FR,weight=w); ct[w]={m:alpha(P,m) for m in ('capm','ff3','ff5m')}; ct[w]['series']=P
    Pc=portfolio(X,FR,mask=X.pay=='Cash'); Ps=portfolio(X,FR,mask=X.pay!='Cash')
    Pe=portfolio(X,FR,mask=X.eq_issuer==1); Pn=portfolio(X,FR,mask=X.eq_issuer==0)
    ct['eq']=alpha(Pe,'ff5m'); ct['noeq']=alpha(Pn,'ff5m'); ct['eq_minus_noeq']=longshort(Pe,Pn)
    ct['cash']=alpha(Pc,'ff5m'); ct['stock']=alpha(Ps,'ff5m'); ct['stock_minus_cash']=longshort(Ps,Pc)
    res['ct']=ct
    res['xsec']=regress(X)
    res['groups']=subgroups(X); res['robust']=robustness(ev,FR,VW,EW,X)
    X['delisted36']=delist_stats(X,FR); res['op_path']=op_path(X,F); res['healy']=healy(X); res['impair']=impairment_curve(X,F)
    Pl=placebo(X,FR,VW); res['placebo_one']=summ(Pl.bhar36); res['placebo']=placebo_draws(X,FR,VW)
    a=__import__('pandas').read_pickle(R+'acq_years.pkl')
    n_me=len(events(ev,FR,min_rs=0,gap=0)); n_rs=len(events(ev,FR,min_rs=0.10,gap=0))
    res['funnel']=dict(acq_years=int(len(a)), nonfin=int(((~a.fin)&(~a.util)).sum()), with201=int(len(ev)), priced=int(ev.cik.isin(list(FR)).sum()),
                       with_me=int(n_me), rs10=int(n_rs), final=int(len(X)), n36=int(X.bhar36.notna().sum()), n60=int(X.bhar60.notna().sum()),
                       by_mar2020=int((X.t<=pd.Period('2020-03','M')).sum()))
    res['repairs']=REPAIRS
    se=ct['ew']['series']; res['ct_range']=dict(first=str(se.index.min()),last=str(se.index.max()),n=int(len(se)))
    return X,FR,VW,EW,res

X, FR, VW, EW, res = run()
pd.to_pickle((X, res), R + "results.pkl")
pd.DataFrame(res["summ"]).T[["n", "mean", "median", "pos", "t", "t_skew", "p_boot", "p_wilcoxon"]]

# %%
rows = []
for w in ("ew", "vw"):
    for m in ("capm", "ff3", "ff5m"):
        r = res["ct"][w][m]; rows.append((w, m, r["alpha_m"], r["t"], r["alpha_y"]))
for k in ("cash", "stock", "stock_minus_cash", "eq", "noeq", "eq_minus_noeq"):
    r = res["ct"][k]; rows.append((k, "ff5m", r["alpha_m"], r["t"], r.get("alpha_y", r["alpha_m"] * 12)))
pd.DataFrame(rows, columns=["portfolio", "model", "alpha / month", "t (Newey-West)", "alpha / year"])

# %%
fig, ax = plt.subplots(1, 2, figsize=(14, 4.5))
P = res["path"]; post = np.hstack([np.zeros((len(P), 1)), P[:, 13:]])[X.bhar36.notna().values]
ax[0].plot(np.nanmean(post, 0), color="#1a1a18", lw=2, label="mean"); ax[0].plot(np.nanmedian(post, 0), color="#b5412b", lw=2, label="median")
ax[0].axhline(0, color="k", lw=.8); ax[0].set_title("BHAR after completion"); ax[0].set_xlabel("months"); ax[0].legend(); ax[0].grid(alpha=.3)
g = pd.DataFrame({k: {"mean": s["mean"], "median": s["median"]} for k, s in res["groups"]["Payment (disclosed)"].items()}).T
g.plot.barh(ax=ax[1], color=["#1a1a18", "#b5412b"]); ax[1].set_title("36-month BHAR by disclosed payment"); ax[1].grid(alpha=.3)
plt.tight_layout(); plt.show()
print("Operating performance (Healy regression):", {k: round(v, 4) for k, v in res["healy"].items()})
print("Goodwill impairments within 5 years:", res["impair"][5])
print("Placebo:", {k: round(v, 3) if isinstance(v, float) else v for k, v in res["placebo"].items()})

# %% [markdown]
# ## 9. Export: the interactive page
# Writes `MA_Long_Run_Performance.html`: the event-time path by deal group, the distribution, calendar-time and robustness
# tables, the cross-section, operating performance and impairments, and a searchable explorer of every deal.

# %%
def pc(x, d=1, sign=False):
    x = x + (1e-9 if x >= 0 else -1e-9)
    s = f'{x*100:+.{d}f}%' if sign else f'{x*100:.{d}f}%'
    return s.replace('-', '&minus;') if '&' not in s else s


def pl(x, d=1, sign=False):
    """Plain-text percentage (no HTML entities), for the deck and Excel."""
    x = x + (1e-9 if x >= 0 else -1e-9)
    return (f'{x*100:+.{d}f}%' if sign else f'{x*100:.{d}f}%').replace('-', '−')


def tt(x):
    return f'{x:.2f}'.replace('-', '&minus;')


def page_text(D, html=True):
    P = pc if html else pl
    T = tt if html else (lambda x: f'{x:.2f}'.replace('-', '−'))
    s36, s60, s12 = D['summ']['36'], D['summ']['60'], D['summ']['12']
    ew, vw = D['ct']['ew']['ff5m'], D['ct']['vw']['ff5m']
    cash, stock, smc = D['ct']['cash'], D['ct']['stock'], D['ct']['stock_minus_cash']
    bp = D['by_pay']
    h, im, pb = D['healy'], D['impair'], D['placebo']
    rb = D['robust']
    rmeans = [v['mean'] for v in rb.values()]
    reg = D['reg']
    trait = [k for k in reg['t'] if not k.startswith('C(') and k != 'Intercept']
    sig = [k for k in trait if abs(reg['t'][k]) >= 1.96]

    signif = lambda t: ('significant at 1%' if abs(t) >= 2.576 else 'significant at 5%' if abs(t) >= 1.96 else
                        'significant only at 10%' if abs(t) >= 1.645 else 'not statistically significant')
    gv, gg = D['groups']['Acquirer valuation'], D['groups']['Goodwill share of price']
    glam, val = gv['Glamour (lowest book-to-market)'], gv['Value (top two quintiles)']

    lead_path = (
        f"The typical acquirer fell behind comparable stocks and kept falling. Twelve months after completion the mean buy-and-hold "
        f"abnormal return, dividends included, is {P(s12['mean'], sign=True)}; at 36 months it is {P(s36['mean'], sign=True)} "
        f"(median {P(s36['median'], sign=True)}, skewness-adjusted t = {T(s36['t_skew'])}), and only {P(s36['pos'], 0)} of acquirers "
        f"beat their benchmark. Deals completed by March 2020 can be followed for five years: the mean reaches {P(s60['mean'], sign=True)}. "
        f"The median sits below the mean because a minority of deals did very well.")

    lead_groups = (
        f"Payment is the one trait that separates deals once returns are measured in calendar time: acquirers that paid partly or "
        f"mostly in stock earned an alpha of {P(stock['alpha_m'], 2, sign=True)} a month (t = {T(stock['t'])}), cash acquirers "
        f"{P(cash['alpha_m'], 2, sign=True)} (t = {T(cash['t'])}), a gap of {P(-smc['alpha_m'], 2)} a month (t = {T(smc['t'])}). "
        f"In event time the sorts point the same way as the literature, with glamour acquirers at {P(glam['mean'], sign=True)} against "
        f"{P(val['mean'], sign=True)} for value acquirers, and deals with the most goodwill doing worst. "
        + ("None of these traits survives the regression once the others are held fixed, so they are better read as warning signs "
           "than as a formula." if not sig else "")
        )

    lead_op = (
        f"Operating returns do not show the synergies that deal announcements promise. The median acquirer earned "
        f"{P(h['pre_med'], sign=True)} above its industry on assets the year before the deal and {P(h['post_med'], sign=True)} "
        f"in the three years after. Part of that fall is mean reversion that any profitable firm shows; the Healy-Palepu-Ruback "
        f"regression isolates the rest, and its intercept of {P(h['a'], 2, sign=True)} (t = {T(h['t_a'])}) says there is no "
        f"abnormal improvement. Meanwhile {P(im['5']['any'], 0)} of acquirers wrote goodwill down within five years, "
        f"{P(im['5']['big'], 0)} by at least a tenth of the deal value.")

    if html:
        verdict = (
            f"<p><b>The base rate is negative.</b> Across {D['funnel']['final']} significant US acquisitions completed from 2012 to 2021, "
            f"the acquirer trailed stocks of similar size and valuation by {P(-s36['mean'], 0)} on average over three years, and {P(1-s36['pos'], 0)} of "
            f"deals ended behind. The same firms, in later three-year windows without a deal, averaged {P(pb['mean'], 0, sign=True)} against the "
            f"same kind of benchmark, although those windows are not a random draw.</p>"
            f"<p><b>Part of it is the benchmark.</b> Calendar-time portfolios, which avoid compounding and overlapping windows, give an "
            f"alpha of {P(ew['alpha_m'], 2, sign=True)} a month equal-weighted (t = {T(ew['t'])}, {signif(ew['t'])}) and "
            f"{P(vw['alpha_m'], 2, sign=True)} value-weighted (t = {T(vw['t'])}, {signif(vw['t'])}) after five factors and momentum. "
            f"Weighted by market value the shortfall is small: it sits mostly in smaller acquirers.</p>"
            f"<p><b>Paying in stock is the clearest warning sign.</b> Acquirers that paid partly or mostly in stock lost about {P(-stock['alpha_y'], 0)} a year against "
            f"the factors; cash acquirers {'were not distinguishable from zero' if abs(cash['t']) < 1.96 else 'lost ' + P(-cash['alpha_y'], 0) + ' a year'}. A board being offered shares, or offering them, should ask "
            f"what the counterparty knows about their value.</p>"
            f"<p><b>Synergies rarely show up in the accounts.</b> Industry-adjusted margins did not improve beyond normal mean reversion, and "
            f"{P(im['5']['any'], 0)} of acquirers impaired goodwill within five years. The integration case deserves more scrutiny than the strategic case.</p>")
    else:
        verdict = None

    limits = [
        "Deals are identified from acquisition spending in XBRL filings and 8-K Item 2.01 dates, not from a commercial deal database. "
        "A year with several deals is treated as one event dated at the first completion, and the target's identity is not recorded.",
        "Deal value comes from cash paid, stock issued and consideration transferred as tagged by each filer; some large stock deals are "
        "tagged incompletely, which is why a separate share-issuance measure is reported next to the disclosed payment mix.",
        f"Prices come from a community-maintained database. {D['funnel']['priced']/D['funnel']['with201']*100:.0f}% of the "
        f"{D['funnel']['with201']:,} candidate deals could be matched to a traded ticker and checked against the public float; the rest "
        "are excluded, which may tilt the sample towards firms that were easier to trace.",
        "The benchmark is a size and book-to-market portfolio. It does not control for industry, and the 2012 to 2024 period favoured "
        "large growth stocks, which the value-weighted and calendar-time tests partly address.",
        f"The placebo uses the same acquirers in three-year windows that start after their deal and contain no other large deal "
        f"(mean {P(pb['mean'], sign=True)}, median {P(pb['median'], sign=True)}, average of {pb['draws']} random draws). Windows before "
        "the deal are left out on purpose: a firm that later makes a large acquisition has usually done well in the meantime, which "
        "would flatter the placebo. Even so, the quiet windows are not a random draw, so the gap is suggestive rather than a clean control.",
        "Dividends come from the same community database, whose amounts are split-adjusted for some securities and not for others; "
        "amounts next to large splits were rescaled to the unadjusted price, and smaller inconsistencies remain.",
        "Long-run abnormal returns show association, not causation: the counterfactual of the same firm without the deal is not observed.",
    ]
    return dict(lead_path=lead_path, lead_groups=lead_groups, lead_op=lead_op, verdict=verdict, limits=limits)

def c(v):
    if isinstance(v,dict): return {str(k):c(x) for k,x in v.items()}
    if isinstance(v,(list,tuple,np.ndarray)): return [c(x) for x in v]
    if isinstance(v,(np.bool_,bool)): return bool(v)
    if isinstance(v,(np.integer,)): return int(v)
    if isinstance(v,(float,np.floating)): return None if not np.isfinite(v) else round(float(v),6)
    if isinstance(v,pd.Period): return str(v)
    return v
SMALL={'OF','AND','THE','FOR','IN','ON','DE'}
KEEP={'IQVIA','ACI','ICON','NCR','CDW','ITT','SPX','RPM','FMC','PTC','MSA','AAON','EPAM','ANSYS','ICU','AMN','WEX','US','USA','II','III','IV','LP','L.P.','N.V.','SE','AG','NV','SA','PLC','LLC','REIT','ICF','PDC','ACCO','ASGN','AZZ','SS&C'}
def _w(w,i):
    core=re.sub(r'[^A-Za-z]','',w)
    if not core or not core.isupper() or w in KEEP or core in KEEP: return w
    if '&' in w or not re.search('[AEIOUY]',core): return w
    if core in SMALL and i>0: return w.lower()
    if w.startswith('MC') and len(w)>3: return 'Mc'+w[2]+w[3:].lower()
    return '-'.join(x[:1]+x[1:].lower() for x in w.split('-'))
def nice(n):
    n=re.sub(r'\s*/[A-Za-z]{2,4}/\s*$','',n.strip())
    words=n.split()
    if n.isupper() or any(len(re.sub(r'[^A-Z]','',w))>=4 and w.isupper() and re.search('[AEIOUY]',w) for w in words):
        return ' '.join(_w(w,i) for i,w in enumerate(words))
    return n
def page_data(X,res):
    P=res['path']; post=np.hstack([np.zeros((len(P),1)),P[:,13:]])
    deals=[]
    for i,r in enumerate(X.itertuples()):
        deals.append(dict(n=nice(r.name),t=str(r.t),dv=r.dv/1e6,rs=r.rs,pay=r.pay,eq=int(r.eq_issuer),
                          bmq=None if not np.isfinite(r.bmq) else int(r.bmq),ser=int(r.serial),ru=r.runup,gw=r.gw_int,sic=None if not np.isfinite(r.sic) else int(r.sic),
                          b12=r.bhar12,b24=r.bhar24,b36=r.bhar36,b60=r.bhar60,p=[round(float(x),4) if np.isfinite(x) else None for x in post[i]],
                          val=bool(r.validated),dl=bool(r.delisted36)))
    ct={w:{m:{k:v for k,v in res['ct'][w][m].items()} for m in ('capm','ff3','ff5m')} for w in ('ew','vw')}
    for k in ('cash','stock','stock_minus_cash','eq','noeq','eq_minus_noeq'): ct[k]=res['ct'][k]
    xs=res['xsec']; reg={'params':xs.params.to_dict(),'t':xs.tvalues.to_dict(),'n':int(xs.nobs),'r2':xs.rsquared}
    out=dict(meta=dict(asof='2025-03-31',start='2012-01',end='2021-12',ct_first=res['ct_range']['first'],ct_last=res['ct_range']['last'],ct_n=res['ct_range']['n'],repairs=res['repairs']),funnel=res['funnel'],summ=res['summ'],summ_ew=res['summ_ew'],summ_mkt=res['summ_mkt'],
             by_pay=res['by_pay'],ct=ct,groups=res['groups'],robust=res['robust'],placebo=res['placebo'],op_path=res['op_path'],healy=res['healy'],impair=res['impair'],
             reg=reg,deals=deals)
    return c(out)

# %% [markdown]
# The page template (HTML, CSS and JavaScript) is kept in the notebook so the project runs from a single file. It loads Plotly and the fonts from public CDNs.

# %%
PAGE_TEMPLATE = r'''<meta charset="utf-8">
<title>Does M&amp;A Create Value?</title>
<meta name="author" content="Alessandro Radice">
<meta name="description" content="Long-run performance of US acquirers, 2012 to 2021: buy-and-hold abnormal returns, calendar-time alphas, operating performance and goodwill impairments for __N__ acquisitions rebuilt from SEC filings.">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js"></script>
<style>
  /* A research-note look: cream paper, ink, one rust accent. Every colour is painted explicitly. */
  :root {
    color-scheme: light;
    --ground: #f7f6f2; --panel: #ffffff; --panel-2: #ecebe5; --rule: #dedcd4; --rule-strong: #c8c7bf;
    --text: #1a1a18; --text-2: #45443f; --muted: #66655e; --faint: #a09f99;
    --acc: #b5412b; --acc-soft: rgba(181, 65, 43, .08); --pos: #2f6f5e; --mid: #8a8983;
    --serif: "Instrument Serif", "Iowan Old Style", "Palatino Linotype", Georgia, serif;
    --sans: "IBM Plex Sans", -apple-system, "Segoe UI", Roboto, sans-serif;
    --mono: "IBM Plex Mono", ui-monospace, "SFMono-Regular", Menlo, Consolas, monospace;
  }
  * { box-sizing: border-box; }
  html { background: var(--ground); }
  body { background: var(--ground); color: var(--text); font-family: var(--sans); font-size: 15px; line-height: 1.55; margin: 0; }
  .wrap { max-width: 1240px; margin: 0 auto; padding-inline: 24px; padding-block: 32px 64px; }
  @media (max-width: 640px) { .wrap { padding-inline: 16px; padding-block: 20px 48px; } }
  .eyebrow { font-family: var(--mono); font-size: 12px; letter-spacing: .08em; text-transform: uppercase; color: var(--muted); }
  h1 { font-family: var(--serif); font-weight: 400; font-size: clamp(42px, 6.4vw, 80px); line-height: 1.0; margin: 10px 0 14px; letter-spacing: -.01em; text-wrap: balance; }
  h1 em { color: var(--acc); font-style: italic; }
  .h1sub { display: block; font-size: .44em; line-height: 1.2; margin-top: 10px; color: var(--text-2); }
  .dek { font-size: 17px; color: var(--text-2); max-width: 66ch; margin: 0; }
  .masthead { display: grid; grid-template-columns: minmax(0, 1.4fr) minmax(0, 1fr); gap: 40px; align-items: end; padding-bottom: 28px; border-bottom: 2px solid var(--text); }
  .brief { border-left: 2px solid var(--acc); padding: 4px 0 4px 18px; }
  .brief .label { font-family: var(--mono); font-size: 11px; letter-spacing: .1em; text-transform: uppercase; color: var(--acc); margin-bottom: 6px; }
  .brief p { margin: 0; font-family: var(--serif); font-size: 22px; line-height: 1.3; }
  @media (max-width: 900px) { .masthead { grid-template-columns: 1fr; gap: 20px; } }
  .findings { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-bottom: 1px solid var(--rule-strong); }
  .finding { padding: 18px 20px 18px 0; }
  .finding + .finding { padding-left: 20px; border-left: 1px solid var(--rule); }
  .finding .v { font-family: var(--mono); font-size: 28px; font-weight: 500; font-variant-numeric: tabular-nums; white-space: nowrap; }
  .finding .v.neg { color: var(--acc); }
  .finding .k { font-size: 13px; color: var(--muted); margin-top: 2px; }
  @media (max-width: 900px) { .findings { grid-template-columns: repeat(2, minmax(0, 1fr)); } .finding:nth-child(3) { border-left: 0; padding-left: 0; } .finding:nth-child(n+3) { border-top: 1px solid var(--rule); } }
  @media (max-width: 640px) { .finding .v { font-size: 20px; } .finding { padding-right: 10px; } .finding + .finding { padding-left: 10px; } }
  section { margin-top: 48px; }
  h2 { font-family: var(--serif); font-weight: 400; font-size: clamp(28px, 3.6vw, 42px); line-height: 1.1; margin: 0 0 8px; text-wrap: balance; }
  h3 { font: 600 12.5px var(--sans); letter-spacing: .06em; text-transform: uppercase; color: var(--text-2); margin: 0 0 2px; }
  .sub { font-size: 13px; color: var(--muted); margin: 0 0 8px; }
  .lead { color: var(--text-2); max-width: 76ch; margin: 0 0 18px; }
  .lead b { color: var(--text); font-weight: 600; }
  .panel { background: var(--panel); border: 1px solid var(--rule); border-radius: 8px; padding: 16px 16px 8px; min-width: 0; }
  .grid2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin-top: 16px; }
  .grid-wide { display: grid; grid-template-columns: minmax(0, 1.55fr) minmax(0, 1fr); gap: 16px; margin-top: 16px; }
  @media (max-width: 1000px) { .grid2, .grid-wide { grid-template-columns: minmax(0, 1fr); } }
  .chart { height: 330px; } #path { height: 400px; } #deal { height: 300px; }
  .seg { display: inline-flex; flex-wrap: wrap; gap: 6px; margin: 4px 0 10px; }
  .seg button { font: 500 12px var(--mono); color: var(--text-2); background: transparent; border: 1px solid var(--rule-strong); border-radius: 999px; padding: 4px 11px; cursor: pointer; }
  .seg button:hover { color: var(--text); border-color: var(--muted); }
  .seg button.on { color: #fff; background: var(--text); border-color: var(--text); }
  .seg button:focus-visible, input:focus-visible { outline: 2px solid var(--acc); outline-offset: 2px; }
  .explain { font-size: 13px; color: var(--text-2); min-height: 40px; margin: 0 0 6px; max-width: 92ch; }
  .tablewrap { overflow-x: auto; margin-top: 12px; }
  table { border-collapse: collapse; width: 100%; font-size: 14px; }
  th, td { padding: 8px 12px; text-align: right; border-bottom: 1px solid var(--rule); white-space: nowrap; font-variant-numeric: tabular-nums; }
  th { font: 500 12px var(--sans); color: var(--muted); letter-spacing: .03em; border-bottom-color: var(--rule-strong); vertical-align: bottom; }
  td { font-family: var(--mono); font-size: 13.5px; }
  th:first-child, td:first-child { text-align: left; font-family: var(--sans); }
  td.l, th.l { text-align: left; font-family: var(--sans); white-space: normal; }
  tr.grp td { font: 600 12px var(--sans); letter-spacing: .05em; text-transform: uppercase; color: var(--text-2); background: var(--panel-2); }
  tr.pick { cursor: pointer; } tr.pick:hover td { background: var(--acc-soft); } tr.sel td { background: var(--acc-soft); }
  .neg { color: var(--acc); } .pos { color: var(--pos); } .dim { color: var(--muted); }
  .search { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin: 6px 0 4px; }
  .search input { font: 14px var(--sans); padding: 7px 10px; border: 1px solid var(--rule-strong); border-radius: 6px; background: var(--panel); color: var(--text); min-width: 0; flex: 1 1 260px; max-width: 420px; }
  .scroll { max-height: 520px; overflow: auto; border: 1px solid var(--rule); border-radius: 8px; background: var(--panel); }
  .scroll th { position: sticky; top: 0; background: var(--panel); z-index: 1; cursor: pointer; }
  .verdict { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 32px; margin-top: 20px; }
  .verdict p { margin: 0 0 12px; color: var(--text-2); }
  .verdict p strong { color: var(--text); font-weight: 600; }
  .verdict .call { font-family: var(--serif); font-size: 26px; line-height: 1.25; color: var(--text); margin: 0 0 12px; }
  @media (max-width: 900px) { .verdict { grid-template-columns: 1fr; gap: 8px; } }
  .method { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 40px; counter-reset: step; margin: 16px 0 0; padding: 0; list-style: none; }
  .method li { counter-increment: step; position: relative; min-width: 0; padding-left: 36px; color: var(--text-2); font-size: 14px; }
  .method li::before { content: counter(step, decimal-leading-zero); position: absolute; left: 0; top: 1px; font: 12px var(--mono); color: var(--acc); }
  .method li b { color: var(--text); font-weight: 600; }
  @media (max-width: 900px) { .method { grid-template-columns: minmax(0, 1fr); } }
  .notes { color: var(--muted); font-size: 13.5px; max-width: 92ch; } .notes li { margin-bottom: 6px; }
  .sources { columns: 2; column-gap: 40px; font-size: 13px; color: var(--muted); padding-left: 18px; }
  .sources li { margin-bottom: 6px; break-inside: avoid; }
  @media (max-width: 800px) { .sources { columns: 1; } }
  a { color: var(--acc); text-underline-offset: 2px; }
  .formula { display: block; max-width: 100%; font-family: var(--mono); font-size: 13.5px; line-height: 1.6; background: var(--panel-2); border-radius: 6px; padding: 10px 14px; margin: 10px 0; overflow-x: auto; white-space: nowrap; }
  footer { margin-top: 48px; padding-top: 16px; border-top: 1px solid var(--rule-strong); display: flex; justify-content: space-between; gap: 16px; flex-wrap: wrap; font-size: 13px; color: var(--muted); }
</style>

<div class="wrap">
  <header class="masthead">
    <div>
      <div class="eyebrow">Radice Capital Partners &middot; M&amp;A research &middot; April 2025</div>
      <h1>Does M&amp;A <em>create value?</em><span class="h1sub">Long-run performance of US acquirers, 2012&ndash;2021</span></h1>
      <p class="dek">Announcement-day returns tell you what the market expects. This page follows __N__ significant acquisitions by US listed companies for three to five years after closing, and measures what actually happened to the acquirer's shareholders, margins and goodwill.</p>
    </div>
    <div class="brief">
      <div class="label">The question</div>
      <p>A board is weighing a large acquisition. Across a decade of US deals, did acquirers beat comparable stocks after closing, and which kinds of deal did better or worse?</p>
    </div>
  </header>

  <div class="findings" id="kpis"></div>

  <section>
    <h2>Three years after closing</h2>
    <p class="lead" id="lead-path"></p>
    <div class="grid-wide">
      <div class="panel">
        <h3>Buy-and-hold abnormal return, months after completion</h3>
        <div class="seg" id="grp"></div>
        <div class="seg" id="stat"></div>
        <p class="explain" id="grp-explain"></p>
        <div id="path"></div>
      </div>
      <div class="panel">
        <h3>36-month abnormal return, deal by deal</h3>
        <p class="sub" id="dist-sub"></p>
        <div class="chart" id="dist"></div>
      </div>
    </div>
  </section>

  <section>
    <h2>Is it the benchmark?</h2>
    <p class="lead">Long-horizon returns are fragile: small errors in the benchmark compound over three years, and a few extreme deals move the average. The result is tested three independent ways.</p>
    <div class="grid2">
      <div class="panel">
        <h3>Calendar-time portfolio alphas</h3>
        <p class="sub">Every month, all acquirers within 36 months of closing; alpha is the return not explained by the factors (Newey-West t-statistics).</p>
        <div class="tablewrap"><table id="ct"></table></div>
      </div>
      <div class="panel">
        <h3>Robustness of the 36-month result</h3>
        <p class="sub">Same measure, different choices. The t-statistic is skewness-adjusted.</p>
        <div class="tablewrap"><table id="rob"></table></div>
      </div>
    </div>
  </section>

  <section>
    <h2>Which deals did worse</h2>
    <p class="lead" id="lead-groups"></p>
    <div class="grid-wide">
      <div class="panel">
        <h3>36-month BHAR by deal and acquirer characteristics</h3>
        <div class="tablewrap"><table id="groups"></table></div>
      </div>
      <div class="panel">
        <h3>What explains the dispersion</h3>
        <p class="sub">Regression of the 36-month BHAR (winsorized at 1%/99%) on deal traits, with industry dummies; standard errors clustered by completion year.</p>
        <div class="tablewrap"><table id="reg"></table></div>
      </div>
    </div>
  </section>

  <section>
    <h2>Margins and goodwill</h2>
    <p class="lead" id="lead-op"></p>
    <div class="grid2">
      <div class="panel">
        <h3>Industry-adjusted operating return on assets</h3>
        <p class="sub">Median operating income / average assets minus the median of the same two-digit SIC industry, fiscal years around the deal.</p>
        <div class="chart" id="op"></div>
      </div>
      <div class="panel">
        <h3>Goodwill impairments after the deal</h3>
        <p class="sub">Cumulative share of acquirers reporting a goodwill impairment, and one of at least 10% of the deal value.</p>
        <div class="chart" id="imp"></div>
      </div>
    </div>
  </section>

  <section>
    <h2>Deal explorer</h2>
    <p class="lead">Every acquisition in the sample. Click a row to see the acquirer's path against its benchmark. Deal value and payment come from the acquirer's XBRL cash-flow and equity statements for the fiscal year of the deal.</p>
    <div class="search"><input id="q" type="search" placeholder="Search an acquirer, e.g. Salesforce" aria-label="Search acquirer"><span class="sub" id="count"></span></div>
    <div class="grid-wide" style="align-items:start">
      <div class="scroll"><table id="deals"></table></div>
      <div class="panel"><h3 id="deal-title">Select a deal</h3><p class="sub" id="deal-sub"></p><div id="deal"></div></div>
    </div>
  </section>

  <section>
    <h2>What it means for a deal team</h2>
    <div class="verdict" id="verdict"></div>
  </section>

  <section>
    <h2>Method</h2>
    <ol class="method">
      <li><b>Deals.</b> Every US filer's acquisition spending, fiscal years 2011 to 2022, from SEC XBRL frames: cash paid for acquisitions, stock issued for acquisitions and consideration transferred. Kept when the deal value is at least $100m, the acquirer filed an 8-K Item 2.01 (completion of an acquisition) that year, and the value is at least 10% of the acquirer's market value. Banks, insurers and utilities are excluded; one deal per acquirer every three years.</li>
      <li><b>Completion date.</b> The 8-K Item 2.01 filing date. Month 1 is the first full month after it.</li>
      <li><b>Prices.</b> Month-end closes from DoltHub, including delisted stocks, adjusted for splits and dividends. Each SEC filer is matched to its ticker and every match is checked: price times shares outstanding must be consistent with the public float on the 10-K cover. Splits missing from the price data, splits recorded a month early and ticker changes are corrected when the company's own share count confirms them.</li>
      <li><b>Benchmark.</b> Each acquirer is assigned at completion to one of the 25 Fama-French portfolios formed on size and book-to-market, using NYSE breakpoints.</li>
      <li><b>BHAR.</b> <span class="formula">BHAR<sub>i,H</sub> = &prod;(1 + r<sub>i,t</sub>) &minus; &prod;(1 + r<sub>b,t</sub>), t = 1..H</span> If the acquirer stops trading, the rest of the window earns the benchmark.</li>
      <li><b>Inference.</b> Skewness-adjusted t-statistic with a bootstrap p-value (Lyon, Barber and Tsai, 1999), Wilcoxon test on the median, and calendar-time portfolio regressions on the Fama-French five factors plus momentum.</li>
      <li><b>Operating performance.</b> Operating income / average assets minus the two-digit SIC industry median from all SEC filers; post-deal (average of years +1 to +3) regressed on pre-deal (year &minus;1), following Healy, Palepu and Ruback (1992).</li>
      <li><b>Point in time.</b> Prices, returns, filings and accounting periods end on 31 March 2025; cover-page data are used only when public by then. Five-year results use deals completed by March 2020.</li>
    </ol>
  </section>

  <section>
    <h2>Limits</h2>
    <ul class="notes" id="limits"></ul>
  </section>

  <section>
    <h2>Sources</h2>
    <ul class="sources">
      <li>SEC EDGAR: XBRL frames API (acquisitions, fundamentals, cover-page shares and float), company submissions (8-K items, filing history), Financial Statement Data Sets (SIC codes).</li>
      <li>DoltHub, post-no-preference/stocks: daily prices, splits and dividends, including delisted securities.</li>
      <li>Kenneth R. French data library: Fama-French factors, momentum, 25 size and book-to-market portfolios, NYSE breakpoints.</li>
      <li>Loughran, T. and Vijh, A. (1997), Do long-term shareholders benefit from corporate acquisitions?, Journal of Finance.</li>
      <li>Barber, B. and Lyon, J. (1997), Detecting long-run abnormal stock returns: the empirical power and specification of test statistics, Journal of Financial Economics.</li>
      <li>Lyon, J., Barber, B. and Tsai, C. (1999), Improved methods for tests of long-run abnormal stock returns, Journal of Finance.</li>
      <li>Mitchell, M. and Stafford, E. (2000), Managerial decisions and long-term stock price performance, Journal of Business.</li>
      <li>Healy, P., Palepu, K. and Ruback, R. (1992), Does corporate performance improve after mergers?, Journal of Financial Economics.</li>
      <li>Moeller, S., Schlingemann, F. and Stulz, R. (2005), Wealth destruction on a massive scale?, Journal of Finance.</li>
    </ul>
  </section>

  <footer><span>Alessandro Radice &middot; Radice Capital Partners &middot; educational project, not investment advice</span><span>Data to 31 March 2025</span></footer>
</div>

<script>
const D = __DATA__;
const TXT = __TEXT__;
const C = { ink: '#1a1a18', acc: '#b5412b', mid: '#8a8983', faint: '#c8c7bf', grid: '#e6e4dc', pos: '#2f6f5e', panel: '#ffffff' };
const pct = (v, n = 1) => v == null || !isFinite(v) ? '–' : ((v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v * 100).toFixed(n) + '%');
const pctu = (v, n = 1) => v == null || !isFinite(v) ? '–' : (v < 0 ? '−' : '') + Math.abs(v * 100).toFixed(n) + '%';
const num = (v, n = 2) => v == null || !isFinite(v) ? '–' : (v < 0 ? '−' : '') + Math.abs(v).toFixed(n);
const usd = v => '$' + (v >= 1000 ? (v / 1000).toFixed(1) + 'bn' : Math.round(v) + 'm');
const cls = v => v < 0 ? 'neg' : v > 0 ? 'pos' : '';
const font = { family: 'IBM Plex Mono, monospace', size: 11, color: '#66655e' };
const base = (extra = {}) => Object.assign({ margin: { l: 56, r: 18, t: 10, b: 44 }, paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', font,
  xaxis: { gridcolor: C.grid, zeroline: false, linecolor: C.ink, ticks: '' }, yaxis: { gridcolor: C.grid, zerolinecolor: C.ink, zerolinewidth: 1, tickformat: '.0%' },
  showlegend: false, hovermode: 'x unified', hoverlabel: { font: { family: 'IBM Plex Mono', size: 12 } } }, extra);
const cfg = { displayModeBar: false, responsive: true };
document.querySelectorAll('[data-n]');
document.title = 'Does M&A Create Value?';

// ---------- KPIs ----------
const S36 = D.summ['36'];
const kpi = [
  [D.deals.length.toLocaleString('en-US'), 'acquisitions, 2012–2021, each at least 10% of the acquirer', ''],
  [pct(S36.mean), 'mean 36-month abnormal return (median ' + pct(S36.median) + ')', 'neg'],
  [pct(D.ct.stock.alpha_y), 'yearly alpha when the deal was paid partly or mostly in stock, five factors + momentum (t = ' + num(D.ct.stock.t, 1) + ')', 'neg'],
  [pctu(D.impair['5'].any, 0), 'of acquirers booked a goodwill impairment within five years', '']];
document.getElementById('kpis').innerHTML = kpi.map(k => `<div class="finding"><div class="v ${k[2]}">${k[0]}</div><div class="k">${k[1]}</div></div>`).join('');
document.getElementById('lead-path').innerHTML = TXT.lead_path;
document.getElementById('lead-groups').innerHTML = TXT.lead_groups;
document.getElementById('lead-op').innerHTML = TXT.lead_op;
document.getElementById('verdict').innerHTML = TXT.verdict;
document.getElementById('limits').innerHTML = TXT.limits.map(x => '<li>' + x + '</li>').join('');

// ---------- event-time paths ----------
const GROUPS = {
  'All deals': d => true, 'Cash': d => d.pay === 'Cash', 'Mixed': d => d.pay === 'Mixed', 'Stock-heavy': d => d.pay === 'Stock-heavy',
  'Glamour acquirers': d => d.bmq === 1, 'Value acquirers': d => d.bmq >= 4, 'Serial acquirers': d => d.ser === 1,
  'Large deals (40%+)': d => d.rs >= 0.4 };
const EXPL = {
  'All deals': 'Every acquisition in the sample. The shaded band is the 95% confidence interval of the mean.',
  'Cash': 'Stock under 5% of the price paid, as disclosed in the cash-flow and equity statements.',
  'Mixed': 'Stock between 5% and 50% of the price paid.',
  'Stock-heavy': 'Stock at least half of the price paid. Theory: managers pay with stock when they think it is overvalued.',
  'Glamour acquirers': 'Lowest book-to-market quintile at completion: the most richly valued acquirers.',
  'Value acquirers': 'Top two book-to-market quintiles at completion.',
  'Serial acquirers': 'Two or more acquisitions of $100m+ in the five previous years.',
  'Large deals (40%+)': 'Deal value at least 40% of the acquirer\'s market value before completion.' };
let G = 'All deals', ST = 'mean';
function agg(list) {
  const H = 37, m = [], md = [], lo = [], hi = [];
  for (let k = 0; k < H; k++) {
    const v = list.map(d => d.p[k]).filter(x => x != null).sort((a, b) => a - b);
    if (!v.length) { m.push(null); md.push(null); lo.push(null); hi.push(null); continue; }
    const mu = v.reduce((a, b) => a + b, 0) / v.length; const sd = Math.sqrt(v.reduce((a, b) => a + (b - mu) ** 2, 0) / Math.max(1, v.length - 1));
    m.push(mu); md.push(v.length % 2 ? v[(v.length - 1) / 2] : (v[v.length / 2 - 1] + v[v.length / 2]) / 2);
    lo.push(mu - 1.96 * sd / Math.sqrt(v.length)); hi.push(mu + 1.96 * sd / Math.sqrt(v.length));
  }
  return { m, md, lo, hi };
}
const months = [...Array(37).keys()];
function drawPath() {
  const ok = D.deals.filter(d => d.b36 != null), all = agg(ok), sel = ok.filter(GROUPS[G]), a = agg(sel);
  const key = ST === 'mean' ? 'm' : 'md';
  const tr = [];
  if (ST === 'mean') tr.push({ x: months.concat(months.slice().reverse()), y: a.hi.concat(a.lo.slice().reverse()), fill: 'toself', fillcolor: 'rgba(181,65,43,.10)', line: { width: 0 }, hoverinfo: 'skip', showlegend: false, type: 'scatter' });
  if (G !== 'All deals') tr.push({ x: months, y: all[key], name: 'All deals', line: { color: C.faint, width: 2 }, type: 'scatter', hovertemplate: '%{y:.1%}' });
  tr.push({ x: months, y: a[key], name: G + ' (' + sel.length + ')', line: { color: G === 'All deals' ? C.ink : C.acc, width: 3 }, type: 'scatter', hovertemplate: '%{y:.1%}' });
  const last = a[key][36];
  Plotly.react('path', tr, base({ margin: { l: 56, r: 64, t: 10, b: 44 }, xaxis: { title: 'Months after completion', dtick: 6, gridcolor: C.grid, zeroline: false, range: [0, 37] }, showlegend: true,
    legend: { orientation: 'h', x: 0, y: 1.08, font: { size: 11 } }, annotations: [{ x: 36, y: last, text: pct(last), showarrow: false, xanchor: 'left', xshift: 6, font: { family: 'IBM Plex Mono', size: 12, color: C.ink } }] }), cfg);
  document.getElementById('grp-explain').textContent = EXPL[G];
  drawDist(sel);
}
function drawDist(sel) {
  const v = sel.map(d => d.b36).filter(x => x != null);
  const cl = v.map(x => Math.max(-1.5, Math.min(1.5, x)));
  const mu = v.reduce((a, b) => a + b, 0) / v.length; const s = v.slice().sort((a, b) => a - b); const md = s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
  Plotly.react('dist', [{ x: cl, type: 'histogram', xbins: { start: -1.5, end: 1.5001, size: 0.1 }, marker: { color: C.faint, line: { color: '#fff', width: 1 } }, hovertemplate: '%{y} deals<extra></extra>' }],
    base({ bargap: 0.02, xaxis: { tickformat: '.0%', gridcolor: C.grid, zeroline: false, title: 'clipped at ±150%' }, yaxis: { gridcolor: C.grid, title: 'deals' }, hovermode: 'closest',
      shapes: [{ type: 'line', x0: mu, x1: mu, yref: 'paper', y0: 0, y1: 1, line: { color: C.ink, width: 2 } }, { type: 'line', x0: md, x1: md, yref: 'paper', y0: 0, y1: 1, line: { color: C.acc, width: 2 } }] }), cfg);
  document.getElementById('dist-sub').innerHTML = `${v.length} deals · mean <b style="color:${C.ink}">${pct(mu)}</b> · median <b style="color:${C.acc}">${pct(md)}</b> · ${pctu(v.filter(x => x > 0).length / v.length, 0)} beat the benchmark`;
}
document.getElementById('grp').innerHTML = Object.keys(GROUPS).map(g => `<button data-g="${g}" class="${g === G ? 'on' : ''}">${g}</button>`).join('');
document.getElementById('stat').innerHTML = ['mean', 'median'].map(s => `<button data-s="${s}" class="${s === ST ? 'on' : ''}">${s[0].toUpperCase() + s.slice(1)}</button>`).join('');
document.getElementById('grp').onclick = e => { const b = e.target.closest('button'); if (!b) return; G = b.dataset.g; document.querySelectorAll('#grp button').forEach(x => x.classList.toggle('on', x === b)); drawPath(); };
document.getElementById('stat').onclick = e => { const b = e.target.closest('button'); if (!b) return; ST = b.dataset.s; document.querySelectorAll('#stat button').forEach(x => x.classList.toggle('on', x === b)); drawPath(); };
drawPath();

// ---------- calendar time & robustness ----------
const ctRow = (lab, a) => `<tr><td>${lab}</td><td class="${cls(a.alpha_m)}">${pct(a.alpha_m, 2)}</td><td class="${cls(a.alpha_y)}">${pct(a.alpha_y, 1)}</td><td>${num(a.t, 2)}</td></tr>`;
document.getElementById('ct').innerHTML = '<tr><th>Portfolio and model</th><th>Alpha / month</th><th>Per year</th><th>t-stat</th></tr>' +
  '<tr class="grp"><td colspan="4">Equal-weighted</td></tr>' + ctRow('CAPM', D.ct.ew.capm) + ctRow('Fama-French 3', D.ct.ew.ff3) + ctRow('FF 5 + momentum', D.ct.ew.ff5m) +
  '<tr class="grp"><td colspan="4">Value-weighted (by market value at completion)</td></tr>' + ctRow('CAPM', D.ct.vw.capm) + ctRow('Fama-French 3', D.ct.vw.ff3) + ctRow('FF 5 + momentum', D.ct.vw.ff5m) +
  '<tr class="grp"><td colspan="4">By payment, FF 5 + momentum, equal-weighted</td></tr>' + ctRow('Cash', D.ct.cash) + ctRow('Mixed or stock-heavy', D.ct.stock) +
  `<tr><td>Stock minus cash</td><td class="${cls(D.ct.stock_minus_cash.alpha_m)}">${pct(D.ct.stock_minus_cash.alpha_m, 2)}</td><td class="${cls(D.ct.stock_minus_cash.alpha_m)}">${pct((1 + D.ct.stock_minus_cash.alpha_m) ** 12 - 1, 1)}</td><td>${num(D.ct.stock_minus_cash.t, 2)}</td></tr>`;
const rr = Object.entries(D.robust).map(([k, v]) => `<tr><td class="l">${k}</td><td>${v.n}</td><td class="${cls(v.mean)}">${pct(v.mean)}</td><td class="${cls(v.median)}">${pct(v.median)}</td><td>${num(v.t_skew, 2)}</td></tr>`).join('');
const pl = D.placebo;
document.getElementById('rob').innerHTML = '<tr><th class="l">Variant</th><th>Deals</th><th>Mean</th><th>Median</th><th>t (skew-adj.)</th></tr>' + rr +
  `<tr class="grp"><td colspan="5">Placebo: the same acquirers, three-year windows with no deal</td></tr><tr><td class="l">Quiet window after the deal (average of ${pl.draws} random draws)</td><td>${pl.n}</td><td class="${cls(pl.mean)}">${pct(pl.mean)}</td><td class="${cls(pl.median)}">${pct(pl.median)}</td><td>${num(pl.t_skew, 2)}</td></tr>`;

// ---------- groups & regression ----------
const gh = '<tr><th class="l">Group</th><th>Deals</th><th>Mean</th><th>Median</th><th>Beat benchmark</th><th>t (skew-adj.)</th></tr>';
document.getElementById('groups').innerHTML = gh + Object.entries(D.groups).map(([g, v]) => `<tr class="grp"><td colspan="6">${g}</td></tr>` +
  Object.entries(v).map(([k, s]) => `<tr><td class="l">${k}</td><td>${s.n}</td><td class="${cls(s.mean)}">${pct(s.mean)}</td><td class="${cls(s.median)}">${pct(s.median)}</td><td>${pctu(s.pos, 0)}</td><td>${num(s.t_skew, 2)}</td></tr>`).join('')).join('');
const RL = { stock_share: 'Stock share of the price', log_rs: 'Log deal size / acquirer value', log_me: 'Log acquirer market value', gw_int: 'Goodwill / deal value',
  serial: 'Serial acquirer (0/1)', runup: 'Stock run-up, 12 months before', glamour: 'Glamour acquirer (0/1)', late: 'Completed 2017–2021 (0/1)' };
document.getElementById('reg').innerHTML = '<tr><th class="l">Variable</th><th>Coefficient</th><th>t-stat</th></tr>' +
  Object.keys(RL).map(k => `<tr><td class="l">${RL[k]}</td><td>${num(D.reg.params[k], 3)}</td><td>${num(D.reg.t[k], 2)}</td></tr>`).join('') +
  `<tr><td class="l dim">Deals / R²</td><td class="dim">${D.reg.n}</td><td class="dim">${num(D.reg.r2, 3)}</td></tr>`;

// ---------- operating & impairment ----------
const yrs = Object.keys(D.op_path).map(Number).sort((a, b) => a - b);
Plotly.newPlot('op', [{ x: yrs, y: yrs.map(y => D.op_path[y].roa_adj_med), mode: 'lines+markers', line: { color: C.ink, width: 3 }, marker: { size: 8 }, hovertemplate: 'year %{x}: %{y:.2%}<extra></extra>' }],
  base({ xaxis: { tickvals: yrs, ticktext: yrs.map(y => y > 0 ? '+' + y : String(y)), title: 'fiscal year relative to the deal', gridcolor: C.grid, zeroline: false },
    yaxis: { tickformat: '.1%', gridcolor: C.grid, zerolinecolor: C.ink, rangemode: 'tozero' }, hovermode: 'closest',
    shapes: [{ type: 'rect', x0: -0.3, x1: 0.3, yref: 'paper', y0: 0, y1: 1, fillcolor: 'rgba(181,65,43,.08)', line: { width: 0 } }] }), cfg);
const ks = Object.keys(D.impair).map(Number);
Plotly.newPlot('imp', [
  { x: ks, y: ks.map(k => D.impair[k].any), name: 'Any impairment', mode: 'lines+markers', line: { color: C.ink, width: 3 }, marker: { size: 8 }, hovertemplate: '%{y:.0%}' },
  { x: ks, y: ks.map(k => D.impair[k].big), name: '10%+ of deal value', mode: 'lines+markers', line: { color: C.acc, width: 3 }, marker: { size: 8 }, hovertemplate: '%{y:.0%}' }],
  base({ xaxis: { dtick: 1, title: 'years after the deal', gridcolor: C.grid, zeroline: false }, yaxis: { tickformat: '.0%', gridcolor: C.grid, range: [0, 0.6] }, showlegend: true, legend: { x: 0, y: 1.1, orientation: 'h' } }), cfg);

// ---------- deal explorer ----------
let SORT = ['dv', -1], SEL = null;
function drawTable() {
  const q = document.getElementById('q').value.trim().toLowerCase();
  let rows = D.deals.map((d, i) => Object.assign({ i }, d)).filter(d => !q || d.n.toLowerCase().includes(q));
  rows.sort((a, b) => { const x = a[SORT[0]], y = b[SORT[0]]; if (x == null) return 1; if (y == null) return -1; return (x > y ? 1 : x < y ? -1 : 0) * SORT[1]; });
  document.getElementById('count').textContent = rows.length + ' deals';
  const hd = [['n', 'Acquirer'], ['t', 'Completed'], ['dv', 'Deal value'], ['rs', 'Deal / acquirer'], ['pay', 'Payment'], ['b12', 'BHAR 12m'], ['b36', 'BHAR 36m']];
  document.getElementById('deals').innerHTML = '<tr>' + hd.map(h => `<th data-k="${h[0]}" class="${h[0] === 'n' ? 'l' : ''}">${h[1]}${SORT[0] === h[0] ? (SORT[1] > 0 ? ' ↑' : ' ↓') : ''}</th>`).join('') + '</tr>' +
    rows.slice(0, 400).map(d => `<tr class="pick ${d.i === SEL ? 'sel' : ''}" data-i="${d.i}"><td class="l">${d.n}</td><td>${d.t}</td><td>${usd(d.dv)}</td><td>${pctu(d.rs, 0)}</td><td>${d.pay}</td><td class="${cls(d.b12)}">${pct(d.b12)}</td><td class="${cls(d.b36)}">${pct(d.b36)}</td></tr>`).join('');
}
function drawDeal(i) {
  SEL = i; const d = D.deals[i]; const ref = agg(D.deals.filter(x => x.b36 != null)).m;
  document.getElementById('deal-title').textContent = d.n;
  document.getElementById('deal-sub').innerHTML = `Completed ${d.t} · ${usd(d.dv)} · ${pctu(d.rs, 0)} of its market value · ${d.pay}${d.dl ? ' · stopped trading within 36 months' : ''}`;
  Plotly.react('deal', [{ x: months, y: ref, line: { color: C.faint, width: 2 }, name: 'All deals, mean', hovertemplate: '%{y:.1%}' },
    { x: months, y: d.p, line: { color: C.acc, width: 3 }, name: d.n, hovertemplate: '%{y:.1%}' }],
    base({ xaxis: { dtick: 12, title: 'months after completion', gridcolor: C.grid, zeroline: false }, showlegend: true, legend: { x: 0, y: 1.15, orientation: 'h' } }), cfg);
  drawTable();
}
document.getElementById('q').oninput = drawTable;
document.getElementById('deals').onclick = e => {
  const th = e.target.closest('th'); if (th) { const k = th.dataset.k; SORT = [k, SORT[0] === k ? -SORT[1] : -1]; drawTable(); return; }
  const tr = e.target.closest('tr.pick'); if (tr) drawDeal(+tr.dataset.i);
};
drawTable();
drawDeal(D.deals.reduce((b, d, i) => d.dv > D.deals[b].dv ? i : b, 0));
</script>
'''

# %%
PD = page_data(X, res)
html = (PAGE_TEMPLATE.replace("__N__", f"{PD['funnel']['final']:,}").replace("__DATA__", json.dumps(PD, separators=(",", ":")))
        .replace("__TEXT__", json.dumps(page_text(PD))))
open("MA_Long_Run_Performance.html", "w", encoding="utf-8").write(html)
print(f"MA_Long_Run_Performance.html written ({len(html) / 1e3:.0f} KB). Open it in any browser.")
try:
    from google.colab import files
    files.download("MA_Long_Run_Performance.html")
except ImportError:
    pass

# %% [markdown]
# ## 10. Export: the Excel model
# Writes `MA_Long_Run_Performance.xlsx`. Monthly returns of every deal, its benchmark and the factors are stored as values;
# every BHAR, summary statistic, alpha (LINEST) and operating regression is a live formula. Change the horizon, benchmark,
# relative-size threshold, payment filter or years on `Inputs` and every sheet recomputes. `Checks` compares Excel with Python.

# %%
LAST=pd.Period('2025-03','M')
def ret_matrices(X,FR,VW,EW,H=60):
    """rows = deals, cols = months 1..H after completion. Firm returns after delisting take the benchmark (as in the BHAR);
    months after March 2025 are blank."""
    n=len(X); F=np.full((n,H),np.nan); B=np.full((n,H),np.nan); BE=np.full((n,H),np.nan); A=np.zeros((n,H),bool)
    for i,r in enumerate(X.itertuples()):
        idx=pd.period_range(r.t+1,r.t+H,freq='M'); ok=np.array([m<=LAST for m in idx])
        b=VW[r.port].reindex(idx).values; be=EW[r.port].reindex(idx).values; f=FR[r.cik].ret.reindex(idx).values
        alive=np.isfinite(f); A[i]=alive&ok
        f=np.where(alive,f,b)
        F[i,ok]=f[ok]; B[i,ok]=b[ok]; BE[i,ok]=be[ok]
    return F,B,BE,A

_,VW,EW=bench_series(X)
X=X.reset_index(drop=True)
Fm,Bm,BEm,Am=ret_matrices(X,FR,VW,EW,60)
OUT='MA_Long_Run_Performance.xlsx'
Fn=lambda **k: Font(name='Arial',size=k.pop('size',10),**k)
BLUE,GREEN,GREY='0000FF','008000','66655E'
HEAD=PatternFill('solid',fgColor='1A1A18'); YEL=PatternFill('solid',fgColor='FFF2CC'); TOT=PatternFill('solid',fgColor='ECEBE5')
PCT='0.0%;\\-0.0%'; PCT2='0.00%;\\-0.00%'; MN='#,##0;\\-#,##0'; NUM='0.00;\\-0.00'; INT='#,##0'
wb=Workbook(); wb.remove(wb.active)
def sheet(title,heading,sub,widths):
    ws=wb.create_sheet(title); ws.sheet_view.showGridLines=False; ws.column_dimensions['A'].width=2.7
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    ws['B2']=heading; ws['B2'].font=Fn(size=14,bold=True); ws['B3']=sub; ws['B3'].font=Fn(color=GREY); return ws
def head(ws,row,labels,col=2):
    for i,t in enumerate(labels):
        c=ws.cell(row,col+i,t); c.font=Fn(bold=True,color='FFFFFF'); c.fill=HEAD; c.alignment=Alignment(horizontal='left' if i==0 else 'right',vertical='center',wrap_text=True)
def put(ws,ref,v,fmt=None,color=None,bold=False,fill=None,al=None):
    c=ws[ref]; c.value=v; c.font=Fn(color=color,bold=bold)
    if fmt: c.number_format=fmt
    if fill: c.fill=fill
    if al: c.alignment=Alignment(horizontal=al)
    return c
def note(ws,ref,t): c=put(ws,ref,t); c.font=Fn(color=GREY); return c
fin=lambda v: float(v) if v is not None and np.isfinite(v) else 'n/a'
N=len(X); r0=7; rN=r0+N-1
# ---------------- Cover ----------------
cv=sheet('Cover','Does M&A Create Value? Long-Run Performance of US Acquirers','Alessandro Radice · data to 31 March 2025 · SEC EDGAR, DoltHub, Kenneth French data library',{'B':34,'C':16,'D':70})
rows=[('Inputs','Horizon, benchmark and filters that drive every live result'),('Deals',f'{N} acquisitions completed 2012-2021, one row each, with the BHAR as a live formula'),
      ('Returns','Monthly returns of each acquirer and of its size/book-to-market benchmark, months 1 to 60 after completion'),
      ('Summary','Mean, median, share positive and t-statistic, overall and by payment method, from the Deals sheet'),
      ('Calendar time','Monthly portfolio of recent acquirers and its alpha against the Fama-French factors (LINEST)'),
      ('Operating','Industry-adjusted operating return on assets before and after, the regression of post on pre, goodwill impairments'),
      ('Checks','Every headline number recomputed in Excel and compared with the Python engine')]
head(cv,5,['Sheet','','What it holds'])
for i,(a,b) in enumerate(rows): put(cv,f'B{6+i}',a,bold=True); put(cv,f'D{6+i}',b)
note(cv,'B14','Colour code: blue = input, black = formula, green = link to another sheet. Educational project, not investment advice.')
put(cv,'B16','Headline results (live)',bold=True)
# ---------------- Inputs ----------------
ip=sheet('Inputs','Inputs','Change the blue cells; every sheet recalculates',{'B':46,'C':16,'D':60})
head(ip,5,['Input','Value','Notes'])
put(ip,'B6','Horizon (months after completion)'); put(ip,'C6',36,color=BLUE,fill=YEL); note(ip,'D6','12, 24, 36 or 60. 60 months only for deals completed by March 2020')
put(ip,'B7','Benchmark (1 = value-weighted, 2 = equal-weighted)'); put(ip,'C7',1,color=BLUE,fill=YEL); note(ip,'D7','Fama-French 25 portfolios formed on size and book-to-market')
put(ip,'B8','Minimum deal size (% of acquirer market value)'); put(ip,'C8',0.10,PCT,color=BLUE,fill=YEL); note(ip,'D8','The sample is built at 10%; raise it to study larger deals')
put(ip,'B9','Payment method filter'); put(ip,'C9','All',color=BLUE,fill=YEL); note(ip,'D9','All, Cash, Mixed or Stock-heavy')
put(ip,'B10','Completion year from'); put(ip,'C10',2012,color=BLUE,fill=YEL)
put(ip,'B11','Completion year to'); put(ip,'C11',2021,color=BLUE,fill=YEL)
dv=DataValidation(type='list',formula1='"12,24,36,60"'); ip.add_data_validation(dv); dv.add('C6')
dv2=DataValidation(type='list',formula1='"1,2"'); ip.add_data_validation(dv2); dv2.add('C7')
dv3=DataValidation(type='list',formula1='"All,Cash,Mixed,Stock-heavy"'); ip.add_data_validation(dv3); dv3.add('C9')
put(ip,'B13','Definitions',bold=True)
defs=['BHAR = (1+r1)(1+r2)...(1+rH) - (1+b1)(1+b2)...(1+bH): the buy-and-hold return of the acquirer minus that of its benchmark portfolio.',
      'Month 1 is the first full month after the month of completion (the 8-K Item 2.01 filing).',
      'If the acquirer stops trading (acquired, delisted), the remaining months earn the benchmark: the BHAR is frozen, not dropped.',
      'Payment: Cash = stock under 5% of the price paid; Stock-heavy = 50% or more; Mixed in between (XBRL cash-flow and equity statements).',
      'Deal size = deal value / acquirer market value at the end of the month before completion.']
for i,d in enumerate(defs): note(ip,f'B{14+i}',d)
# ---------------- Returns ----------------
rt=wb.create_sheet('Returns'); rt.sheet_view.showGridLines=False
rt['B2']='Returns'; rt['B2'].font=Fn(size=14,bold=True); rt['B3']='Monthly total returns; acquirer after delisting = benchmark; blank after March 2025'; rt['B3'].font=Fn(color=GREY)
C_F,C_B,C_E=4,4+60+1,4+120+2
for blk,c0,lab in [(Fm,C_F,'Acquirer'),(Bm,C_B,'Benchmark, value-weighted'),(BEm,C_E,'Benchmark, equal-weighted')]:
    rt.cell(5,c0,lab).font=Fn(bold=True)
    for k in range(60):
        c=rt.cell(6,c0+k,k+1); c.font=Fn(bold=True,color='FFFFFF'); c.fill=HEAD
    for i in range(N):
        for k in range(60):
            v=blk[i,k]
            if np.isfinite(v): rt.cell(r0+i,c0+k,round(float(v),6)).number_format=PCT2
rt.cell(6,2,'Deal').font=Fn(bold=True,color='FFFFFF'); rt.cell(6,2).fill=HEAD; rt.cell(6,3,'Traded in month 1').font=Fn(bold=True,color='FFFFFF'); rt.cell(6,3).fill=HEAD
for i in range(N):
    rt.cell(r0+i,2,f'=Deals!B{r0+i}').font=Fn(color=GREEN); rt.cell(r0+i,3,int(Am[i,0])).font=Fn(color=BLUE)
rt.column_dimensions['B'].width=30; rt.freeze_panes='D7'
# ---------------- Deals ----------------
dl=sheet('Deals','Deals','One row per acquisition. Grey columns are data; the BHAR columns are live formulas on the Returns sheet',
         {'B':34,'C':11,'D':12,'E':11,'F':12,'G':10,'H':9,'I':9,'J':9,'K':11,'L':12,'M':12,'N':12,'O':12,'P':14})
head(dl,6,['Acquirer','Completed','Deal value ($m)','Deal / acquirer value','Payment','Stock share','BM quintile','Serial (2+ prior)','New shares 10%+','Run-up (12m)','BHAR (selected)','Included','BHAR in filter','Pre ROA (adj.)','Post ROA (adj., avg. 1-3)'])
for i,r in enumerate(X.itertuples()):
    row=r0+i
    put(dl,f'B{row}',r.name.title() if r.name.isupper() else r.name,color=BLUE); put(dl,f'C{row}',str(r.t),color=BLUE)
    put(dl,f'D{row}',round(r.dv/1e6,1),MN,color=BLUE); put(dl,f'E{row}',round(float(r.rs),4),PCT,color=BLUE); put(dl,f'F{row}',r.pay,color=BLUE)
    put(dl,f'G{row}',fin(r.stock_share),PCT,color=BLUE); put(dl,f'H{row}',int(r.bmq) if np.isfinite(r.bmq) else 'n/a',color=BLUE)
    put(dl,f'I{row}',int(r.serial),color=BLUE); put(dl,f'J{row}',int(r.eq_issuer),color=BLUE); put(dl,f'K{row}',fin(r.runup),PCT,color=BLUE)
    fr=f"Returns!{L(C_F)}{row}"; br=f"Returns!{L(C_B)}{row}"; er=f"Returns!{L(C_E)}{row}"
    rng=lambda a: f"OFFSET({a},0,0,1,Inputs!$C$6)"
    bench=f"IF(Inputs!$C$7=1,EXP(SUMPRODUCT(LN(1+{rng(br)}))),EXP(SUMPRODUCT(LN(1+{rng(er)}))))"
    put(dl,f'L{row}',f'=IF(OR(Returns!$C{row}=0,COUNT({rng(br)})<Inputs!$C$6),"n/a",EXP(SUMPRODUCT(LN(1+{rng(fr)})))-{bench})',PCT)
    put(dl,f'M{row}',f'=IF(AND(ISNUMBER(L{row}),E{row}>=Inputs!$C$8,OR(Inputs!$C$9="All",F{row}=Inputs!$C$9),VALUE(LEFT(C{row},4))>=Inputs!$C$10,VALUE(LEFT(C{row},4))<=Inputs!$C$11),1,0)',INT)
    put(dl,f'N{row}',f'=IF(M{row}=1,L{row},"")',PCT)
    put(dl,f'O{row}',fin(r.pre_adj),PCT,color=BLUE); put(dl,f'P{row}',fin(r.post_adj) if r.post_n>=2 else 'n/a',PCT,color=BLUE)
dl.freeze_panes='C7'
# ---------------- Summary ----------------
sm=sheet('Summary','Summary','Long-run abnormal returns of the acquirers selected on the Inputs sheet',{'B':34,'C':14,'D':14,'E':14,'F':14,'G':14})
put(sm,'B5','Horizon (months)'); put(sm,'C5','=Inputs!C6',color=GREEN); put(sm,'B6','Benchmark'); put(sm,'C6','=IF(Inputs!C7=1,"Value-weighted","Equal-weighted")',color=GREEN)
head(sm,8,['Group','Deals','Mean BHAR','Median BHAR','Share positive','t-statistic'])
N_=f'Deals!$N${r0}:$N${rN}'; F_=f'Deals!$F${r0}:$F${rN}'; M_=f'Deals!$M${r0}:$M${rN}'
groups=[('All selected deals',None),('Cash',"Cash"),('Mixed',"Mixed"),('Stock-heavy',"Stock-heavy")]
for j,(lab,g) in enumerate(groups):
    row=9+j; put(sm,f'B{row}',lab,bold=(g is None))
    if g is None:
        put(sm,f'C{row}',f'=SUM({M_})',INT); put(sm,f'D{row}',f'=AVERAGE({N_})',PCT); put(sm,f'E{row}',f'=MEDIAN({N_})',PCT)
        put(sm,f'F{row}',f'=COUNTIFS({M_},1,{N_},">0")/C{row}',PCT); put(sm,f'G{row}',f'=D{row}/(STDEV({N_})/SQRT(C{row}))',NUM)
    else:
        col=f'Deals!$Q${r0}:$Q${rN}'.replace('Q',L(17+j-1))
        put(sm,f'C{row}',f'=COUNTIFS({M_},1,{F_},"{g}")',INT); put(sm,f'D{row}',f'=AVERAGEIFS({N_},{M_},1,{F_},"{g}")',PCT)
        put(sm,f'E{row}',f'=MEDIAN({col})',PCT); put(sm,f'F{row}',f'=COUNTIFS({M_},1,{F_},"{g}",{N_},">0")/C{row}',PCT)
        put(sm,f'G{row}',f'=D{row}/(STDEV({col})/SQRT(C{row}))',NUM)
        # helper column on Deals for the group median / stdev
        dl.cell(6,17+j-1,f'In {g}').font=Fn(bold=True,color='FFFFFF'); dl.cell(6,17+j-1).fill=HEAD
        for i in range(N): dl.cell(r0+i,17+j-1,f'=IF(AND(M{r0+i}=1,F{r0+i}="{g}"),L{r0+i},"")').number_format=PCT
note(sm,'B14','With Inputs at 36 months, value-weighted, 10%, All and 2012-2021, the first row reproduces the Python results (Checks sheet).')
# ---------------- Calendar time ----------------
ctp=res['ct']['ew']['series']; f=factors()
ct=sheet('Calendar time','Calendar-time portfolio','Each month, the equal-weighted return of every acquirer within 36 months of completion (at least 10 firms)',
         {'B':10,'C':12,'D':8,'E':11,'F':11,'G':11,'H':11,'I':11,'J':11,'K':11,'L':12})
head(ct,6,['Month','Portfolio','Firms','Mkt-RF','SMB','HML','RMW','CMA','MOM','RF','Excess'])
d=ctp.join(f,how='inner'); c0=7
for i,(m,r) in enumerate(d.iterrows()):
    row=c0+i; put(ct,f'B{row}',str(m)); put(ct,f'C{row}',round(r.rp,6),PCT2,color=BLUE); put(ct,f'D{row}',int(r.n),INT,color=BLUE)
    for k,cc in zip(['Mkt-RF','SMB','HML','RMW','CMA','MOM','RF'],'EFGHIJK'): put(ct,f'{cc}{row}',round(float(r[k]),6),PCT2,color=BLUE)
    put(ct,f'L{row}',f'=C{row}-K{row}',PCT2)
cN=c0+len(d)-1
put(ct,'N6','Alpha (LINEST)',bold=True); head(ct,7,['Model','Alpha / month','Annualized'],col=14)
Y=f'L{c0}:L{cN}'
for j,(lab,xr,k) in enumerate([('CAPM',f'E{c0}:E{cN}',1),('Fama-French 3',f'E{c0}:G{cN}',3),('FF 5 + momentum',f'E{c0}:J{cN}',6)]):
    row=8+j; put(ct,f'N{row}',lab); put(ct,f'O{row}',f'=INDEX(LINEST({Y},{xr}),1,{k+1})',PCT2); put(ct,f'P{row}',f'=(1+O{row})^12-1',PCT)
ct.column_dimensions['N'].width=18; ct.column_dimensions['O'].width=14; ct.column_dimensions['P'].width=12
note(ct,'N12','Point estimates only; the t-statistics in the report use Newey-West standard errors (3 lags).')
# ---------------- Operating ----------------
op=sheet('Operating','Operating performance and goodwill','Operating income / average assets minus the median of the same two-digit SIC industry and year',{'B':44,'C':14,'D':14,'E':40})
head(op,5,['Measure','Value','','Notes'])
O_=f'Deals!$O${r0}:$O${rN}'; P_=f'Deals!$P${r0}:$P${rN}'
put(op,'B6','Deals with pre and post data'); put(op,'C6',f'=SUMPRODUCT(ISNUMBER({O_})*ISNUMBER({P_}))',INT)
put(op,'B7','Median pre (year -1)'); put(op,'C7',f'=MEDIAN({O_})',PCT)
put(op,'B8','Median post (average of years +1 to +3)'); put(op,'C8',f'=MEDIAN({P_})',PCT)
note(op,'E7','Medians over all deals with data in that year')
put(op,'B10','Regression: post = a + b × pre',bold=True)
put(op,'B11','a (abnormal change)'); put(op,'B12','b (persistence)')
# LINEST needs complete pairs: helper columns
op.cell(5,7,'pre').font=Fn(bold=True); op.cell(5,8,'post').font=Fn(bold=True)
pairs=X[(X.pre_adj.notna())&(X.post_adj.notna())&(X.post_n>=2)]
lo_p,hi_p=pairs.pre_adj.quantile([0.01,0.99]); lo_q,hi_q=pairs.post_adj.quantile([0.01,0.99])
for i,r in enumerate(pairs.itertuples()):
    put(op,f'G{6+i}',round(float(np.clip(r.pre_adj,lo_p,hi_p)),6),PCT2,color=BLUE); put(op,f'H{6+i}',round(float(np.clip(r.post_adj,lo_q,hi_q)),6),PCT2,color=BLUE)
pe=6+len(pairs)-1
put(op,'C11',f'=INDEX(LINEST(H6:H{pe},G6:G{pe}),1,2)',PCT2); put(op,'C12',f'=INDEX(LINEST(H6:H{pe},G6:G{pe}),1,1)',NUM)
note(op,'E11','Winsorized at 1% and 99% (columns G and H)')
put(op,'B14','Goodwill impairment within 5 years',bold=True)
imp=res['impair']
for j,k in enumerate(sorted(imp)):
    put(op,f'B{15+j}',f'By year {k}: any impairment / 10%+ of deal value'); put(op,f'C{15+j}',round(imp[k]['any'],6),PCT,color=BLUE); put(op,f'D{15+j}',round(imp[k]['big'],6),PCT,color=BLUE)
# ---------------- Checks ----------------
ck=sheet('Checks','Checks','Excel against the Python engine (inputs at their defaults)',{'B':46,'C':14,'D':14,'E':12,'F':10})
head(ck,5,['Check','Excel','Python','Difference','OK'])
S=res['summ'][36]; h=res['healy']
checks=[('Deals in the 36-month sample','=Summary!C9',S['n']),('Mean 36-month BHAR','=Summary!D9',S['mean']),('Median 36-month BHAR','=Summary!E9',S['median']),
        ('Share positive','=Summary!F9',S['pos']),('Mean BHAR, cash deals','=Summary!D10',res['by_pay']['Cash']['mean']),('Mean BHAR, stock-heavy deals','=Summary!D12',res['by_pay']['Stock-heavy']['mean']),
        ('FF5 + momentum alpha / month','=\'Calendar time\'!O10',res['ct']['ew']['ff5m']['alpha_m']),('Operating: regression intercept','=Operating!C11',h['a']),('Operating: regression slope','=Operating!C12',h['b'])]
for j,(lab,fx,pv) in enumerate(checks):
    row=6+j; put(ck,f'B{row}',lab); put(ck,f'C{row}',fx,'0.0000',color=GREEN); put(ck,f'D{row}',round(float(pv),6),'0.0000',color=BLUE)
    put(ck,f'E{row}',f'=C{row}-D{row}','0.000000'); put(ck,f'F{row}',f'=IF(ABS(E{row})<0.00005,"OK","CHECK")')
put(ck,f'B{7+len(checks)}','All checks',bold=True); put(ck,f'F{7+len(checks)}',f'=IF(COUNTIF(F6:F{5+len(checks)},"OK")={len(checks)},"OK","CHECK")',bold=True)
# cover live headline
for j,(lab,fx,fmt) in enumerate([('Mean 36-month BHAR','=Summary!D9',PCT),('Median 36-month BHAR','=Summary!E9',PCT),('FF5 + momentum alpha, annualized',"='Calendar time'!P10",PCT),('Checks',f"=Checks!F{7+len(checks)}",None)]):
    put(cv,f'B{17+j}',lab); put(cv,f'C{17+j}',fx,fmt,color=GREEN)
wb.properties.creator='Alessandro Radice'; wb.properties.lastModifiedBy='Alessandro Radice'; wb.properties.title='Does M&A Create Value? Long-Run Performance of US Acquirers'
wb.save(OUT)
print(f'{OUT} written: open it in Excel, which calculates every formula on opening.')

try:
    from google.colab import files
    files.download(OUT)
except ImportError:
    pass

