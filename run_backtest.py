#!/usr/bin/env python3
"""
备兑认购策略回测
50ETF月度卖出近月平值/虚值认购期权
"""
import yfinance as yf
import akshare as ak
import pandas as pd
import numpy as np
from scipy.stats import norm
from pathlib import Path
import warnings, time
warnings.filterwarnings('ignore')

BASE = Path("/Users/castle/Desktop/space for claude")
CACHE_DIR = BASE / ".backtest_cache"
CACHE_DIR.mkdir(exist_ok=True)

# ============================================================
# 0. Helpers
# ============================================================
def bs_price(S, K, T, r, sigma, option_type='call'):
    if T <= 0: return max(0, S - K) if option_type == 'call' else max(0, K - S)
    d1 = (np.log(S/K) + (r + sigma**2/2)*T) / (sigma*np.sqrt(T))
    d2 = d1 - sigma*np.sqrt(T)
    if option_type == 'call':
        return S*norm.cdf(d1) - K*np.exp(-r*T)*norm.cdf(d2)
    else:
        return K*np.exp(-r*T)*norm.cdf(-d2) - S*norm.cdf(-d1)

def calc_metrics(rets, rfr=0.025):
    rets = rets.dropna()
    if len(rets) < 6: return {}
    ann_ret = rets.mean() * 12
    ann_vol = rets.std() * np.sqrt(12)
    sharpe = (ann_ret - rfr) / ann_vol if ann_vol > 0 else 0
    cum = (1 + rets).cumprod()
    peak = cum.expanding().max()
    dd = (cum - peak) / peak
    max_dd = dd.min()
    calmar = ann_ret / abs(max_dd) if max_dd != 0 else 0
    down = rets[rets < 0]
    down_std = down.std() * np.sqrt(12) if len(down) > 0 else ann_vol
    sortino = (ann_ret - rfr) / down_std if down_std > 0 else 0
    win_rate = (rets > 0).mean()
    cum_ret = (1 + rets).prod() - 1
    return dict(periods=len(rets), ann_return=ann_ret, ann_vol=ann_vol,
                sharpe=sharpe, max_dd=max_dd, calmar=calmar, sortino=sortino,
                win_rate=win_rate, cum_return=cum_ret)

# ============================================================
# 1. Load data
# ============================================================
print("Loading data...")

# 50ETF from yfinance
etf50_file = CACHE_DIR / "50etf_yf.parquet"
if etf50_file.exists():
    etf50 = pd.read_parquet(etf50_file)
else:
    ticker = yf.Ticker("510050.SS")
    etf50 = ticker.history(start="2015-02-09", end="2026-06-19")
    etf50.to_parquet(etf50_file)
etf50.index = pd.to_datetime(etf50.index).tz_localize(None)
print(f"50ETF: {len(etf50)} days, {etf50.index[0].date()} to {etf50.index[-1].date()}")

# 300ETF
etf300_file = CACHE_DIR / "300etf_yf.parquet"
if etf300_file.exists():
    etf300 = pd.read_parquet(etf300_file)
else:
    ticker300 = yf.Ticker("510300.SS")
    etf300 = ticker300.history(start="2015-02-09", end="2026-06-19")
    etf300.to_parquet(etf300_file)
etf300.index = pd.to_datetime(etf300.index).tz_localize(None)
print(f"300ETF: {len(etf300)} days")

# IV data
iv_file = CACHE_DIR / "50etf_iv.parquet"
if iv_file.exists():
    iv_df = pd.read_parquet(iv_file)
else:
    for attempt in range(3):
        try:
            iv_df = ak.index_option_50etf_qvix()
            break
        except Exception as e:
            print(f"  IV retry {attempt+1}: {e}")
            time.sleep(5)
    iv_df.to_parquet(iv_file)
iv_df['date'] = pd.to_datetime(iv_df['date'])
iv_df = iv_df.set_index('date').sort_index()
iv_raw = iv_df['close'] / 100.0
print(f"IV: {len(iv_raw)} days, range {iv_raw.min():.1%} - {iv_raw.max():.1%}")

# ============================================================
# 2. Monthly covered call backtest
# ============================================================
print("\n=== 月度备兑认购策略回测 ===")
rfr = 0.025

def run_covered_call(etf_df, iv_series, name, otm_pct=0.0):
    """Run monthly ATM/OTM covered call backtest"""
    prices = etf_df['Close']
    common = prices.index.intersection(iv_series.index)
    prices = prices.loc[common]
    iv = iv_series.loc[common]

    # Monthly rebalance dates
    monthly = []
    for y in range(prices.index[0].year, prices.index[-1].year + 1):
        for m in range(1, 13):
            target = pd.Timestamp(year=y, month=m, day=15)
            valid = prices.index[prices.index >= target]
            if len(valid) > 0: monthly.append(valid[0])
    monthly = sorted(set(d for d in monthly if d >= prices.index[0] and d <= prices.index[-1]))

    results = []
    for i in range(len(monthly) - 1):
        entry_d, exit_d = monthly[i], monthly[i+1]
        mask = (prices.index >= entry_d) & (prices.index <= exit_d)
        pp = prices[mask]
        if len(pp) < 2: continue

        S0, S1 = pp.iloc[0], pp.iloc[-1]
        K = S0 * (1 + otm_pct)
        T = (exit_d - entry_d).days / 365.0
        try:
            sigma = iv.loc[entry_d]
        except:
            sigma = iv.iloc[iv.index.get_indexer([entry_d], method='nearest')[0]]
        premium = bs_price(S0, K, T, rfr, sigma, 'call')
        payoff = -max(0, S1 - K) + premium
        etf_ret = (S1 - S0) / S0
        strat_ret = etf_ret + payoff / S0
        results.append(dict(date=entry_d, etf_ret=etf_ret, strat_ret=strat_ret,
                            premium_pct=premium/S0, sigma=sigma, T_days=(exit_d-entry_d).days))

    return pd.DataFrame(results)

# Run for 50ETF ATM and OTM
r50_atm = run_covered_call(etf50, iv_raw, "50ETF", otm_pct=0.0)
r50_otm2 = run_covered_call(etf50, iv_raw, "50ETF", otm_pct=0.02)
r50_otm3 = run_covered_call(etf50, iv_raw, "50ETF", otm_pct=0.03)

# ============================================================
# 3. Results
# ============================================================
print(f"\n回测区间: {r50_atm['date'].iloc[0].date()} ~ {r50_atm['date'].iloc[-1].date()}")
print(f"共 {len(r50_atm)} 个月度周期\n")

labels = ['买入持有(同期)', '备兑ATM', '备兑OTM(102%)', '备兑OTM(103%)']
datasets = [r50_atm['etf_ret'], r50_atm['strat_ret'], r50_otm2['strat_ret'], r50_otm3['strat_ret']]

print(f"{'策略':<20} {'年化回报':>8} {'年化波动':>8} {'最大回撤':>8} {'Sharpe':>7} {'Calmar':>7} {'Sortino':>7} {'胜率':>7}")
print("-" * 80)
for label, data in zip(labels, datasets):
    m = calc_metrics(data)
    if m:
        print(f"{label:<20} {m['ann_return']:>7.1%}  {m['ann_vol']:>7.1%}  {m['max_dd']:>7.1%}  "
              f"{m['sharpe']:>6.2f}  {m['calmar']:>6.2f}  {m['sortino']:>6.2f}  {m['win_rate']:>6.1%}")

# Premium statistics
print(f"\n月均权利金: ATM={r50_atm['premium_pct'].mean():.2%}, OTM(102%)={r50_otm2['premium_pct'].mean():.2%}, OTM(103%)={r50_otm3['premium_pct'].mean():.2%}")
print(f"年化权利金: ATM={r50_atm['premium_pct'].mean()*12:.2%}, OTM(102%)={r50_otm2['premium_pct'].mean()*12:.2%}, OTM(103%)={r50_otm3['premium_pct'].mean()*12:.2%}")

# Yearly breakdown
print("\n--- 分年表现 (备兑ATM) ---")
r50_atm['year'] = pd.to_datetime(r50_atm['date']).dt.year
yr = r50_atm.groupby('year').agg(
    etf_ann=('etf_ret', lambda x: x.mean()*12),
    strat_ann=('strat_ret', lambda x: x.mean()*12),
    prem_avg=('premium_pct', 'mean'),
    strat_win=('strat_ret', lambda x: (x>0).mean()),
    etf_win=('etf_ret', lambda x: (x>0).mean()),
    n=('etf_ret', 'count'),
)
print(f"{'年份':<6} {'ETF年化':>8} {'备兑年化':>8} {'月均权利金':>9} {'备兑胜率':>7} {'ETF胜率':>7} {'期数':>5}")
for y, row in yr.iterrows():
    print(f"{int(y):<6} {row['etf_ann']:>7.1%}  {row['strat_ann']:>7.1%}  {row['prem_avg']:>8.2%}  {row['strat_win']:>6.1%}  {row['etf_win']:>6.1%}  {int(row['n']):>5}")

print("\nDone.")
