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

# ============================================================
# BUFFER STRATEGY BACKTEST
# ============================================================
print("\n" + "=" * 80)
print("沪深300 Buffer策略回测")
print("=" * 80)

# Load 300 index data
idx300_file = CACHE_DIR / "300index_yf.parquet"
if idx300_file.exists():
    idx300 = pd.read_parquet(idx300_file)
else:
    ticker300 = yf.Ticker("000300.SS")
    idx300 = ticker300.history(start="2019-12-01", end="2026-06-19")
    idx300.to_parquet(idx300_file)
idx300.index = pd.to_datetime(idx300.index).tz_localize(None)

# Load 300 ETF IV (more reliable than index IV)
iv300_file = CACHE_DIR / "300etf_iv.parquet"
if iv300_file.exists():
    iv300_df = pd.read_parquet(iv300_file)
else:
    iv300_df = ak.index_option_300etf_qvix()
    iv300_df.to_parquet(iv300_file)
iv300_df['date'] = pd.to_datetime(iv300_df['date'])
iv300_df = iv300_df.set_index('date').sort_index()
iv300 = iv300_df['close'] / 100.0

# Clean data
idx_prices = idx300['Close'].dropna()
common = idx_prices.index.intersection(iv300.index)
idx_prices = idx_prices.loc[common]
iv_300 = iv300.loc[common]
# Remove zero/negative prices
idx_prices = idx_prices[idx_prices > 1000]  # 沪深300 should be >1000
iv_300 = iv_300[iv_300 > 0.05]  # Remove very low IV
common = idx_prices.index.intersection(iv_300.index)
idx_prices = idx_prices.loc[common]
iv_300 = iv_300.loc[common]

print(f"300 Index: {len(idx_prices)} days, {idx_prices.index[0].date()} to {idx_prices.index[-1].date()}")
print(f"300 IV: range {iv_300.min():.1%} - {iv_300.max():.1%}")

# Quarterly rebalance
quarterly_dates = []
for y in range(idx_prices.index[0].year, idx_prices.index[-1].year + 1):
    for m in [3, 6, 9, 12]:
        target = pd.Timestamp(year=y, month=m, day=15)
        valid = idx_prices.index[idx_prices.index >= target]
        if len(valid) > 0:
            quarterly_dates.append(valid[0])
quarterly_dates = sorted(set(d for d in quarterly_dates if d >= idx_prices.index[0] and d <= idx_prices.index[-1]))
print(f"Quarterly rebalance dates: {len(quarterly_dates)} periods")

# Run buffer backtest for different buffer depths
buffer_depths = [0.02, 0.04, 0.06, 0.08, 0.10]
rfr = 0.025

def run_buffer(etf_df, iv_series, buffer_pct):
    """
    Buffer strategy:
    - Long index
    - Buy ATM put (protection)
    - Sell OTM put at (1-buffer%) strike (defines buffer floor)
    - Sell OTM call at 105% strike (to offset cost)
    Quarterly rebalance.
    """
    prices = etf_df
    iv = iv_series

    results = []
    for i in range(len(quarterly_dates) - 1):
        entry_d = quarterly_dates[i]
        exit_d = quarterly_dates[i + 1]

        # Get entry and exit prices
        mask = (prices.index >= entry_d) & (prices.index <= exit_d)
        pp = prices[mask]
        if len(pp) < 2: continue

        S0, S1 = pp.iloc[0], pp.iloc[-1]
        T = (exit_d - entry_d).days / 365.0

        try:
            sigma = iv.loc[entry_d]
        except:
            sigma = iv.iloc[iv.index.get_indexer([entry_d], method='nearest')[0]]

        # Option strikes
        K_put_atm = S0  # ATM protective put
        K_put_buffer = S0 * (1 - buffer_pct)  # Buffer boundary put
        K_call_otm = S0 * 1.05  # 5% OTM call to offset cost

        # Price options
        p_atm = bs_price(S0, K_put_atm, T, rfr, sigma, 'put')
        p_buffer = bs_price(S0, K_put_buffer, T, rfr, sigma, 'put')
        c_otm = bs_price(S0, K_call_otm, T, rfr, sigma, 'call')

        # Buffer structure P&L at expiry
        # - Long put at ATM: max(0, K_put_atm - S1) - p_atm
        # + Short put at buffer: -(max(0, K_put_buffer - S1) - p_buffer)
        # + Short call: -(max(0, S1 - K_call_otm) - c_otm)
        put_atm_payoff = max(0, K_put_atm - S1) - p_atm
        put_buffer_payoff = -(max(0, K_put_buffer - S1) - p_buffer)
        call_payoff = -(max(0, S1 - K_call_otm) - c_otm)

        option_pnl = put_atm_payoff + put_buffer_payoff + call_payoff
        hedge_cost_pct = (p_atm - p_buffer - c_otm) / S0  # net cost

        index_ret = (S1 - S0) / S0
        strat_ret = index_ret + option_pnl / S0

        # Did index breach the buffer?
        breached = S1 < K_put_buffer

        results.append(dict(
            date=entry_d, index_ret=index_ret, strat_ret=strat_ret,
            buffer_pct=buffer_pct, hedge_cost=hedge_cost_pct,
            breached=breached, sigma=sigma, T_days=(exit_d-entry_d).days
        ))

    return pd.DataFrame(results)

# Run for each buffer depth
buffer_results = {}
for b in buffer_depths:
    print(f"\nBuffer {b*100:.0f}%...")
    df = run_buffer(idx_prices, iv_300, b)
    buffer_results[b] = df

# Print comparison
print("\n" + "=" * 80)
print(f"{'策略':<22} {'年化回报':>8} {'年化波动':>8} {'最大回撤':>8} {'Sharpe':>7} {'对冲成本':>8} {'突破月数':>7}")
print("-" * 80)

# Buy-and-hold (over same quarterly periods)
bh_rets_300 = []
for i in range(len(quarterly_dates) - 1):
    entry_d = quarterly_dates[i]
    exit_d = quarterly_dates[i + 1]
    mask = (idx_prices.index >= entry_d) & (idx_prices.index <= exit_d)
    pp = idx_prices[mask]
    if len(pp) < 2: continue
    bh_rets_300.append((pp.iloc[-1] - pp.iloc[0]) / pp.iloc[0])

m_bh_300 = calc_metrics(pd.Series(bh_rets_300))
if m_bh_300:
    print(f"{'买入持有(同期)':<22} {m_bh_300['ann_return']:>7.1%}  {m_bh_300['ann_vol']:>7.1%}  "
          f"{m_bh_300['max_dd']:>7.1%}  {m_bh_300['sharpe']:>6.2f}  {'—':>8}  {'—':>7}")

for b in buffer_depths:
    df = buffer_results[b]
    m = calc_metrics(df['strat_ret'])
    avg_cost = df['hedge_cost'].mean()
    breaches = df['breached'].sum()
    if m:
        print(f"{'Buffer '+str(int(b*100))+'%':<22} {m['ann_return']:>7.1%}  {m['ann_vol']:>7.1%}  "
              f"{m['max_dd']:>7.1%}  {m['sharpe']:>6.2f}  {avg_cost:>7.2%}  {breaches:>6}")

print("\nDone — Buffer backtest complete.")
