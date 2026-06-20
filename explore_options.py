#!/usr/bin/env python3
"""
50ETF期权数据探索脚本
目的：了解合约链结构、行权价档位、到期日分布，为回测策略设计提供依据
用法：先替换 TOKEN 为有效凭证，然后 python explore_options.py
"""

import rqdatac as rq
import pandas as pd
import numpy as np

# ============================================================
# 1. 初始化 (替换为你的有效凭证)
# ============================================================
TOKEN = "你的有效token"
rq.init(TOKEN)

# ============================================================
# 2. 50ETF 标的历史行情
# ============================================================
print("=" * 80)
print("一、50ETF 标的历史行情")
print("=" * 80)

etf = rq.get_price('510050.XSHG', '20150101', '20260618', fields=['close', 'volume'])
etf.index = pd.to_datetime(etf.index)
etf['year'] = etf.index.year
print(f"数据范围: {etf.index[0].date()} ~ {etf.index[-1].date()}")
print(f"交易日数: {len(etf)}")
print(f"价格: {etf['close'].min():.3f} ~ {etf['close'].max():.3f}")
print(f"日均成交额: {etf['volume'].mean():.0f} 手")
print(f"\n年度统计:")
for yr in range(2015, 2027):
    y = etf[etf['year'] == yr]['close']
    if len(y) > 0:
        print(f"  {yr}: 均价 {y.mean():.3f}, 最低 {y.min():.3f}, 最高 {y.max():.3f}, 波动率 {y.pct_change().std()*np.sqrt(252):.1%}")

# ============================================================
# 3. 期权合约链结构
# ============================================================
print("\n" + "=" * 80)
print("二、期权合约链结构")
print("=" * 80)

# 获取所有存续过的50ETF期权合约
all_contracts = rq.options.get_contracts('510050.XSHG')
print(f"当前存续合约数: {len(all_contracts)}")

# 分析合约属性
contracts_info = []
for c in all_contracts[:500]:  # 只取前500只避免超时
    try:
        prop = rq.options.get_contract_property(c)
        contracts_info.append({
            '合约代码': c,
            '行权价': prop.get('strike_price', np.nan),
            '到期日': pd.to_datetime(prop.get('maturity_date', None)),
            '期权类型': '认购' if 'C' in str(c) else '认沽',
        })
    except:
        pass

if contracts_info:
    df_c = pd.DataFrame(contracts_info)
    df_c['到期年'] = df_c['到期日'].dt.year
    df_c['到期月'] = df_c['到期日'].dt.month

    print(f"\n合约样本 (前10条):")
    print(df_c.head(10).to_string())

    print(f"\n合约到期日分布 (按年):")
    print(df_c.groupby('到期年').size())

    print(f"\n合约到期日分布 (按月):")
    print(df_c.groupby('到期月').size())

    print(f"\n行权价分布 (按期权类型):")
    for opt_type in ['认购', '认沽']:
        sub = df_c[df_c['期权类型'] == opt_type]['行权价']
        if len(sub) > 0:
            print(f"  {opt_type}: {len(sub)}只, {sub.min():.3f} ~ {sub.max():.3f}")

# ============================================================
# 4. 行权价档位和间隔
# ============================================================
print("\n" + "=" * 80)
print("三、行权价档位分析（取最近一个到期日的合约）")
print("=" * 80)

if contracts_info:
    df_c = pd.DataFrame(contracts_info)
    # 取最近到期日
    latest_expiry = df_c['到期日'].min()
    near_contracts = df_c[df_c['到期日'] == latest_expiry]

    print(f"最近到期日: {latest_expiry.date()}, 合约数: {len(near_contracts)}")

    for opt_type in ['认购', '认沽']:
        sub = near_contracts[near_contracts['期权类型'] == opt_type].sort_values('行权价')
        if len(sub) > 0:
            print(f"\n  {opt_type}:")
            strikes = sub['行权价'].tolist()
            # 计算行权价间隔
            diffs = np.diff(strikes)
            print(f"    行权价范围: {strikes[0]:.4f} ~ {strikes[-1]:.4f}")
            print(f"    档位数: {len(strikes)}")
            print(f"    间隔: min={diffs.min():.4f}, max={diffs.max():.4f}, median={np.median(diffs):.4f}")
            print(f"    行权价列表: {[f'{s:.3f}' for s in strikes[:10]]}...")

# ============================================================
# 5. 期权日线数据样本
# ============================================================
print("\n" + "=" * 80)
print("四、期权日线数据样本 (最近一个月)")
print("=" * 80)

if all_contracts and len(all_contracts) > 0:
    # 取交易量最大的前5只合约
    sample_contracts = []
    for c in all_contracts[:20]:
        try:
            d = rq.get_price(c, '20240501', '20240601', fields=['close', 'volume', 'open_interest'])
            if len(d) > 0:
                avg_vol = d['volume'].mean()
                sample_contracts.append((c, avg_vol, len(d)))
        except:
            pass

    sample_contracts.sort(key=lambda x: x[1], reverse=True)

    for c, vol, days in sample_contracts[:5]:
        try:
            d = rq.get_price(c, '20240501', '20240601', fields=['close', 'volume', 'open_interest'])
            if len(d) > 0:
                print(f"\n  合约: {c}")
                print(f"    交易日: {days}天, 日均成交量: {vol:.0f}手")
                print(f"    收盘价: {d['close'].min():.4f} ~ {d['close'].max():.4f}")
                print(f"    日均持仓: {d['open_interest'].mean():.0f}手")
                print(f"    最近5天:")
                print(d.tail(5).to_string())
        except:
            pass

# ============================================================
# 6. 关键约束汇总
# ============================================================
print("\n" + "=" * 80)
print("五、对回测策略的关键约束")
print("=" * 80)

print("""
1. 期权期限约束：
   - 近月、次月、随后两个季月 = 最多约9个月
   - 无法做12个月 outcome period 的 Buffer
   - Covered Call 用近月合约足够

2. 行权价约束：
   - 交易所挂牌标准化行权价，不能自定义
   - 虚值程度只能用最近的标准档位近似
   - Covered Call: 取最接近标的+2~3%的标准行权价
   - Buffer: 保护区间需要根据实际行权价反算

3. 权利金比例约束（公募）：
   - 支付+收取权利金总额 ≤ 净资产10%
   - 未平仓合约面值 ≤ 净资产20%

4. 卖出认沽约束：
   - 需持有行权所需全额现金
   - Buffer底部卖出认沽占用资金成本需计入

5. 流动性约束：
   - 需检查非近月/深度虚值合约的成交量
   - 回测时应剔除日均成交量<100手的合约
""")
PYEOF