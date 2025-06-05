import pandas as pd
import numpy as np
import os

# ---------------- 方法2：替代分配 ----------------
def add_replacement_quality(df: pd.DataFrame) -> pd.DataFrame:
    """
    替代分析：评价占比上升的队列 vs 被替代的平均质量
    """
    grow = df[df['评价权重差值'] > 0].copy()
    loss = df[df['评价权重差值'] < 0].copy()

    if grow.empty or loss.empty:
        df['替代平均质量'] = np.nan
        df['净替代质量差'] = np.nan
        return df

    total_loss = loss['评价权重差值'].abs().sum()
    s_replaced = (loss['现期服务满意度'] * loss['评价权重差值'].abs()).sum() / total_loss

    df['替代平均质量'] = np.nan
    df['净替代质量差'] = np.nan
    df.loc[grow.index, '替代平均质量'] = s_replaced
    df.loc[grow.index, '净替代质量差'] = df.loc[grow.index, '现期服务满意度'] - s_replaced
    return df

# ---------------- 留一法净影响 ----------------
def add_leave_one_out(df: pd.DataFrame) -> pd.DataFrame:
    """
    留一法净影响 = (剔除后现期满意度 - 基期整体) - (原现期整体 - 基期整体)
    """
    cur_total_score = df['现期_总评分'].sum()
    cur_total_eval = df['现期_总评价量'].sum()
    base_total_score = df['基期_总评分'].sum()
    base_total_eval = df['基期_总评价量'].sum()

    S_cur_all = cur_total_score / cur_total_eval if cur_total_eval != 0 else np.nan
    S_base_all = base_total_score / base_total_eval if base_total_eval != 0 else np.nan

    excl_score = cur_total_score - df['现期_总评分']
    excl_eval = cur_total_eval - df['现期_总评价量']
    S_cur_excl = np.where(excl_eval != 0, excl_score / excl_eval, np.nan)

    df['留一法净影响'] = (S_cur_excl - S_base_all) - (S_cur_all - S_base_all)
    return df

# ---------------- 方法1：运营标签 ----------------
def add_op_tag(df: pd.DataFrame, delta_thresh=0.005, replace_thresh=0.01, dominance_ratio=1.5) -> pd.DataFrame:
    ΔS = df['服务满意度差值']
    ΔW = df['评价权重差值']
    impact = df['总影响值']
    replace_gap = df['净替代质量差']

    # 初始化标签
    tags = np.full(len(df), '非主要变动影响', dtype=object)

    # 1 核心贡献
    mask = (ΔS > delta_thresh) & (ΔW > delta_thresh) & (impact > 0) & (replace_gap > replace_thresh)
    tags[mask] = '核心贡献者（服务提升，评价流量提升）'

    # 2 服务提升+抢占
    mask = (ΔS > delta_thresh) & (ΔW > delta_thresh) & (impact > 0) & (replace_gap < -replace_thresh)
    tags[mask] = '服务提升，但抢了更多评价流量'

    # 3 服务提升+流量减少
    mask = (ΔS > delta_thresh) & (ΔW < -delta_thresh) & (impact > 0)
    tags[mask] = '服务提升，但流量丢失（建议关注回流）'

    # 4 服务下降+正影响+优化
    mask = (ΔS < -delta_thresh) & (ΔW > delta_thresh) & (impact > 0) & (replace_gap > replace_thresh)
    tags[mask] = '服务下降但仍属正向贡献'

    # 5 服务下降+抢占
    mask = (ΔS < -delta_thresh) & (ΔW > delta_thresh) & (impact > 0) & (replace_gap < -replace_thresh)
    tags[mask] = '服务下降评价流量却增加（需关注）'

    # 6 服务下降+抢占+负影响
    mask = (ΔS < -delta_thresh) & (ΔW > delta_thresh) & (impact <= 0)
    tags[mask] = '服务下降评价流量却增加（需关注）'

    # ✅ 7 服务下降+流量下降 → 同时判断主导因素
    mask = (ΔS < 0) & (ΔW < 0) & (impact < 0)#门槛值有问题，如果权重变动非常小低于门槛值，但是满意度变动大，也会导致最终标签是无主要影响，所以取消门槛值，防止误杀
    sub = df.loc[mask]
    service_impact = sub['服务满意度影响值'].abs()
    weight_impact = sub['评价权重影响值'].abs()

    sub_tags = np.where(
        service_impact > dominance_ratio * weight_impact,
        '服务下降，主要因服务质量下降',
        np.where(
            weight_impact > dominance_ratio * service_impact,
            '服务下降，主要因流量下降',
            '服务下降，服务+流量双重影响'
        )
    )
    tags[mask] = sub_tags

    # 8 数据不足
    mask = (ΔS.abs() <= delta_thresh) & (ΔW.abs() <= delta_thresh) & (impact.abs() <= delta_thresh)
    tags[mask] = '基期或现期数据不足'

    # 返回标签
    df['运营评价标签'] = tags
    return df


# ---------------- 正负缩放 ----------------
def sign_split_compress(df):
    """
    对影响值进行正负缩放，确保结果闭合
    """
    Δ = df['总影响值'].sum()
    pos_mask = df['总影响值'] > 0
    neg_mask = df['总影响值'] < 0
    L_pos = df.loc[pos_mask, '总影响值'].sum()
    L_neg = df.loc[neg_mask, '总影响值'].abs().sum()
    Δ_pos = max(0, Δ)
    Δ_neg = max(0, -Δ)
    k_pos = Δ_pos / L_pos if L_pos else 0
    k_neg = Δ_neg / L_neg if L_neg else 0

    arr = np.where(
        df['总影响值'] > 0, df['总影响值'] * k_pos,
        np.where(df['总影响值'] < 0, df['总影响值'] * k_neg, 0.0)
    ).astype(float)
    df['转化影响值'] = np.round(arr, 6)
    return df

# ---------------- 核心分析 ----------------
def zy_satis_analy(cur_df, base_df, dims, min_eva, enable_method2):
    """
    核心分析逻辑：二阶分解 + 替代分析 + 留一法
    """
    agg = ['总评分', '总评价量']
    cur = cur_df.groupby(dims)[agg].sum().rename(columns=lambda c: f'现期_{c}')
    base = base_df.groupby(dims)[agg].sum().rename(columns=lambda c: f'基期_{c}')
    g = cur.join(base, how='outer').fillna(0).reset_index()

    g['基期服务满意度'] = g['基期_总评分'] / g['基期_总评价量'].replace(0, pd.NA)
    g['现期服务满意度'] = g['现期_总评分'] / g['现期_总评价量'].replace(0, pd.NA)
    g[['基期服务满意度', '现期服务满意度']] = g[['基期服务满意度', '现期服务满意度']].ffill(axis=1)

    g['基期评价权重'] = g['基期_总评价量'] / g['基期_总评价量'].sum()
    g['现期评价权重'] = g['现期_总评价量'] / g['现期_总评价量'].sum()

    g['服务满意度差值'] = g['现期服务满意度'] - g['基期服务满意度']
    g['评价权重差值'] = g['现期评价权重'] - g['基期评价权重']

    g['服务满意度影响值'] = g['基期评价权重'] * g['服务满意度差值']
    g['评价权重影响值'] = g['基期服务满意度'] * g['评价权重差值']
    g['交互项影响值'] = g['服务满意度差值'] * g['评价权重差值']
    g['总影响值'] = g[['服务满意度影响值', '评价权重影响值', '交互项影响值']].sum(axis=1)

    g = g[g['现期_总评价量'] >= min_eva]

    # 替代分析
    if enable_method2:
        g = add_replacement_quality(g)

    # 留一法
    g = add_leave_one_out(g)

    # 标签
    g = add_op_tag(g)

    return g.sort_values('总影响值', key=abs, ascending=False)

# ---------------- Notebook接口 ----------------
def run_analysis(data_path, cur_start, cur_end,
                 base_start, base_end, core_dims,
                 min_eva=0, excel_out="result.xlsx",
                 enable_method2=True):
    """
    主分析入口
    """
    data = pd.read_excel(data_path)
    data['dt'] = pd.to_datetime(data['dt'])

    cur_df = data[data['dt'].between(cur_start, cur_end)]
    base_df = data[data['dt'].between(base_start, base_end)]

    results = {}
    for dims in core_dims:
        df_out = zy_satis_analy(cur_df, base_df, dims, min_eva, enable_method2)
        df_out = sign_split_compress(df_out)
        results[tuple(dims)] = df_out

    with pd.ExcelWriter(excel_out, engine='xlsxwriter') as w:
        for dims, table in results.items():
            sheet = (" + ".join(dims))[:31]
            table.to_excel(w, sheet_name=sheet, index=False)

    print(f"\n✅ 完成！共输出 {len(results)} 张表 → {os.path.abspath(excel_out)}")
    return results
