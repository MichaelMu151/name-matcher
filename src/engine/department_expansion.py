# -*- coding: utf-8 -*-
"""
codebook + 数据集 处理脚本（全列返回 + 边界不拆分 + 任职起止临时补充版）

流程：
0) 临时补充：styr/stmo 为空 → 填年鉴开始年月；fiyr/fimo 为空 → 填年鉴结束年月
   - 用 _bf_styr_stmo / _bf_fiyr_fimo 记录哪些行是补充的
1) 处理（无变化检查、年份扩充、月份拆分等）
2) 恢复：把被补充的 styr/stmo、fiyr/fimo 还回空值（拆分节点产生的边界保留）
3) 输出：原始全列 + audit_result + 拆分辅助列
"""

import pandas as pd
import numpy as np
from pathlib import Path

# =========================
# 1. 配置区
# =========================
CFG = {
    # 文件路径
    "codebook_path": "/Users/mushiquan/Desktop/RA for Prof.Zhu/广东省官僚/vibe-coding/wname_matcher/codebook.xlsx",
    "data_path": "/Users/mushiquan/Desktop/RA for Prof.Zhu/广州市、深圳市、广东省更新[2026.10.15]/Re__广州深圳数据已更新，广东年鉴或许已经完成/2008-2012_广州市组织史_人工复核完成.xlsx",
    "output_path": "/Users/mushiquan/Desktop/RA for Prof.Zhu/广州市、深圳市、广东省更新[2026.10.15]/Re__广州深圳数据已更新，广东年鉴或许已经完成/2008-2012_广州市组织史_变化已经扩充.xlsx",
    "codebook_sheet": 0,
    "data_sheet": 0,

    # codebook 列名
    "code_name": "w_name",
    "code_sector": "s1_sector",
    "code_wrank": "wrank",
    "code_year_from": "year_valid_from",
    "code_month_from": "month_valid_from",
    "code_year_to": "year_valid_to",
    "code_month_to": "month_valid_to",
    "code_city": "cityname",
    "code_dist": "distname",
    "code_alias": "alias",

    # 数据列名
    "data_name": "w_name",
    "data_name_orig": "w_name_o",
    "data_sector": "s1_sector",
    "data_wrank20": "wrank20",
    "data_wrank_t": "wrank_t",
    "data_year": "year",
    "data_pname": "p_name",
    "data_styr": "styr",
    "data_stmo": "stmo",
    "data_fiyr": "fiyr",
    "data_fimo": "fimo",
    "data_city": "cityname",
    "data_dist": "distname",

    # 年鉴起止时间：由界面输入；为空时在处理函数中按数据 year 自动推断
    "yearbook_year_start": None,
    "yearbook_month_start": 1,
    "yearbook_year_end": None,
    "yearbook_month_end": 12,

    # 是否启用临时补充
    "enable_backfill": True,

    # 有有效期的部门在扩充完成后的最后一步，用覆盖 2020 年的 codebook
    # 段补填 wrank20；无有效期部门不改 wrank20。
    "fill_wrank20_from_periods": True,
    "wrank20_target_year": 2020,
}

PERIOD_COLS = [
    CFG["code_year_from"], CFG["code_month_from"],
    CFG["code_year_to"], CFG["code_month_to"],
]

PENDING_WRANK_VALUES = {"00待处理", "待处理", "00待出"}


# =========================
# 2. 工具函数
# =========================
def is_blank(x):
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    s = str(x).strip()
    return s == "" or s.lower() in ("nan", "none", "<na>", "nat")


def norm_name(x):
    if is_blank(x):
        return pd.NA
    return str(x).strip()


def to_int_scalar(x):
    if is_blank(x):
        return pd.NA
    try:
        return int(float(x))
    except Exception:
        return pd.NA


def to_int_series(s):
    return pd.to_numeric(s, errors="coerce").astype("Int64")


def read_table(path, sheet=0):
    p = Path(path).expanduser()  # 【改】展开 ~
    if not p.exists():
        raise FileNotFoundError(f"文件不存在：{p}")
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    return pd.read_excel(p, sheet_name=sheet)


def split_multi_value(v):
    if is_blank(v):
        return []
    s = str(v)
    for sep in ["|", ";", "；", ",", "，", "、"]:
        s = s.replace(sep, "|")
    return [x.strip() for x in s.split("|") if x.strip()]


def reorder_full_columns(df, original_cols, preferred_extra=None):
    if preferred_extra is None:
        preferred_extra = []
    cols = []
    cols.extend(original_cols)
    for c in preferred_extra:
        if c in df.columns and c not in cols:
            cols.append(c)
    for c in df.columns:
        if c not in cols:
            cols.append(c)
    return df.reindex(columns=cols)


def row_in_scope(row, cr, cfg=CFG):
    for code_col, data_col in [
        (cfg["code_city"], cfg["data_city"]),
        (cfg["code_dist"], cfg["data_dist"]),
    ]:
        if code_col not in cr.index or data_col not in row.index:
            continue
        cv = cr[code_col]
        dv = row[data_col]
        if is_blank(cv):
            continue
        if is_blank(dv):
            return False
        if str(dv).strip() not in split_multi_value(cv):
            return False
    return True


def build_alias_map(cb, cfg=CFG):
    mp = {}
    for _, r in cb.iterrows():
        std = norm_name(r.get(cfg["code_name"], pd.NA))
        if pd.isna(std):
            continue
        mp[std] = std
    for _, r in cb.iterrows():
        std = norm_name(r.get(cfg["code_name"], pd.NA))
        if pd.isna(std):
            continue
        alias_val = r.get(cfg["code_alias"], pd.NA)
        for a in split_multi_value(alias_val):
            if a not in mp:
                mp[a] = std
    return mp


def prepare_codebook(codebook, cfg=CFG):
    cb = codebook.copy()
    for c in PERIOD_COLS:
        if c in cb.columns:
            cb[c] = to_int_series(cb[c])
    cb["_has_period"] = cb.apply(
        lambda r: any(
            (c in cb.columns) and (not is_blank(r[c]))
            for c in PERIOD_COLS
        ),
        axis=1,
    )
    cb["_name_key"] = cb[cfg["code_name"]].map(norm_name)
    return cb


def resolve_name_key(row, code_groups, alias_map, cfg=CFG):
    n = norm_name(row.get(cfg["data_name"], pd.NA))
    if (not pd.isna(n)) and n in code_groups:
        return n
    no = norm_name(row.get(cfg["data_name_orig"], pd.NA))
    if (not pd.isna(no)) and no in alias_map:
        mapped = alias_map[no]
        if mapped in code_groups:
            return mapped
    return n


def get_segments_for_row(row, code_sub, cfg=CFG):
    out = []
    if code_sub is None or len(code_sub) == 0:
        return out
    for _, cr in code_sub.iterrows():
        if row_in_scope(row, cr, cfg):
            out.append(cr)
    return out


def fill_from_code(row, cr, cfg=CFG):
    out = row.copy()
    out[cfg["data_name"]] = cr[cfg["code_name"]]
    out[cfg["data_sector"]] = cr[cfg["code_sector"]]
    out[cfg["data_wrank_t"]] = cr[cfg["code_wrank"]]
    return out


def mark_pending_wrank(out, cr, cfg=CFG):
    wrt = cr.get(cfg["code_wrank"], pd.NA)
    if (not is_blank(wrt)) and str(wrt).strip() in PENDING_WRANK_VALUES:
        out["audit_result"] = "变动后wrank为00待处理，需人工"
        return True
    return False


# =========================
# 【新增】任职起止临时补充 / 恢复
# =========================
def backfill_employment_dates(data, cfg=CFG):
    """
    临时补充 styr/stmo（空 → 年鉴开始年月）和 fiyr/fimo（空 → 年鉴结束年月）。
    记录补充标记，便于处理完后恢复。
    """
    data = data.copy()
    styr, stmo = cfg["data_styr"], cfg["data_stmo"]
    fiyr, fimo = cfg["data_fiyr"], cfg["data_fimo"]
    ys, ms = cfg.get("yearbook_year_start"), cfg.get("yearbook_month_start")
    ye, me = cfg.get("yearbook_year_end"), cfg.get("yearbook_month_end")

    for c in [styr, stmo, fiyr, fimo]:
        if c not in data.columns:
            data[c] = pd.NA

    has_start = (ys is not None) and (ms is not None)
    has_end = (ye is not None) and (me is not None)

    if has_start:
        mask = data[styr].apply(is_blank) | data[stmo].apply(is_blank)
        data["_bf_styr_stmo"] = mask
        data.loc[mask, styr] = ys
        data.loc[mask, stmo] = ms
    else:
        data["_bf_styr_stmo"] = False

    if has_end:
        mask = data[fiyr].apply(is_blank) | data[fimo].apply(is_blank)
        data["_bf_fiyr_fimo"] = mask
        data.loc[mask, fiyr] = ye
        data.loc[mask, fimo] = me
    else:
        data["_bf_fiyr_fimo"] = False

    return data


def restore_employment_dates(df, cfg=CFG):
    """
    恢复临时补充的 styr/stmo、fiyr/fimo。
    只恢复“值仍等于补充值”的单元格：
      - 若该行被拆分到节点段，节点会替换 styr/stmo 或 fiyr/fimo，
        此时值不等于补充值，保留节点边界；
      - 未拆分的行或边界未被替换的行，还原为空。
    """
    if df is None or len(df) == 0:
        return df
    if "_bf_styr_stmo" not in df.columns and "_bf_fiyr_fimo" not in df.columns:
        return df

    df = df.copy()
    styr, stmo = cfg["data_styr"], cfg["data_stmo"]
    fiyr, fimo = cfg["data_fiyr"], cfg["data_fimo"]
    ys, ms = cfg.get("yearbook_year_start"), cfg.get("yearbook_month_start")
    ye, me = cfg.get("yearbook_year_end"), cfg.get("yearbook_month_end")

    def eq_int(col, target):
        def f(x):
            if is_blank(x):
                return False
            v = to_int_scalar(x)
            if pd.isna(v):
                return False
            return int(v) == int(target)
        return df[col].apply(f)

    if "_bf_styr_stmo" in df.columns and ys is not None and ms is not None:
        mask = df["_bf_styr_stmo"].fillna(False).astype(bool) & eq_int(styr, ys) & eq_int(stmo, ms)
        df.loc[mask, [styr, stmo]] = pd.NA

    if "_bf_fiyr_fimo" in df.columns and ye is not None and me is not None:
        mask = df["_bf_fiyr_fimo"].fillna(False).astype(bool) & eq_int(fiyr, ye) & eq_int(fimo, me)
        df.loc[mask, [fiyr, fimo]] = pd.NA

    for c in ["_bf_styr_stmo", "_bf_fiyr_fimo"]:
        if c in df.columns:
            df = df.drop(columns=[c])

    return df


def fill_wrank20_from_periods(data, codebook, cfg=CFG):
    """Fill wrank20 from the validity segment covering the target year.

    This is deliberately a final post-processing step.  It only considers
    codebook rows with at least one validity field and never changes rows whose
    department has no validity segments.  If two segments cover the target
    year because of a month-node overlap, the segment with the latest start
    month wins, which selects the post-transition level.
    """
    if not cfg.get("fill_wrank20_from_periods", True):
        return data
    if data is None or data.empty or cfg["data_wrank20"] not in data.columns:
        return data
    target_year = int(cfg.get("wrank20_target_year", 2020))

    def _period_key(year, month=None):
        year_value = to_int_scalar(year)
        if year_value is None:
            return None
        month_value = to_int_scalar(month)
        return year_value * 12 + (month_value if month_value is not None else 1)
    aliases = build_alias_map(codebook, cfg)
    groups = {key: group for key, group in codebook.groupby("_name_key", dropna=False)}
    output = data.copy()
    for index, row in output.iterrows():
        key = norm_name(row.get(cfg["data_name"]))
        if key not in groups:
            key = aliases.get(norm_name(row.get(cfg["data_name_orig"])), key)
        if key not in groups:
            continue
        candidates = []
        for _, code_row in groups[key].iterrows():
            if not bool(code_row.get("_has_period")) or not row_in_scope(row, code_row, cfg):
                continue
            year_from = to_int_scalar(code_row.get(cfg["code_year_from"]))
            year_to = to_int_scalar(code_row.get(cfg["code_year_to"]))
            if year_from is not None and target_year < year_from:
                continue
            if year_to is not None and target_year > year_to:
                continue
            candidates.append(code_row)
        if candidates:
            selected = max(
                candidates,
                key=lambda item: _period_key(
                    item.get(cfg["code_year_from"]),
                    item.get(cfg["code_month_from"]),
                ) if _period_key(item.get(cfg["code_year_from"]), item.get(cfg["code_month_from"])) is not None else -10**9,
            )
            rank = selected.get(cfg["code_wrank"])
            if not is_blank(rank):
                output.at[index, cfg["data_wrank20"]] = rank
    return output


# =========================
# 3. 主处理函数
# =========================
def process_codebook_data(codebook, data, cfg=CFG):
    # 不把任何一套年鉴年份写死在模块中。界面通常会传入明确范围；
    # 直接调用模块时则从数据中的 year 推断起止年，月份使用 1/12。
    cfg = dict(cfg)
    years = pd.to_numeric(data.get(cfg["data_year"], pd.Series(dtype=object)), errors="coerce").dropna()
    if len(years):
        if cfg.get("yearbook_year_start") in (None, ""):
            cfg["yearbook_year_start"] = int(years.min())
        if cfg.get("yearbook_year_end") in (None, ""):
            cfg["yearbook_year_end"] = int(years.max())
    if cfg.get("yearbook_month_start") in (None, ""):
        cfg["yearbook_month_start"] = 1
    if cfg.get("yearbook_month_end") in (None, ""):
        cfg["yearbook_month_end"] = 12
    cb = prepare_codebook(codebook, cfg)
    data = data.copy()

    # 【新增】保留用户原始列（用于输出时保证全列）
    original_cols = list(data.columns)

    for c in [
        cfg["data_name"], cfg["data_sector"], cfg["data_wrank20"],
        cfg["data_wrank_t"], cfg["data_year"], cfg["data_pname"],
        cfg["data_styr"], cfg["data_stmo"], cfg["data_fiyr"], cfg["data_fimo"],
    ]:
        if c not in data.columns:
            data[c] = pd.NA

    # 【新增】步骤 1：临时补充任职起止
    if cfg.get("enable_backfill", True):
        data = backfill_employment_dates(data, cfg)

    alias_map = build_alias_map(cb, cfg)
    code_groups = {
        k: g for k, g in cb.groupby("_name_key", dropna=False)
    }

    data["_name_key"] = data.apply(
        lambda r: resolve_name_key(r, code_groups, alias_map, cfg),
        axis=1,
    )

    processed = []
    issues = []
    split_log = []

    for name_key, sub in data.groupby("_name_key", dropna=False):
        if pd.isna(name_key) or name_key not in code_groups:
            for _, r in sub.iterrows():
                out = r.copy()
                out["audit_result"] = "未匹配到codebook"
                processed.append(out)
                issues.append(out)
            continue

        code_sub = code_groups[name_key]
        valid_rows = code_sub[code_sub["_has_period"]].copy()

        # =====================================================
        # 情况 A：无有效期字段
        # =====================================================
        if valid_rows.empty:
            for _, r in sub.iterrows():
                matched = get_segments_for_row(r, code_sub, cfg)
                if not matched:
                    out = r.copy()
                    out["audit_result"] = "未匹配到codebook适用范围"
                    processed.append(out)
                    issues.append(out)
                    continue

                cr = matched[0]
                out = fill_from_code(r, cr, cfg)

                wr20 = out.get(cfg["data_wrank20"], pd.NA)
                wrt = out.get(cfg["data_wrank_t"], pd.NA)

                # 分流 1：没有任何有效期记录时，2020 年级别就是当前
                # codebook 级别。覆盖前先保留原值，便于后续审计追溯。
                if not is_blank(wrt):
                    out["wrank20_before_stage"] = wr20
                    out[cfg["data_wrank20"]] = wrt
                    wr20 = wrt

                if (not is_blank(wrt)) and str(wrt).strip() in PENDING_WRANK_VALUES:
                    out["audit_result"] = "wrank为00待处理，需人工"
                    issues.append(out)
                elif (not is_blank(wr20)) and (not is_blank(wrt)) and \
                        str(wr20).strip() != str(wrt).strip():
                    out["audit_result"] = "级别没变但是wrank不统一"
                    issues.append(out)
                else:
                    out["audit_result"] = "正常"

                processed.append(out)
            continue

        # =====================================================
        # 情况 B：有有效期字段
        # =====================================================
        is_month_level = any(
            (not is_blank(r.get(cfg["code_month_from"], pd.NA))) or
            (not is_blank(r.get(cfg["code_month_to"], pd.NA)))
            for _, r in valid_rows.iterrows()
        )

        # B1：只有年份变化
        if not is_month_level:
            for _, r in sub.iterrows():
                y = to_int_scalar(r[cfg["data_year"]])
                if pd.isna(y):
                    out = r.copy()
                    out["audit_result"] = "有变动但year为空，无法按年份填充"
                    processed.append(out)
                    issues.append(out)
                    continue

                matched = get_segments_for_row(r, valid_rows, cfg)
                covered = []
                for cr in matched:
                    yf = cr[cfg["code_year_from"]]
                    yt = cr[cfg["code_year_to"]]
                    if (not is_blank(yf)) and (not is_blank(yt)) and \
                            int(yf) <= int(y) <= int(yt):
                        covered.append(cr)

                if len(covered) == 0:
                    out = r.copy()
                    out["audit_result"] = "有变动但year不在任何有效期"
                    processed.append(out)
                    issues.append(out)
                    continue

                cr = covered[0]
                out = fill_from_code(r, cr, cfg)

                pending = mark_pending_wrank(out, cr, cfg)
                if pending:
                    issues.append(out)
                elif len(covered) > 1:
                    out["audit_result"] = "有变动但year匹配多个有效期，需人工"
                    issues.append(out)
                else:
                    out["audit_result"] = "正常"

                processed.append(out)
            continue

        # B2：有具体月份
        group_cols = [
            cfg["data_pname"], cfg["data_name"],
            cfg["data_styr"], cfg["data_stmo"],
            cfg["data_fiyr"], cfg["data_fimo"],
        ]
        group_cols = [c for c in group_cols if c in sub.columns]

        for gkey, g in sub.groupby(group_cols, dropna=False):
            first = g.iloc[0]

            segs = get_segments_for_row(first, valid_rows, cfg)
            if not segs:
                for _, r in g.iterrows():
                    out = r.copy()
                    out["audit_result"] = "有变动但未匹配到codebook适用范围"
                    processed.append(out)
                    issues.append(out)
                continue

            oy1 = to_int_scalar(first[cfg["data_styr"]])
            om1 = to_int_scalar(first[cfg["data_stmo"]])
            oy2 = to_int_scalar(first[cfg["data_fiyr"]])
            om2 = to_int_scalar(first[cfg["data_fimo"]])

            if pd.isna(oy1) or pd.isna(oy2):
                for _, r in g.iterrows():
                    out = r.copy()
                    out["audit_result"] = "任职起止年份缺失，无法按月拆分"
                    processed.append(out)
                    issues.append(out)
                continue

            month_missing = pd.isna(om1) or pd.isna(om2)
            om1 = 1 if pd.isna(om1) else int(om1)
            om2 = 12 if pd.isna(om2) else int(om2)

            orig_start = (int(oy1), int(om1))
            orig_end = (int(oy2), int(om2))

            overlaps = []
            for cr in segs:
                sy = cr[cfg["code_year_from"]]
                sm = cr[cfg["code_month_from"]]
                ty = cr[cfg["code_year_to"]]
                tm = cr[cfg["code_month_to"]]

                seg_start = None
                if not is_blank(sy):
                    seg_start = (int(sy), int(sm) if not is_blank(sm) else 1)

                seg_end = None
                if not is_blank(ty):
                    seg_end = (int(ty), int(tm) if not is_blank(tm) else 12)

                int_start = max(orig_start, seg_start) if seg_start is not None else orig_start
                int_end = min(orig_end, seg_end) if seg_end is not None else orig_end

                if int_start <= int_end:
                    overlaps.append((cr, int_start, int_end))

            if not overlaps:
                for _, r in g.iterrows():
                    out = r.copy()
                    out["audit_result"] = "任职区间不在任何有效期段内"
                    processed.append(out)
                    issues.append(out)
                continue

            # 边界不拆分：过滤零长度边界段
            if len(overlaps) > 1:
                non_zero = [(cr, s, e) for cr, s, e in overlaps if s < e]
                if non_zero:
                    overlaps = non_zero
                else:
                    overlaps = [overlaps[-1]]

            split_flag = 1 if len(overlaps) > 1 else 0

            for cr, int_start, int_end in overlaps:
                for idx, orig_row in g.iterrows():
                    y = to_int_scalar(orig_row[cfg["data_year"]])
                    if pd.isna(y):
                        continue

                    if int(int_start[0]) <= int(y) <= int(int_end[0]):
                        out = orig_row.copy()

                        out[cfg["data_styr"]] = int_start[0]
                        out[cfg["data_stmo"]] = int_start[1]
                        out[cfg["data_fiyr"]] = int_end[0]
                        out[cfg["data_fimo"]] = int_end[1]

                        out[cfg["data_name"]] = cr[cfg["code_name"]]
                        out[cfg["data_sector"]] = cr[cfg["code_sector"]]
                        out[cfg["data_wrank_t"]] = cr[cfg["code_wrank"]]

                        out["_split_flag"] = split_flag
                        out["_segment_start"] = f"{int_start[0]}-{int_start[1]:02d}"
                        out["_segment_end"] = f"{int_end[0]}-{int_end[1]:02d}"

                        pending = mark_pending_wrank(out, cr, cfg)
                        if pending:
                            issues.append(out)
                        elif month_missing:
                            out["audit_result"] = "月份缺失，已按默认月份拆分，需复核"
                        else:
                            out["audit_result"] = "正常"

                        processed.append(out)

                        log_row = orig_row.copy()
                        log_row["段开始"] = f"{int_start[0]}-{int_start[1]:02d}"
                        log_row["段结束"] = f"{int_end[0]}-{int_end[1]:02d}"
                        log_row["wrank_t_填充"] = cr[cfg["code_wrank"]]
                        log_row["split_flag"] = split_flag
                        split_log.append(log_row)

    processed_df = pd.DataFrame(processed)
    issues_df = pd.DataFrame(issues)
    split_log_df = pd.DataFrame(split_log)

    processed_df = processed_df.drop(columns=["_name_key"], errors="ignore")
    issues_df = issues_df.drop(columns=["_name_key"], errors="ignore")
    split_log_df = split_log_df.drop(columns=["_name_key"], errors="ignore")

    # Final step of block 3: only validity-changing departments receive a
    # 2020 wrank20 value.  No-validity departments retain their input value.
    processed_df = fill_wrank20_from_periods(processed_df, cb, cfg)
    issues_df = fill_wrank20_from_periods(issues_df, cb, cfg)

    # 【新增】步骤 3：恢复临时补充的任职起止
    if cfg.get("enable_backfill", True):
        processed_df = restore_employment_dates(processed_df, cfg)
        issues_df = restore_employment_dates(issues_df, cfg)
        # 拆分日志也恢复，日志里仍保留 段开始/段结束 作为拆分记录
        split_log_df = restore_employment_dates(split_log_df, cfg)

    # 保证全列返回且列顺序稳定
    processed_df = reorder_full_columns(
        processed_df,
        original_cols=original_cols,
        preferred_extra=[
            "audit_result",
            "_split_flag",
            "_segment_start",
            "_segment_end",
        ],
    )

    issues_df = reorder_full_columns(
        issues_df,
        original_cols=original_cols,
        preferred_extra=["audit_result"],
    )

    if len(split_log_df) > 0:
        split_log_df = reorder_full_columns(
            split_log_df,
            original_cols=original_cols,
            preferred_extra=["段开始", "段结束", "wrank_t_填充", "split_flag"],
        )

    return processed_df, issues_df, split_log_df


# =========================
# 4. 卡片摘要
# =========================
def build_card_summary(processed_df, issues_df, split_log_df, original_row_count):
    issue_counts = (
        issues_df["audit_result"].value_counts()
        if "audit_result" in issues_df.columns else pd.Series(dtype=int)
    )

    split_group_count = 0
    if "split_flag" in split_log_df.columns and len(split_log_df) > 0:
        cols = [c for c in [CFG["data_pname"], CFG["data_name"],
                            CFG["data_styr"], CFG["data_stmo"],
                            CFG["data_fiyr"], CFG["data_fimo"]]
                if c in split_log_df.columns]
        try:
            split_group_count = split_log_df[split_log_df["split_flag"] == 1][cols].drop_duplicates().shape[0]
        except Exception:
            split_group_count = 0

    return pd.DataFrame([
        {"指标": "原始数据行数", "值": original_row_count},
        {"指标": "处理总行数", "值": len(processed_df)},
        {"指标": "新增拆分行数", "值": len(processed_df) - original_row_count},
        {"指标": "问题行数", "值": len(issues_df)},
        {"指标": "级别没变但是wrank不统一", "值": int(issue_counts.get("级别没变但是wrank不统一", 0))},
        {"指标": "变动后wrank为00待处理", "值": int(issue_counts.get("变动后wrank为00待处理，需人工", 0))},
        {"指标": "wrank为00待处理（无变化）", "值": int(issue_counts.get("wrank为00待处理，需人工", 0))},
        {"指标": "未匹配到codebook", "值": int(issue_counts.get("未匹配到codebook", 0))},
        {"指标": "发生拆分的任职组数", "值": split_group_count},
    ])


def process_files(codebook_path, data_path, output_path, cfg=CFG):
    """Run the module outside the GUI and export the four-stage workbook."""
    codebook = read_table(codebook_path, cfg.get("codebook_sheet", 0))
    data = read_table(data_path, cfg.get("data_sheet", 0))
    processed_df, issues_df, split_log_df = process_codebook_data(codebook, data, cfg)
    summary_df = build_card_summary(processed_df, issues_df, split_log_df, len(data))
    output_path = Path(output_path).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        processed_df.to_excel(writer, sheet_name="处理结果", index=False)
        issues_df.to_excel(writer, sheet_name="问题清单", index=False)
        split_log_df.to_excel(writer, sheet_name="拆分日志", index=False)
        summary_df.to_excel(writer, sheet_name="卡片摘要", index=False)
    return processed_df, issues_df, split_log_df, summary_df


# =========================
# 5. 主入口
# =========================
if __name__ == "__main__":
    codebook = read_table(CFG["codebook_path"], CFG["codebook_sheet"])
    data = read_table(CFG["data_path"], CFG["data_sheet"])

    processed_df, issues_df, split_log_df = process_codebook_data(codebook, data)

    summary_df = build_card_summary(
        processed_df, issues_df, split_log_df,
        original_row_count=len(data),
    )

    output_path = Path(CFG["output_path"]).expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        processed_df.to_excel(writer, sheet_name="处理结果", index=False)
        issues_df.to_excel(writer, sheet_name="问题清单", index=False)
        split_log_df.to_excel(writer, sheet_name="拆分日志", index=False)
        summary_df.to_excel(writer, sheet_name="卡片摘要", index=False)

    print(f"处理完成：{output_path}")
    print(f"原始数据行数：{len(data)}")
    print(f"处理结果行数：{len(processed_df)}")
    print(f"问题行数：{len(issues_df)}")
    print(f"拆分日志行数：{len(split_log_df)}")
