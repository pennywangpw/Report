from __future__ import annotations
import sqlite3
from modules.database import get_connection


# ── 工具：解析選項字串 ─────────────────────────────────────────────────────────
def parse_option(option: str) -> tuple[str, str, str]:
    """'Site A / 2025-Q1 / Actual' → ('Site A', '2025-Q1', 'Actual')"""
    parts = [p.strip() for p in option.split("/")]
    return parts[0], parts[1], parts[2]


# ── 從 DB 取單筆記錄 ───────────────────────────────────────────────────────────
def get_record(site_id: str, year_month: str, data_type: str) -> dict | None:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT site_id, year_month, data_type,
               revenue, cogs, opex, gross_profit
        FROM financial_master
        WHERE site_id=? AND year_month=? AND data_type=?
        """,
        (site_id, year_month, data_type),
    )
    row = cursor.fetchone()
    conn.close()
    if row is None:
        return None
    return dict(row)


# ── 自動判斷對比類型 ───────────────────────────────────────────────────────────
def detect_comparison_type(base: dict, comp: dict) -> str:
    """
    回傳:
      'site_vs_site' — 同季度跨廠
      'yoy'          — 同廠不同年度
      'avb'          — Actual vs Budget
      'custom'       — 其他
    """
    same_site  = base["site_id"]    == comp["site_id"]
    same_ym    = base["year_month"] == comp["year_month"]
    same_dtype = base["data_type"]  == comp["data_type"]

    if not same_site and same_ym and same_dtype:
        return "site_vs_site"
    if same_site and not same_ym:
        return "yoy"
    if same_site and same_ym and not same_dtype:
        return "avb"
    return "custom"


COMPARISON_LABELS = {
    "site_vs_site": "跨廠比較（Site A vs. Site B）",
    "yoy":          "年度成長分析（YoY）/季增率（QoQ)",
    "avb":          "預算執行率（Actual vs. Budget）",
    "custom":       "自訂對比",
}

COMPARISON_LABELS_EN = {
    "site_vs_site": "Cross-Site Comparison (Site A vs. Site B)",
    "yoy":          "Year-over-Year Growth Analysis",
    "avb":          "Actual vs. Budget Execution Rate",
    "custom":       "Custom Comparison",
}


# ── 核心計算 ───────────────────────────────────────────────────────────────────
def _safe_pct(numerator: float, denominator: float) -> float | None:
    """計算百分比變動；分母為 0 時回傳 None"""
    if denominator == 0:
        return None
    return numerator / denominator * 100


def compare(base: dict, comp: dict, opex_threshold: float = 20.0) -> dict:
    """
    計算雙方財務指標差異，回傳完整比較結果 dict。
    base / comp 均為 financial_master 的 row dict。
    opex_threshold: OpEx% 警示閾值（百分比，預設 20%）。
    """
    def metrics(rec: dict) -> dict:
        rev  = rec["revenue"]     or 0.0
        cogs = rec["cogs"]        or 0.0
        opex = rec["opex"]        or 0.0
        gp   = rec["gross_profit"] if rec["gross_profit"] is not None else (rev - cogs)
        gm_pct   = _safe_pct(gp,   rev)
        opex_pct = _safe_pct(opex, rev)
        return {
            "revenue": rev, "cogs": cogs, "opex": opex, "gross_profit": gp,
            "gross_margin_pct": gm_pct,
            "opex_pct": opex_pct,
        }

    bm = metrics(base)
    cm = metrics(comp)

    def delta(key):
        return cm[key] - bm[key]

    def delta_pct(key):
        return _safe_pct(delta(key), bm[key]) if bm[key] != 0 else None

    # OpEx 超標預警
    b_alert = bm["opex_pct"] is not None and bm["opex_pct"] > opex_threshold
    c_alert = cm["opex_pct"] is not None and cm["opex_pct"] > opex_threshold

    ctype = detect_comparison_type(base, comp)

    return {
        "comparison_type": ctype,
        "comparison_label":    COMPARISON_LABELS[ctype],
        "comparison_label_en": COMPARISON_LABELS_EN[ctype],
        "base": base,
        "comp": comp,
        "base_metrics": bm,
        "comp_metrics": cm,
        # 差額 & 成長率
        "revenue_delta":      delta("revenue"),
        "revenue_delta_pct":  delta_pct("revenue"),
        "gp_delta":           delta("gross_profit"),
        "gp_delta_pct":       delta_pct("gross_profit"),
        "opex_delta":         delta("opex"),
        "opex_pct_delta":     (
            (cm["opex_pct"] or 0) - (bm["opex_pct"] or 0)
        ),
        # 預警
        "opex_threshold": opex_threshold,
        "base_opex_alert": b_alert,
        "comp_opex_alert": c_alert,
    }
