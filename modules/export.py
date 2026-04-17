from __future__ import annotations
import io
from datetime import datetime

import pandas as pd
from fpdf import FPDF


# ── 工具 ──────────────────────────────────────────────────────────────────────
def _fmt(val, suffix="", decimals=1) -> str:
    if val is None:
        return "N/A"
    return f"{val:,.{decimals}f}{suffix}"


def _pct_arrow(val) -> str:
    if val is None:
        return "N/A"
    sign = "+" if val >= 0 else "-"
    return f"{sign}{abs(val):.1f}%"


# ── B-4a：匯出 Excel ──────────────────────────────────────────────────────────
def to_excel(result: dict) -> bytes:
    base = result["base"]
    comp = result["comp"]
    bm   = result["base_metrics"]
    cm   = result["comp_metrics"]

    base_label = f"{base['site_id']} / {base['year_month']} / {base['data_type']}"
    comp_label = f"{comp['site_id']} / {comp['year_month']} / {comp['data_type']}"

    rows = [
        ("營收 Revenue",        bm["revenue"],       cm["revenue"],
         result["revenue_delta"],  result["revenue_delta_pct"]),
        ("銷貨成本 COGS",        bm["cogs"],          cm["cogs"],
         cm["cogs"] - bm["cogs"],  None),
        ("毛利 Gross Profit",   bm["gross_profit"],  cm["gross_profit"],
         result["gp_delta"],       result["gp_delta_pct"]),
        ("毛利率 GM%",
         bm["gross_margin_pct"], cm["gross_margin_pct"],
         (cm["gross_margin_pct"] or 0) - (bm["gross_margin_pct"] or 0), None),
        ("營業費用 OpEx",        bm["opex"],          cm["opex"],
         result["opex_delta"],     None),
        ("費用率 OpEx%",
         bm["opex_pct"],          cm["opex_pct"],
         result["opex_pct_delta"], None),
    ]

    df = pd.DataFrame(rows, columns=["指標", base_label, comp_label, "增減額", "增減率%"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="財務對比")
        wb  = writer.book
        ws  = writer.sheets["財務對比"]

        # 格式
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#1F4E79",
                                  "font_color": "white", "border": 1})
        num_fmt  = wb.add_format({"num_format": "#,##0.0", "border": 1})
        pct_fmt  = wb.add_format({"num_format": "0.00%",   "border": 1})
        cell_fmt = wb.add_format({"border": 1})
        alert_fmt = wb.add_format({"font_color": "red", "bold": True, "border": 1})

        # 寫標題
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, hdr_fmt)

        # 寫資料
        for row_idx, row in enumerate(rows, start=1):
            ws.write(row_idx, 0, row[0], cell_fmt)
            for col_offset, val in enumerate(row[1:], start=1):
                fmt = pct_fmt if "%" in row[0] else num_fmt
                if val is None:
                    ws.write(row_idx, col_offset, "N/A", cell_fmt)
                else:
                    ws.write(row_idx, col_offset, val, fmt)

        # 欄寬
        ws.set_column(0, 0, 22)
        ws.set_column(1, 4, 18)

        # 摘要區
        summary_row = len(rows) + 2
        ws.write(summary_row, 0, "對比類型", hdr_fmt)
        ws.write(summary_row, 1, result["comparison_label"], cell_fmt)
        ws.write(summary_row + 1, 0, "OpEx 警示閾值", hdr_fmt)
        ws.write(summary_row + 1, 1, f"{result['opex_threshold']}%", cell_fmt)
        ws.write(summary_row + 2, 0, "產生時間", hdr_fmt)
        ws.write(summary_row + 2, 1, datetime.now().strftime("%Y-%m-%d %H:%M"), cell_fmt)

    buf.seek(0)
    return buf.read()


# ── B-4b：匯出 PDF ────────────────────────────────────────────────────────────
class _PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "Financial Analysis Report", align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "", 9)
        self.cell(0, 6, datetime.now().strftime("Generated: %Y-%m-%d %H:%M"),
                  align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def to_pdf(result: dict) -> bytes:
    base = result["base"]
    comp = result["comp"]
    bm   = result["base_metrics"]
    cm   = result["comp_metrics"]

    base_label = f"{base['site_id']} {base['year_month']} {base['data_type']}"
    comp_label = f"{comp['site_id']} {comp['year_month']} {comp['data_type']}"

    pdf = _PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # 對比類型標題
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 8, f"  {result['comparison_label_en']}", fill=True,
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)

    # 比較對象
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(40, 7, "Base:")
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, base_label, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(40, 7, "Compare:")
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, comp_label, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # 表格
    headers = ["Metric", base_label[:22], comp_label[:22], "Delta", "Delta%"]
    col_w   = [50, 32, 32, 28, 28]

    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(31, 78, 121)
    pdf.set_text_color(255, 255, 255)
    for h, w in zip(headers, col_w):
        pdf.cell(w, 7, h, border=1, fill=True)
    pdf.ln()
    pdf.set_text_color(0, 0, 0)

    rows = [
        ("Revenue",        bm["revenue"],          cm["revenue"],
         result["revenue_delta"],     result["revenue_delta_pct"]),
        ("COGS",           bm["cogs"],             cm["cogs"],
         cm["cogs"] - bm["cogs"],     None),
        ("Gross Profit",   bm["gross_profit"],     cm["gross_profit"],
         result["gp_delta"],          result["gp_delta_pct"]),
        ("Gross Margin%",  bm["gross_margin_pct"], cm["gross_margin_pct"],
         (cm["gross_margin_pct"] or 0) - (bm["gross_margin_pct"] or 0), None),
        ("OpEx",           bm["opex"],             cm["opex"],
         result["opex_delta"],        None),
        ("OpEx%",          bm["opex_pct"],         cm["opex_pct"],
         result["opex_pct_delta"],    None),
    ]

    for i, row in enumerate(rows):
        fill = i % 2 == 0
        pdf.set_fill_color(240, 245, 255) if fill else pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "B" if "%" in row[0] else "", 9)
        pdf.cell(col_w[0], 6, row[0], border=1, fill=fill)

        is_pct = "%" in row[0]
        for j, (val, w) in enumerate(zip(row[1:], col_w[1:]), start=1):
            txt = _fmt(val, "%") if is_pct else _fmt(val)
            if j == 4 and not is_pct:   # Delta%
                txt = _pct_arrow(val)
            # 紅色標示負值 delta
            if j in (3, 4) and val is not None and val < 0:
                pdf.set_text_color(200, 0, 0)
            else:
                pdf.set_text_color(0, 0, 0)
            pdf.cell(w, 6, txt, border=1, fill=fill)
        pdf.set_text_color(0, 0, 0)
        pdf.ln()

    pdf.ln(5)

    # OpEx 預警
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, f"OpEx Alert (threshold: {result['opex_threshold']:.0f}%)",
             new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)

    for label, rec, alert, pct in [
        (base_label, base, result["base_opex_alert"],  bm["opex_pct"]),
        (comp_label, comp, result["comp_opex_alert"],  cm["opex_pct"]),
    ]:
        status = "!! EXCEEDED" if alert else "OK"
        pdf.set_text_color(200, 0, 0) if alert else pdf.set_text_color(0, 128, 0)
        pdf.cell(0, 6, f"  {label[:30]}: OpEx% = {_fmt(pct, '%')}  =>  {status}",
                 new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)

    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf.read()
