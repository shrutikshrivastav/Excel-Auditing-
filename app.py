import os
import io
import json
import uuid
import hashlib
import warnings
import traceback
from datetime import datetime
from collections import Counter

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import re

warnings.filterwarnings("ignore")

app = Flask(__name__, static_folder=".", template_folder=".")
CORS(app)

# ── colour palette ──────────────────────────────────────────────────────────
C = {
    "bg_dark":    "1A0A0A",
    "bg_header":  "2D0808",
    "accent":     "C0392B",
    "accent2":    "E74C3C",
    "gold":       "D4AC0D",
    "text_light": "F5F5F5",
    "text_dim":   "AAAAAA",
    "row_alt":    "2A1010",
    "green":      "27AE60",
    "orange":     "E67E22",
    "red":        "C0392B",
    "white":      "FFFFFF",
    "border":     "5C1A1A",
}

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=11, color="F5F5F5", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def _border():
    side = Side(border_style="thin", color=C["border"])
    return Border(left=side, right=side, top=side, bottom=side)

def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


# ── helpers ─────────────────────────────────────────────────────────────────
def detect_column_type(series: pd.Series):
    s = series.dropna()
    if len(s) == 0:
        return "empty"
    if pd.api.types.is_numeric_dtype(s):
        return "numeric"
    try:
        pd.to_datetime(s, infer_datetime_format=True, errors="raise")
        return "date"
    except Exception:
        pass
    return "text"


def find_salary_cols(df):
    kw = ["salary", "sal", "wage", "pay", "ctc", "compensation", "income"]
    return [c for c in df.columns
            if any(k in c.lower() for k in kw) and pd.api.types.is_numeric_dtype(df[c])]


def find_profit_cols(df):
    kw = ["profit", "margin", "earnings", "net", "gain", "revenue", "sales"]
    return [c for c in df.columns
            if any(k in c.lower() for k in kw) and pd.api.types.is_numeric_dtype(df[c])]


def find_amount_cols(df):
    kw = ["amount", "total", "price", "cost", "value", "fee", "charge"]
    return [c for c in df.columns
            if any(k in c.lower() for k in kw) and pd.api.types.is_numeric_dtype(df[c])]


def check_business_logic(df):
    issues = []

    # revenue > cost => profit check
    rev_cols = [c for c in df.columns if "revenue" in c.lower() and pd.api.types.is_numeric_dtype(df[c])]
    cost_cols = [c for c in df.columns if "cost" in c.lower() and pd.api.types.is_numeric_dtype(df[c])]
    profit_cols = find_profit_cols(df)

    if rev_cols and cost_cols and profit_cols:
        r, co, p = rev_cols[0], cost_cols[0], profit_cols[0]
        expected = df[r] - df[co]
        diff = (df[p] - expected).abs()
        bad = diff[diff > 0.01]
        if not bad.empty:
            issues.append({
                "type": "logic_error",
                "severity": "high",
                "description": f"Profit mismatch: {p} ≠ {r} − {co}",
                "affected_rows": bad.index.tolist()[:20],
                "count": len(bad)
            })

    # discount > 100%
    disc_cols = [c for c in df.columns if "discount" in c.lower() and pd.api.types.is_numeric_dtype(df[c])]
    for dc in disc_cols:
        bad = df[df[dc] > 100]
        if not bad.empty:
            issues.append({
                "type": "invalid_value",
                "severity": "high",
                "description": f"Discount > 100% in column '{dc}'",
                "affected_rows": bad.index.tolist()[:20],
                "count": len(bad)
            })

    # age out of range
    age_cols = [c for c in df.columns if "age" == c.lower() and pd.api.types.is_numeric_dtype(df[c])]
    for ac in age_cols:
        bad = df[(df[ac] < 0) | (df[ac] > 120)]
        if not bad.empty:
            issues.append({
                "type": "invalid_value",
                "severity": "medium",
                "description": f"Unrealistic age values in '{ac}'",
                "affected_rows": bad.index.tolist()[:20],
                "count": len(bad)
            })

    return issues


def detect_statistical_anomalies(df):
    anomalies = []
    for col in df.select_dtypes(include=[np.number]).columns:
        s = df[col].dropna()
        if len(s) < 10:
            continue
        mean, std = s.mean(), s.std()
        if std == 0:
            continue
        z = ((s - mean) / std).abs()
        outliers = z[z > 3]
        if not outliers.empty:
            anomalies.append({
                "column": col,
                "count": len(outliers),
                "rows": outliers.index.tolist()[:10],
                "values": s[outliers.index].tolist()[:10]
            })
    return anomalies


def detect_suspicious_edits(df):
    flags = []
    for col in df.select_dtypes(include=[np.number]).columns:
        s = df[col].dropna()
        if len(s) < 5:
            continue
        # look for suspiciously round numbers
        round_mask = s.apply(lambda x: x % 1000 == 0 and x != 0)
        round_rows = s[round_mask]
        ratio = len(round_rows) / len(s)
        if 0 < ratio < 0.15 and len(round_rows) >= 2:
            flags.append({
                "column": col,
                "type": "suspiciously_round",
                "description": f"{len(round_rows)} suspiciously round values (multiples of 1000) in '{col}'",
                "rows": round_rows.index.tolist()[:10]
            })
    return flags


def full_audit(df: pd.DataFrame, filename: str):
    report = {
        "filename": filename,
        "timestamp": datetime.now().isoformat(),
        "shape": {"rows": len(df), "cols": len(df.columns)},
        "columns": list(df.columns),
        "col_types": {c: detect_column_type(df[c]) for c in df.columns},
        "missing": {},
        "duplicates": {},
        "negative_profits": {},
        "salary_issues": {},
        "business_logic_errors": [],
        "statistical_anomalies": [],
        "suspicious_edits": [],
        "kpis": {},
    }

    # Missing values
    missing = df.isnull().sum()
    report["missing"] = {
        col: {"count": int(cnt), "pct": round(cnt / len(df) * 100, 2)}
        for col, cnt in missing.items() if cnt > 0
    }

    # Duplicates
    dup_mask = df.duplicated()
    dup_rows = df[dup_mask]
    report["duplicates"] = {
        "count": int(dup_mask.sum()),
        "rows": dup_rows.index.tolist()[:20]
    }

    # Negative profits
    for col in find_profit_cols(df):
        neg = df[df[col] < 0]
        if not neg.empty:
            report["negative_profits"][col] = {
                "count": len(neg),
                "rows": neg.index.tolist()[:20],
                "min": float(neg[col].min())
            }

    # Salary totals
    for col in find_salary_cols(df):
        s = df[col].dropna()
        report["salary_issues"][col] = {
            "total": float(s.sum()),
            "mean": float(s.mean()),
            "median": float(s.median()),
            "min": float(s.min()),
            "max": float(s.max()),
            "zero_count": int((s == 0).sum()),
            "negative_count": int((s < 0).sum())
        }

    # Business logic
    report["business_logic_errors"] = check_business_logic(df)

    # Statistical anomalies
    report["statistical_anomalies"] = detect_statistical_anomalies(df)

    # Suspicious edits
    report["suspicious_edits"] = detect_suspicious_edits(df)

    # KPIs
    num_df = df.select_dtypes(include=[np.number])
    kpis = {}
    for col in num_df.columns:
        s = num_df[col].dropna()
        if len(s) == 0:
            continue
        kpis[col] = {
            "sum": round(float(s.sum()), 2),
            "mean": round(float(s.mean()), 2),
            "median": round(float(s.median()), 2),
            "std": round(float(s.std()), 2),
            "min": round(float(s.min()), 2),
            "max": round(float(s.max()), 2),
            "count": int(s.count()),
            "null_count": int(num_df[col].isnull().sum())
        }
    report["kpis"] = kpis

    # Severity score
    total_issues = (
        len(report["missing"]) +
        report["duplicates"]["count"] +
        sum(v["count"] for v in report["negative_profits"].values()) +
        len(report["business_logic_errors"]) +
        len(report["statistical_anomalies"]) +
        len(report["suspicious_edits"])
    )
    if total_issues == 0:
        report["health_score"] = 100
    elif total_issues < 5:
        report["health_score"] = 85
    elif total_issues < 15:
        report["health_score"] = 65
    elif total_issues < 30:
        report["health_score"] = 40
    else:
        report["health_score"] = 20

    return report


# ── Excel report builder ─────────────────────────────────────────────────────
def style_header_row(ws, row, cols, fill_hex, font_size=11):
    for col_idx in range(1, cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.fill = _fill(fill_hex)
        cell.font = _font(bold=True, size=font_size, color=C["white"])
        cell.alignment = _center()
        cell.border = _border()


def write_section_title(ws, row, col, text, col_span=6):
    cell = ws.cell(row=row, column=col, value=text)
    cell.fill = _fill(C["bg_header"])
    cell.font = _font(bold=True, size=13, color=C["accent2"])
    cell.alignment = _left()
    ws.merge_cells(start_row=row, start_column=col,
                   end_row=row, end_column=col + col_span - 1)


def build_cover_sheet(wb, audit, df):
    ws = wb.create_sheet("📋 Audit Summary")
    ws.sheet_properties.tabColor = "C0392B"
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18

    # Title block
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 50
    ws.row_dimensions[3].height = 10

    ws.merge_cells("A2:D2")
    title_cell = ws["A2"]
    title_cell.value = "EXCEL AUDIT & REPORTING SYSTEM"
    title_cell.fill = _fill(C["accent"])
    title_cell.font = Font(bold=True, size=22, color=C["white"], name="Calibri")
    title_cell.alignment = _center()

    ws.row_dimensions[4].height = 30
    ws.merge_cells("A4:D4")
    sub = ws["A4"]
    sub.value = f"File: {audit['filename']}   |   Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}"
    sub.fill = _fill(C["bg_dark"])
    sub.font = _font(italic=True, size=11, color=C["text_dim"])
    sub.alignment = _center()

    # Metrics row
    ws.row_dimensions[6].height = 60

    metrics = [
        ("Total Rows", audit["shape"]["rows"], C["accent"]),
        ("Total Columns", audit["shape"]["cols"], C["bg_header"]),
        ("Health Score", f"{audit['health_score']}%",
         C["green"] if audit["health_score"] > 70 else C["orange"] if audit["health_score"] > 40 else C["red"]),
        ("Issues Found",
         sum([len(audit["missing"]), audit["duplicates"]["count"],
              len(audit["business_logic_errors"]), len(audit["statistical_anomalies"])]),
         C["orange"]),
    ]
    for i, (label, val, color) in enumerate(metrics, 1):
        ws.merge_cells(start_row=6, start_column=i, end_row=6, end_column=i)
        c = ws.cell(row=6, column=i)
        c.value = f"{label}\n{val}"
        c.fill = _fill(color)
        c.font = Font(bold=True, size=14, color=C["white"], name="Calibri")
        c.alignment = _center()
        c.border = _border()

    # Issue breakdown
    row = 8
    write_section_title(ws, row, 1, "  ISSUE BREAKDOWN", 4)
    row += 1

    headers = ["Category", "Details", "Severity", "Status"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(C["bg_header"])
        c.font = _font(bold=True, color=C["accent2"])
        c.alignment = _center()
        c.border = _border()
        ws.row_dimensions[row].height = 20
    row += 1

    issues_data = []
    if audit["missing"]:
        issues_data.append(("Missing Values",
                             f"{sum(v['count'] for v in audit['missing'].values())} cells across {len(audit['missing'])} columns",
                             "MEDIUM", "⚠ Review"))
    if audit["duplicates"]["count"] > 0:
        issues_data.append(("Duplicate Rows",
                             f"{audit['duplicates']['count']} duplicate rows detected",
                             "HIGH", "✖ Fix"))
    for col, info in audit["negative_profits"].items():
        issues_data.append(("Negative Profits",
                             f"{info['count']} rows in '{col}' (min: {info['min']:,.2f})",
                             "HIGH", "✖ Fix"))
    for err in audit["business_logic_errors"]:
        issues_data.append(("Business Logic",
                             err["description"],
                             err["severity"].upper(), "✖ Fix"))
    for a in audit["statistical_anomalies"]:
        issues_data.append(("Statistical Anomaly",
                             f"{a['count']} outliers in '{a['column']}'",
                             "MEDIUM", "⚠ Review"))
    for s in audit["suspicious_edits"]:
        issues_data.append(("Suspicious Edit",
                             s["description"],
                             "LOW", "ℹ Note"))

    if not issues_data:
        issues_data.append(("No Issues", "Dataset passed all checks", "NONE", "✔ Clean"))

    sev_color = {"HIGH": C["red"], "MEDIUM": C["orange"], "LOW": C["gold"], "NONE": C["green"]}

    for alt, (cat, det, sev, status) in enumerate(issues_data):
        bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
        row_data = [cat, det, sev, status]
        ws.row_dimensions[row].height = 22
        for ci, val in enumerate(row_data, 1):
            c = ws.cell(row=row, column=ci, value=val)
            if ci == 3:
                c.fill = _fill(sev_color.get(sev, C["bg_dark"]))
                c.font = _font(bold=True, size=10, color=C["white"])
            else:
                c.fill = _fill(bg)
                c.font = _font(size=10)
            c.alignment = _left() if ci == 2 else _center()
            c.border = _border()
        row += 1

    return ws


def build_cleaned_data_sheet(wb, df, audit):
    ws = wb.create_sheet("🧹 Cleaned Data")
    ws.sheet_properties.tabColor = "27AE60"
    ws.freeze_panes = "A2"

    dup_rows = set(audit["duplicates"]["rows"])
    missing_rows = set()
    for col_info in audit["missing"].values():
        pass  # we'll mark cells individually

    # Column widths
    for ci, col in enumerate(df.columns, 1):
        try:
            col_str_len = df[col].fillna("").astype(str).str.len().max()
            col_str_len = int(col_str_len) if pd.notna(col_str_len) else 10
        except Exception:
            col_str_len = 10
        max_len = max(len(str(col)), col_str_len)
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 30)

    # Header
    ws.row_dimensions[1].height = 28
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.fill = _fill(C["accent"])
        c.font = _font(bold=True, size=11, color=C["white"])
        c.alignment = _center()
        c.border = _border()

    # Data rows
    null_fill = _fill("5C1A1A")
    dup_fill = _fill("3D2800")
    norm_fill_a = _fill(C["bg_dark"])
    norm_fill_b = _fill(C["row_alt"])

    for ri, (idx, row_data) in enumerate(df.iterrows(), 2):
        ws.row_dimensions[ri].height = 18
        is_dup = idx in dup_rows
        for ci, (col, val) in enumerate(row_data.items(), 1):
            c = ws.cell(row=ri, column=ci)
            if pd.isna(val):
                c.value = "—"
                c.fill = null_fill
                c.font = _font(italic=True, color=C["text_dim"], size=10)
            else:
                c.value = val
                if is_dup:
                    c.fill = dup_fill
                    c.font = _font(italic=True, color=C["gold"], size=10)
                else:
                    c.fill = norm_fill_a if ri % 2 == 0 else norm_fill_b
                    c.font = _font(size=10)
            c.alignment = _left()
            c.border = _border()

    # Auto-filter
    ws.auto_filter.ref = ws.dimensions

    return ws


def build_kpi_sheet(wb, audit, df):
    ws = wb.create_sheet("📊 KPI Dashboard")
    ws.sheet_properties.tabColor = "D4AC0D"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 18

    row = 1
    ws.row_dimensions[row].height = 40
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "KEY PERFORMANCE INDICATORS"
    c.fill = _fill(C["accent"])
    c.font = Font(bold=True, size=18, color=C["white"], name="Calibri")
    c.alignment = _center()

    row = 2
    ws.row_dimensions[row].height = 10

    row = 3
    headers = ["Column", "Sum", "Mean", "Median", "Std Dev", "Min", "Max"]
    ws.row_dimensions[row].height = 22
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = _fill(C["bg_header"])
        c.font = _font(bold=True, color=C["accent2"])
        c.alignment = _center()
        c.border = _border()

    row = 4
    for alt, (col, kpi) in enumerate(audit["kpis"].items()):
        ws.row_dimensions[row].height = 20
        bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
        vals = [col, kpi["sum"], kpi["mean"], kpi["median"],
                kpi["std"], kpi["min"], kpi["max"]]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.fill = _fill(bg)
            c.font = _font(size=10, bold=(ci == 1))
            c.alignment = _center() if ci > 1 else _left()
            c.border = _border()
            if ci > 1 and isinstance(val, (int, float)):
                c.number_format = '#,##0.00'
        row += 1

    # Salary block
    if audit["salary_issues"]:
        row += 1
        write_section_title(ws, row, 1, "  SALARY ANALYSIS", 7)
        row += 1
        s_headers = ["Column", "Total", "Mean", "Median", "Min", "Max", "Zeros"]
        ws.row_dimensions[row].height = 22
        for ci, h in enumerate(s_headers, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = _fill(C["gold"])
            c.font = _font(bold=True, color="1A0A0A")
            c.alignment = _center()
            c.border = _border()
        row += 1
        for col, info in audit["salary_issues"].items():
            ws.row_dimensions[row].height = 20
            row_vals = [col, info["total"], info["mean"], info["median"],
                        info["min"], info["max"], info["zero_count"]]
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(C["bg_dark"])
                c.font = _font(size=10)
                c.alignment = _center() if ci > 1 else _left()
                c.border = _border()
                if ci > 1 and isinstance(val, (int, float)):
                    c.number_format = '#,##0.00'
            row += 1

    return ws


def build_error_sheet(wb, audit, df):
    ws = wb.create_sheet("❌ Error Analysis")
    ws.sheet_properties.tabColor = "C0392B"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 30

    row = 1
    ws.row_dimensions[1].height = 38
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "ERROR & ANOMALY ANALYSIS"
    c.fill = _fill(C["accent"])
    c.font = Font(bold=True, size=17, color=C["white"], name="Calibri")
    c.alignment = _center()

    def section(title, col_span=5):
        nonlocal row
        row += 1
        ws.row_dimensions[row].height = 26
        write_section_title(ws, row, 1, f"  {title}", col_span)
        row += 1

    # Missing values
    section("MISSING VALUES")
    if audit["missing"]:
        hdrs = ["Column", "Missing Count", "% Missing", "Severity", "Recommendation"]
        ws.row_dimensions[row].height = 20
        for ci, h in enumerate(hdrs, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = _fill(C["bg_header"])
            c.font = _font(bold=True, color=C["accent2"])
            c.alignment = _center()
            c.border = _border()
        row += 1
        for alt, (col, info) in enumerate(audit["missing"].items()):
            ws.row_dimensions[row].height = 18
            sev = "HIGH" if info["pct"] > 20 else "MEDIUM" if info["pct"] > 5 else "LOW"
            rec = "Drop column" if info["pct"] > 50 else "Impute median/mode" if info["pct"] > 10 else "Fill with default"
            bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
            for ci, val in enumerate([col, info["count"], f"{info['pct']}%", sev, rec], 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(bg) if ci != 4 else _fill(
                    C["red"] if sev == "HIGH" else C["orange"] if sev == "MEDIUM" else C["gold"])
                c.font = _font(size=10, bold=(ci == 4))
                c.alignment = _center() if ci != 5 else _left()
                c.border = _border()
            row += 1
    else:
        c = ws.cell(row=row, column=1, value="✔ No missing values detected")
        c.fill = _fill(C["green"])
        c.font = _font(bold=True, color=C["white"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c.alignment = _left()
        row += 1

    # Duplicates
    section("DUPLICATE ROWS")
    cnt = audit["duplicates"]["count"]
    rows_list = audit["duplicates"]["rows"][:10]
    c = ws.cell(row=row, column=1, value=f"Total Duplicates: {cnt}")
    c.fill = _fill(C["red"] if cnt > 0 else C["green"])
    c.font = _font(bold=True, color=C["white"], size=11)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    c.alignment = _left()
    if rows_list:
        c2 = ws.cell(row=row, column=3, value=f"First duplicate rows: {', '.join(map(str, rows_list))}")
        c2.fill = _fill(C["bg_dark"])
        c2.font = _font(size=10, italic=True)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
        c2.alignment = _left()
    row += 1

    # Negative profits
    section("NEGATIVE PROFIT VALUES")
    if audit["negative_profits"]:
        for alt, (col, info) in enumerate(audit["negative_profits"].items()):
            ws.row_dimensions[row].height = 20
            bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
            data = [col, f"{info['count']} rows", f"Min: {info['min']:,.2f}",
                    "HIGH", f"Rows: {', '.join(map(str, info['rows'][:5]))}"]
            for ci, val in enumerate(data, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(C["red"]) if ci == 4 else _fill(bg)
                c.font = _font(size=10, bold=(ci == 4))
                c.alignment = _left()
                c.border = _border()
            row += 1
    else:
        c = ws.cell(row=row, column=1, value="✔ No negative profit values detected")
        c.fill = _fill(C["green"])
        c.font = _font(bold=True, color=C["white"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c.alignment = _left()
        row += 1

    # Business logic
    section("BUSINESS LOGIC ERRORS")
    if audit["business_logic_errors"]:
        for alt, err in enumerate(audit["business_logic_errors"]):
            ws.row_dimensions[row].height = 20
            bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
            sev_c = C["red"] if err["severity"] == "high" else C["orange"]
            data = [err["type"], err["description"], str(err["count"]) + " rows",
                    err["severity"].upper(), str(err["affected_rows"][:5])]
            for ci, val in enumerate(data, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(sev_c) if ci == 4 else _fill(bg)
                c.font = _font(size=10, bold=(ci == 4))
                c.alignment = _left()
                c.border = _border()
            row += 1
    else:
        c = ws.cell(row=row, column=1, value="✔ No business logic errors detected")
        c.fill = _fill(C["green"])
        c.font = _font(bold=True, color=C["white"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c.alignment = _left()
        row += 1

    # Statistical anomalies
    section("STATISTICAL ANOMALIES (Z-Score > 3)")
    if audit["statistical_anomalies"]:
        for alt, a in enumerate(audit["statistical_anomalies"]):
            ws.row_dimensions[row].height = 20
            bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
            vals_str = ", ".join(f"{v:,.2f}" for v in a["values"][:3])
            data = [a["column"], f"{a['count']} outliers", "Z > 3.0", "MEDIUM", vals_str]
            for ci, val in enumerate(data, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(C["orange"]) if ci == 4 else _fill(bg)
                c.font = _font(size=10, bold=(ci == 4))
                c.alignment = _left()
                c.border = _border()
            row += 1
    else:
        c = ws.cell(row=row, column=1, value="✔ No statistical anomalies detected")
        c.fill = _fill(C["green"])
        c.font = _font(bold=True, color=C["white"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c.alignment = _left()
        row += 1

    # Suspicious edits
    section("SUSPICIOUS MANUAL EDITS")
    if audit["suspicious_edits"]:
        for alt, s in enumerate(audit["suspicious_edits"]):
            ws.row_dimensions[row].height = 20
            bg = C["row_alt"] if alt % 2 == 0 else C["bg_dark"]
            data = [s["column"], s["description"], s["type"], "LOW",
                    str(s["rows"][:5])]
            for ci, val in enumerate(data, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.fill = _fill(C["gold"]) if ci == 4 else _fill(bg)
                c.font = _font(size=10, bold=(ci == 4), color="1A0A0A" if ci == 4 else C["text_light"])
                c.alignment = _left()
                c.border = _border()
            row += 1
    else:
        c = ws.cell(row=row, column=1, value="✔ No suspicious edits detected")
        c.fill = _fill(C["green"])
        c.font = _font(bold=True, color=C["white"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        c.alignment = _left()
        row += 1

    return ws


def build_chart_sheet(wb, df, audit):
    ws = wb.create_sheet("📈 Charts & Visuals")
    ws.sheet_properties.tabColor = "D4AC0D"

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "DATA VISUALIZATIONS"
    c.fill = _fill(C["accent"])
    c.font = Font(bold=True, size=18, color=C["white"], name="Calibri")
    c.alignment = _center()
    ws.row_dimensions[1].height = 38

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    chart_row = 3

    charts_added = 0

    # Bar chart: first numeric col or KPI summary
    if len(num_cols) >= 1 and len(df) >= 3:
        col_name = num_cols[0]
        # Write mini data table for chart
        data_start_row = chart_row
        ws.cell(row=data_start_row, column=1, value="Index")
        ws.cell(row=data_start_row, column=2, value=col_name)
        for i, val in enumerate(df[col_name].dropna().head(20).values, 1):
            ws.cell(row=data_start_row + i, column=1, value=i)
            ws.cell(row=data_start_row + i, column=2, value=float(val))

        chart = BarChart()
        chart.type = "col"
        chart.title = f"Distribution — {col_name}"
        chart.style = 10
        chart.y_axis.title = col_name
        chart.x_axis.title = "Record"
        chart.width = 18
        chart.height = 12

        data_ref = Reference(ws, min_col=2, min_row=data_start_row,
                             max_row=data_start_row + min(20, len(df)))
        chart.add_data(data_ref, titles_from_data=True)
        ws.add_chart(chart, f"D{chart_row}")
        chart_row += 22
        charts_added += 1

    # Null distribution pie chart
    if audit["missing"]:
        pie_start = chart_row
        ws.cell(row=pie_start, column=1, value="Column")
        ws.cell(row=pie_start, column=2, value="Missing %")
        for i, (col, info) in enumerate(audit["missing"].items(), 1):
            ws.cell(row=pie_start + i, column=1, value=col)
            ws.cell(row=pie_start + i, column=2, value=info["pct"])

        pie = PieChart()
        pie.title = "Missing Values Distribution"
        pie.style = 10
        pie.width = 16
        pie.height = 12

        data_ref = Reference(ws, min_col=2, min_row=pie_start,
                             max_row=pie_start + len(audit["missing"]))
        labels_ref = Reference(ws, min_col=1, min_row=pie_start + 1,
                               max_row=pie_start + len(audit["missing"]))
        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(labels_ref)
        ws.add_chart(pie, f"D{chart_row}")
        chart_row += 22
        charts_added += 1

    # Line chart: second numeric col if exists
    if len(num_cols) >= 2 and len(df) >= 5:
        col_name = num_cols[1]
        line_start = chart_row
        ws.cell(row=line_start, column=1, value="Index")
        ws.cell(row=line_start, column=2, value=col_name)
        for i, val in enumerate(df[col_name].dropna().head(30).values, 1):
            ws.cell(row=line_start + i, column=1, value=i)
            ws.cell(row=line_start + i, column=2, value=float(val))

        line = LineChart()
        line.title = f"Trend — {col_name}"
        line.style = 10
        line.y_axis.title = col_name
        line.x_axis.title = "Record"
        line.width = 18
        line.height = 12

        data_ref = Reference(ws, min_col=2, min_row=line_start,
                             max_row=line_start + min(30, len(df)))
        line.add_data(data_ref, titles_from_data=True)
        ws.add_chart(line, f"D{chart_row}")
        charts_added += 1

    if charts_added == 0:
        c = ws.cell(row=3, column=1, value="Not enough numeric data to generate charts.")
        c.font = _font(italic=True, color=C["text_dim"])

    return ws


def build_insights_sheet(wb, audit, df):
    ws = wb.create_sheet("💡 Insights & Recommendations")
    ws.sheet_properties.tabColor = "27AE60"
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 65

    ws.row_dimensions[1].height = 38
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "INSIGHTS & RECOMMENDATIONS"
    c.fill = _fill(C["accent"])
    c.font = Font(bold=True, size=18, color=C["white"], name="Calibri")
    c.alignment = _center()

    row = 2
    ws.row_dimensions[row].height = 10

    insights = []

    score = audit["health_score"]
    if score >= 85:
        insights.append(("✔", "Dataset Health", "Excellent — dataset is clean and ready for analysis."))
    elif score >= 65:
        insights.append(("⚠", "Dataset Health", f"Good — minor issues detected. Score: {score}%. Address warnings before reporting."))
    elif score >= 40:
        insights.append(("✖", "Dataset Health", f"Poor — several issues require attention. Score: {score}%. Review all error flags before use."))
    else:
        insights.append(("✖", "Dataset Health", f"Critical — dataset has serious integrity problems. Score: {score}%. Do not use for decisions without thorough cleansing."))

    if audit["duplicates"]["count"] > 0:
        insights.append(("⚠", "Duplicate Rows",
                          f"Remove {audit['duplicates']['count']} duplicate rows. Use pandas drop_duplicates() or Excel 'Remove Duplicates' feature."))

    for col, info in audit["missing"].items():
        if info["pct"] > 50:
            insights.append(("✖", f"Drop '{col}'",
                              f"Column has {info['pct']}% missing values. Consider dropping it entirely."))
        elif info["pct"] > 10:
            insights.append(("⚠", f"Impute '{col}'",
                              f"{info['pct']}% missing. Impute with median (numeric) or mode (categorical)."))
        else:
            insights.append(("ℹ", f"Fill '{col}'",
                              f"Only {info['pct']}% missing. Fill with a sensible default or forward-fill."))

    for col, info in audit["negative_profits"].items():
        insights.append(("✖", f"Negative Profits: '{col}'",
                          f"{info['count']} negative profit entries. Verify if these are returns/refunds or data entry errors."))

    for err in audit["business_logic_errors"]:
        insights.append(("✖", "Logic Error",
                          f"{err['description']}. Recalculate derived columns from source data."))

    if audit["statistical_anomalies"]:
        cols = [a["column"] for a in audit["statistical_anomalies"]]
        insights.append(("⚠", "Outliers Detected",
                          f"Columns {', '.join(cols)} contain statistical outliers (Z > 3). Verify these are real data points, not input errors."))

    if audit["suspicious_edits"]:
        insights.append(("ℹ", "Manual Edit Flags",
                          f"{len(audit['suspicious_edits'])} columns contain suspiciously round numbers that may indicate manual overrides. Verify against source systems."))

    if not insights:
        insights.append(("✔", "All Clear", "Dataset passed all automated checks successfully."))

    icon_color = {"✔": C["green"], "⚠": C["orange"], "✖": C["red"], "ℹ": C["gold"]}

    row = 3
    for alt, (icon, title, text) in enumerate(insights):
        ws.row_dimensions[row].height = 36
        c1 = ws.cell(row=row, column=1, value=icon)
        c1.fill = _fill(icon_color.get(icon, C["bg_dark"]))
        c1.font = Font(bold=True, size=14, color=C["white"], name="Calibri")
        c1.alignment = _center()
        c1.border = _border()

        c2 = ws.cell(row=row, column=2, value=title)
        c2.fill = _fill(C["row_alt"] if alt % 2 == 0 else C["bg_dark"])
        c2.font = _font(bold=True, size=11)
        c2.alignment = _left()
        c2.border = _border()

        c3 = ws.cell(row=row, column=3, value=text)
        c3.fill = _fill(C["row_alt"] if alt % 2 == 0 else C["bg_dark"])
        c3.font = _font(size=10)
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c3.border = _border()

        row += 1

    return ws


def generate_excel_report(df: pd.DataFrame, audit: dict) -> io.BytesIO:
    wb = openpyxl.Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    build_cover_sheet(wb, audit, df)
    build_cleaned_data_sheet(wb, df, audit)
    build_kpi_sheet(wb, audit, df)
    build_error_sheet(wb, audit, df)
    build_chart_sheet(wb, df, audit)
    build_insights_sheet(wb, audit, df)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── routes ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    with open("index.html", "rb") as f:
        html_bytes = f.read()
    return html_bytes, 200, {"Content-Type": "text/html; charset=utf-8"}


@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["file"]
        filename = file.filename

        file_bytes = io.BytesIO(file.read())

        if filename.lower().endswith(".csv"):
            # Try multiple encodings for CSV
            for enc in ("utf-8", "latin-1", "cp1252", "iso-8859-1"):
                try:
                    file_bytes.seek(0)
                    df = pd.read_csv(file_bytes, encoding=enc)
                    break
                except (UnicodeDecodeError, Exception):
                    continue
            else:
                return jsonify({"error": "Could not decode CSV file. Try saving as UTF-8."}), 400
        elif filename.lower().endswith((".xlsx", ".xls")):
            try:
                file_bytes.seek(0)
                engine = "openpyxl" if filename.lower().endswith(".xlsx") else "xlrd"
                df = pd.read_excel(file_bytes, engine=engine)
            except Exception as xe:
                return jsonify({"error": f"Excel read error: {str(xe)}"}), 400
        else:
            return jsonify({"error": "Only CSV, XLSX, or XLS files are supported."}), 400

        if df.empty:
            return jsonify({"error": "Uploaded file is empty."}), 400

        audit = full_audit(df, filename)

        # Build summary for frontend
        summary = {
            "filename": filename,
            "rows": audit["shape"]["rows"],
            "cols": audit["shape"]["cols"],
            "health_score": audit["health_score"],
            "missing_count": sum(v["count"] for v in audit["missing"].values()),
            "missing_cols": len(audit["missing"]),
            "duplicate_count": audit["duplicates"]["count"],
            "negative_profit_count": sum(v["count"] for v in audit["negative_profits"].values()),
            "logic_errors": len(audit["business_logic_errors"]),
            "anomalies": len(audit["statistical_anomalies"]),
            "suspicious": len(audit["suspicious_edits"]),
            "columns": audit["columns"][:20],
            "col_types": dict(list(audit["col_types"].items())[:20]),
        }

        # Store df and audit in a temp file keyed by hash
        key = hashlib.md5(f"{filename}{datetime.now().isoformat()}".encode()).hexdigest()[:12]
        os.makedirs("/tmp/audit_cache", exist_ok=True)
        df.to_parquet(f"/tmp/audit_cache/{key}_df.parquet", index=True)
        with open(f"/tmp/audit_cache/{key}_audit.json", "w") as f_out:
            json.dump(audit, f_out, default=str)

        summary["report_key"] = key
        return jsonify(summary)

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/download/<key>", methods=["GET"])
def download(key):
    try:
        # Sanitize key
        if not re.match(r'^[a-f0-9]{12}$', key):
            return jsonify({"error": "Invalid key"}), 400

        df_path = f"/tmp/audit_cache/{key}_df.parquet"
        audit_path = f"/tmp/audit_cache/{key}_audit.json"

        if not os.path.exists(df_path) or not os.path.exists(audit_path):
            return jsonify({"error": "Report expired. Please re-upload your file."}), 404

        df = pd.read_parquet(df_path)
        with open(audit_path) as f:
            audit = json.load(f)

        buf = generate_excel_report(df, audit)
        fname = f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=fname
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

