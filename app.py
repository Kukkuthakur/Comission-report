"""
Commission Report Builder — Streamlit app
Handles real .xls/.xlsx, HTML tables masquerading as .xls, and Google Drive / Sheets links.
"""
import io
import re
import urllib.request
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_PERIOD = "September 2026"
ROWS_PER_PAGE  = 38
BLOCK_OVERHEAD = 6
XL_THRESHOLD   = 20

FONT_TITLE   = Font(name="Calibri", size=14, bold=True)
FONT_REF     = Font(name="Calibri", size=12, bold=True)
FONT_PERIOD  = Font(name="Calibri", size=10, italic=True)
FONT_HEADER  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_BODY    = Font(name="Calibri", size=10)
FONT_TOTAL   = Font(name="Calibri", size=10, bold=True)
FONT_THANKS  = Font(name="Calibri", size=9, italic=True)

FILL_HEADER  = PatternFill("solid", fgColor="305496")
FILL_TOTAL   = PatternFill("solid", fgColor="D9E1F2")
FILL_REF     = PatternFill("solid", fgColor="FCE4D6")

THIN   = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")

HEADERS    = ["S.No", "Date", "Patient Name", "Investigation Done",
              "Investigation Charge", "Ambulance", "Discount",
              "Percent Cut", "Rate"]
COL_WIDTHS = [6, 12, 18, 30, 12, 11, 10, 11, 10]


def is_html_bytes(b):
    head = b[:1024].lstrip().lower()
    return head.startswith(b"<") or b"<html" in head or b"<table" in head


def clean_money(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def fmt_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return pd.to_datetime(v, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return str(v)


def compute_rate(cut, disc, amb):
    a = amb if amb == 100 else 0
    return max(cut - disc - a, 0.0)


def load_source(raw_bytes, filename="file.xls"):
    if isinstance(raw_bytes, str):
        raw_bytes = raw_bytes.encode("utf-8", errors="ignore")

    if is_html_bytes(raw_bytes):
        try:
            tables = pd.read_html(io.BytesIO(raw_bytes))
        except Exception as e:
            raise ValueError(f"Could not parse HTML table: {e}")
        if not tables:
            raise ValueError("No <table> found in the HTML file.")
        df = max(tables, key=lambda t: t.shape[0] * t.shape[1])
    else:
        engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
        try:
            df = pd.read_excel(io.BytesIO(raw_bytes), engine=engine)
        except Exception:
            alt = "openpyxl" if engine == "xlrd" else "xlrd"
            df = pd.read_excel(io.BytesIO(raw_bytes), engine=alt)

    df.columns = [str(c).strip() for c in df.columns]

    # ---- column alias mapping ----
    # App expects canonical names; source file uses different labels.
    # Order matters: the first existing alias wins.
    aliases = {
        "PatientName": ["PatientName", "Patient Name"],
        "BillDate":    ["BillDate", "Bill Date", "Date"],
        "ReferBy":     ["ReferBy", "Refer By", "Referrer"],
        "TestName":    ["TestName", "Test Name", "Investigation"],
        "PatientRate": ["PatientRate", "Patient Rate", "Charge",
                        "Investigation Charge"],
        "DiscPercent": ["DiscPercent", "Disc", "Discount"],
        "CutRate":     ["CutRate", "PercentCut", "Percent Cut"],
        "Ambulance":   ["Ambulance", "Ambulance Charge"],
    }
    rename = {}
    for canonical, options in aliases.items():
        for opt in options:
            if opt in df.columns and opt != canonical:
                rename[opt] = canonical
                break
    if rename:
        df = df.rename(columns=rename)

    required = ["PatientName", "BillDate", "ReferBy", "TestName",
                "PatientRate", "DiscPercent", "CutRate", "Ambulance"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}\nFound: {list(df.columns)}")

    # Drop source-provided Rate/Profit so we don't collide with our computed Rate
    for c in ["Rate", "Profit"]:
        if c in df.columns:
            df = df.drop(columns=[c])

    for c in ["PatientRate", "DiscPercent", "CutRate", "Ambulance"]:
        df[c] = df[c].apply(clean_money)

    df["PatientName"] = df["PatientName"].astype(str).str.strip()
    df["ReferBy"]     = df["ReferBy"].astype(str).str.strip()
    df["TestName"]    = df["TestName"].astype(str).str.strip()
    df["BillDate"]    = df["BillDate"].apply(fmt_date)

    df = df[df["ReferBy"].notna() & (df["ReferBy"] != "") & (df["ReferBy"] != "nan")]

    df["Rate"] = df.apply(
        lambda r: compute_rate(r["CutRate"], r["DiscPercent"], r["Ambulance"]),
        axis=1,
    )
    return df.reset_index(drop=True)


def fetch_from_gdrive(url_or_id):
    s = url_or_id.strip()
    if not s:
        raise ValueError("Empty link.")

    fid = None
    for pat in (r"/file/d/([A-Za-z0-9_-]+)",
                r"/spreadsheets/d/([A-Za-z0-9_-]+)",
                r"[?&]id=([A-Za-z0-9_-]+)"):
        m = re.search(pat, s)
        if m:
            fid = m.group(1)
            break
    if not fid:
        fid = s

    if "docs.google.com/spreadsheets" in s or "/spreadsheets/" in s:
        export_url = f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx"
    else:
        export_url = f"https://drive.google.com/uc?export=download&id={fid}"

    req = urllib.request.Request(export_url,
                                 headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
    except Exception as e:
        raise ValueError(f"Fetch failed: {e}")

    if not data or len(data) < 100:
        raise ValueError("Downloaded file is empty or too small. "
                         "Make sure the file is shared as 'Anyone with the link'.")
    return data


def pack_referrers(ref_groups):
    sheets = []
    small, medium, large, xl = [], [], [], []

    for name, g in ref_groups:
        n = len(g)
        if n >= XL_THRESHOLD:
            xl.append((name, g))
        elif n >= 10:
            large.append((name, g))
        elif n >= 3:
            medium.append((name, g))
        else:
            small.append((name, g))

    for i, (name, g) in enumerate(xl, 1):
        sheets.append((f"XL{i}", [(name, g)]))

    def fill(bucket, cap, prefix):
        out, cur, used = [], [], 0
        for name, g in bucket:
            need = BLOCK_OVERHEAD + len(g)
            if used + need > cap and cur:
                out.append(cur)
                cur, used = [], 0
            cur.append((name, g))
            used += need
        if cur:
            out.append(cur)
        for i, grp in enumerate(out, 1):
            sheets.append((f"{prefix}{i}", grp))

    fill(large,  ROWS_PER_PAGE, "L")
    fill(medium, ROWS_PER_PAGE, "M")
    fill(small,  ROWS_PER_PAGE, "P")
    return sheets


def write_index(ws, ref_groups, sheet_map, period_text):
    ws["A1"] = f"Commission Report Index — {period_text}"
    ws["A1"].font = FONT_TITLE
    ws.merge_cells("A1:D1")

    for i, h in enumerate(["Referrer", "Line Items", "Total Rate (₹)", "Go to Statement"], 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = CENTER
        c.border = BORDER

    row = 4
    for name, g in ref_groups:
        sn, fr = sheet_map[name]
        ws.cell(row=row, column=1, value=name).border = BORDER
        ws.cell(row=row, column=2, value=len(g)).border = BORDER
        c = ws.cell(row=row, column=3, value=float(g["Rate"].sum()))
        c.border = BORDER
        c.number_format = "#,##0"
        link = ws.cell(row=row, column=4, value="Go ►")
        link.hyperlink = f"#'{sn}'!A{fr}"
        link.font = Font(color="0563C1", underline="single")
        link.border = BORDER
        link.alignment = CENTER
        row += 1

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.freeze_panes = "A4"


def write_block(ws, start_row, referrer, df, period_text):
    r = start_row

    ws.cell(row=r, column=1, value=referrer)
    ws.cell(row=r, column=1).font = FONT_REF
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    for c in range(1, 10):
        ws.cell(row=r, column=c).fill = FILL_REF
        ws.cell(row=r, column=c).border = BORDER
    r += 1

    ws.cell(row=r, column=1, value=period_text)
    ws.cell(row=r, column=1).font = FONT_PERIOD
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    r += 1

    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = CENTER
        c.border = BORDER
    r += 1

    for i, (_, row) in enumerate(df.iterrows(), 1):
        amb = row["Ambulance"]
        amb_disp = amb if amb == 100 else ""
        vals = [
            i,
            row["BillDate"],
            row["PatientName"],
            row["TestName"],
            row["PatientRate"],
            amb_disp,
            row["DiscPercent"],
            row["CutRate"],
            row["Rate"],
        ]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = FONT_BODY
            c.border = BORDER
            c.alignment = (
                CENTER if j in (1, 2, 6)
                else RIGHT if j in (5, 7, 8, 9)
                else LEFT
            )
            if j in (5, 6, 7, 8, 9):
                c.number_format = "#,##0"
        r += 1

    totals = [
        "Total", "", "", "",
        float(df["PatientRate"].sum()),
        float(df.loc[df["Ambulance"] == 100, "Ambulance"].sum()),
        float(df["DiscPercent"].sum()),
        float(df["CutRate"].sum()),
        float(df["Rate"].sum()),
    ]
    for j, v in enumerate(totals, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = FONT_TOTAL
        c.fill = FILL_TOTAL
        c.border = BORDER
        c.alignment = CENTER if j == 1 else RIGHT
        if j >= 5:
            c.number_format = "#,##0"
    r += 1

    ws.cell(row=r, column=1, value="Thanks & Regards")
    ws.cell(row=r, column=1).font = FONT_THANKS
    r += 2
    return r, start_row


def build_workbook(df, ref_groups, packed, period_text):
    wb = Workbook()
    wb.remove(wb.active)
    idx = wb.create_sheet("Index")
    sheet_map = {}

    for sn, group in packed:
        ws = wb.create_sheet(sn)
        for i, w in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True

        row = 1
        for referrer, g in group:
            row, fr = write_block(ws, row, referrer, g, period_text)
            sheet_map[referrer] = (sn, fr)

    write_index(idx, ref_groups, sheet_map, period_text)
    return wb


st.set_page_config(page_title="Commission Report Builder",
                   page_icon="📊", layout="centered")

st.title("📊 Commission Report Builder")
st.caption("Load your flat Excel file (.xls / .xlsx / HTML-exported .xls) "
           "from an upload or a Google Drive / Sheets link.")

st.markdown(
    "**Expected columns in the source file:**  \n"
    "`PatientName · BillDate · ReferBy · TestName · PatientRate · "
    "Disc · PercentCut · Ambulance`"
)

if "raw_bytes"   not in st.session_state: st.session_state.raw_bytes   = None
if "source_name" not in st.session_state: st.session_state.source_name = "source.xls"
if "report_buf"  not in st.session_state: st.session_state.report_buf  = None
if "summary"     not in st.session_state: st.session_state.summary     = None

period_text = st.text_input("Report period", value=DEFAULT_PERIOD)

st.subheader("Option 1 — Upload")
uploaded = st.file_uploader("Choose your Excel file",
                            type=["xls", "xlsx", "html", "htm"])
if uploaded is not None:
    st.session_state.raw_bytes   = uploaded.read()
    st.session_state.source_name = uploaded.name
    st.session_state.report_buf  = None

st.subheader("Option 2 — Google Drive / Sheets link")
st.caption("Share the file as 'Anyone with the link', then paste the link here.")
drive_input = st.text_input("Drive / Sheets link or file ID", value="")

if st.button("Fetch from Drive"):
    if not drive_input.strip():
        st.error("Paste a link first.")
    else:
        with st.spinner("Downloading from Google…"):
            try:
                st.session_state.raw_bytes   = fetch_from_gdrive(drive_input)
                st.session_state.source_name = "drive_file.xls"
                st.session_state.report_buf  = None
                st.success("File fetched successfully.")
            except Exception as e:
                st.error(f"Drive fetch failed: {e}")

if st.session_state.raw_bytes is not None:
    st.info(f"Loaded file: {st.session_state.source_name} "
            f"({len(st.session_state.raw_bytes):,} bytes)")
    if st.button("Generate Report", type="primary"):
        with st.spinner("Building report…"):
            try:
                df = load_source(st.session_state.raw_bytes,
                                 st.session_state.source_name)
            except Exception as e:
                st.error(f"Could not read file: {e}")
                st.stop()

            if df.empty:
                st.warning("No rows found after cleaning. Check your file.")
                st.stop()

            ref_groups = sorted(
                df.groupby("ReferBy", sort=False),
                key=lambda x: x[1]["Rate"].sum(),
                reverse=True,
            )
            packed = pack_referrers(ref_groups)

            packed_names = {n for _, group in packed for n, _ in group}
            expected = {n for n, _ in ref_groups}
            if packed_names != expected:
                st.error("Internal error: some referrers were dropped during packing.")
                st.stop()

            wb = build_workbook(df, ref_groups, packed, period_text)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            st.session_state.report_buf = buf.getvalue()
            st.session_state.summary = {
                "rows": int(len(df)),
                "referrers": int(df["ReferBy"].nunique()),
                "total_rate": float(df["Rate"].sum()),
                "top": [
                    (n, len(g), float(g["Rate"].sum()))
                    for n, g in ref_groups[:15]
                ],
            }

if st.session_state.report_buf is not None:
    s = st.session_state.summary
    st.success(f"✅ Loaded **{s['rows']}** rows across **{s['referrers']}** referrers.")
    st.metric("Grand total Rate (₹)", f"{s['total_rate']:,.0f}")

    st.download_button(
        label="⬇️ Download report (.xlsx)",
        data=st.session_state.report_buf,
        file_name="Commission_Report_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Preview referrers (top 15 by total Rate)"):
        preview = pd.DataFrame({
            "Referrer": [r for r, _, _ in s["top"]],
            "Line Items": [n for _, n, _ in s["top"]],
            "Total Rate (₹)": [f"{t:,.0f}" for _, _, t in s["top"]],
        })
        st.dataframe(preview, hide_index=True, use_container_width=True)

st.divider()
st.caption("Rate = PercentCut − Disc − Ambulance (ambulance only when 100), floored at 0.")
