import streamlit as st
import pdfplumber
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from io import BytesIO

st.set_page_config(
    page_title="PDF to Excel",
    page_icon="📄",
    layout="centered",
)

# ── Minimal, clean CSS ────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Remove default top padding */
.block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 680px; }

/* Step indicator */
.steps {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin: 0 0 2rem;
    font-size: 13px;
}
.step { display: flex; align-items: center; gap: 6px; color: #9ca3af; }
.step.active { color: #111827; font-weight: 500; }
.step.done { color: #16a34a; }
.step-dot {
    width: 22px; height: 22px; border-radius: 50%;
    border: 1.5px solid #d1d5db;
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 500; color: #9ca3af;
    flex-shrink: 0;
}
.step.active .step-dot { border-color: #6b7280; color: #111827; background: #f3f4f6; }
.step.done .step-dot { border-color: #16a34a; background: #dcfce7; color: #16a34a; }
.step-line { width: 32px; height: 1px; background: #e5e7eb; }

/* File badge */
.file-badge {
    display: flex; align-items: center; gap: 10px;
    background: #f9fafb;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 1rem;
    font-size: 14px;
    color: #374151;
}
.file-badge .icon { font-size: 18px; }
.file-badge .name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

/* Result box */
.result-box {
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-radius: 10px;
    padding: 1.25rem;
    margin-top: 1rem;
}
.result-box .result-title {
    display: flex; align-items: center; gap: 8px;
    color: #15803d; font-weight: 500; font-size: 15px;
    margin-bottom: 12px;
}

/* Make the primary button full width and clean */
.stDownloadButton > button, .stButton > button {
    width: 100% !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    padding: 0.55rem 1rem !important;
}

/* Section label */
.section-label {
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.05em;
    color: #6b7280;
    text-transform: uppercase;
    margin: 1.5rem 0 0.5rem;
}

/* Hide Streamlit branding */
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_title(t):
    return re.sub(r'[\/*?:\[\]]', '', t).strip()[:31]

def clean_cell(v):
    return "" if v is None else str(v).replace("\n", " ").strip()

def split_assertions(text):
    if not text:
        return []
    parts = re.split(r'\s+and\s+|\s*&\s*|\s*/\s*|,\s*|;\s*', text, flags=re.IGNORECASE)
    return [p.strip() for p in parts if p.strip()]

def extract_header_info(all_tables):
    info = {
        "entity_name": "", "audit_period": "",
        "assessed_name": "", "assessed_designation": "", "assessed_date": "",
        "reviewed_name": "", "reviewed_designation": "", "reviewed_date": "",
    }
    for table in all_tables:
        if not table:
            continue
        for ri, row in enumerate(table):
            rt = [clean_cell(c) for c in row]
            for i, cell in enumerate(rt):
                lc = cell.lower()
                if "name of the entity" in lc:
                    vals = [x for x in rt[i+1:] if x]
                    if vals:
                        info["entity_name"] = vals[0]
                    elif ri + 1 < len(table):
                        nxt = [clean_cell(c) for c in table[ri+1]]
                        v = [x for x in nxt if x]
                        if v:
                            info["entity_name"] = v[0]
                    if not info["entity_name"]:
                        a = re.sub(r'name of the entity', '', cell, flags=re.IGNORECASE).strip(" :-")
                        if a:
                            info["entity_name"] = a
                elif "period of audit" in lc:
                    vals = [x for x in rt[i+1:] if x]
                    if vals:
                        info["audit_period"] = vals[0]
                    elif ri + 1 < len(table):
                        nxt = [clean_cell(c) for c in table[ri+1]]
                        v = [x for x in nxt if x]
                        if v:
                            info["audit_period"] = v[0]
                    if not info["audit_period"]:
                        m = re.search(r'(\d{4}[-/]\d{2,4})', " ".join(rt))
                        if m:
                            info["audit_period"] = m.group(1)
                elif lc in ["name:", "name"]:
                    vals = [x for x in rt[i+1:] if x]
                    if len(vals) >= 1: info["assessed_name"] = vals[0]
                    if len(vals) >= 2: info["reviewed_name"] = vals[1]
                elif lc in ["designation:", "designation"]:
                    vals = [x for x in rt[i+1:] if x]
                    if len(vals) >= 1: info["assessed_designation"] = vals[0]
                    if len(vals) >= 2: info["reviewed_designation"] = vals[1]
                elif lc in ["date:", "date"]:
                    vals = [x for x in rt[i+1:] if x]
                    if len(vals) >= 1: info["assessed_date"] = vals[0]
                    if len(vals) >= 2: info["reviewed_date"] = vals[1]
    return info


def generate_excel(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

    header_info = extract_header_info(all_tables)
    relevant_tables = [t for t in all_tables if t and len(t[0]) >= 5]
    parsed_data = {}
    current_class = None
    junk_keys = [
        "Material", "Risks", "Name:", "Date:", "Designation:",
        "Name of the Entity", "Assessed by", "Reviewed",
        "Reviewed & agreed by", "Period of Audit", "Awp", "AWP",
        "Not significant", "1",
    ]

    for table in relevant_tables:
        for row in table:
            if not row or all(c is None or str(c).strip() == '' for c in row):
                continue
            if row[0] is not None:
                raw = clean_cell(row[0])
                if raw and all(not raw.lower().startswith(j.lower()) for j in junk_keys):
                    current_class = clean_title(raw)
                    if current_class not in parsed_data:
                        parsed_data[current_class] = []
            risk          = clean_cell(row[1]) if len(row) > 1 else ""
            assertion_raw = clean_cell(row[3]) if len(row) > 3 else ""
            ctrl_activity = clean_cell(row[4]) if len(row) > 4 else ""
            assertions = split_assertions(assertion_raw) or [""]
            substantive = ""
            for idx in [11, 10, 9, 8, 7, 6, 5]:
                if len(row) > idx and clean_cell(row[idx]):
                    substantive = clean_cell(row[idx])
                    break
            if current_class:
                parsed_data[current_class].append({
                    "Risk": risk,
                    "Assertions": assertions,
                    "Control Activity": ctrl_activity,
                    "Substantive Testing Procedures": substantive,
                })

    if not parsed_data:
        st.warning("⚠️ No audit tables found. The PDF may use a different format or contain scanned images.")

    # ── Style helpers ─────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)
    TH = Side(style="thin")
    MD = Side(style="medium")

    def tb():
        return Border(left=TH, right=TH, top=TH, bottom=TH)

    def outer_med(ws, r1, r2, c1, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = Border(
                    left=MD if c == c1 else TH,
                    right=MD if c == c2 else TH,
                    top=MD if r == r1 else TH,
                    bottom=MD if r == r2 else TH,
                )

    TITLE  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    BLUE   = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    BHDR   = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    GREY   = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    YELLOW = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    PINK   = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
    GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    WHITE  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    STRIPE = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")

    def sc(cell, bold=False, fill=None, ha="left", va="center",
           wrap=True, sz=10, color="000000"):
        cell.font      = Font(bold=bold, size=sz, color=color, name="Arial")
        cell.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
        cell.border    = tb()
        if fill:
            cell.fill = fill

    for sheet_name, records in parsed_data.items():
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_view.showGridLines = False

        seen_risks = {}
        for entry in records:
            rt = entry["Risk"].strip()
            if not rt:
                continue
            if rt not in seen_risks:
                seen_risks[rt] = []
            for a in entry["Assertions"]:
                if a and a not in seen_risks[rt]:
                    seen_risks[rt].append(a)

        rac = []
        for risk_text, assertions in seen_risks.items():
            for a in (assertions or [""]):
                rac.append((risk_text, a))
        if not rac:
            rac = [("Risk", "Relevant Assertions")]

        DYN_START = 7
        REM_COL   = DYN_START + len(rac)
        LAST_COL  = REM_COL

        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 26
        ws.column_dimensions["D"].width = 28
        ws.column_dimensions["E"].width = 32
        ws.column_dimensions["F"].width = 20
        ws.column_dimensions["G"].width = 26
        ws.column_dimensions["H"].width = 14
        ws.column_dimensions["I"].width = 14
        for ci in range(DYN_START, REM_COL):
            ws.column_dimensions[get_column_letter(ci)].width = 22
        ws.column_dimensions[get_column_letter(REM_COL)].width = 28

        # Row 1 — Title
        ws.row_dimensions[1].height = 28
        ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=LAST_COL)
        ws["B1"] = f"AWP 5.2 Performing Substantive Audit Procedures - {sheet_name}"
        sc(ws["B1"], bold=True, sz=13, ha="center", fill=TITLE, color="FFFFFF")

        # Rows 2-3 — Entity / Period
        for r, lbl, key in [
            (2, "Name of the Entity :", "entity_name"),
            (3, "Period of Audit :", "audit_period"),
        ]:
            ws.row_dimensions[r].height = 22
            ws.cell(r, 2, lbl)
            sc(ws.cell(r, 2), bold=True, fill=BLUE, wrap=False)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=LAST_COL)
            ws.cell(r, 3, header_info[key])
            for c in range(2, LAST_COL + 1):
                ws.cell(r, c).border = tb()
                ws.cell(r, c).font   = Font(size=10, name="Arial")
                if c >= 3:
                    ws.cell(r, c).alignment = Alignment(
                        wrap_text=False, vertical="center", horizontal="left"
                    )
            sc(ws.cell(r, 2), bold=True, fill=BLUE, wrap=False)
        outer_med(ws, 2, 3, 2, LAST_COL)

        ws.row_dimensions[4].height = 6

        # Rows 5-8 — Assessed / Reviewed signatures
        for r in range(5, 9):
            ws.row_dimensions[r].height = 22

        for merge, val, ref in [
            ("B5:C5", "Assessed by", "B5"),
            ("D5:E5", "Signature", "D5"),
            ("F5:G5", "Reviewed & agreed by", "F5"),
            ("H5:I5", "Signature", "H5"),
        ]:
            ws.merge_cells(merge)
            sc(ws[ref], bold=True, fill=BHDR, ha="center", color="FFFFFF", sz=10)
            ws[ref] = val

        for r, lbl, akey, rkey in [
            (6, "Name", "assessed_name", "reviewed_name"),
            (7, "Designation", "assessed_designation", "reviewed_designation"),
            (8, "Date", "assessed_date", "reviewed_date"),
        ]:
            ws.cell(r, 2, lbl)
            sc(ws.cell(r, 2), bold=True, fill=GREY, wrap=False)
            ws.cell(r, 3, header_info[akey])
            ws.cell(r, 3).alignment = Alignment(wrap_text=False, vertical="center", horizontal="left")
            ws.cell(r, 3).font      = Font(size=10, name="Arial")
            ws.cell(r, 3).border    = tb()
            ws.cell(r, 6, lbl)
            sc(ws.cell(r, 6), bold=True, fill=GREY, wrap=False)
            ws.cell(r, 7, header_info[rkey])
            ws.cell(r, 7).alignment = Alignment(wrap_text=False, vertical="center", horizontal="left")
            ws.cell(r, 7).font      = Font(size=10, name="Arial")
            ws.cell(r, 7).border    = tb()

        ws.merge_cells(start_row=6, start_column=4, end_row=8, end_column=5)
        ws.merge_cells(start_row=6, start_column=8, end_row=8, end_column=9)

        for r in range(5, 9):
            for c in range(2, 10):
                ws.cell(r, c).border = tb()
        outer_med(ws, 5, 8, 2, 9)

        ws.row_dimensions[9].height = 6

        # Row 10 — Step 1 header
        ws.row_dimensions[10].height = 20
        ws.merge_cells(start_row=10, start_column=2, end_row=10, end_column=LAST_COL)
        ws["B10"] = "STEP 1 : Trace risks, control activity, substantive audit procedures and relevant audit assertions"
        sc(ws["B10"], bold=True, sz=10, fill=PINK, ha="left", color="FFFFFF")

        # Row 11 — Sheet label
        ws.row_dimensions[11].height = 18
        ws.merge_cells(start_row=11, start_column=2, end_row=11, end_column=LAST_COL)
        ws["B11"] = f"Significant COTABD:  {sheet_name}"
        sc(ws["B11"], bold=True, sz=10, fill=BLUE, ha="left")

        # Row 12 — Step 1 column headers
        ws.row_dimensions[12].height = 32
        for ci, hdr in [
            (2, "Risk Description"),
            (3, "Relevant Assertions"),
            (4, "Control Activity"),
            (5, "Substantive Testing Procedures"),
        ]:
            sc(ws.cell(12, ci, hdr), bold=True, fill=BHDR, ha="center", color="FFFFFF", sz=10)
        outer_med(ws, 12, 12, 2, 5)

        # Step 1 data rows
        s1_start = 13
        for i, entry in enumerate(records):
            row_fill = STRIPE if i % 2 == 1 else WHITE
            ws.row_dimensions[s1_start + i].height = 48
            for ci, val in zip(
                [2, 3, 4, 5],
                [
                    entry["Risk"],
                    " / ".join(entry["Assertions"]),
                    entry["Control Activity"],
                    entry["Substantive Testing Procedures"],
                ],
            ):
                c = ws.cell(s1_start + i, ci, val)
                c.font      = Font(size=10, name="Arial")
                c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                c.border    = tb()
                c.fill      = row_fill

        s1_end = s1_start + max(len(records) - 1, 0)
        if records:
            outer_med(ws, s1_start, s1_end, 2, 5)

        # ── Step 2 ────────────────────────────────────────────────────────────
        s2_title = s1_end + 2
        ws.row_dimensions[s2_title - 1].height = 6
        ws.row_dimensions[s2_title].height     = 20
        ws.merge_cells(start_row=s2_title, start_column=2, end_row=s2_title, end_column=LAST_COL)
        ws.cell(s2_title, 2, "STEP 2 : Substantive audit procedures performed")
        sc(ws.cell(s2_title, 2), bold=True, fill=PINK, ha="left", color="FFFFFF", sz=10)

        h1 = s2_title + 1
        h2 = s2_title + 2
        ws.row_dimensions[h1].height = 36
        ws.row_dimensions[h2].height = 36

        for ci, hdr in [
            (2, "Sl\nNo"),
            (3, "Date"),
            (4, "Voucher\nNo."),
            (5, "Voucher\nAmount (Nu.)"),
            (6, "Details"),
            (REM_COL, "Remarks"),
        ]:
            ws.merge_cells(start_row=h1, start_column=ci, end_row=h2, end_column=ci)
            sc(ws.cell(h1, ci, hdr), bold=True, fill=BHDR, ha="center", color="FFFFFF", sz=10)

        risk_groups = {}
        col_cur = DYN_START
        for rt, _ in rac:
            risk_groups.setdefault(rt, []).append(col_cur)
            col_cur += 1

        for rt, cols in risk_groups.items():
            col_s, col_e = cols[0], cols[-1]
            if col_s != col_e:
                ws.merge_cells(start_row=h1, start_column=col_s, end_row=h1, end_column=col_e)
            ws.cell(h1, col_s).value     = rt
            ws.cell(h1, col_s).font      = Font(bold=True, size=10, name="Arial")
            ws.cell(h1, col_s).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.cell(h1, col_s).fill      = YELLOW
            ws.cell(h1, col_s).border    = tb()
            for ci2 in cols:
                ws.cell(h1, ci2).border = tb()

        col_cur = DYN_START
        for _, asrt in rac:
            ws.cell(h2, col_cur).value     = asrt
            ws.cell(h2, col_cur).font      = Font(bold=True, size=10, name="Arial")
            ws.cell(h2, col_cur).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.cell(h2, col_cur).fill      = YELLOW
            ws.cell(h2, col_cur).border    = tb()
            col_cur += 1

        # Data rows (20 blank rows with Yes/No validation)
        dr_start = h2 + 1
        dr_end   = dr_start + 19
        ynv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
        ws.add_data_validation(ynv)

        for i, r in enumerate(range(dr_start, dr_end + 1)):
            rf = STRIPE if i % 2 == 1 else WHITE
            ws.row_dimensions[r].height = 18
            for c in range(2, LAST_COL + 1):
                ws.cell(r, c).border    = tb()
                ws.cell(r, c).fill      = rf
                ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            ws.cell(r, 2, i + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(r, 2).font = Font(size=10, name="Arial")
            ws.cell(r, 2).fill = rf
            for c in range(DYN_START, REM_COL):
                ynv.add(ws.cell(r, c))

        outer_med(ws, s2_title, dr_end, 2, LAST_COL)

        # Overall conclusion row
        conc = dr_end + 3
        ws.row_dimensions[dr_end + 2].height = 6
        ws.row_dimensions[conc].height = 60
        ws.cell(conc, 2, "Overall Conclusion:")
        ws.cell(conc, 2).font      = Font(bold=True, size=10, name="Arial")
        ws.cell(conc, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.cell(conc, 2).fill      = GREEN
        ws.cell(conc, 2).border    = tb()
        ws.merge_cells(start_row=conc, start_column=3, end_row=conc, end_column=LAST_COL)
        ws.cell(conc, 3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(conc, 3).fill      = WHITE
        for ci in range(3, LAST_COL + 1):
            ws.cell(conc, ci).border = tb()
        outer_med(ws, conc, conc, 2, LAST_COL)

    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    return excel_file


# ── Step indicator ────────────────────────────────────────────────────────────

def step_indicator(current: int):
    steps = ["Upload", "Convert", "Download"]
    html = '<div class="steps">'
    for i, label in enumerate(steps, 1):
        if i < current:
            cls = "step done"
            dot = "✓"
        elif i == current:
            cls = "step active"
            dot = str(i)
        else:
            cls = "step"
            dot = str(i)
        html += f'<div class="{cls}"><div class="step-dot">{dot}</div>{label}</div>'
        if i < len(steps):
            html += '<div class="step-line"></div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ── App layout ────────────────────────────────────────────────────────────────

st.markdown("## 📄 PDF to Excel")
st.caption("Upload an audit report PDF and get a formatted Excel workbook.")

uploaded_pdf = st.file_uploader(
    "Drop your PDF here or click to browse",
    type=["pdf"],
    label_visibility="collapsed",
)

if uploaded_pdf is None:
    step_indicator(1)
    st.info("Upload a PDF file above to get started.", icon="📂")
else:
    # Show file badge
    st.markdown(
        f'<div class="file-badge">'
        f'<span class="icon">📄</span>'
        f'<span class="name">{uploaded_pdf.name}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Output filename
    default_name = re.sub(r'\.pdf$', '', uploaded_pdf.name, flags=re.IGNORECASE)
    file_name_input = st.text_input(
        "Output filename (no extension needed)",
        value=default_name or "FINAL_Audit_Workbook",
        placeholder="e.g. BCTA_Audit_2024",
    )

    step_indicator(2)

    if st.button("Extract to Excel", type="primary", use_container_width=True):
        with st.status("Processing your PDF…", expanded=True) as status:
            st.write("📖 Reading PDF tables…")
            excel_output = generate_excel(uploaded_pdf)
            st.write("🗂️ Building workbook structure…")
            st.write("🎨 Applying formatting…")
            status.update(label="Done!", state="complete", expanded=False)

        safe_name = re.sub(r'[\\/*?:"<>|]', '_', file_name_input).strip() or "FINAL_Audit_Workbook"

        step_indicator(3)

        st.markdown(
            '<div class="result-box">'
            '<div class="result-title">✅ Excel generated successfully</div>',
            unsafe_allow_html=True,
        )
        st.download_button(
            label=f"⬇️ Download {safe_name}.xlsx",
            data=excel_output,
            file_name=f"{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.markdown("</div>", unsafe_allow_html=True)
