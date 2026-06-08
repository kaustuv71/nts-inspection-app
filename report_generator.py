"""NTS Report Generator - Read Excel data and produce branded PDF reports."""
import os, sys, math
from pathlib import Path
from datetime import datetime

def read_workbook(excel_path: str) -> dict:
    """Read the intermediate Excel workbook and return parsed data."""
    from openpyxl import load_workbook
    wb = load_workbook(excel_path, data_only=True)

    data = {
        "general": {},
        "summary": {},
        "aql": {},
        "skus": [],
        "packaging": [],
        "functional_tests": [],
        "factory_review": {},
        "conclusion": "",
    }

    # ── Parse Page 2 (General + Summary + AQL) ──
    ws = wb["Page 2"]
    general_fields = [
        ("Client:", "client"), ("Destination:", "destination"),
        ("Supplier:", "supplier"), ("Factory Address:", "factory_address"),
        ("Product Name:", "product_name"), ("SKU Designs:", "sku_designs"),
        ("Inspection Date:", "inspection_date"), ("Inspection Duration:", "inspection_duration"),
        ("Order Quantity:", "order_quantity"), ("Sampled Quantity:", "sampled_quantity"),
        ("Inspection Location:", "inspection_location"), ("Report ID:", "report_id"),
        ("Inspectors:", "inspectors"),
    ]
    for row in ws.iter_rows(min_row=6, max_row=30, values_only=False):
        label_cell = row[0]
        val_cell = row[2] if len(row) > 2 else None
        if label_cell and label_cell.value:
            label = str(label_cell.value).strip()
            for prefix, key in general_fields:
                if label.startswith(prefix):
                    data["general"][key] = str(val_cell.value or "").strip()
                    break

    # Summary categories (rows after "II. INSPECTION SUMMARY")
    cat_keys = ["A", "B", "C", "D", "E", "F", "G", "H"]
    cat_labels = [
        "A. Quantity", "B. Shape and Size Consistency",
        "C. Workmanship", "D. Product Specification",
        "E. Packing", "F. Marking & Labeling",
        "G. Client Special Requirements", "H. Factory Review",
    ]
    for row in ws.iter_rows(min_row=34, max_row=60, values_only=False):
        label_val = str(row[0].value or "").strip() if row[0] else ""
        for ck, cl in zip(cat_keys, cat_labels):
            if label_val.startswith(cl.replace(".", "")) or label_val == cl:
                status = "passed"
                if row[3] and str(row[3].value).strip().upper() == "FAILED":
                    status = "failed"
                elif row[4] and str(row[4].value).strip().upper() == "FAILED":
                    status = "failed"
                findings = str(row[5].value or "") if len(row) > 5 and row[5] else ""
                data["summary"][ck] = {"status": status, "findings": findings}
                break

    # AQL (rows after summary, ~row 50)
    aql_fields = {
        "Inspection Standard:": "standard",
        "Sampling Plan:": "sampling_plan",
        "Inspection Level:": "level",
    }
    for row in ws.iter_rows(min_row=50, max_row=65, values_only=False):
        label_val = str(row[0].value or "").strip() if row[0] else ""
        for prefix, key in aql_fields.items():
            if label_val.startswith(prefix):
                data["aql"][key] = str(row[3].value or "").strip() if len(row) > 3 else ""
                break
    # critical/major/minor from specific columns
    for row in ws.iter_rows(min_row=50, max_row=65, values_only=False):
        if row[7] and str(row[7].value).strip() == "Critical":
            data["aql"]["critical"] = str(row[7].value or "")
            data["aql"]["major"] = str(row[8].value or "") if len(row) > 8 else ""
            data["aql"]["minor"] = str(row[9].value or "") if len(row) > 9 else ""
            break

    # ── Parse SKU pages ──
    sku_sheets = [s for s in wb.sheetnames if s.startswith("Page ") and s != "Page 2"]
    for sheet_name in sku_sheets:
        ws_sku = wb[sheet_name]
        sku = {"name": "", "fields": {}}
        for row in ws_sku.iter_rows(min_row=6, max_row=40, values_only=False):
            label = str(row[0].value or "").strip() if row[0] else ""
            val = str(row[1].value or "").strip() if len(row) > 1 and row[1] else ""
            if label.startswith("III."):
                sku["name"] = label.split(":")[-1].strip() if ":" in label else ""
            elif label:
                sku["fields"][label] = val
        if sku["name"] or sku["fields"]:
            data["skus"].append(sku)

    # ── Parse Last page (Packaging + Tests + Factory + Conclusion) ──
    last_sheet = [s for s in wb.sheetnames if s.startswith(f"Page {len(sku_sheets) + 3}")]
    if last_sheet:
        ws_last = wb[last_sheet[0]]
        section = None
        for row in ws_last.iter_rows(min_row=6, max_row=200, values_only=False):
            label = str(row[0].value or "").strip() if row[0] else ""
            val3 = str(row[2].value or "").strip() if len(row) > 2 and row[2] else ""
            val4 = str(row[3].value or "").strip() if len(row) > 3 and row[3] else ""

            if "PACKAGING" in label.upper():
                section = "packaging"
                continue
            elif "FUNCTIONAL" in label.upper():
                section = "functional"
                continue
            elif "MEDIA" in label.upper() and "LINK" in label.upper():
                section = "media"
                continue
            elif "FACTORY" in label.upper() or "FACTOR" in label.upper():
                section = "factory"
                continue
            elif "CONCLUSION" in label.upper():
                section = "conclusion"
                continue

            if section == "packaging" and label and label not in ("Item",):
                data["packaging"].append({
                    "item": label,
                    "findings": val3,
                    "status": val4,
                })
            elif section == "functional" and label and label not in ("Test/Check",):
                data["functional_tests"].append({
                    "test": label,
                    "performed": str(row[1].value or "").strip() if len(row) > 1 and row[1] else "",
                    "result": val3,
                    "status": val4,
                })
            elif section == "media" and label and label not in ("Description",):
                data.setdefault("media_links", []).append(val3)
            elif section == "factory":
                if "cooperation" in label.lower():
                    data["factory_review"]["cooperation"] = val3
                elif "workers" in label.lower():
                    data["factory_review"]["workers"] = val3
                elif "opinion" in label.lower():
                    data["factory_review"]["opinion"] = val3
            elif section == "conclusion" and label:
                data["conclusion"] = label

    return data


def build_report(excel_path: str, photo_dir: str, photo_specs: dict | None = None) -> str:
    """Build a professional PDF report from the Excel workbook and photos.

    photo_specs: dict mapping photo path -> caption string (for defect descriptions in red).
    Returns the path to the generated PDF.
    """
    from fpdf import FPDF

    data = read_workbook(excel_path)
    g = data.get("general", {})
    insp_id = Path(excel_path).stem.replace("_data", "")
    output_dir = Path(excel_path).parent
    pdf_path = output_dir / f"NTS_Report_{insp_id}.pdf"

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── Color Constants ──
    NAVY = (15, 30, 55)
    GOLD = (198, 156, 75)
    WHITE = (255, 255, 255)
    LIGHT_GRAY = (245, 243, 238)
    DARK_TEXT = (40, 40, 40)

    def header_block():
        pdf.set_fill_color(*NAVY)
        pdf.rect(0, 0, 210, 30, "F")
        # Logo on the left
        logo_path = Path(__file__).parent / "nts_logo.png"
        if logo_path.exists():
            pdf.image(str(logo_path), x=8, y=3, w=24)
        # Report ID on the right
        rid = g.get("report_id", "")
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_xy(160, 5)
        pdf.cell(40, 5, rid, ln=True, align="R")
        # Company name
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_xy(38, 5)
        pdf.cell(115, 5, "NEPAL TRADE SOLUTIONS", ln=True, align="L")
        # Tagline
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(*WHITE)
        pdf.set_xy(38, 12)
        pdf.cell(115, 4, "Quality Inspection Report | Mandev Marg, Byasi-10, Bhaktapur, Nepal", ln=True, align="L")
        pdf.set_text_color(*WHITE)
        pdf.set_font("Helvetica", "I", 6)
        pdf.set_xy(38, 18)
        pdf.cell(115, 4, "www.nepalts.com", ln=True, align="L")

    def footer_block():
        pdf.set_fill_color(*NAVY)
        pdf.rect(0, pdf.h - 12, 210, 12, "F")
        pdf.set_y(pdf.h - 10)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(0, 5, f"NTS Report | Page {pdf.page_no()}", align="C")

    # ── Cover Page ──
    pdf.add_page()
    header_block()
    footer_block()

    # Large logo on cover
    logo_path = Path(__file__).parent / "nts_logo.png"
    if logo_path.exists():
        pdf.ln(28)
        pdf.image(str(logo_path), x=55, y=pdf.get_y(), w=100)
        pdf.ln(42)
    else:
        pdf.ln(50)

    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 28)
    pdf.cell(0, 15, "INSPECTION REPORT", ln=True, align="C")

    # Report ID prominent below title
    report_id = g.get("report_id", "")
    if report_id:
        pdf.ln(4)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 10, report_id, ln=True, align="C")

    pdf.set_text_color(*DARK_TEXT)
    pdf.set_font("Helvetica", "", 11)
    pdf.ln(6)
    cover_fields = [
        ("Client", g.get("client", "-")),
        ("Product", g.get("product_name", "-")),
        ("Supplier", g.get("supplier", "-")),
        ("Inspection Date", g.get("inspection_date", "-")),
        ("Inspectors", g.get("inspectors", "-")),
    ]
    for label, val in cover_fields:
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(35, 7, label + ":", align="R")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 7, "  " + val, ln=True)

    pdf.ln(10)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "NEPAL TRADE SOLUTIONS", ln=True, align="C")
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*DARK_TEXT)
    pdf.cell(0, 5, "Mandev Marg, Byasi-10, Bhaktapur, Nepal", ln=True, align="C")
    pdf.ln(4)
    pdf.set_fill_color(*GOLD)
    pdf.rect(60, pdf.get_y(), 90, 0.5, "F")
    pdf.ln(5)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "I", 8)
    pdf.cell(0, 5, "This report is prepared solely for the use of the client named above", ln=True, align="C")
    pdf.cell(0, 5, "and contains confidential information.", ln=True, align="C")

    # ── Page 2: General Information ──
    pdf.add_page()
    header_block()
    footer_block()

    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.ln(20)
    pdf.cell(0, 10, "I. GENERAL INFORMATION", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(4)

    rows = [
        ("Client:", g.get("client", "")),
        ("Destination:", g.get("destination", "")),
        ("Supplier:", g.get("supplier", "")),
        ("Factory Address:", g.get("factory_address", "")),
        ("Product Name:", g.get("product_name", "")),
        ("SKU Designs:", g.get("sku_designs", "")),
        ("Inspection Date:", g.get("inspection_date", "")),
        ("Duration:", g.get("inspection_duration", "")),
        ("Order Quantity:", g.get("order_quantity", "")),
        ("Sampled Quantity:", g.get("sampled_quantity", "")),
        ("Location:", g.get("inspection_location", "")),
        ("Report ID:", g.get("report_id", "")),
        ("Inspectors:", g.get("inspectors", "")),
    ]
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*DARK_TEXT)
    for label, val in rows:
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(45, 6, label)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 6, val, ln=True)
        if label == "Factory Address:":
            pdf.ln(1)

    # ── Page 3: Summary ──
    pdf.add_page()
    header_block()
    footer_block()

    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.ln(20)
    pdf.cell(0, 10, "II. INSPECTION SUMMARY", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)

    # Summary table
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*WHITE)
    pdf.cell(55, 7, "Category", 1, 0, "C", True)
    pdf.cell(12, 7, "Status", 1, 0, "C", True)
    pdf.cell(0, 7, "Notes / Findings Summary", 1, 1, "C", True)

    cat_titles = {
        "A": "A. Quantity",
        "B": "B. Shape & Size Consistency",
        "C": "C. Workmanship",
        "D": "D. Product Specification",
        "E": "E. Packing",
        "F": "F. Marking & Labeling",
        "G": "G. Client Requirements",
        "H": "H. Factory Review",
    }
    pdf.set_text_color(*DARK_TEXT)
    for ck, cl in cat_titles.items():
        s = data.get("summary", {}).get(ck, {})
        status = s.get("status", "pending").upper()
        findings = s.get("findings", "")

        pdf.set_font("Helvetica", "", 8)

        # Calculate height needed
        lines_needed = max(1, len(str(findings)) // 70 + 1)
        row_h = max(6, lines_needed * 4)

        y_before = pdf.get_y()
        if y_before + row_h > 270:
            pdf.add_page()
            header_block()
            footer_block()
            pdf.set_fill_color(*NAVY)
            pdf.set_text_color(*NAVY)
            pdf.set_font("Helvetica", "B", 14)
            pdf.ln(20)
            pdf.cell(0, 10, "II. INSPECTION SUMMARY (cont.)", ln=True)
            pdf.set_draw_color(*GOLD)
            pdf.line(15, pdf.get_y(), 195, pdf.get_y())
            pdf.ln(4)
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(*NAVY)
            pdf.set_text_color(*WHITE)
            pdf.cell(55, 7, "Category", 1, 0, "C", True)
            pdf.cell(12, 7, "Status", 1, 0, "C", True)
            pdf.cell(0, 7, "Notes / Findings Summary", 1, 1, "C", True)
            pdf.set_text_color(*DARK_TEXT)
            y_before = pdf.get_y()

        pdf.set_font("Helvetica", "", 8)
        pdf.cell(55, row_h, cl, 1, 0, "L")
        pdf.set_font("Helvetica", "B", 8)
        status_color = (21, 87, 36) if status == "PASSED" else (114, 28, 36)
        pdf.set_text_color(*status_color)
        pdf.cell(12, row_h, "PASSED" if status == "PASSED" else "FAILED", 1, 0, "C")
        pdf.set_text_color(*DARK_TEXT)
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(0, row_h, findings, 1, 1, "L")

    # AQL Section
    aql = data.get("aql", {})
    pdf.ln(8)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "AQL Details (Based on the inspection of the finished products)", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(4)
    pdf.set_text_color(*DARK_TEXT)
    pdf.set_font("Helvetica", "", 9)
    aql_rows = [
        ("Inspection Standard:", aql.get("standard", "ANSI/ASQ Z1.4-2008"), ""),
        ("Sampling Plan:", aql.get("sampling_plan", "Normal, Single"), ""),
        ("Inspection Level:", aql.get("level", "II"), ""),
        ("", "", f"Critical:  {aql.get('critical', 'Not Allowed')}    Major:  {aql.get('major', '2.5')}    Minor:  {aql.get('minor', '4.5')}"),
    ]
    for label, val, extra in aql_rows:
        if label:
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(35, 6, label)
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(0, 6, val, ln=True)
        else:
            pdf.cell(0, 6, extra, ln=True)
    # Extra AQL rows from general data
    order_qty = g.get("order_quantity", "")
    sample_qty = g.get("sampled_quantity", "")
    for label, val in [("Order Quantity:", order_qty), ("Available Quantity:", order_qty),
                       ("Sample Size:", sample_qty)]:
        if val:
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(35, 6, label)
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(0, 6, val, ln=True)

    # ── SKU Pages ──
    for idx, sku in enumerate(data.get("skus", [])):
        pdf.add_page()
        header_block()
        footer_block()

        sku_name = sku.get("name", f"SKU {idx + 1}")
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 14)
        pdf.ln(20)
        pdf.cell(0, 10, f"III. SKU INSPECTION", ln=True)
        pdf.set_draw_color(*GOLD)
        pdf.line(15, pdf.get_y(), 195, pdf.get_y())
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(*GOLD)
        pdf.cell(0, 8, sku_name, ln=True)
        pdf.set_text_color(*DARK_TEXT)
        pdf.ln(2)

        fields = sku.get("fields", {})
        field_order = [
            "Ordered Quantity", "Found in Warehouse", "Missing Quantity",
            "Sampled Units", "FNSKU", "Defects/Findings",
            "Reference Categories",
        ]
        for key in field_order:
            val = fields.get(key, "")
            if val:
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(50, 6, key + ":")
                pdf.set_font("Helvetica", "", 9)
                pdf.cell(0, 6, val, ln=True)

        # Measurements
        meas_keys = [k for k in fields if k not in field_order and k != "SKU Name"]
        if meas_keys:
            pdf.ln(3)
            pdf.set_fill_color(*NAVY)
            pdf.set_text_color(*WHITE)
            pdf.set_font("Helvetica", "B", 9)
            for mk in meas_keys:
                pdf.cell(55, 6, "Measurement", 1, 0, "C", True)
                pdf.cell(0, 6, "Value", 1, 1, "C", True)
                pdf.set_text_color(*DARK_TEXT)
                pdf.set_font("Helvetica", "", 9)
                pdf.cell(55, 6, mk, 1)
                pdf.cell(0, 6, fields.get(mk, ""), 1, 1)
                pdf.ln(1)
                pdf.set_text_color(*WHITE)
                pdf.set_font("Helvetica", "B", 9)
            pdf.set_text_color(*DARK_TEXT)

        # Photos for this SKU
        photos = []
        if photo_dir:
            photo_path = Path(photo_dir)
            if photo_path.exists():
                # Match photos by naming convention insp_id_skuIdx_...
                photos = sorted(photo_path.glob(f"{insp_id}_{idx}_*.jpg"))
                photos += sorted(photo_path.glob(f"{insp_id}_{idx}_*.jpeg"))

        pdf.ln(2)
        if photos:
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(0, 6, "Photos:", ln=True)
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            img_w = 55
            img_h = 40
            gap = 3
            max_per_row = 3
            for pi, photo_file in enumerate(photos):
                col = pi % max_per_row
                row = pi // max_per_row
                x = x_start + col * (img_w + gap)
                y = y_start + row * (img_h + gap)
                cap_h = 6
                if y + img_h + cap_h > 270:
                    break
                try:
                    pdf.image(str(photo_file), x=x, y=y, w=img_w, h=img_h)
                    # Caption in red below photo
                    fname = Path(photo_file).name
                    if photo_specs:
                        for pkey, cap in photo_specs.items():
                            if fname in pkey or pkey.endswith(fname):
                                if cap:
                                    pdf.set_font("Helvetica", "I", 6)
                                    pdf.set_text_color(200, 30, 30)
                                    pdf.set_xy(x, y + img_h + 1)
                                    pdf.multi_cell(img_w, 3, cap[:80])
                                    pdf.set_text_color(*DARK_TEXT)
                                break
                except:
                    pass
        else:
            pdf.set_font("Helvetica", "I", 8)
            pdf.set_text_color(*GOLD)
            pdf.cell(0, 6, "No photos uploaded for this SKU.", ln=True)
            pdf.set_text_color(*DARK_TEXT)

    # ── Last Pages: Packaging + Tests + Factory + Conclusion ──
    pdf.add_page()
    header_block()
    footer_block()

    # IV. Packaging
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.ln(20)
    pdf.cell(0, 10, "IV. PACKAGING, MARKING & LABELING", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)

    if data.get("packaging"):
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*WHITE)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(55, 7, "Item", 1, 0, "C", True)
        pdf.cell(80, 7, "Findings", 1, 0, "C", True)
        pdf.cell(0, 7, "Status", 1, 1, "C", True)
        pdf.set_text_color(*DARK_TEXT)

        for pkg in data["packaging"]:
            pdf.set_font("Helvetica", "", 8)
            findings = pkg.get("findings", "")
            lines = max(1, len(findings) // 45 + 1)
            rh = max(6, lines * 4)
            pdf.cell(55, rh, pkg.get("item", ""), 1, 0, "L")
            pdf.cell(80, rh, findings, 1, 0, "L")
            status = pkg.get("status", "").upper()
            sc = (21, 87, 36) if status in ("PASS", "PASSED") else (114, 28, 36)
            pdf.set_text_color(*sc)
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(0, rh, status, 1, 1, "C")
            pdf.set_text_color(*DARK_TEXT)
    else:
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(0, 6, "No packaging data recorded.", ln=True)

    # V. Functional Tests
    pdf.ln(8)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "V. FUNCTIONAL TESTS", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)

    if data.get("functional_tests"):
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*WHITE)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(40, 7, "Test", 1, 0, "C", True)
        pdf.cell(15, 7, "Done", 1, 0, "C", True)
        pdf.cell(90, 7, "Result", 1, 0, "C", True)
        pdf.cell(0, 7, "Status", 1, 1, "C", True)
        pdf.set_text_color(*DARK_TEXT)

        for ft in data["functional_tests"]:
            pdf.set_font("Helvetica", "", 8)
            result = ft.get("result", "")
            lines = max(1, len(result) // 50 + 1)
            rh = max(6, lines * 4)
            pdf.cell(40, rh, ft.get("test", ""), 1, 0, "L")
            pdf.cell(15, rh, ft.get("performed", ""), 1, 0, "C")
            pdf.cell(90, rh, result, 1, 0, "L")
            status = ft.get("status", "").upper()
            sc = (21, 87, 36) if status in ("PASS", "PASSED") else (114, 28, 36)
            pdf.set_text_color(*sc)
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(0, rh, status, 1, 1, "C")
            pdf.set_text_color(*DARK_TEXT)
    else:
        pdf.set_font("Helvetica", "I", 8)
        pdf.cell(0, 6, "No functional test data recorded.", ln=True)

    # VI. Media Links
    media_links = data.get("media_links", [])
    if media_links and any(media_links):
        pdf.ln(6)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, "VI. MEDIA & ATTACHMENTS", ln=True)
        pdf.set_draw_color(*GOLD)
        pdf.line(15, pdf.get_y(), 195, pdf.get_y())
        pdf.ln(4)
        media_labels = ["Sound Tests", "Warehouse & Inspection Clips", "Drop Test Clips", "Product Images"]
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*WHITE)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(55, 7, "Description", 1, 0, "C", True)
        pdf.cell(0, 7, "Link", 1, 1, "C", True)
        pdf.set_text_color(*DARK_TEXT)
        for i, link in enumerate(media_links):
            if link:
                label = media_labels[i] if i < len(media_labels) else f"Attachment {i+1}"
                pdf.set_font("Helvetica", "", 8)
                pdf.cell(55, 6, label, 1, 0, "L")
                pdf.cell(0, 6, link, 1, 1, "L")
        pdf.ln(3)

    # VII. Factory Review
    pdf.ln(4 if (media_links and any(media_links)) else 8)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "VII. FACTORY REVIEW", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)

    fr = data.get("factory_review", {})
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*DARK_TEXT)
    pdf.cell(0, 6, f"Factory Cooperation:  {fr.get('cooperation', '-')}", ln=True)
    pdf.cell(0, 6, f"Number of Workers:  {fr.get('workers', '-')}", ln=True)
    opinion = fr.get("opinion", "")
    pdf.cell(0, 6, "Inspector's Opinion on Shipment:", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.multi_cell(0, 5, opinion if opinion else "(No opinion recorded)")

    # VIII. Conclusion
    pdf.ln(8)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "VIII. CONCLUSION", ln=True)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(5)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(*DARK_TEXT)
    conclusion = data.get("conclusion", "")
    pdf.multi_cell(0, 5, conclusion if conclusion else "(No conclusion entered)")

    # ── Signatures & Stamp ──
    pdf.ln(15)

    def draw_stamp(x, y, size=36):
        """Draw an NTS company stamp circle."""
        r = size / 2
        cx = x + r
        cy = y + r
        # Outer circle
        pdf.set_line_width(1.5)
        pdf.set_draw_color(*NAVY)
        pdf.circle(x, y, r)
        # Inner circle
        pdf.set_draw_color(*GOLD)
        pdf.circle(x + 2, y + 2, r - 4)
        pdf.set_line_width(0.3)
        # Text inside stamp
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 7)
        pdf.set_xy(cx - 14, cy - 10)
        pdf.cell(28, 3, "NEPAL TRADE", ln=True, align="C")
        pdf.set_xy(cx - 14, cy - 7)
        pdf.cell(28, 3, "SOLUTIONS", ln=True, align="C")
        # Separator line
        pdf.set_draw_color(*NAVY)
        pdf.line(cx - 10, cy - 3, cx + 10, cy - 3)
        pdf.set_font("Helvetica", "B", 6)
        pdf.set_xy(cx - 14, cy - 1)
        pdf.cell(28, 3, "INSPECTION", ln=True, align="C")
        # Separator line
        pdf.line(cx - 10, cy + 3, cx + 10, cy + 3)
        pdf.set_font("Helvetica", "", 6)
        pdf.set_xy(cx - 14, cy + 5)
        pdf.cell(28, 3, "APPROVED", ln=True, align="C")
        pdf.set_font("Helvetica", "I", 5)
        pdf.set_xy(cx - 14, cy + 8)
        pdf.cell(28, 3, datetime.now().strftime("%Y-%m-%d"), ln=True, align="C")

    # Left: Inspector signature
    pdf.set_draw_color(*GOLD)
    pdf.set_text_color(*DARK_TEXT)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(85, 5, "____________________________", ln=True)
    pdf.cell(85, 5, "Inspector: Kaustuv Guragain", ln=True)
    pdf.ln(2)
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(*GOLD)
    pdf.cell(85, 4, f"Date: {datetime.now().strftime('%Y-%m-%d')}", ln=True)

    # Right: Company stamp
    pdf.set_text_color(*DARK_TEXT)
    stamp_x = 120
    stamp_y = pdf.get_y() + 2
    draw_stamp(stamp_x, stamp_y, 36)
    pdf.set_y(stamp_y + 42)

    # Reviewed by line
    pdf.ln(3)
    pdf.set_draw_color(*GOLD)
    pdf.set_text_color(*DARK_TEXT)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(85, 5, "____________________________", ln=True)
    pdf.cell(85, 5, "Reviewed by:", ln=True)
    pdf.ln(2)
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(*GOLD)
    pdf.cell(85, 4, f"Date: {datetime.now().strftime('%Y-%m-%d')}", ln=True)

    pdf.ln(6)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(*GOLD)
    pdf.cell(0, 6, f"Report generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", ln=True)

    # ── End of Report ──
    pdf.ln(10)
    pdf.set_draw_color(*GOLD)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(3)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "End Of The Report", ln=True, align="C")
    pdf.ln(2)
    pdf.set_text_color(*DARK_TEXT)
    pdf.set_font("Helvetica", "I", 7)
    pdf.multi_cell(0, 3.5, "Note: This report presents our observations and findings based on the inspection conducted on the date stated. "
                 "It is prepared solely for the use of the client named above and contains confidential information. "
                 "Nepal Trade Solutions shall not be liable for any reliance placed on this report by third parties.")

    pdf.output(str(pdf_path))
    return str(pdf_path)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python report_generator.py <excel_path> [photo_dir]")
        sys.exit(1)
    excel = sys.argv[1]
    photos = sys.argv[2] if len(sys.argv) > 2 else "photos"
    result = build_report(excel, photos)
    print(f"Report generated: {result}")
