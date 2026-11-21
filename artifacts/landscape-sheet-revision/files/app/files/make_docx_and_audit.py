import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_memo_docx(output_path):
    doc = docx.Document()
    
    # Page setup - 1 inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)
    
    # Title / Company Letterhead
    p_comp = doc.add_paragraph()
    p_comp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_comp = p_comp.add_run("SUNDIAL ENGINEERING & SURVEY")
    run_comp.font.size = Pt(16)
    run_comp.font.bold = True
    run_comp.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D) # Navy
    p_comp.paragraph_format.space_after = Pt(2)
    
    p_addr = doc.add_paragraph()
    p_addr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_addr = p_addr.add_run("4720 South Rainbow Boulevard, Suite 210 · Las Vegas, Nevada 89103\nTelephone: (702) 555-0142 · NV Certificate of Authority No. C-3874")
    run_addr.font.size = Pt(8.5)
    run_addr.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    p_addr.paragraph_format.space_after = Pt(12)
    
    # Divider rule
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="18" w:space="1" w:color="1A365D"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)
    
    # MEMORANDUM Title
    p_memo = doc.add_paragraph()
    run_memo = p_memo.add_run("MEMORANDUM")
    run_memo.font.size = Pt(14)
    run_memo.font.bold = True
    run_memo.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    p_memo.paragraph_format.space_after = Pt(10)
    
    # Memo Header Table
    tbl_header = doc.add_table(rows=5, cols=2)
    tbl_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_header.autofit = False
    col_widths = [Inches(1.2), Inches(5.3)]
    
    headers_data = [
        ("TO:", "Garrett Nokes, Vice President, Land Development, Callaway Land Partners, LLC"),
        ("FROM:", "Max Braddock, Project Manager, Sundial Engineering & Survey"),
        ("APPROVED BY:", "Jonathan R. Kessler, PE (NV License #14203), Principal-in-Charge, Sundial Engineering & Survey"),
        ("DATE:", "October 31, 2026"),
        ("SUBJECT:", "Desert Sage Ranch Phase 3 (Job No. 2201-03) — Evaluation of Comment L-4 Street Tree Revision, Resubmittal Schedule Commitment (11/30/2026), Entitlement Runway (NRS 278.360), and Critical Prerequisite Actions")
    ]
    
    for r_idx, (label, val) in enumerate(headers_data):
        row = tbl_header.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = col_widths[0]
        c1.width = col_widths[1]
        set_cell_margins(c0, 40, 40, 50, 50)
        set_cell_margins(c1, 40, 40, 50, 50)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)
        if label == "SUBJECT:":
            r1.font.bold = True
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Project References Box
    tbl_ref = doc.add_table(rows=1, cols=1)
    c_ref = tbl_ref.rows[0].cells[0]
    set_cell_background(c_ref, "F0F4F8")
    set_cell_margins(c_ref, 80, 80, 120, 120)
    p_ref = c_ref.paragraphs[0]
    p_ref.paragraph_format.space_after = Pt(0)
    r_ref = p_ref.add_run("PROJECT REFERENCES: Clark County Improvement Plans: IP-26-100288 · Final Map: FM-26-100290 · Related Tentative Map: TM-22-500041 · APN: 177-27-201-001 (Portion)")
    r_ref.font.size = Pt(8.5)
    r_ref.font.bold = True
    r_ref.font.color.rgb = RGBColor(0x2C, 0x52, 0x82)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # Sections Helper
    def add_sec_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p_bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2B6CB0"/></w:pBdr>')
        p._p.get_or_add_pPr().append(p_bdr)
        
    def add_callout(text, hex_bg="FFF5F5", border_color="E53E3E"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tbl.rows[0].cells[0]
        set_cell_background(c, hex_bg)
        set_cell_margins(c, 100, 100, 150, 150)
        tcPr = c._tc.get_or_add_tcPr()
        bdr = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
        tcPr.append(bdr)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.font.size = Pt(9.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 1. EXECUTIVE SUMMARY
    add_sec_heading("1. EXECUTIVE SUMMARY & SCHEDULE COMMITMENT DETERMINATION")
    
    doc.add_paragraph(
        "This memorandum responds directly to your email of October 22, 2026, requesting that Sundial Engineering & Survey (\"Sundial\") "
        "and Ocotillo Landscape Studio (\"Ocotillo\") treat Clark County Public Works (CCPW) First-Round Review Comment L-4 (street tree canopy) "
        "as a routine clerical redline, adopt the legal exemption theory authored by Vogel & Marsh LLP for Comment L-2 (park turf classification), "
        "and commit to delivering the Phase 3 resubmittal package to Civil Plans Examiner Dale Yeoman, PE, by November 30, 2026, "
        "in order to preserve final map recordation by December 18, 2026."
    )
    
    add_callout(
        "SCHEDULE COMMITMENT FINDING:\n"
        "Sundial cannot commit to the November 30, 2026 resubmittal date on the basis of a clerical drafting redline, "
        "nor can we certify that a resubmittal executed on that premise will achieve County approval in time to record by December 18, 2026. "
        "Treating Comment L-4 as a clerical markup without resolving the underlying physical, utility, and safety clearances guarantees a second-round "
        "plan check rejection. Furthermore, submitting counsel's turf memo will be rejected outright by the County, and missing the statutory "
        "presentation deadline under NRS 278.360 on December 23, 2026 will terminate the entire tentative map by operation of law.",
        "FFF5F5", "C53030"
    )
    
    doc.add_paragraph(
        "A complete, legally compliant, and approvable resubmittal on November 30, 2026 remains technically feasible, but only if Callaway "
        "immediately executes three critical policy and budget actions by November 4, 2026, enabling our design team to execute the inter-agency "
        "clearances and technical reconciliations detailed below."
    )

    # 2. QUANTITATIVE ANALYSIS
    add_sec_heading("2. QUANTITATIVE ANALYSIS: COMMENT L-4 VS. CURRENT LANDSCAPE SHEETS")
    
    doc.add_paragraph(
        "Clark County Public Works General Comment L-4 (Dale Yeoman, PE, letter dated October 9, 2026) states:\n"
        "\"Street frontage planting shall meet the Title 30 ch. 30.64 standard of one (1) tree per twenty (20) linear feet of street frontage. "
        "Based on 4,160 linear feet of frontage, 208 trees are required; 58 trees are shown. Provide revised planting plan and a revised annual "
        "irrigation demand with the resubmittal.\""
    )
    
    doc.add_paragraph(
        "A rigorous audit of Sheet L-2 (Typical Streetscape Sections & Streetscape Plan), Sheet L-1 (Plant Schedule), and the Phase 3 Turf and "
        "Planting Tabulation confirms the exact frontage measurements and tree deficits across all three dedicated rights-of-way:"
    )
    
    # Table of Frontages
    tbl_front = doc.add_table(rows=5, cols=5)
    tbl_front.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_widths = [Inches(2.2), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.0)]
    f_headers = ["Frontage Segment / Street", "Length (LF)", "Trees Shown", "Required (1/20)", "Deficit"]
    f_rows = [
        ("Windmill Lane (North frontage, Ph 3 ext.)", "2,560 LF", "34", "128", "+94"),
        ("Phase 3 Interior Collector (Both sides)", "1,180 LF", "16", "59", "+43"),
        ("Rainbow Blvd. Entry (East frontage)", "420 LF", "8", "21", "+13"),
        ("TOTAL DEDICATED STREET FRONTAGE", "4,160 LF", "58", "208", "+150")
    ]
    
    for c_idx, h_text in enumerate(f_headers):
        cell = tbl_front.rows[0].cells[c_idx]
        cell.width = f_widths[c_idx]
        set_cell_background(cell, "1A365D")
        set_cell_margins(cell, 60, 60, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        
    for r_idx, r_data in enumerate(f_rows):
        row = tbl_front.rows[r_idx + 1]
        bg = "F7FAFC" if r_idx % 2 == 1 else "FFFFFF"
        if r_idx == 3: bg = "EDF2F7"
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = f_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if r_idx == 3:
                r.font.bold = True
                if c_idx == 4: r.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
                
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    doc.add_paragraph(
        "Project-Wide Tree Totals: In addition to the 58 street trees, the current Sheet L-1 schedule includes 6 Red Push Pistache (PIE) "
        "at the Rainbow Blvd. entry monument (Sheet L-4) and 10 Evergreen Elm (ULP) park shade trees (Sheet L-3), totaling 74 trees. "
        "Adding 150 street trees increases the total project tree count from 74 to 224 trees—a 258% increase in street trees and a 203% "
        "increase overall. On Sheet L-2, trees were drawn schematically at average intervals of 75 feet on Windmill Lane, 74 feet on the collector, "
        "and 52.5 feet on Rainbow Blvd. Squeezing 208 trees into these corridors at 20-foot intervals represents a major spatial re-design, not a clerical redline."
    )

    # 3. PHYSICAL CONSTRAINTS
    add_sec_heading("3. PHYSICAL, SUBSURFACE, AND SAFETY CONSTRAINTS ON TREE PLACEMENT")
    
    doc.add_paragraph(
        "The assumption that 150 additional trees can simply be drafted onto Sheet L-2 ignores severe subsurface, wet utility, and safety encumbrances:"
    )
    
    constraints = [
        ("Recorded 20-Foot Dry Utility Easement (Book 175, Page 60; SBT-2026-04188 Exception 4(b)): ",
         "A blanket 20-foot private dry utility easement runs along the north boundary of Phase 3 in favor of NV Energy, Southwest Gas, CenturyLink, "
         "and Cox Communications. Sheet L-2 General Note 5 strictly mandates: 'Contractor shall coordinate with dry utility joint trench and shall "
         "not plant trees within any dedicated easement without written permission of the utility.' Advisory Comment A-1 reinforces that modifying "
         "this private easement is an independent legal process requiring formal utility joinder. Tree pits and root barriers cannot encroaching on this corridor."),
        ("Dry Utility Joint Trench Envelope & Above-Ground Appurtenances (Sheets DU-1 / DU-2): ",
         "The composite dry utility plan establishes a 36-inch wide by 42-inch deep joint trench containing electric primary/secondary conduits, gas mains, "
         "and communications cables. Above ground, six pad-mounted transformers (TX-1 through TX-6), a Southwest Gas regulator station (GRS-1), and numerous "
         "telecom pedestals occupy the landscape strips. Strict working clearances (10 feet in front of transformer doors, 3 feet on sides) disqualify large frontage sections."),
        ("Mandatory Root Barriers (Sheet L-1, Detail 6): ",
         "Detail 6 requires a 24-inch deep by 10-foot long mechanical root barrier at every street tree adjacent to a public sidewalk. Installing 150 additional "
         "root barriers within narrow 5-foot to 8-foot landscape strips creates unavoidable vertical and horizontal clashes with shallow lateral conduits."),
        ("CCFD Fire Hydrant Sight Clearances (Sheet FA-1): ",
         "Fourteen public fire hydrants (FH-1 through FH-14) are located along the Phase 3 frontage. CCFD standards require an unobstructed 3-to-5-foot clear "
         "radial zone around each hydrant and clear visibility from roadway travel lanes. Trees cannot be placed within hydrant influence zones."),
        ("Storm Drain Line A Conflict (Sheet C-8): ",
         "Line A (36-inch to 60-inch RCP) parallels the north boundary at a nominal 10-foot offset. Deep tree pits (twice root ball diameter) jeopardize pipe "
         "structural bedding and County maintenance access."),
        ("Intersection Sight Visibility Triangles (Title 30 § 30.52.025; TM Condition 6): ",
         "Sight visibility triangles at Rainbow Blvd., Windmill Ln., and interior street intersections strictly prohibit any plantings or structures exceeding 24 inches in height."),
        ("Clark County Shade Canopy Guidance (CP-PW-GM-2025-04) & Alternative Standards: ",
         "County guidance requires a formal Shade Coverage Exhibit demonstrating 30% coverage of streetscape area at 15-year canopy. Where subsurface utilities "
         "or sight zones prevent 20-foot tree spacing, Clark County Code § 30.64.050 provides an Alternative Standards procedure to propose larger-canopy species "
         "at wider spacing. Comprehensive Planning approval of an Alternative Standards proposal is discretionary and requires written justification.")
    ]
    
    for bold_prefix, text in constraints:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(4)
        r_b = p.add_run(f"• {bold_prefix}")
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
        r_t = p.add_run(text)
        r_t.font.size = Pt(9.5)

    # 4. HYDRAULIC REPERCUSSIONS
    add_sec_heading("4. HYDRAULIC, WATER BUDGET, AND METER SIZING REPERCUSSIONS")
    
    doc.add_paragraph(
        "Comment L-4 explicitly requires submitting a revised annual irrigation demand with the resubmittal. Adding 150 trees substantially "
        "expands water consumption and alters system hydraulics:"
    )
    
    # Table of Water Demand
    tbl_water = doc.add_table(rows=5, cols=4)
    tbl_water.alignment = WD_TABLE_ALIGNMENT.CENTER
    w_widths = [Inches(2.5), Inches(1.5), Inches(1.5), Inches(1.0)]
    w_headers = ["Landscape Component", "First Submittal (08/21)", "Revised Set (11/30)", "Variance"]
    w_rows = [
        ("Spray Turf (Parcel A Park)", "18,400 SF (2.32 AF/YR)", "18,400 SF (2.32 AF/YR)", "0.00 AF/YR"),
        ("Drip Planting (Common Areas)", "22,200 SF (0.51 AF/YR)", "22,200 SF (0.51 AF/YR)", "0.00 AF/YR"),
        ("Drip Trees (2,000 gal/tree/yr)", "74 Trees (0.45 AF/YR)", "224 Trees (1.38 AF/YR)", "+0.92 AF/YR"),
        ("TOTAL ANNUAL APPLIED WATER", "3.28 AF/YR (1,068,791 gal)", "4.21 AF/YR (1,368,791 gal)", "+300,000 gal")
    ]
    
    for c_idx, h_text in enumerate(w_headers):
        cell = tbl_water.rows[0].cells[c_idx]
        cell.width = w_widths[c_idx]
        set_cell_background(cell, "1A365D")
        set_cell_margins(cell, 60, 60, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        
    for r_idx, r_data in enumerate(w_rows):
        row = tbl_water.rows[r_idx + 1]
        bg = "F7FAFC" if r_idx % 2 == 1 else "FFFFFF"
        if r_idx == 3: bg = "EDF2F7"
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if r_idx == 3:
                r.font.bold = True
                if c_idx == 3: r.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
                
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    doc.add_paragraph(
        "Meter Sizing and Hydraulic Integrity: Adding 150 trees requires 150 to 300 additional drip emitters. Sheet L-1 currently specifies one 1½-inch "
        "LVVWD meter at Parcel A and one ¾-inch meter at the entry monument. Ocotillo must calculate instantaneous peak flow (GPM) to verify whether "
        "the 1½-inch meter and existing 2-inch backflow preventer can deliver required flow within mandatory SNWA watering windows without excessive pressure drop. "
        "Both DSR_Irrigation_Demand_Worksheet.xlsx and DSR_Ph3_Water_Budget.xlsx must be formally recalculated and resubmitted."
    )

    # 5. FINANCIAL EXPOSURE
    add_sec_heading("5. FINANCIAL EXPOSURE: ENGINEER'S ESTIMATE & PERFORMANCE BOND")
    
    doc.add_paragraph(
        "Per Clark County Public Works Advisory Comment A-6, the itemized Engineer's Estimate forms the legal basis of the performance security "
        "to be posted prior to permit issuance. The current estimate totals $5,357,000.00 (Resubmittal Tracker Line B-01). Incorporating Comment L-4 "
        "under the Clark County Uniform Unit Cost Schedule results in the following mandatory budget adjustments:"
    )
    
    # Table of Costs
    tbl_cost = doc.add_table(rows=6, cols=6)
    tbl_cost.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_widths = [Inches(0.8), Inches(2.3), Inches(0.8), Inches(0.8), Inches(0.8), Inches(1.0)]
    c_headers = ["Item", "Description", "Orig Qty", "Rev Qty", "Unit Price", "Extended Cost"]
    c_rows = [
        ("7.09", "Tree, 24\" box, planted/staked/mulched", "64 EA", "214 EA", "$485.00", "+$72,750.00"),
        ("7.42", "Point-source drip emitter assembly", "750 EA", "950 EA", "$12.50", "+$2,500.00"),
        ("Detail 6", "Root barrier, 24\" D x 10' L (public walk)", "58 EA", "208 EA", "$65.00", "+$9,750.00"),
        ("7.40", "Lateral line, Class 200 PVC, 1-inch", "3,500 LF", "4,200 LF", "$3.15", "+$2,205.00"),
        ("7.35", "Remote control valve, 1-inch, with box", "40 EA", "44 EA", "$275.00", "+$1,100.00")
    ]
    
    for c_idx, h_text in enumerate(c_headers):
        cell = tbl_cost.rows[0].cells[c_idx]
        cell.width = c_widths[c_idx]
        set_cell_background(cell, "1A365D")
        set_cell_margins(cell, 60, 60, 60, 60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx not in [1] else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h_text)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        
    for r_idx, r_data in enumerate(c_rows):
        row = tbl_cost.rows[r_idx + 1]
        bg = "F7FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = c_widths[c_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 50, 50, 60, 60)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx not in [1] else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    doc.add_paragraph(
        "Revised Performance Bond Basis: Incorporating these direct additions ($88,305.00) increases the Phase 3 Engineer's Estimate from "
        "$5,357,000.00 to approximately $5,446,000.00 (rounded up to the next $1,000 per Schedule Sec. 12.02). Jonathan R. Kessler, PE, "
        "must re-seal the estimate. Callaway must notify Aegis Surety Company to increase the penal sum of the Phase 3 Performance Bond by $89,000.00."
    )

    # 6. TURF DILEMMA
    add_sec_heading("6. THE INTERTWINED TURF DILEMMA: COMMENT L-2, L-3 VS. COUNSEL'S EXEMPTION")
    
    doc.add_paragraph(
        "Your email directs Sundial to \"use Adrienne's memo on the turf issue and move on.\" Adopting that position guarantees a second-round rejection "
        "from Clark County Public Works and Comprehensive Planning."
    )
    
    doc.add_paragraph(
        "Vogel & Marsh LLP's memorandum (dated 10/20/2026) asserts that because Desert Sage Ranch Phase 3 is zoned R-2 (single-family residential), "
        "the entire subdivision—including the HOA common-area park (Parcel A)—is exempt from Nevada AB 356 and SNWA nonfunctional turf conversion rules."
    )
    
    add_callout(
        "REGULATORY CONFLICT — SNWA COMPLIANCE BULLETIN WRC-2024-07 (SECTION 2):\n"
        "\"The single-family residential carve-out is lot-based, not project-based. HOA-owned or HOA-managed common-area parcels of any description, "
        "including park parcels, tot lots, and greenbelts... remain fully subject to the prohibition even where they lie within a subdivision the zoning "
        "of which is single-family residential. Turf lying outside recorded lot boundaries is not within the carve-out. The operative distinction is "
        "the lot line, not the zoning district.\"\n\n"
        "SNWA BULLETIN SECTION 3.1 & TITLE 30 § 30.64.030(j):\n"
        "Playground Turf qualifies as functional ONLY if it lies within a designated play area containing INSTALLED AND OPERABLE PLAYGROUND AMENITIES "
        "(specifically: slides, swings, and climbing structures). Because Callaway value-engineered out all play amenities in July 2026, the 18,400 SF "
        "of turf on Sheet L-3 is classified as NONFUNCTIONAL TURF. Continuing to irrigate nonfunctional turf after 12/31/2026 constitutes a violation "
        "of state statute and SNWA service rules, subjecting the HOA to excess-use rates and water shut-off.",
        "FFF5F5", "C53030"
    )
    
    doc.add_paragraph(
        "Callaway Must Direct One of Two Compliant Paths Forward by November 4, 2026:\n"
        "• Option A (Recommended — Amenity Reinstatement): Reintroduce the playground amenities (swings, slides, climbing structures) onto Sheet L-3. "
        "This immediately qualifies the 18,400 SF turf area as functional Playground Turf under SNWA § 3.1, closes Comments L-2 and L-3, and directly "
        "satisfies the Phase 1/2 HOA homeowners' demand (Sagebrook Community Management letter of 10/17/2026), defusing planned protests at the TAB.\n"
        "• Option B (Turf Conversion): Direct Ocotillo to eliminate the 18,400 SF turf lawn from Sheet L-3, replacing it with a drip-irrigated desert palette "
        "and decomposed granite. This reduces annual irrigation demand by 2.32 AF/YR but requires preparing a formal turf conversion plan."
    )

    # 7. TAB CALENDAR
    add_sec_heading("7. CONDITION 14 DESIGN REVIEW & ENTERPRISE TAB CALENDAR TRAP")
    
    doc.add_paragraph(
        "Tentative Map TM-22-500041 Condition 14 mandates public hearing Design Review before the Enterprise Town Advisory Board (TAB) and Planning "
        "Commission for any \"significant change to approved plans, including landscape and streetscape treatments in common areas.\" Advisory Comment A-3 "
        "reminds Callaway that revisions altering approved plan character trigger additional review cycles."
    )
    
    doc.add_paragraph(
        "The Enterprise TAB 2026 Calendar creates an insurmountable schedule barrier if a Design Review is triggered:\n"
        "• Meeting 11 (November 18, 2026): Application submittal cutoff (21 days prior) was OCTOBER 28, 2026 (CLOSED THREE DAYS AGO).\n"
        "• Meeting 12 (December 16, 2026): Application submittal deadline is November 25, 2026. The TAB hearing is December 16, 2026, and recommendation "
        "transmittal to the Planning Commission / BCC would trail into January 2027."
    )
    
    add_callout(
        "CRITICAL PROCESS RISK:\n"
        "If Comprehensive Planning determines that quadrupling street trees or modifying park amenities constitutes a 'significant change' requiring "
        "TAB Design Review, approval prior to December 18, 2026 is procedurally impossible, causing the tentative map to terminate under NRS 278.360. "
        "To avoid TAB referral, Sundial must pre-coordinate with Principal Planner Odette Marchetti to confirm in writing that increasing street trees "
        "represents technical code compliance, and Callaway must adopt Option A to restore the approved tentative map park concept.",
        "FFF5F5", "C53030"
    )

    # 8. STATUTORY ENTITLEMENT RUNWAY
    add_sec_heading("8. STATUTORY ENTITLEMENT RUNWAY: NRS 278.360 REALITY VS. TITLE ADVICE")
    
    doc.add_paragraph(
        "On October 26, 2026, Title Officer Neil Petrosyan (Silver Butte Title) emailed you that under NRS 278.360 the tentative map has four years from "
        "approval (to 03/14/2026) plus a 2-year extension to 03/14/2028, advising: \"you've got plenty of runway... wouldn't sweat the December recordation date.\""
    )
    
    add_callout(
        "CORRECTION OF GOVERNING STATUTORY DEADLINE (NRS 278.360):\n"
        "Neil Petrosyan's advice is factually and legally incorrect. NRS 278.360 and the TM-22-500041 Notice of Approval govern presentation deadlines "
        "for successive final maps in a series:\n"
        "1. The 4-year clock from tentative approval applies ONLY to the first final map in a series. That requirement was satisfied when Phases 1 & 2 "
        "recorded on December 23, 2024 (Book 175, Page 60).\n"
        "2. For all successive final maps, NRS 278.360 explicitly dictates:\n"
        "   \"Each successive final map in the same series must be presented within two (2) years after the date on which the immediately preceding "
        "final map is recorded. Failure to timely present a map terminates all proceedings relating to the subdivision.\"\n"
        "3. Two years from December 23, 2024 is DECEMBER 23, 2026.\n"
        "4. NRS 278.360 expressly prohibits pre-expiration extensions: 'An extension shall not be granted before the applicable presentation period has expired.' "
        "(Comprehensive Planning formally denied Callaway's premature extension request on 12/30/2024 on this exact statutory basis).\n\n"
        "CONCLUSION: If FM-26-100290 is not presented to Clark County on or before DECEMBER 23, 2026, Tentative Map TM-22-500041 terminates automatically "
        "by operation of law. All entitlement rights to the 60 lots will be permanently extinguished.",
        "FFF5F5", "9B2C2C"
    )

    # 9. OTHER INTER-AGENCY BLOCKERS
    add_sec_heading("9. AUDIT OF OTHER INTER-AGENCY BLOCKERS ON THE RESUBMITTAL TRACKER")
    
    doc.add_paragraph(
        "In addition to landscape, several critical agency clearances on the Resubmittal Tracker are unresolved and will block resubmittal intake:"
    )
    
    blockers = [
        ("Expired Sewer Capacity Confirmation (Comment U-2; Condition 22): ",
         "Confirmation CCWRD-SCC-2025-1187 (issued 09/30/2025) explicitly expired on SEPTEMBER 30, 2026. It is currently EXPIRED. Comment U-2 mandates "
         "current confirmations at resubmittal. An emergency written renewal request must be filed with Vince Alarid (CCWRD) immediately."),
        ("CCFD Permanent Second Access Required (Comment F-1; Condition 11; Sheet FA-1): ",
         "Sheet FA-1 shows only one permanent connection to Windmill Lane. Keynote 4 confirms that the temporary Phase 2 construction stub must be removed "
         "upon Phase 3 buildout per CCFD File FPB-2023-0417. CCFD requires two permanent access points for developments exceeding 30 lots. Sundial must "
         "coordinate an acceptable permanent secondary access design or emergency access easement with Capt. Emmett Roark prior to resubmittal."),
        ("Landscape Architect Professional Seal (Comment L-1): ",
         "All revised sheets L-1 through L-4 must bear the professional wet seal and signature of Bea Cortland, PLA (NV LA #742)."),
        ("Owner's Certificate & Statutory Notary (Comment M-5): ",
         "Callaway must designate the authorized corporate signatory and execute the Owner's Certificate on FM-26-100290 using current Nevada statutory notary acknowledgment language."),
        ("Stand-Alone ESC Sheet & NDEP NOI (Comment G-8): ",
         "Sundial must produce a dedicated Erosion and Sediment Control sheet referencing active coverage under NDEP Permit NVR100000."),
        ("Detention Reconciliation & Drainage Profiles (Comments G-1, G-3, G-4, G-5): ",
         "Sundial must finalize the C-1 drainage narrative, reconcile the 6.4 AF C-9 basin summary with the Master Drainage Study, detail the spillway apron, and add HGL/EGL profiles on C-8.")
    ]
    
    for bold_prefix, text in blockers:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(4)
        r_b = p.add_run(f"• {bold_prefix}")
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
        r_t = p.add_run(text)
        r_t.font.size = Pt(9.5)

    # 10. ACTION PLAN
    add_sec_heading("10. COMPREHENSIVE ACTION PLAN: WHAT HAS TO HAPPEN FIRST")
    
    doc.add_paragraph(
        "To achieve an approvable resubmittal by November 30, 2026, and protect the December 23, 2026 statutory entitlement cliff, the following "
        "four-phase action plan must be strictly executed:"
    )
    
    phases = [
        ("Phase I: Immediate Client Direction & Agency Clearances (November 2 – November 6)", [
            ("Client Park Turf Decision (Due 11/04 — Garrett Nokes): ", "Formally direct Ocotillo to adopt Option A (reinstate play amenities to qualify as functional playground turf) or Option B (full desert conversion)."),
            ("Surety Notification (Due 11/04 — Garrett Nokes): ", "Authorize an $89,000 increase in the Phase 3 Performance Bond basis ($5,446,000) and notify Aegis Surety Company."),
            ("Owner's Signatory Designation (Due 11/04 — Garrett Nokes): ", "Designate authorized corporate signatory and statutory notary for FM-26-100290."),
            ("Sewer Capacity Renewal (Due 11/02 — Max Braddock): ", "Submit emergency written application to Vince Alarid (CCWRD) to renew expired confirmation CCWRD-SCC-2025-1187."),
            ("CCFD Second Access Meeting (Due 11/04 — Max Braddock): ", "Meet with Capt. Emmett Roark to establish an acceptable permanent secondary access design or emergency access easement."),
            ("Planning Sign-Off on Administrative Review (Due 11/06 — Max Braddock): ", "Secure written confirmation from Odette Marchetti that landscape revisions are technical code compliance and will not be referred to Enterprise TAB."),
            ("Sagebrook HOA Outreach (Due 11/06 — Garrett Nokes): ", "Respond to Sagebrook Community Management confirming reinstatement of play amenities, neutralizing homeowner protests at the 11/18 TAB meeting.")
        ]),
        ("Phase II: Technical Coordination, Hydraulic Modeling & Alternative Standards (November 4 – November 13)", [
            ("Spatial Conflict Layout (Due 11/10 — Bea Cortland & Priya Sarwal): ", "Overlay 208 street trees against Sheet DU-1/DU-2 joint trench, 14 fire hydrants, storm drain Line A, and sight triangles."),
            ("Alternative Standards Application (Due 11/13 — Bea Cortland, if needed): ", "If utility corridors prevent 20-foot spacing, prepare formal Title 30 § 30.64.050 Alternative Standards justification for staff approval."),
            ("Hydraulic & Water Demand Recalculation (Due 11/13 — Bea Cortland): ", "Update DSR_Irrigation_Demand_Worksheet.xlsx and DSR_Ph3_Water_Budget.xlsx (+300,000 gal/yr tree allowance); verify 1½-inch meter and backflow sizing."),
            ("Civil & Drainage Reconciliations (Due 11/13 — Priya Sarwal): ", "Finalize C-1 drainage narrative, C-9 basin summary (6.4 AF), C-8 HGL/EGL profiles, spillway apron details, and stand-alone ESC sheet.")
        ]),
        ("Phase III: Sheet Finalization, Sealing & QA/QC (November 13 – November 20)", [
            ("Landscape Sheet Finalization (Due 11/18 — Bea Cortland): ", "Reconcile L-1 schedule counts with L-2/L-3 plans (Comment L-6); dimension 10-foot curb setbacks (Comment L-5); apply professional seal and signature (Comment L-1)."),
            ("Survey & Mapping Finalization (Due 11/18 — Marcus Halvorsen, PLS): ", "Plot Schedule B exceptions from Title Report SBT-2026-04188 on FM-26-100290; reconcile Basis of Bearings with Phase 1 & 2 map; finalize closure sheets."),
            ("Estimate Re-Sealing (Due 11/20 — Jonathan R. Kessler, PE): ", "Update itemized estimate ($5,446,000) and issue sealed PDF."),
            ("Internal QA/QC Back-Check (Due 11/20 — Priya Sarwal & Max Braddock): ", "Execute complete Sundial QC-05 check across all 13 sheets and comment response matrix.")
        ]),
        ("Phase IV: Packaging, Expedited Review Request & Submittal (November 20 – November 30)", [
            ("Submittal Assembly (Due 11/25 — Max Braddock): ", "Assemble annotated comment matrix, red-lined plan sets, clean revised sheets, updated water budget, renewal letters, and title report."),
            ("Expedited Review Negotiation (Due 11/25 — Max Braddock): ", "Dale Yeoman's letter notes standard second-round review takes 20 business days (which would push approval to January 2027). Sundial and Callaway must meet with Dale Yeoman to negotiate an expedited 10-to-12 business day review window, predicated on delivering a clean, fully pre-coordinated package with zero open agency defects."),
            ("Counter Delivery (Due 11/30 — Sundial): ", "Hand-deliver bond sets and upload digital files to Clark County Development Review counter under IP-26-100288 and FM-26-100290.")
        ])
    ]
    
    for phase_title, items in phases:
        p_ph = doc.add_paragraph()
        p_ph.paragraph_format.space_before = Pt(8)
        p_ph.paragraph_format.space_after = Pt(4)
        p_ph.paragraph_format.keep_with_next = True
        r_ph = p_ph.add_run(phase_title)
        r_ph.font.bold = True
        r_ph.font.size = Pt(10)
        r_ph.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        
        for prefix, desc in items:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.left_indent = Inches(0.2)
            p_item.paragraph_format.space_after = Pt(3)
            r_ip = p_item.add_run(f"• {prefix}")
            r_ip.font.bold = True
            r_ip.font.size = Pt(9)
            r_id = p_item.add_run(desc)
            r_id.font.size = Pt(9)

    # SIGNATURE BLOCK
    doc.add_paragraph().paragraph_format.space_before = Pt(14)
    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.keep_with_next = True
    p_sig.add_run("Respectfully submitted,\n\n").font.size = Pt(10)
    r_firm = p_sig.add_run("SUNDIAL ENGINEERING & SURVEY\n\n\n\n")
    r_firm.font.bold = True
    r_firm.font.size = Pt(10)
    
    tbl_sig = doc.add_table(rows=1, cols=2)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sig.autofit = False
    sig_widths = [Inches(3.2), Inches(3.3)]
    
    c_s0, c_s1 = tbl_sig.rows[0].cells[0], tbl_sig.rows[0].cells[1]
    c_s0.width = sig_widths[0]
    c_s1.width = sig_widths[1]
    
    p_s0 = c_s0.paragraphs[0]
    p_s0.add_run("__________________________________________\n").font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    r_mb = p_s0.add_run("Max Braddock\n")
    r_mb.font.bold = True
    p_s0.add_run("Project Manager\nDirect: (702) 555-0144\nEmail: mbraddock@sundialensv.com")
    
    p_s1 = c_s1.paragraphs[0]
    p_s1.add_run("__________________________________________\n").font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    r_jk = p_s1.add_run("Jonathan R. Kessler, PE\n")
    r_jk.font.bold = True
    p_s1.add_run("Principal-in-Charge / Engineer of Record\nNevada Professional Engineer License No. 14203")
    
    doc.add_paragraph().paragraph_format.space_before = Pt(12)
    p_cc = doc.add_paragraph()
    r_cc = p_cc.add_run("cc:  Adrienne Vogel, Esq. — Vogel & Marsh LLP\n     Bea Cortland, PLA — Ocotillo Landscape Studio\n     Sundial Project File 2201-03 (Desert Sage Ranch Phase 3)")
    r_cc.font.size = Pt(8.5)
    r_cc.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    
    doc.save(output_path)
    print(f"Successfully generated DOCX: {output_path}")

create_memo_docx('output/MEMORANDUM_Callaway_Land_Partners_2026-10-31.docx')
