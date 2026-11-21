import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

os.makedirs('output', exist_ok=True)

# -------------------------------------------------------------
# 1. GENERATE MARKDOWN MEMORANDUM
# -------------------------------------------------------------
memo_md = """# SUNDIAL ENGINEERING & SURVEY
**4720 South Rainbow Boulevard, Suite 210 · Las Vegas, Nevada 89103**  
**Telephone: (702) 555-0142 · NV Certificate of Authority No. C-3874**

---

### **MEMORANDUM**

**TO:** Garrett Nokes, Vice President, Land Development, Callaway Land Partners, LLC  
**FROM:** Max Braddock, Project Manager, Sundial Engineering & Survey  
**REVIEWED & APPROVED BY:** Jonathan R. Kessler, PE (NV License #14203), Principal-in-Charge, Sundial Engineering & Survey  
**DATE:** October 31, 2026  
**SUBJECT:** Desert Sage Ranch Phase 3 (Job No. 2201-03) — Evaluation of Comment L-4 Street Tree Revision, Resubmittal Schedule Commitment (11/30/2026), Entitlement Runway (NRS 278.360), and Critical Prerequisite Actions  
**PROJECT REFERENCES:** Clark County Improvement Plans: IP-26-100288 · Final Map: FM-26-100290 · Related Tentative Map: TM-22-500041 · Assessor's Parcel No.: 177-27-201-001 (Portion)

---

### 1. EXECUTIVE SUMMARY & SCHEDULE COMMITMENT DETERMINATION

This memorandum responds directly to your instruction of October 22, 2026, requesting that Sundial Engineering & Survey ("Sundial") and Ocotillo Landscape Studio ("Ocotillo") treat Clark County Public Works (CCPW) First-Round Review Comment L-4 (street frontage tree canopy) as a routine clerical redline, adopt the legal exemption theory authored by Vogel & Marsh LLP for Comment L-2 (park turf classification), and commit to delivering the Phase 3 resubmittal package to Civil Plans Examiner Dale Yeoman, PE, by November 30, 2026, to preserve final map recordation by December 18, 2026.

After conducting a rigorous technical cross-check against the current Phase 3 landscape sheets (L-1 through L-4), the civil improvement plans (C-1 through C-9), the dry utility composite plan (DU-1/DU-2), the preliminary title report (SBT-2026-04188), and the running Resubmittal Schedule and Task Tracker, Sundial's professional determination is as follows:

1. **Sundial cannot commit to the November 30, 2026 resubmittal date on the basis of a clerical drafting redline, nor can we certify that a resubmittal executed on that premise will achieve County approval in time to record by December 18, 2026.**
2. Treating Comment L-4 as a clerical markup without resolving the underlying physical, utility, and safety clearances guarantees a second-round plan rejection. Comment L-4 requires adding **150 street frontage trees** (increasing the count from 58 to 208 trees across 4,160 linear feet). Dropping these trees into dedicated rights-of-way creates direct conflicts with the recorded 20-foot dry utility easement along the north boundary, NV Energy transformers, Southwest Gas mains, 14 Clark County Fire Department (CCFD) fire hydrants, storm drain Line A, and mandatory intersection sight visibility triangles.
3. Submitting the Vogel & Marsh legal memorandum to close Comment L-2 (park turf) will be **rejected outright** by CCPW and Clark County Comprehensive Planning. Southern Nevada Water Authority (SNWA) Compliance Bulletin WRC-2024-07 explicitly establishes that the single-family turf exemption is strictly *lot-based*, not *project-based*, and does not apply to Homeowner Association (HOA) common areas regardless of zoning. Because Callaway value-engineered out all playground amenities (swings, slides, climbing structures) in July 2026, the 18,400 square feet of turf in Parcel A is legally classified as **nonfunctional turf** under Nevada Assembly Bill 356 and Clark County Code § 30.64.030(j), carrying a mandatory conversion deadline of December 31, 2026.
4. **Correction of Entitlement Runway (NRS 278.360):** The informal email advice you received from Silver Butte Title & Escrow on October 26, 2026—asserting that the tentative map remains protected until March 2028—is legally incorrect. Under NRS 278.360 and the TM-22-500041 Conditions of Approval, successive final maps must be presented within **two (2) years of the recordation date of the immediately preceding final map**. Phases 1 & 2 recorded on December 23, 2024 (Book 175, Page 60). Therefore, the Phase 3 final map must be presented on or before **December 23, 2026**. Failure to timely present terminates the entire tentative map by operation of law. The December 18 target is not an internal preference; it is our only buffer against a catastrophic statutory forfeiture.
5. **Path Forward:** A complete, legally compliant, and approvable resubmittal on November 30, 2026 **remains achievable**, but only if Callaway immediately authorizes three critical policy and budget actions by **November 4, 2026**, enabling our design team to execute the inter-agency clearances and technical reconciliations detailed in Section 7 of this memorandum.

---

### 2. QUANTITATIVE ANALYSIS: COMMENT L-4 VS. CURRENT LANDSCAPE SHEETS

Clark County Public Works General Comment L-4 (Dale Yeoman, PE, letter dated October 9, 2026) states:
> *"Street frontage planting shall meet the Title 30 ch. 30.64 standard of one (1) tree per twenty (20) linear feet of street frontage. Based on 4,160 linear feet of frontage, 208 trees are required; 58 trees are shown. Provide revised planting plan and a revised annual irrigation demand with the resubmittal."*

A line-by-line audit of Sheet L-2 (Typical Streetscape Sections & Streetscape Plan) and the Phase 3 Common Area Turf and Planting Tabulation establishes the following frontage breakdown:

| Frontage Segment | Governing Right-of-Way | Frontage Length (LF) | Trees Shown (08/21/2026 Set) | Required Trees (1 per 20 LF) | Net Deficit (Trees to Add) |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **Windmill Lane Frontage** (North side, Phase 3 extension) | Windmill Lane (80' Collector) | 2,560 LF | 34 | 128 | **+94** |
| **Phase 3 Interior Collector** (Both sides of roadway) | Desert Sage Parkway (60' Collector) | 1,180 LF | 16 | 59 | **+43** |
| **Rainbow Boulevard Entry** (East side landscape strip) | Rainbow Boulevard (100' Arterial) | 420 LF | 8 | 21 | **+13** |
| **TOTAL DEDICATED FRONTAGE** | — | **4,160 LF** | **58** | **208** | **+150** |

#### Plant Schedule Discrepancies (Sheets L-1 through L-4)
- **Current Tree Count Across Entire Set:** The plant schedule on Sheet L-1 currently specifies 74 total trees: 58 street frontage trees (18 Desert Willow [CHL], 22 Palo Verde [PAR], 18 Mesquite [PBR]); 6 Red Push Pistache [PIE] at the Rainbow Blvd. entry monument (Sheet L-4); and 10 Evergreen Elm [ULP] park shade trees (Sheet L-3).
- **Revised Tree Count:** Adding 150 street trees increases the street frontage subtotal from 58 to 208 trees, and the project-wide total from 74 to **224 trees** (a 258% increase in street trees; 203% increase overall).
- **Drafting Spacing vs. Code Standard:** On the first submittal, trees were depicted schematically at average spacings of 75 feet on Windmill Lane, 74 feet on the interior collector, and 52.5 feet on Rainbow Boulevard, despite sheet notes citing 20-foot spacing. Compressing 208 trees into these corridors at actual 20-foot intervals cannot be accomplished as a drafting redline; it requires a comprehensive spatial and civil re-engineering of the right-of-way corridors.

---

### 3. PHYSICAL, SUBSURFACE, AND SAFETY CONSTRAINTS ON TREE PLACEMENT

Arbitrarily placing 150 additional trees into the landscape strips creates direct spatial conflicts with existing and proposed infrastructure:

```
+----------------------------------------------------------------------------------------------------+
| TYPICAL FRONTAGE CONFLICT PROFILE (NORTH BOUNDARY & DEDICATED STREETSCAPES)                        |
|                                                                                                    |
|  [ R/W or Boundary Line ]                                                                          |
|       |                                                                                            |
|       |<----------------- 20-FOOT RECORDED DRY UTILITY EASEMENT (Book 175, Page 60) ------------->|
|       |                                                                                            |
|       |   +---------------------------------------+                                                |
|       |   | 36" W x 42" D JOINT TRENCH (DU-1/2)   |  <--- Electric, Gas, CenturyLink, Cox          |
|       |   +---------------------------------------+                                                |
|       |                                                                                            |
|       |          * PROPOSED 24" BOX TREE PIT (Detail 1/L-1: 4'-0" to 5'-0" excavation width)       |
|       |         ***  CONFLICT: Root balls and required 24" D x 10' L Root Barriers (Detail 6/L-1)  |
|       |        ***** directly encroach into joint-trench envelope and shallow conduits!           |
|       |          |                                                                                 |
|       |<-- 10' ->| nominal offset                                                                  |
|       |                                                                                            |
|       |   (=======================================)                                                |
|       |   ( 36" to 60" RCP STORM DRAIN - LINE A   )  <--- C-8 alignment parallels boundary         |
|       |   (=======================================)                                                |
|       |                                                                                            |
|       |   [O] FIRE HYDRANT CLEAR ZONE (FA-1) <--- 14 hydrants require 3'-5' clear radius           |
|       |   |>| SIGHT TRIANGLE (Title 30 / C-1) <--- Max 24" vegetation height at intersections       |
+----------------------------------------------------------------------------------------------------+
```

1. **Recorded 20-Foot Dry Utility Easement (Book 175, Page 60; SBT-2026-04188 Exception 4(b)):**
   - A blanket 20-foot private dry utility easement runs along the north boundary of Phase 3, granted to NV Energy, Southwest Gas, CenturyLink, and Cox Communications on the Phase 1 & 2 recorded final map.
   - Sheet L-2 General Note 5 explicitly dictates: *"Contractor shall coordinate with dry utility joint trench (per composite plan) and shall not plant trees within any dedicated easement without written permission of the utility."*
   - Advisory Comment A-1 confirms that modifying or encroaching upon this easement requires separate legal joinder of all four utility purveyors. Root barriers (Detail 6/L-1: 24 inches deep by 10 feet long) cannot be installed without trenching through franchise utilities.
2. **Above-Ground Utility Structures:** The Dry Utility Composite (Sheets DU-1 and DU-2) establishes six pad-mounted transformers (TX-1 through TX-6), a Southwest Gas regulator station (GRS-1), and numerous telecom splice pedestals (PED-#) within the landscape strips. Strict working clearances (10 feet in front of transformer doors, 3 feet on sides) eliminate substantial frontage segments from tree planting.
3. **CCFD Fire Hydrant Obstructions (Sheet FA-1):** Fourteen public fire hydrants (FH-1 through FH-14) are located along the Phase 3 frontage. CCFD design standards require an unobstructed 3-to-5-foot radial clear zone around every hydrant and a clear sight window from approaching fire apparatus. Trees cannot be placed within hydrant influence zones.
4. **Storm Drain Line A Encroachment (Sheet C-8):** Line A consists of 36-inch to 60-inch reinforced concrete pipe running west-to-east at a nominal 10-foot southerly offset from the boundary. Tree excavations (twice root ball diameter) and deep root systems jeopardize pipe bedding, backfill compaction, and County maintenance access.
5. **Intersection Sight Visibility Triangles (Clark County Code § 30.52.025; TM Condition 6):** Minimum 30-foot to 50-foot sight visibility triangles at the Rainbow Blvd., Windmill Ln., and interior street intersections legally prohibit any plantings or structures exceeding 24 inches in finish height.
6. **Clark County Shade Canopy Guidance (CP-PW-GM-2025-04) & Alternative Standards:**
   - County guidance requires a formal **Shade Coverage Exhibit** demonstrating 30% streetscape coverage at 15-year canopy.
   - If subsurface utilities, hydrants, and sight triangles prevent mathematical placement of 208 trees at 20-foot intervals, Clark County Code § 30.64.050 allows an **Alternative Standards Application**. Under this mechanism, the applicant may propose higher-canopy species at wider spacing, or cluster planting where corridors permit. Comprehensive Planning approval of an Alternative Standards proposal is discretionary and requires a written submittal.

---

### 4. HYDRAULIC, WATER BUDGET, AND METER SIZING REPERCUSSIONS

Dale Yeoman's Comment L-4 explicitly requires submitting a revised annual irrigation demand with the resubmittal. Adding 150 trees dramatically expands water demand and hydraulic loading:

```
====================================================================================================
ANNUAL IRRIGATION DEMAND RECONCILIATION — DESERT SAGE RANCH PHASE 3
====================================================================================================
Landscape Component               First Submittal (08/21/2026)         Revised Resubmittal (11/30/2026)
----------------------------------------------------------------------------------------------------
Spray Turf (Parcel A Park)        18,400 SF @ 5.5 FT/YR = 2.32 AF/YR   18,400 SF @ 5.5 FT/YR = 2.32 AF/YR
Drip Planting (Common Areas)      22,200 SF @ 1.0 FT/YR = 0.51 AF/YR   22,200 SF @ 1.0 FT/YR = 0.51 AF/YR
Drip Trees (Per-Tree Rate)        74 Trees @ 2,000 GPY  = 0.45 AF/YR   224 Trees @ 2,000 GPY = 1.38 AF/YR
                                  (148,000 Gallons/Year)               (448,000 Gallons/Year)
----------------------------------------------------------------------------------------------------
TOTAL ANNUAL APPLIED WATER        3.28 ACRE-FEET / YEAR                4.21 ACRE-FEET / YEAR
                                  (1,068,791 Gallons/Year)             (1,368,791 Gallons/Year)
----------------------------------------------------------------------------------------------------
NET ANNUAL WATER INCREASE:        +300,000 GALLONS / YEAR (+0.92 ACRE-FEET / YEAR; +28.4% OVER BASELINE)
====================================================================================================
```

#### Hydraulic Infrastructure Impacts
1. **Meter Sizing & Pressure Delivery:** Sheet L-1 currently specifies one 1½-inch LVVWD landscape meter serving Parcel A and one ¾-inch LVVWD meter serving the Rainbow Blvd. entry monument. Adding 150 trees requires adding 150 to 300 pressure-compensating drip emitters. Ocotillo must calculate instantaneous peak flow (GPM) to verify whether the 1½-inch meter and existing 2-inch reduced pressure backflow preventer (Item 7.32) can deliver required volume within mandatory SNWA watering windows without excessive friction loss.
2. **Distribution Sizing:** Additional lateral lines (Class 200 PVC, Item 7.40) and remote control valves (Item 7.35) must be incorporated into Sheet L-2 and reflected on civil utility interfaces.
3. **Worksheet Revisions:** Ocotillo must formally revise both `DSR_Irrigation_Demand_Worksheet.xlsx` and `DSR_Ph3_Water_Budget.xlsx`.

---

### 5. FINANCIAL EXPOSURE: ENGINEER'S ESTIMATE & PERFORMANCE BOND

Clark County Public Works Advisory Comment A-6 states that the itemized Engineer's Estimate forms the legal basis of the performance security required prior to permit issuance. 

The current Phase 3 Engineer's Estimate (dated 10/10/2026, sealed by Jonathan R. Kessler, PE) totals **$5,357,000.00** (Resubmittal Tracker Line B-01). Incorporating Comment L-4 requires the following cost adjustments under the Clark County Uniform Unit Cost Schedule:

| Schedule Item | Description | Original Qty | Revised Qty | Unit | Unit Price | Extended Cost Increase |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| **7.09** | Tree, 24" box, planted/staked/mulched | 64 EA | 214 EA | EA | $485.00 | **+$72,750.00** |
| **7.42** | Point-source drip emitter assembly | 750 EA | 950 EA | EA | $12.50 | **+$2,500.00** |
| **Detail 6** | Root barrier, 24" D x 10' L (public walk) | 58 EA | 208 EA | EA | $65.00 | **+$9,750.00** |
| **7.40** | Lateral line, Class 200 PVC, 1-inch | 3,500 LF | 4,200 LF | LF | $3.15 | **+$2,205.00** |
| **7.35** | Remote control valve, 1-inch, with box | 40 EA | 44 EA | EA | $275.00 | **+$1,100.00** |
| **NET CONSTRUCTION INCREASE** | — | — | — | — | — | **+$88,305.00** |

- **Revised Bond Requirement:** Incorporating these adjustments raises the estimated Phase 3 construction total to approximately **$5,446,000.00** (rounded up to the nearest $1,000 per Schedule Sec. 12.02).
- **Administrative Action:** Jonathan R. Kessler, PE, must re-seal the estimate. Callaway must immediately notify its surety broker (Aegis Surety Company) to increase the penal sum of the Phase 3 Performance Bond by **$89,000.00** prior to executing the Subdivision Improvement Agreement (SIA).

---

### 6. THE INTERTWINED TURF DILEMMA: COMMENT L-2, L-3 VS. COUNSEL'S EXEMPTION

Your email instructs Sundial to *"use Adrienne's memo on the turf issue and move on."* Relying on that memo will lead to immediate second-round rejection.

```
+----------------------------------------------------------------------------------------------------+
| REGULATORY CONFLICT: COUNSEL OPINION VS. GOVERNING SNWA / CLARK COUNTY REGULATIONS                 |
|                                                                                                    |
| VOGEL & MARSH LLP POSITION (10/20/2026):                                                           |
| Desert Sage Ranch Phase 3 is zoned R-2 (single-family residential). AB 356 applies to              |
| non-single-family properties. Therefore, the single-family designation exempts the entire          |
| subdivision, including the common-area park. No turf conversion required; Sheet L-1 note is valid.|
|                                                                                                    |
|                                     VS.                                                            |
|                                                                                                    |
| SOUTHERN NEVADA WATER AUTHORITY BULLETIN WRC-2024-07 (SECTION 2):                                  |
| "The single-family residential carve-out is lot-based, not project-based. HOA-owned common-area    |
| parcels of any description, including park parcels... remain fully subject to the prohibition even |
| where they lie within a subdivision the zoning of which is single-family residential. Turf lying   |
| outside recorded lot boundaries is not within the carve-out. The operative distinction is the lot  |
| line, not the zoning district."                                                                    |
|                                                                                                    |
| SNWA BULLETIN WRC-2024-07 (SECTION 3.1) & CLARK COUNTY CODE § 30.64.030(j):                        |
| Playground Turf qualifies as functional ONLY if it lies within a designated play area containing    |
| INSTALLED AND OPERABLE PLAYGROUND AMENITIES (slides, swings, climbing structures).                 |
|                                                                                                    |
| FACTUAL STATUS ON SHEET L-3:                                                                       |
| Callaway VE'd out all play equipment in July 2026. Sheet L-3 shows 18,400 SF of passive turf lawn. |
| Under SNWA rules, PASSIVE LAWN IN AN HOA PARK IS NONFUNCTIONAL TURF. Continued irrigation after    |
| 12/31/2026 constitutes a statutory violation, triggering excess-use penalties and flow restrictors.|
+----------------------------------------------------------------------------------------------------+
```

#### Plan Check Consequences
Dale Yeoman's Comment L-2 explicitly requires a functional turf classification tabulation, and Comment L-3 requires amending the L-1 compliance note to cite Title 30 § 30.64.030(j). If Sundial submits Vogel & Marsh's memorandum asserting that the park is exempt because of R-2 zoning, CCPW and Comprehensive Planning will reject the response as non-compliant with state and county law.

#### Required Action: Callaway Must Choose Between Two Approvals Paths:
- **Option A (Strongly Recommended — Amenity Reinstatement):** Reintroduce the playground amenities (swings, slides, climbing structures) onto Sheet L-3 within a designated play area boundary. This immediately establishes the turf as functional "Playground Turf" under SNWA Bulletin § 3.1, fully satisfies Comments L-2 and L-3, and eliminates the nonfunctional turf designation. Crucially, this also resolves the homeowner dispute raised by Sagebrook Community Management in their letter of October 17, 2026, defusing planned public protests at the Town Advisory Board.
- **Option B (Turf Conversion):** Authorize Ocotillo to eliminate the 18,400 SF cool-season turf lawn from Sheet L-3, replacing it with a water-smart desert palette and decomposed granite. This eliminates 2.32 acre-feet/year of irrigation demand but requires preparing a formal turf conversion plan and redesigning the park.

---

### 7. CONDITION 14 DESIGN REVIEW & ENTERPRISE TAB CALENDAR TRAP

Tentative Map TM-22-500041 Condition of Approval 14 establishes:
> *"Any significant change to approved plans, including but not limited to landscape and streetscape treatments in common areas, shall be subject to design review as a public hearing before the applicable Town Advisory Board and Planning Commission prior to staff-level action."*

Advisory Comment A-3 reiterates that revisions materially altering approved character require additional public hearings.

#### The Enterprise Town Advisory Board (TAB) Calendar Reality
Review of the official Enterprise TAB 2026 Calendar reveals a fatal scheduling trap:
- **Meeting 11 (November 18, 2026):** Agenda posting date: November 8, 2026. Application submittal deadline (21 calendar days prior): **October 28, 2026**.  
  **THIS SUBMITTAL DEADLINE EXPIRED THREE DAYS AGO.**
- **Meeting 12 (December 16, 2026):** Application submittal deadline: November 25, 2026. The TAB hearing is on December 16, 2026. Forwarding recommendations to the Planning Commission and Board of County Commissioners would occur in **January 2027**.

```
+----------------------------------------------------------------------------------------------------+
| THE ENTERPRISE TAB / CONDITION 14 TIMELINE CONFLICT                                                |
|                                                                                                    |
| 10/28/2026: Submittal deadline for 11/18 TAB meeting CLOSED (Three days ago)                      |
| 10/31/2026: TODAY                                                                                  |
| 11/18/2026: Enterprise TAB Meeting (Homeowners planning public comment on Phase 3 park)           |
| 11/25/2026: Submittal deadline for 12/16 TAB meeting                                               |
| 11/30/2026: TARGET RESUBMITTAL TO CLARK COUNTY PUBLIC WORKS                                        |
| 12/16/2026: Enterprise TAB Hearing (Earliest available date if Design Review triggered)           |
| 12/18/2026: CALLAWAY TARGET RECORDATION DATE (MISSED!)                                             |
| 12/23/2026: NRS 278.360 STATUTORY MAP PRESENTATION EXPIRATION CLIFF! (ENTITLEMENT TERMINATES)     |
| 01/2027:    Planning Commission / BCC Final Action (TOO LATE — MAP ALREADY DEAD)                   |
+----------------------------------------------------------------------------------------------------+
```

#### Critical Strategy to Avoid TAB Referral
If Comprehensive Planning determines that quadrupling street trees or removing park play amenities constitutes a "significant change" under Condition 14, **the Phase 3 final map cannot record in 2026, causing the entire tentative map to expire on December 23, 2026.**

To prevent this:
1. Max Braddock must meet with Principal Planner Odette Marchetti by **November 6, 2026**, to confirm in writing that increasing street trees to 208 represents technical compliance with Title 30 rather than a discretionary redesign.
2. Callaway must choose Option A (reinstating play amenities), presenting the park revision as restoring the tentative-map approved park concept rather than creating a material change.
3. Callaway must execute immediate homeowner outreach with Sagebrook Community Management prior to November 18 to resolve the amenity dispute, ensuring homeowners do not trigger a TAB investigation during public comment.

---

### 8. STATUTORY ENTITLEMENT RUNWAY: NRS 278.360 REALITY VS. TITLE ADVICE

On October 26, 2026, Title Officer Neil Petrosyan (Silver Butte Title) sent you the following email:
> *"Yours was approved 03/14/2022, so the base clock runs to 03/14/2026 — but the statute also allows a two-year extension on top of that, so effectively you're covered out to 03/14/2028. That's a comfortable window... I wouldn't sweat the December recordation date from a TM standpoint..."*

**THIS ADVICE IS FACTUALLY INCORRECT UNDER NEVADA LAW AND THREATENS THE SURVIVAL OF THE PROJECT.**

#### Governing Law: NRS 278.360 and TM-22-500041 Conditions of Approval
NRS 278.360 dictates the presentation schedule for phased subdivision maps:
1. The **first final map** in a series must be presented within four (4) years of tentative map approval (03/14/2022 to 03/14/2026). This condition was satisfied when Desert Sage Ranch Phases 1 & 2 recorded on **December 23, 2024** (Book 175, Page 60).
2. For **successive final maps** in the series, NRS 278.360 states:
   > *"Each successive final map in the same series must be presented within two (2) years after the date on which the immediately preceding final map is recorded. Failure to timely present a map terminates all proceedings relating to the subdivision."*
3. Two years from the recording date of Phases 1 & 2 (12/23/2024) is **DECEMBER 23, 2026**.
4. NRS 278.360 explicitly prohibits granting pre-expiration extensions: *"An extension shall not be granted before the applicable presentation period has expired."* (Comprehensive Planning formally rejected Callaway's premature extension request on December 30, 2024, on this precise statutory ground).

**Conclusion:** If the Phase 3 final map is not presented to Clark County on or before **December 23, 2026**, Tentative Map TM-22-500041 terminates automatically by operation of law. All entitlement rights to the 60 lots will be permanently lost. The December 18 target provides only three business days of margin.

---

### 9. AUDIT OF OTHER INTER-AGENCY BLOCKERS ON THE RESUBMITTAL TRACKER

A review of the Resubmittal Tracker reveals several additional critical-path blockers that must be cleared concurrently:

1. **Expired Sewer Capacity Confirmation (Comment U-2; Condition 22):**
   - Confirmation CCWRD-SCC-2025-1187 (issued 09/30/2025) explicitly states in Paragraph 3: *"This capacity confirmation is issued for a period of one (1) year... and expires on 09/30/2026. Expired confirmations require re-review of downstream conditions and reissuance before plan approval."*
   - **This confirmation is currently EXPIRED.** It lapsed 31 days ago. Comment U-2 mandates current letters at resubmittal. An emergency renewal request must be filed with Vince Alarid (CCWRD) immediately.
2. **CCFD Permanent Second Point of Access (Comment F-1; Condition 11; Sheet FA-1):**
   - Sheet FA-1 depicts only one permanent point of access to Windmill Lane (Keynote 1). Keynote 4 acknowledges that the temporary Phase 2 construction stub must be removed upon Phase 3 buildout per CCFD File FPB-2023-0417 (Capt. Emmett Roark).
   - CCFD requires two permanent access points for developments exceeding 30 lots. Sundial must coordinate an emergency access easement or alternative second point of access with Capt. Roark prior to resubmittal.
3. **Landscape Architect Seal (Comment L-1):** All revised sheets L-1 through L-4 must bear the professional wet seal and signature of Bea Cortland, PLA (NV LA #742).
4. **Owner's Certificate & Statutory Notary (Comment M-5):** Callaway must designate the current authorized corporate signatory and execute the Owner's Certificate and Dedication on FM-26-100290 using current Nevada statutory notary acknowledgment language.
5. **Stand-Alone ESC Sheet & NDEP NOI (Comment G-8):** Sundial must prepare a dedicated Erosion and Sediment Control sheet referencing active coverage under NDEP Permit NVR100000.
6. **Drainage & Detention Revisions (Comments G-1, G-3, G-4, G-5):** Reconcile the 6.4 acre-foot detention allocation on Sheet C-9 with the Master Drainage Study, detail the spillway apron, and add HGL/EGL profiles on C-8.

---

### 10. COMPREHENSIVE ACTION PLAN: WHAT HAS TO HAPPEN FIRST

To achieve an approvable resubmittal by November 30, 2026, and secure final map presentation prior to the December 23 statutory cliff, the following four-phase action plan must be strictly executed:

```
+----------------------------------------------------------------------------------------------------+
| MASTER RESUBMITTAL & CLEARANCE TIMELINE                                                            |
|                                                                                                    |
| [PHASE I: NOV 2 - NOV 6]       [PHASE II: NOV 4 - NOV 13]      [PHASE III: NOV 13 - 20] [PHASE IV] |
| Client Direction & Intake      Technical Design & Conflicts    Sheet Finalization & QC  Submittal  |
|                                                                                                    |
| • Park Option A/B direction    • Dry utility conflict overlay  • Ocotillo seals L-set   • Package  |
| • CCWRD sewer renewal request  • Hydraulic & water demand calc • Civil C-1..C-9 sealed    delivered |
| • CCFD access meeting (Roark)  • Alternative Standards prep    • Estimate re-sealed       to CCPW  |
| • Planning sign-off (Marchetti)• Reconcile plant schedules     • QA/QC back-check       • Expedited|
| • Surety notified (+$89k bond) • Draft response matrix         • Title exceptions on map  review   |
|                                                                                           begins   |
|   11/02                  11/06 11/09                     11/13 11/16              11/20   11/30    |
+----------------------------------------------------------------------------------------------------+
```

#### Phase I: Immediate Client Direction & Agency Clearances (November 2 – November 6)
1. **Client Park Turf Decision (Due 11/04 — Garrett Nokes):** Direct Ocotillo to adopt Option A (reinstate play amenities to qualify as functional playground turf) or Option B (full desert conversion).
2. **Surety Notification (Due 11/04 — Garrett Nokes):** Authorize an $89,000 increase in the Phase 3 Performance Bond basis ($5,446,000) and notify Aegis Surety Company.
3. **Owner's Signatory Designation (Due 11/04 — Garrett Nokes):** Provide authorized corporate signatory information and notary acknowledgment for FM-26-100290.
4. **Sewer Capacity Renewal (Due 11/02 — Max Braddock):** Submit emergency written application to Vince Alarid (CCWRD) to renew CCWRD-SCC-2025-1187.
5. **CCFD Second Access Coordination (Due 11/04 — Max Braddock):** Meet with Capt. Emmett Roark to establish an acceptable permanent secondary access design or emergency access easement.
6. **Comprehensive Planning Administrative Sign-Off (Due 11/06 — Max Braddock):** Secure written concurrence from Odette Marchetti that landscape revisions are technical code compliance and will not be referred to Enterprise TAB.
7. **Sagebrook HOA Outreach (Due 11/06 — Garrett Nokes):** Respond to Sagebrook Community Management's letter, confirming reinstatement of play amenities to neutralize homeowner protests at the 11/18 TAB meeting.

#### Phase II: Technical Coordination, Hydraulic Modeling & Alternative Standards (November 4 – November 13)
8. **Spatial Conflict Layout (Due 11/10 — Bea Cortland & Priya Sarwal):** Overlay 208 street trees against Sheet DU-1/DU-2 joint trench, 14 fire hydrants, storm drain Line A, and sight triangles.
9. **Alternative Standards Application (Due 11/13 — Bea Cortland, if needed):** If utility corridors prevent 20-foot spacing, prepare formal Title 30 § 30.64.050 Alternative Standards justification for staff approval.
10. **Hydraulic & Water Demand Recalculation (Due 11/13 — Bea Cortland):** Update `DSR_Irrigation_Demand_Worksheet.xlsx` and `DSR_Ph3_Water_Budget.xlsx` (+300,000 gal/yr tree allowance); verify 1½-inch meter and backflow sizing.
11. **Civil & Drainage Reconciliations (Due 11/13 — Priya Sarwal):** Finalize C-1 drainage narrative, C-9 basin summary (6.4 AF), C-8 HGL/EGL profiles, spillway apron details, and stand-alone ESC sheet.

#### Phase III: Sheet Finalization, Sealing & QA/QC (November 13 – November 20)
12. **Landscape Sheet Finalization (Due 11/18 — Bea Cortland):** Reconcile L-1 schedule counts with L-2/L-3 plans (Comment L-6); dimension 10-foot curb setbacks (Comment L-5); apply professional seal and signature (Comment L-1).
13. **Survey & Mapping Finalization (Due 11/18 — Marcus Halvorsen, PLS):** Plot Schedule B exceptions from Title Report SBT-2026-04188 on FM-26-100290; reconcile Basis of Bearings with Phase 1 & 2 map; finalize closure sheets.
14. **Estimate Re-Sealing (Due 11/20 — Jonathan R. Kessler, PE):** Update itemized estimate ($5,446,000) and issue sealed PDF.
15. **Internal QA/QC Back-Check (Due 11/20 — Priya Sarwal & Max Braddock):** Execute complete Sundial QC-05 check across all 13 sheets and comment response matrix.

#### Phase IV: Packaging, Expedited Review Request & Submittal (November 20 – November 30)
16. **Submittal Assembly (Due 11/25 — Max Braddock):** Assemble annotated comment matrix, red-lined plan sets, clean revised sheets, updated water budget, renewal letters, and title report.
17. **Expedited Review Negotiation (Due 11/25 — Max Braddock):** Dale Yeoman's letter notes standard second-round review takes 20 business days (which would push approval to January 2027). Sundial and Callaway must meet with Dale Yeoman to negotiate an expedited **10-to-12 business day review window**, predicated on delivering a clean, fully pre-coordinated package with zero open agency defects.
18. **Counter Delivery (Due 11/30 — Sundial):** Hand-deliver bond sets and upload digital files to Clark County Development Review counter under IP-26-100288 and FM-26-100290.

---

### SIGNATURE BLOCK

Respectfully submitted,

**SUNDIAL ENGINEERING & SURVEY**

\
__________________________________________  
**Max Braddock**  
Project Manager  
Direct: (702) 555-0144  
Email: mbraddock@sundialensv.com  

\
__________________________________________  
**Jonathan R. Kessler, PE**  
Principal-in-Charge / Engineer of Record  
Nevada Professional Engineer License No. 14203  

---
*cc: Adrienne Vogel, Esq. — Vogel & Marsh LLP*  
*Bea Cortland, PLA — Ocotillo Landscape Studio*  
*Sundial Project File 2201-03 (Desert Sage Ranch Phase 3)*  
"""

with open('output/MEMORANDUM_Callaway_Land_Partners_2026-10-31.md', 'w') as f:
    f.write(memo_md)

print("Saved output/MEMORANDUM_Callaway_Land_Partners_2026-10-31.md")
