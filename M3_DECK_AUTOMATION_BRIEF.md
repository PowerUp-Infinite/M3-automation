# M3 Deck Automation — Claude Code Project Brief
*v2 — updated: single sheet source, corrected transition table logic, variable sheet names*

## What This Project Does
Automates the generation of a client Portfolio Transition deck (PowerPoint .pptx) from a client-specific Excel workbook. You provide a client name and their Excel file; the script outputs a fully populated, ready-to-share deck.

---

## Files You Will Work With

Place both of these in your project folder:

- `template.pptx` — the master deck (Hari Vootori's completed deck is your template)
- `client_data.xlsx` — the client's Excel workbook

The script will be called like this:
```bash
python generate_deck.py --client "Priya Sharma" --excel "priya_data.xlsx" --output "priya_deck.pptx"
```

---

## Excel Sheet Structure — CRITICAL RULES

### Rule 1: Only ONE sheet is ever used — the PF_Curation sheet
Every single piece of data for every slide comes from the one sheet whose name starts with `PF_Curation`. Do NOT read any other sheet. The sheet name will vary per client file (e.g. `PF_Curation_10032026`, `PF_Curation_15042026`) so always find it dynamically:

```python
def get_curation_sheet(wb):
    for name in wb.sheetnames:
        if name.startswith("PF_Curation"):
            return wb[name]
    raise ValueError("No PF_Curation sheet found in workbook")
```

### Rule 2: Always open Excel with data_only=True
```python
import openpyxl
wb = openpyxl.load_workbook(excel_path, data_only=True)
ws = get_curation_sheet(wb)
```
This ensures you read computed values, not raw formulas.

### Rule 3: Row numbers are NOT fixed — detect sections by their header rows
The PF_Curation sheet has multiple sections separated by blank rows. Each client's file may have different numbers of schemes, so sections will start at different rows. Always scan for section headers rather than hardcoding row numbers.

---

## PF_Curation Sheet Layout

The sheet has 4 sections, each separated by blank rows. Here is how to detect each one:

### Section 1 — Risk Group Pivot (sell-side summary)
**Detect by:** Row where column A = `"Row Labels"` AND column B = `"Total Selected Value"`

Structure after the header row:
- Risk Group rows: column A = e.g. `"1) Aggressive"`, `"2) Balanced"`, `"3) Conservative"`, `"Hybrid"`, `"Debt Like"`, `"Global"` — subtotal rows
- Subcategory rows: column A = e.g. `"MID_CAP"`, `"FLEXI_CAP"` — indent level 1
- Individual scheme rows: column A = fund name — have non-None value in `Action` column
- Grand Total row: column A = `"Grand Total"`

Columns in this section:
```
A: Row Labels
B: Total Selected Value
C: % Allocation
D: Sum of MG
E: Exit Load
F: STCG
G: Total Impact (EL+STCG)
H: LTCG
I: Fund Rating
J: Action        <- "Sell completely", "Retain completely", "Partial sell", "Partial retain"
K: Reason
```

**Key data from this section:**
- Grand Total row → column H (LTCG) and column F (STCG) → for tax liability on Slide 9

---

### Section 2 — Scheme-Level Sell/Retain Detail
**Detect by:** First row where column A = `"FUND_NAME"` AND column C = `"RISK_GROUP_L0"`

Columns in this section:
```
A: FUND_NAME
B: FOLIO_NUMBER
C: RISK_GROUP_L0
D: UPDATED_BROAD_CATEGORY_GROUP
E: UPDATED_SUBCATEGORY
F: TOTAL_UNITS
G: TOTAL_VALUE
H: Scheme % of SPF Value
I: Action            <- "Sell completely", "Retain completely", "Partial sell", "Partial retain"
J: Reason
K: Redemption Factor
L: UNITS_REDEEMED
M: Redemption Value  <- use this as Sell Value
N: Redemption Value as % SPF Value
O: UNITS_RETAINED
P: Retained Value as % of SPF
Q: Retained Value    <- use this as Retain Value
R: EXIT_LOAD
S: Max STCG
T: EL+STCG
```

One row per scheme. Ends at the Grand Total row (column A = "Grand Total").

**Data used for:** Slides 17 (Complete Sell), 19 (Complete Retain), and partial slides if applicable.

---

### Section 3 — Allocation Timeline (transition plan per risk group)
**Detect by:** Row where column A = `"Row Labels"` AND column B = `"Sum of TOTAL_VALUE"`

Columns in this section (detect positions from header row — do NOT hardcode):
```
Row Labels                   <- Risk Group name e.g. "1) Aggressive"
Sum of TOTAL_VALUE           <- Current Value
Sum of Scheme % of SPF Value
Sum of Redemption Value
Sum of Redemption Value as % SPF Value
Sum of Retained Value as % of SPF
Sum of Retained Value
Sum of EXIT_LOAD
Sum of Max STCG
Sum of EL+STCG
Sum of Max LTCG
Sum of (EL+STCG)/SPFV
Sum of Sold at D0            <- fraction of SPF sold at this tranche
Sum of Sold at D30
Sum of Sold at D60
Sum of Sold at D90
Sum of Sold at D120
Sum of Sold at D150
Sum of Amount Sold at D0     <- rupee amount sold at this tranche
Sum of Amount Sold at D30
Sum of Amount Sold at D60
Sum of Amount Sold at D90
Sum of Amount Sold at D120
Sum of Amount Sold at D150
```

Rows: one per Risk Group, plus Grand Total at the end.

---

### Section 4 — Buy Plan (new scheme allocations + SIP data)
**Detect by:** The SECOND row where column A = `"FUND_NAME"` (Section 2 is the first, Section 4 is the second occurrence)

Columns in this section (detect positions from header row — do NOT hardcode):
```
FUND_NAME
RISK_GROUP_L0
UPDATED_BROAD_CATEGORY_GROUP
UPDATED_SUBCATEGORY
Retained Value as % of PF
Buy as % of PFV
Buy Value
Total Value as % of PF
Total Value
SIP Allocation %
SIP Amount
Action              <- "Fresh buy", "Buy more", "Retain completely"
Reason
Buy Plan
Buy % of PFV at D0
Buy % of PFV at D30
Buy % of PFV at D60
Buy % of PFV at D90
Buy % of PFV at D120
Buy % of PFV at D150
```

**Data used for:** Slide 4 (SIP summary), Slides 5/6/7 (SIP schemes), Slides 13/14/15 (corpus schemes), Slide 21 (Fresh Buy/Buy More), and the transition plan calculation.

---

## Transition Plan Calculation — CORRECT LOGIC

This is the most critical section. Read carefully.

### What the transition table in Slide 9 shows:
For each risk group, the table shows the portfolio value and % allocation at Current state, and the ideal state after each tranche: D0, D30, D60, D90, D120, D150.

The value at each milestone represents the ideal portfolio state at that point — accounting for BOTH what was sold AND what was bought.

### How to calculate — use the Allocation columns from Section 4:

In Section 4, each scheme row has columns: `Buy % of PFV at D0`, `Buy % of PFV at D30`, ..., `Buy % of PFV at D150`

These represent the **cumulative target allocation fraction** of total portfolio value for that scheme at each milestone. This already encodes the full picture (sold + bought + retained) and IS the ideal allocation at each point.

**For each risk group at each milestone Dx:**
```
Ideal % at Dx = SUM of (Buy % of PFV at Dx) for all schemes in that risk group (Section 4)
Ideal Value at Dx = Ideal % at Dx × Total Portfolio Value
```

**Total Portfolio Value** = Grand Total row, `Sum of TOTAL_VALUE` column, from Section 3.

**Current value** = Section 3, `Sum of TOTAL_VALUE` column, for each Risk Group row.
**Current %** = Section 3, `Sum of Scheme % of SPF Value` column, × 100 for each Risk Group row.

### Why this approach:
The `Buy % of PFV at Dx` columns in Section 4 are the computed final allocation answer for each tranche. They already net out the selling and buying across all tranches. Attempting to reconstruct it from separate sell + buy amounts is error-prone and unnecessary — the columns are already there.

### Equity row:
Equity = sum of Aggressive + Balanced + Conservative. Compute this from their individual values — do NOT look for an "Equity" row in the Excel.

### Hybrid, Debt-Like, Global rows:
Read directly from Section 3 (current) and aggregate Section 4 schemes by their `RISK_GROUP_L0`.

### Total row:
Always = sum of all risk groups. Should equal `Sum of TOTAL_VALUE` from Grand Total in Section 3.

---

## Slides to Automate — Exact Instructions

### SLIDE 1 — Cover
- Find the text shape containing `"Hari Vootori"` and replace with `{client_name}`
- Do not touch the date on this slide

### SLIDE 2 — Welcome
- Find text containing `"Hari"` (the greeting) and replace with first name only
- Find the date text (pattern: number + month name + year, e.g. `"9 mar 2026"`) and replace with today formatted as `"D Mon YYYY"` e.g. `"15 Mar 2026"`

### SLIDE 3 — Skip entirely

---

### SLIDE 4 — Ideal SIP Strategy (summary table)

**Data source:** Section 4 of PF_Curation

Build a hierarchical table grouped by risk group → subcategory:

```
| Portfolio      | Allocation          | # of funds |
|----------------|---------------------|------------|
| Equity         | [sum equity SIP]    | [count]    |
|  Aggressive    | [sum aggr SIP]      | [count]    |
|   Mid Cap      | [SIP amount]        | [count]    |
|   Small Cap    | [SIP amount]        | [count]    |
|  Balanced      | [sum bal SIP]       | [count]    |
|   Flexi Cap    | [SIP amount]        | [count]    |
|   Value&Contra | [SIP amount]        | [count]    |
|  Conservative  | [sum cons SIP]      | [count]    |
|   Large Cap    | [SIP amount]        | [count]    |
| Hybrid         | [sum hybrid SIP]    | [count]    |
|   Dynamic Asset| [SIP amount]        | [count]    |
|   Multi Asset  | [SIP amount]        | [count]    |
| Total          | [total SIP amount]  | [total]    |
```

The exact rows depend on this client's data. Add or remove rows as needed — never show empty rows.

**Formatting:**
- SIP Amount: `40K`, `15K`, `7.5K` (divide by 1000, drop .0 suffix)
- Allocation %: `SIP Allocation %` × 100, append `%`

**Subcategory display name mapping:**
```
MID_CAP                  -> Mid Cap
SMALL_CAP                -> Small Cap
FLEXI_CAP                -> Flexi Cap
VALUE_AND_CONTRA         -> Value & Contra
LARGE_CAP                -> Large Cap
DYNAMIC_ASSET_ALLOCATION -> Dynamic Asset
MULTI_ASSET_ALLOCATION   -> Multi Asset
AGGRESSIVE_ALLOCATION    -> Aggressive Hybrid
COMMODITY_GOLD           -> Gold
COMMODITY_SILVER         -> Silver
INDEX_LARGE_CAP          -> Large Cap Index
ELSS_TAX_SAVINGS         -> ELSS
```

Find the main data table on Slide 4 (not the three icon-callout boxes at the bottom). Add/remove rows by copying/removing XML. Total row is always last.

---

### SLIDES 5, 6, 7 — SIP Strategy at Scheme Level

**Data source:** Section 4 of PF_Curation

Template: Slide 5 = Aggressive, Slide 6 = Balanced, Slide 7 = Conservative + Hybrid.

**Per scheme row, update:**
- Scheme Name
- Monthly Allocation % = `SIP Allocation %` × 100, e.g. `7.5%`
- Leave all other columns blank (Upside/Downside, Power Rank, 5Y Rolling Return, AUM) — will be automated later

**Rules:**
- One risk group per table. Never mix.
- Maintain spacing between scheme rows as in template
- More schemes than template: add rows (deep copy XML from existing data row)
- Fewer schemes: delete extra rows
- Zero schemes for a risk group: delete that table block entirely
- Conservative and Hybrid are on the same slide (Slide 7): apply add/remove independently to each table

---

### SLIDE 8 — DELETE THIS SLIDE
Remove Slide 8 entirely for every client. No exceptions.

**After deleting Slide 8, all subsequent slide indices shift by -1. Always find slides by scanning for known text content, never by hardcoded index.**

---

### SLIDE 9 — Transition Plan Table

**Data sources:**
- Section 3 → Current value per risk group
- Section 4 → `Buy % of PFV at Dx` columns → ideal allocation at each milestone
- Section 1 Grand Total → LTCG and STCG for tax liability text

**Table structure:**
```
| Portfolio Allocation | Current    | Ideal-D0   | Ideal-D30  | Ideal-D60  | Ideal-D90  | Ideal-D120 | Ideal-D150 |
| Equity               | 1.5Cr 98%  | 1.5Cr 96%  | ...
|  Aggressive          | 62.8L 40%  | ...
|  Balanced            | 48.1L 30%  | ...
|  Conservative        | 43.7L 28%  | ...
| Hybrid               | 3.4L   2%  | ...
| Total                | 1.6Cr 100% | ...
```

Each cell = `[formatted value]  [%]` — match whatever combined format the template uses.

Add/remove rows to match this client's risk groups. If no Hybrid, remove that row. If they have Debt-Like or Global, add it. Total row always last.

**Tax Liability text (below/beside the table):**
- Find the text shape on this slide containing "Tax Liability"
- Replace with: `Tax Liability: INR {format_inr(ltcg + stcg)}`
- Sub-line: `LTCG {format_inr(ltcg)} + STCG {format_inr(stcg)}`

---

### SLIDES 10, 11, 12 — Skip entirely

---

### SLIDES 13, 14, 15 — Corpus Strategy at Scheme Level

Same logic as Slides 5/6/7 except:
- **Data source:** Section 4 of PF_Curation (same sheet — no other sheet used)
- Use `Total Value as % of PF` (not SIP Allocation %) for the allocation column
- Use `Total Value` (not SIP Amount) for the value column
- Column header on these slides says `% Allocation` not `% Monthly Allocation`
- All other rules identical: one risk group per table, add/remove rows, leave Upside/Downside/PowerRank/AUM/5Y blank

---

### SLIDES 16, 18, 20 — Section header slides — Skip entirely

---

### SLIDE 17 — Complete Sell
**Source:** Section 2, filter `Action == "Sell completely"`
**Columns:** `Scheme Name | Sell Value | Reason`
- Sell Value = `Redemption Value` via `format_inr()`
- Add/remove rows to match count

### SLIDE 19 — Complete Retain
**Source:** Section 2, filter `Action == "Retain completely"`
**Columns:** `Scheme Name | Retained Value | Reason`
- Retained Value = `Retained Value` column via `format_inr()`

### SLIDE 21 — Fresh Buy & Buy More
**Source:** Section 4, filter `Action == "Fresh buy"` and separately `Action == "Buy more"`
**Columns:** `Scheme Name | Buy Value | Reason`
- Buy Value = `Buy Value` column via `format_inr()`
- Fresh buy and Buy more get separate tables (as in template)
- If no Buy More schemes exist, remove/skip that table

**Also handle if present:**
- Section 2, `Action == "Partial sell"` → separate table, `Scheme Name | Partial Sell Value | Reason`
- Section 2, `Action == "Partial retain"` → separate table, `Scheme Name | Retained Value | Reason`
- These get their own slides between 17 and 19 if they exist (duplicate template slide format)

---

## Number Formatting — Use This Everywhere

```python
def format_inr(value):
    if value is None:
        return "-"
    cr = 10_000_000   # 1 Crore
    L  = 100_000      # 1 Lakh
    K  = 1_000        # 1 Thousand

    if value >= cr:
        return f"{value/cr:.1f}Cr"
    elif value >= L:
        return f"{value/L:.1f}L"
    elif value >= K:
        n = round(value / K, 1)
        return f"{n:g}K"   # :g strips trailing zeros e.g. 50.0 -> 50
    else:
        return f"{round(value)}"
```

For the SIP table (Slide 4) where the template shows `50K` without ₹, use the above without ₹ prefix.
For money values elsewhere, prepend `₹` as shown in the template.

---

## How to Manipulate the PowerPoint

Use `python-pptx`. Key patterns:

### Text replacement (handles split runs)
```python
def replace_text_preserving_format(slide, old_text, new_text):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full_text = "".join(run.text for run in para.runs)
            if old_text in full_text:
                new_full = full_text.replace(old_text, new_text)
                if para.runs:
                    para.runs[0].text = new_full
                    for run in para.runs[1:]:
                        run.text = ""
```

### Add table row (deep copy, preserves formatting)
```python
from pptx.oxml.ns import qn
import copy

def add_table_row(table, copy_from_row_idx):
    tbl = table._tbl
    rows = tbl.findall(qn('a:tr'))
    source_row = rows[copy_from_row_idx]
    new_row = copy.deepcopy(source_row)
    # Clear text in all cells of new row
    for tc in new_row.findall(qn('a:tc')):
        for r in tc.findall('.//' + qn('a:r')):
            t = r.find(qn('a:t'))
            if t is not None:
                t.text = ""
    # Insert before last row (the total row)
    rows[-1].addprevious(new_row)
```

### Delete table row
```python
def delete_table_row(table, row_idx):
    tbl = table._tbl
    rows = tbl.findall(qn('a:tr'))
    tbl.remove(rows[row_idx])
```

### Set cell text (preserves font/color)
```python
def set_cell_text(cell, text):
    tf = cell.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = str(text)
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run().text = str(text)
```

### Delete a slide
```python
def delete_slide(prs, slide_index):
    slide_id_list = prs.slides._sldIdLst
    sldId = slide_id_list[slide_index]
    rId = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    slide_id_list.remove(sldId)
    del prs.part.related_parts[rId]
```

### Find slide by text content
```python
def find_slide_by_text(prs, search_text):
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and search_text in shape.text_frame.text:
                return i
    return None
```

---

## Section Detection Reference Implementation

```python
def detect_sections(ws):
    """
    Scan PF_Curation sheet and return the 1-based start row of each section.
    Returns dict: {section1, section2, section3, section4}
    """
    sections = {}
    fund_name_rows = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        a = row[0] if row else None
        b = row[1] if len(row) > 1 else None

        if a == "Row Labels" and b == "Total Selected Value":
            sections['section1'] = row_idx
        elif a == "Row Labels" and b == "Sum of TOTAL_VALUE":
            sections['section3'] = row_idx
        elif a == "FUND_NAME":
            fund_name_rows.append(row_idx)

    if len(fund_name_rows) >= 1:
        sections['section2'] = fund_name_rows[0]
    if len(fund_name_rows) >= 2:
        sections['section4'] = fund_name_rows[1]

    return sections
```

Then read each section by iterating from its header row until you hit a blank row or "Grand Total" in column A.

**For reading column positions dynamically from a header row:**
```python
def col_map(header_row):
    """Return dict of column_name -> 0-based index from a header row tuple."""
    return {v: i for i, v in enumerate(header_row) if v is not None}
```

Use `col_map` on every section header row. Never hardcode `row[12]`.

---

## Build Order — Phases

Build one phase at a time. Test after each before proceeding.

**Phase 1 — Project Setup**
1. Create `m3_deck_automation/` folder
2. `requirements.txt`: `python-pptx==0.6.23`, `openpyxl==3.1.2`
3. `pip install -r requirements.txt`
4. `generate_deck.py` with argparse: `--client`, `--excel`, `--output`
5. `excel_reader.py` and `deck_writer.py` stubs
6. Test: open template.pptx, save as output, no errors

**Phase 2 — Name & Date (Slides 1 & 2)**
Update client name and date. Test visually.

**Phase 3 — Delete Slide 8**
Find and delete. Confirm slide count drops by 1.

**Phase 4 — SIP Summary Table (Slide 4)**
Read Section 4, build hierarchy, update table with add/remove rows.

**Phase 5 — SIP Scheme Slides (Slides 5, 6, 7)**
Group Section 4 schemes by risk group, update scheme tables.

**Phase 6 — Transition Plan Table (Slide 9)**
Section 3 for current values + Section 4 Buy % columns for milestone allocations. Tax from Section 1.

**Phase 7 — Corpus Strategy Slides (Slides 13, 14, 15)**
Same as Phase 5 but use Total Value % and Total Value columns from Section 4.

**Phase 8 — Sell/Retain/Buy Tables (Slides 17, 19, 21)**
Filter Section 2 and Section 4 by Action type, update tables.

**Phase 9 — End-to-End Test**
Run with Hari's file, confirm output matches template. Run with modified data to verify dynamic rows.

---

## Common Pitfalls

1. **Slide index drift** — always find slides by text content after deleting Slide 8.
2. **Split runs** — use the concatenate-then-replace pattern for text replacement.
3. **Deep copy required** — `copy.deepcopy()` on XML rows, never shallow copy.
4. **Transition table** — use `Buy % of PFV at Dx` from Section 4, not reconstructed from sell+buy amounts.
5. **None = blank row** — when scanning for section boundaries, `None` in column A means blank/separator.
6. **Dynamic column positions** — always use `col_map()` on the section header row.
7. **Font on added rows** — copy from a middle data row, not the header or total row.
8. **Multiple tables per slide** — use `[s for s in slide.shapes if s.has_table]`, then identify by first-cell content or dimensions.

---

## Starting Prompt for Claude Code

> "I want to build a Python script that automates a client investment deck (PowerPoint) from an Excel workbook. I have a complete project brief in `M3_DECK_AUTOMATION_BRIEF.md` in this folder. Please read that file fully before writing any code. Then let's start with Phase 1 — project setup. My template deck is `template.pptx` and a sample Excel is `client_data.xlsx`, both in this folder."

After each phase succeeds: "Phase N is working and tested. Let's move to Phase N+1."

---

*Brief v2 — after full analysis of Hari_Vootori_Transition_Plan.pptx and Hari_Vootori_transition_plan_100326.xlsx*
