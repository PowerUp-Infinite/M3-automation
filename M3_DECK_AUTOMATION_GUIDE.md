# M3 Deck Automation — How It Works
> A plain-English guide for the team. Start here.

---

## 1. What Does This Script Do?

Given a **client Excel file** and a set of **reference CSV/XLSX files**, it automatically fills in a PowerPoint template and saves a ready-to-present deck.

---

## 2. End-to-End Pipeline

```
python app.py --client "Ramesh Shah" --excel client_data.xlsx
                           │
              ┌────────────▼────────────┐
              │      app.py : main()    │
              └────────────┬────────────┘
                           │
           ┌───────────────┼───────────────┐
           │               │               │
           ▼               ▼               ▼
   read_excel()   load_reference_data()   generate_deck()
   (Excel file)   (4 reference files)    (writes PowerPoint)
```

---

## 3. Reading the Client Excel (`excel_reader.py`)

The script looks for a sheet whose name starts with **PF_Curation**.

Inside that sheet it detects **4 sections** by searching for specific header rows:

```
┌─────────────────────────────────────────────────────────────────┐
│  PF_Curation Sheet                                              │
│                                                                 │
│  SECTION 1  ← found when col-A = "Row Labels",                 │
│               col-B = "Total Selected Value"                    │
│               → Current portfolio BY CATEGORY                  │
│               → Contains: Exit Load, LTCG, STCG per category   │
│                                                                 │
│  SECTION 2  ← found at FIRST "FUND_NAME" header row            │
│               → Current portfolio BY SCHEME (master table)     │
│               → KEY COLUMNS: FUND_NAME, RISK_GROUP_L0,         │
│                 Action, Redemption Value, Retained Value        │
│                                                                 │
│  SECTION 3  ← found when col-A = "Row Labels",                 │
│               col-B = "Sum of TOTAL_VALUE"  (old format)       │
│               OR "Sum of Current Value Amount" (new format)     │
│               → Current + transition pivot BY RISK GROUP        │
│                                                                 │
│  SECTION 4  ← found at LAST "FUND_NAME" header row             │
│               → Ideal portfolio BY SCHEME                      │
│               → KEY COLUMNS: FUND_NAME, RISK_GROUP_L0,         │
│                 SIP Amount, Buy Value Amount,                   │
│                 Cumm Buy Amount M1–M6 (or D0–D150)              │
└─────────────────────────────────────────────────────────────────┘
```

> **Old vs New Excel format** is auto-detected:
> - Old: uses `Sum of TOTAL_VALUE`, `Cummulative Buy Amount at D0..D150`, `SIP Amount`, `Buy Value`
> - New: uses `Sum of Current Value Amount`, `Cumm Buy Amount in M1..M6`, `SIP Allocation Amount`, `Buy Value Amount`

---

## 4. Loading Reference Data (`reference_data.py`)

Four files are read from the project folder:

| File | Columns Used | What It Gives Us |
|------|-------------|------------------|
| `AUM_31Jan.csv` | ISIN, AUM, FUND_NAME | AUM in Crores; name↔ISIN maps |
| `Powerranking.csv` | ISIN, POWERRANK | Power Rank integer |
| `upside_downside_mar.xlsx` | col B=ISIN, col D=downside, col E=upside | Upside \| Downside capture ratios |
| `Rolling_Returns_Mar.csv` | ENTITYID, ROLLING_PERIOD=60, RETURN_VALUE | 5-Year rolling return per scheme AND per category |

All four are merged into a single `lookup[ISIN]` dict plus a `rr_category[SUBCATEGORY]` dict.

> **ISIN cross-validation**: when filling scheme slides, if the ISIN from the Excel maps to a *different* fund name in the AUM file, the script discards the Excel ISIN and falls back to name-based lookup (handles scrambled ISINs in client files).

---

## 5. What Slides Get Populated

```
Template Slide          Data Source          Key Logic
─────────────────────────────────────────────────────────────────────
Slide 1  (Cover)        client_name          Replace "Hari Vootori"
Slide 2  (Welcome)      client_name, date    Replace first name + date
Slide 3  (Recap)        client_name          Replace first name
Slide 4  (SIP table)    Section 4            SIP %, amounts by RG
Slides 5-9 (SIP schemes)Section 4 + Ref     Per-scheme: %, U|D, Rank, 5Y, AUM
Slide 10 (SIP sep.)     —                   Separator slide, untouched
Slide 11 (Transition)   Sections 1,2,3,4    6-milestone allocation table
Slides 12-19(Corpus)    Section 4 + Ref     Same per-scheme data as SIP
Sell slide              Section 2            Complete + Partial Sell tables
Retain slide            Section 2            Complete + Partial Retain tables
Buy slide               Section 4            Fresh Buy + Buy More tables
─────────────────────────────────────────────────────────────────────
Note: slides without data are automatically deleted.
```

---

## 6. How the Transition Table Is Built (Slide 11)

This is the most complex slide. It shows:
- **Current column**: how much money sits in each Risk Group today
- **D0 → D150 columns**: what the ideal allocation looks like at each monthly milestone

### Step-by-step logic

```
CURRENT COLUMN (what the client holds today, per Risk Group):
  Source: Section 2 (current portfolio, scheme-level)
  For each scheme row:
    current_value = Retained Value + Redemption Value
    (sell completely → Redemption = all, Retained = 0
     partial sell   → Redemption + Retained = full current value
     retain         → Retained = full current value, Redemption = 0)
  rg_current[RG] = SUM(current_value) for all schemes in that Risk Group

IDEAL COLUMNS (D0 → D150, what the portfolio looks like at each milestone):
  Source: Section 4 (ideal portfolio, scheme-level)
  Section 4 has 'Allocation M1', 'Allocation M2', ..., 'Allocation M6' per scheme.
  These columns already contain the NET allocation at each milestone:
    - Retained schemes: constant value across all milestones
    - Partial sell:     reduced value (current - redemption) after sell milestone
    - Fresh Buy:        growing cumulative buy amounts
    - Buy More:         existing retained + growing cumulative buys
    - Sell completely:  goes to 0 after sell milestone
  rg_ideal[RG][D0]  = SUM(Allocation M1) for all section4 schemes in that RG
  rg_ideal[RG][D30] = SUM(Allocation M2)
  rg_ideal[RG][D60] = SUM(Allocation M3)
  rg_ideal[RG][D90] = SUM(Allocation M4)
  rg_ideal[RG][D120]= SUM(Allocation M5)
  rg_ideal[RG][D150]= SUM(Allocation M6)

TOTAL PORTFOLIO VALUE (denominator for %):
  Priority: Section 1 Grand Total "Total Selected Value"
          > Section 3 Grand Total
          > SUM(rg_current) from Section 2

All values shown as:  "INR value | X%"
```

> **Key insight**: Section 4's `Allocation M1-M6` columns are the single source of truth for the ideal portfolio. They already account for all sell/buy/retain actions per scheme. The code simply sums them by Risk Group — no manual retained + buy calculations needed.

---

## 7. How Scheme Slides Are Filled

```
Section 4 (ideal portfolio, filtered by non-zero SIP Amount or corpus %)
    │
    ▼ Group by RISK_GROUP_L0 → UPDATED_SUBCATEGORY
    │
    For each Risk Group:
      Find the matching template slide
      Split schemes into pages of 4
      ┌─ For each scheme (slot on slide):
      │   1. Get ISIN from Section 4
      │   2. Cross-validate ISIN vs fund name in AUM reference
      │      → Mismatch? Use name lookup instead
      │   3. Look up in reference data:
      │      - AUM (from AUM_31Jan.csv)
      │      - Power Rank (from Powerranking.csv)
      │      - Upside | Downside (from upside_downside_mar.xlsx)
      │      - 5Y Return + category average → alpha = scheme 5Y - category avg
      │        (from Rolling_Returns_Mar.csv, ROLLING_PERIOD=60)
      │   4. Write into 6 table columns:
      │      [Fund Name] [Alloc%] [Up|Down] [Rank] [5Y(alpha)] [AUM]
      └─ Clone slide if more than 4 schemes in one RG
```

---

## 8. Sell / Retain / Buy Action Slides

All three slides are driven by the `Action` column:

```
SELL SLIDE  ← data from Section 2
  Action = "sell completely" → "Complete Sell" table → value = Redemption Value
  Action = "partial sell"    → "Partial Sell"  table → value = Redemption Value

RETAIN SLIDE ← data from Section 2
  Action = "retain completely" → "Complete Retain" table → value = Retained Value
  Action = "partial retain"    → "Partial Retain"  table → value = Retained Value

BUY SLIDE ← data from Section 4
  Action = "fresh buy"  → "Fresh Buy"  table → value = Buy Value Amount
  Action = "buy more"   → "Buy More"   table → value = Buy Value Amount
```

All tables auto-expand (clone slides) if there are more than 15 schemes per section.

---

## 9. Adding a New Client

1. Get the client's Excel (must have a `PF_Curation*` sheet with all 4 sections)
2. Ensure the reference files are up to date in the project folder
3. Run:
   ```bash
   python app.py --client "Client Name" --excel path/to/file.xlsx
   ```
4. Output saved as `ClientName_deck.pptx` in the same folder

---

## 10. Common Issues & What They Mean

| Error / Symptom | Cause | Fix |
|-----------------|-------|-----|
| `ValueError: Could not detect sections: ['s3']` | s3 header not found (new Excel uses different column B header) | Already handled: detector accepts both `Sum of TOTAL_VALUE` and `Sum of Current Value Amount` |
| `ZeroDivisionError: division by zero` | New-format s3 has current values = 0 (buy-only client) | Fallback to section1 total, then section2 sum |
| Partial sell not in transition table | s3 pivot missing partial sell entries | Fixed: now reads `Retained Value` from section2 directly |
| Scheme shows all dashes (—) | ISIN scrambled or fund name not in reference files | ISIN cross-validation + name fallback handles this; check reference files if still missing |
| Wrong slide found (wrong RG data) | Recap slide accidentally matched (no table shapes) | Fixed: `_find_scheme_slide_for_rg` requires at least one table shape |

---

## 11. File Map

```
C:\PowerUpInfinite\M3\
├── app.py                          ← Entry point (run this)
├── Template.pptx                   ← Master slide template
├── AUM_31Jan.csv                   ← Reference: AUM + fund names
├── Powerranking.csv                ← Reference: Power ranks
├── upside_downside_mar.xlsx        ← Reference: Capture ratios
├── Rolling_Returns_Mar.csv         ← Reference: 5Y rolling returns
└── m3_deck_automation/
    ├── excel_reader.py             ← Reads client Excel → 4 sections
    ├── reference_data.py           ← Loads & merges reference files
    └── deck_writer.py              ← All slide-population logic
```
