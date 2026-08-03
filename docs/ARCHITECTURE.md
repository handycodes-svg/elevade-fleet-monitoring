# ELEVADE Fleet Monitoring — MR2 Screening Architecture

**System Version:** 1.0  
**Last Updated:** 31 July 2026  
**Scope:** AirAsia Indonesia A320 fleet (PK-AXD → PK-AZY, ~25–27 aircraft)

---

## Overview

Automated pipeline untuk **screen ACARS defects terhadap OPEN MR2 daftar** pesawat. Sistem ini:

1. **Imports** CSV exports dari ELEVADE (Warning & Fault Messages)
2. **Correlates** Warning-Fault pairs by aircraft registration, time window, ATA Chapter, keyword overlap
3. **Tracks** defect trends per aircraft (NEW / WATCH / STABLE / no recent)
4. **Outputs** "Screening MR2" sheet dengan status MR2 lookup (MR2 OPEN ↔ ACARS message match)
5. **Emails** daily summary @ 17:00 WIB → MOC team (repetitive + current items only)

---

## Data Flow
---

## Key Components

### 1. **CSV Import & Parsing** (`parseElevadeCsv`)
- **Input:** ELEVADE CSV export (13-column format)
- **Header Detection:** Scan first 30 rows; stop when "Message", "Occurence", "A/C Reg" found
- **Date Parsing:** Tolerant format detection
- **Output:** Normalized array of message objects

### 2. **Data Store** (`_DATA_STORE` sheet)
- **Purpose:** Accumulate 30-day rolling history of all ACARS messages
- **Structure:** 13 columns matching config.COLUMNS
- **Hidden:** Yes (system use only)
- **Pruning:** Automatic on each import (keeps data from cutoff date onward)

### 3. **Deduplication** (`dedupRows`)
- **Key:** `Reg||Message||ATA||Occurence` (timestamp-aware)
- **Goal:** Prevent double-counting if same CSV imported twice
- **Retention:** 30 days from maximum data date in set

### 4. **MR2 Index** (`loadMr2List`, `buildMr2Index`)
- **Source:** External spreadsheet (spreadsheet ID in config)
- **Scope:** OPEN MR2 ONLY (stops reading at "OPEN MR3" marker)
- **Extraction:**
  - A/C Reg: "PK-AXD" or "AXD" from column A
  - ATA Chapter: From reference column or description
  - Description tokens: Filtered (≥3 char, stop words removed)
  - Unit extraction: APU, ENG1/2/3, L/R WING, L/R TK, CTR TANK
- **Index Structure:** `{ Reg → [ {desc, ata, tokens, units}, ... ] }`

### 5. **Correlation Logic** (`correlationOk`, `descKeywordMatch`, `ataRelated`)
- **Warning-Fault Match Criteria:**
  1. Same A/C Reg (exact)
  2. Same flight number
  3. Same date
  4. Time within ±10 minutes
  5. ATA Chapter compatible (77↔71/75/73/70 allowed)
  6. Description keyword overlap (at least 1 common ≥2-char word beyond stop list)

### 6. **MR2 Status Lookup** (`lookupMr2Status`)
- **Strict Registration Isolation:** Only check MR2 items for same A/C Reg
- **Token Matching:** Extract ≥3-char words (stop-filtered) from ACARS message
- **Minimum Threshold:** ≥ 2 keywords must overlap with any MR2 description token set
- **ATA Compatibility:** If MR2 has ATA Chapter, message ATA must be compatible
- **Unit Conflict Check:** If both message and MR2 have unit specifications, they must overlap
- **Result:** "MR2" or "" (empty)

### 7. **Trend Calculation** (`trendOf`)
- **NEW:** First occurrence ≤ 3 days ago → "NEW Nx"
- **NO RECENT:** Last occurrence > 21 days ago → "no recent"
- **STABLE:** Otherwise → "stable Nx"

### 8. **Recommended Action** (`recommendedAction`)
- **Input:** count, trend, priority, hasMr2 flag
- **Decision Tree:**
  - If MR2 OPEN → "MONITOR (MR2 OPEN)" [watch]
  - Else if High + NEW → "INVESTIGATE NOW" [urgent]
  - Else if count ≥ 10 + stable → "SCHEDULE INSPECTION (chronic)" [high]
  - Else if NEW → "MONITOR next flights" [watch]
  - Else if no recent → "MONITOR only (may be cleared)" [low]
  - Default → "ROUTINE monitoring" [low]

### 9. **Output Sheet** (`buildMr2ScreeningSheet`)
- **Layout per Aircraft:**
- **Columns:** Type/Item | ATA | Source | Count | Priority | Trend | Last Occ. | Action | Status MR2
- **Styling:**
  - Header: Dark blue bg, white text
  - A/C band: Dark blue, bold, sortable divider
  - WARNING rows: Peach bg (#FCE4D6), priority color
  - FAULT rows: Light blue bg (#DDEBF7), bold ATA
  - MR2 cells: Orange highlight if "MR2"
  - Action colors: Red (urgent) / Amber (high) / Yellow (watch) / Green (low)
- **Frozen rows:** Headers stay visible when scrolling

### 10. **Email Summary** (`sendScreeningSummaryEmail`)
- **Trigger:** Daily @ 17:00 WIB (TZ: Asia/Jakarta)
- **Filter:**
  - REPETITIVE: count ≥ 2
  - CURRENT: last occurred within last 3 days
  - Exclude: messages in `EMAIL_EXCLUDE_MESSAGES` (e.g., "BRAKES HOT")
- **Sort:** WARNING first → FAULT; newest first within each
- **Badges:**
  - Red "TODAY" → occurred same day + repetitive (needs immediate attention)
  - Orange "MR2 OPEN" → matches OPEN MR2 item
- **Recipients:** 12 MOC team members (configurable)
- **Subject:** "ELEVADE Monitoring per [date]"

### 11. **File Management** (`moveToProcessed`)
- **After import:** Processed CSV files moved to `processed/` subfolder
- **Prevents:** Re-import of same file

---

## Configuration

**See `config.json` for:**
- Google Sheets IDs & sheet names
- Drive folder ID (ELEVADE_CSV_INBOX)
- Email recipients & trigger hour
- Correlation thresholds (time window, keyword minimum)
- ATA equivalence map
- Spurious message exclusion list
- Stop words & subsystem keywords

---

## Known Limitations & Design Decisions

1. **MR2 Only:** System screens OPEN MR2 items exclusively; MR3 and non-MEL categories are ignored.
   - **Why:** MOC workflow prioritizes deferrable vs. must-fix items separately.

2. **Strict Registration Isolation:** MR2 lookup ONLY checks items for the exact same A/C Reg.
   - **Why:** Prevents data bleed between aircraft (critical for maintenance scheduling).

3. **ATA Chapter Flexibility (77 ↔ 71/75/73/70):**
   - **Why:** Common in flight control defects; same subsystems often span multiple chapters.

4. **Description Keyword Filtering:** Stop words ignored; minimum 3 characters.
   - **Why:** Avoid false positives from common aviation terms.

5. **Email Repetitive Filter:** Only count ≥ 2 occurrences + last ≤ 3 days.
   - **Why:** Single-flight glitches are noise; MOC needs fleet-wide pattern alerts.

---

## Maintenance & Troubleshooting

### Common Issues

**"Tidak ada CSV" error:**
- Check folder ID in config (should be ELEVADE_CSV_INBOX, not parent or sibling folder).
- Verify CSV file extensions are `.csv` or MIME type is `text/csv`.

**No data appearing in sheet:**
- Confirm data store isn't empty: click menu → "Show data store info".
- Check date range: data older than 30 days from max data date is pruned.

**MR2 lookup failing (all empty Status MR2 cells):**
- Run menu → "Diagnose MR2 source": see if external spreadsheet opens & tab exists.
- Check A/C Reg extraction: column A format must contain recognizable code (PK-AXD or AXD).

**Email not sending:**
- Confirm `CONFIG.EMAIL_RECIPIENTS` has valid emails.
- Check trigger is active: menu → "Diagnose" → see if trigger registered.
- Email only sends if ≥1 item passes filters (no repetitive + current = silent).

---

## Future Enhancements

- Dashboard (Looker Studio) with MR2 coverage metrics
- Slack integration for real-time alerts
- Defect-delay correlation analysis (parts replacement patterns)
- Airbus FSR dossier auto-linking (by ATA + symptom)
- Multi-fleet support (extend to A330, Boeing 737, etc.)

---

## File Structure (GitHub Repo)
---

**For questions or updates:** Contact MOC / Tech Services  
**Last sync with production:** 31 Jul 2026, 19:00 WIB
