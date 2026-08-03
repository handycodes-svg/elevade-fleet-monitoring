# elevade-fleet-monitoring
Automated ELEVADE ACARS defect tracking &amp; email summaries
# ELEVADE Fleet Monitoring

Automated system untuk track ELEVADE ACARS defects dan generate daily email summaries ke IAA MOC.

## Features
- CSV import dari ELEVADE ACARS exports
- Warning-Fault correlation (ATA-scoped, ±10 min co-occurrence)
- Per-aircraft defect trend tracking (RISING FAST / WATCH / STABLE / GONE)
- Daily email summaries (7:30 AM, sorted by recency)
- Nuisance message exclusion (BRAKES HOT, dll)

## Tech Stack
- Google Apps Script
- Google Sheets
- ELEVADE API (CSV export)

## Setup
1. Clone repo ini
2. Copy Google Apps Script kode ke Apps Script editor
3. Configure trigger (19:00 WIB daily rebuild)
4. Add email recipients di config sheet
5. Deploy!

## Files
- `elevade-fleet-monitoring.gs` — Main Apps Script
- `config.json` — Settings (recipients, aircraft list, ATA codes)
- `docs/` — Architecture & troubleshooting

## Status
- ✅ ACARS CSV import
- ✅ Defect correlation logic
- ✅ Email automation
- 🔄 MR2/MR3 integration (WIP)
