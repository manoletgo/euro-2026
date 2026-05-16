---
description: Sync info_notes.txt → ROWS, then regenerate the Euro 2026 itinerary outputs
argument-hint: (no args)
allowed-tools: Read, Edit, Write, Bash, Glob, Grep
---

Regenerate the Euro 2026 itinerary outputs. Follow the editing workflow in `CLAUDE.md` exactly:

## Step 1 — Sync `info_notes.txt` → `ROWS`

1. Read `travel_info_data/info_notes.txt` (the source of truth).
2. Identify the most recently timestamped `output/euro_2026_itinerary_<YYYY-MM-DD_HHMM>.xlsx` and treat its rows as the retention baseline.
3. Reconcile the `ROWS` list in `scripts/build_itinerary.py` so that:
   - All rows from `info_notes.txt` are reflected, in **strict chronological order**.
   - Rows that exist only in the latest xlsx (resolved Activity routing, Tips rows, filled-in details) are **preserved unchanged** unless `info_notes.txt` overrides them or contains a `<remove: ...>` marker.
   - Inline `<...>` placeholders in `info_notes.txt` are resolved into concrete data per `CLAUDE.md` (read referenced files/URLs, apply positional inference for time and location, follow the URL-handling rules for Google Maps short links and Trainline deep links).
   - Every `Activity` row has full routing per the "Activity routing" section: source inferred from preceding row, Google Maps directions URL in Notes, full destination address in Location/Details, travel duration, and a derived departure time.
   - Each destination has exactly one `Tips` row at the top of the stay covering transport / attractions / exploration as specified.
4. Edit `scripts/build_itinerary.py` to update the `ROWS` list. Do not touch the schema, color map, or rendering code unless explicitly required.

## Step 2 — Run the generator

From the repo root, run:

```
python scripts/build_itinerary.py
```

If the venv exists, prefer `.venv/bin/python scripts/build_itinerary.py`.

## Step 3 — Verify

Confirm both files were written:

- `output/euro_2026_itinerary_<new-timestamp>.xlsx` (new file)
- `output/euro_2026_itinerary.pdf` (overwritten)

Report a short summary of what changed in `ROWS` (rows added, modified, removed) and any `<...>` placeholders that were resolved during the sync.
