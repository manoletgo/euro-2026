# Euro 2026 Trip — Project Conventions

This repo holds travel documents and a generated itinerary for a May 2026 Europe trip.

## Source of truth

`travel_info_data/info_notes.txt` is the source of truth for the itinerary. Any
attached files referenced from it (e.g. `Norwegian Air Manolet.pdf`) live in the
same `travel_info_data/` directory and are also authoritative.

`info_notes.txt` is written in **strict chronological order**. All rows in the
output must preserve that order. Never reorder or insert rows out of sequence.
When the user changes `info_notes.txt`, treat the new content as the spec and
regenerate the outputs.

### Retention from the latest xlsx

When regenerating, **carry forward all rows from the most recently timestamped
`output/euro_2026_itinerary_<...>.xlsx`** unless `info_notes.txt` overrides
them. `info_notes.txt` is the source of truth for anything it mentions — any
row it adds, modifies, or marks for removal takes effect on the next run.
Rows that exist only in the latest xlsx — e.g. resolved Activity routing,
Tips rows, or details that were filled in during a prior run but never written
back to `info_notes.txt` — must be preserved unchanged.

Silence in `info_notes.txt` means "keep". To remove a row that lives only in
the xlsx, call it out explicitly in `info_notes.txt` with a `<remove: ...>`
marker (e.g. `<remove: 2026-05-20 Fisherman's Bastion activity>`).

### `<...>` markers

Inline placeholders wrapped in angle brackets (e.g.
`<Place here flight to Stockholm. refer to Norwegian Air Manolet.pdf>`) are
**instructions to follow**, not data. Read the referenced file/URL and inline the
resulting facts at that location in the itinerary. Do not put the literal `<...>`
text into the spreadsheet.

### Positional inference

A placeholder's position in `info_notes.txt` carries implicit meaning:

- **Time**: treat the event as occurring between the timestamps of the rows
  immediately above and below it. If only one bound exists, use it as the
  anchor and leave the other end open.
- **Location**: inherit the city or place established by the most recent
  preceding row, unless the placeholder explicitly names a different one.
  Do not assume a location change unless the content of the placeholder
  or the next concrete row indicates one.

### URL handling

- **Google Maps short links** (`maps.app.goo.gl/...`): the redirect URL contains
  the resolved street address. Pull the address from the redirect target.
- **Trainline deep links** (`app.trainline.com/...`): these redirect to a
  Branch.io page that exposes journey JSON only when fetched with a link‑preview
  user agent. Use `curl -A "facebookexternalhit/1.1" <url>`, then base64‑decode
  the payload after `link-<digits>-` in the `twitter:app:url:iphone` meta tag.
  The decoded JSON has a `trip.legs[]` array with `departureDate`, `arrivalDate`,
  `originStation.name`, `destinationStation.name`, `timeTableId`, and `carrier.name`.

## Output

Run from the repo root:

```
python scripts/build_itinerary.py
```

Each run writes two files:

- `output/euro_2026_itinerary_<YYYY-MM-DD_HHMM>.xlsx` — a **new** timestamped
  file per run (e.g. `euro_2026_itinerary_2026-05-10_1845.xlsx`). Older xlsx
  files are retained as version history. The file with the most recent
  timestamp is the "latest" and is the one that retention (above) reads from.
- `output/euro_2026_itinerary.pdf` — **overwritten** every run; no timestamp
  in the filename. The PDF is always assumed to be the latest output.

Both files share the same data, schema, and color coding — defined once at the
top of `scripts/build_itinerary.py`. Do not hand‑edit the outputs.

First-time setup: `python -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt`.

## Schema (10 columns, fixed order)

| # | Header              | Notes                                                        |
|---|---------------------|--------------------------------------------------------------|
| 1 | Date                | `YYYY-MM-DD`                                                 |
| 2 | Day                 | 3-letter weekday (`Mon`, `Tue`, ...)                         |
| 3 | Time (Depart)       | 24h `HH:MM`, blank for non-transit rows                      |
| 4 | City / Route        | Single city for lodging/activity; `From → To` for transit    |
| 5 | Type                | One of the categories below — drives row color               |
| 6 | Location / Details  | Address for lodging/activity; carrier + train/flight number for transit |
| 7 | Departure           | `Station/Airport — HH:MM` for transit; `Check-in: <date>` for lodging; `From: <address>` for activity |
| 8 | Arrival             | `Station/Airport — HH:MM` for transit; `Check-out: <date>` for lodging; `HH:MM (est.)` for activity |
| 9 | Duration            | `Xh Ym` for transit; `N nights` for lodging; travel time to attraction for activity (e.g. `25 min`) |
| 10| Notes               | Booking refs, ticket numbers; Google Maps directions URL for activity rows (`maps.google.com/dir/...`) + travel mode summary (e.g. `Metro L1 → walk 3 min`) |

## Activity routing

Every `Activity` row represents a single attraction or destination. Claude must
research and include the most efficient route to get there.

### Source inference

- The **source** of each activity is inherited from the preceding row:
  - If the preceding row is a `Lodging` row, the source is that accommodation's
    address.
  - If the preceding row is another `Activity`, the source is that attraction.
  - The first activity of each day always departs from the Airbnb/lodging where
    we are staying that night.

### Routing

- Look up the most efficient route on Google Maps (walking, transit, or a mix —
  whichever is fastest given the time of day and distance).
- Record the route as a Google Maps directions URL in the **Notes** column using
  the format:
  `https://www.google.com/maps/dir/?api=1&origin=<source>&destination=<dest>&travelmode=<mode>`
- Fill in **Location / Details** with the full address of the destination,
  resolved from Google Maps.
- Fill in **Duration** with the estimated travel time from Google Maps
  (e.g. `25 min`, `1h 10m`).
- Fill in **Time (Depart)** with the estimated departure time needed to arrive
  at the attraction at a reasonable hour, based on the preceding row's end time.

### What "most efficient" means

Prefer the option that minimises total travel time. When two options are
within 5 minutes of each other, prefer the one with fewer transfers or
less walking. Note the chosen mode in the **Notes** column alongside the URL
(e.g. `Metro line 1 → walk 3 min | maps.google.com/...`).

## Tips rows

For every destination — whether it's a day trip or a multi-night stay — insert
a single `Tips` row at the **top of the stay**, immediately before the first
Lodging or Activity row for that destination. One row per destination, not per
day. The row gives at-a-glance guidance for making the most of the place.

The tips must cover three areas, in this order:

1. **Transportation access** — what to buy for getting around the city
   (e.g. multi-day metro pass, regional rail card, integrated city card that
   bundles transit).
2. **Attraction maximisation** — what passes, combo tickets, or booking
   windows give the best value or skip-the-line access (e.g. museum day pass,
   official advance-booking site, free-entry days, timed-entry tips).
3. **Exploration** — what's worth seeing, eating, and shopping for beyond the
   headline attractions (signature dishes, neighbourhoods to wander, markets,
   notable shops).

### Column usage for a Tips row

- **Date / Day**: first date (and weekday) of the stay in that destination.
- **Time (Depart) / Departure / Arrival / Duration**: leave blank.
- **City / Route**: destination city.
- **Type**: `Tips`.
- **Location / Details**: **transport tip** only — concise, with price if
  relevant. Example: `GVB 72-hour pass (€24) — trams + metro + buses`.
- **Notes**: **attractions** and **exploration** tips, labeled and
  pipe-separated. Example:
  `Attractions: I amsterdam City Card €60/24h — Rijksmuseum + canal cruise included | Explore: Jordaan district, broodje haring at Stubbe's, vintage shops on Haarlemmerstraat`.

Tips should be specific and actionable (named passes, prices, neighbourhoods,
dishes). Avoid generic advice like "use public transport" or "try local food".

## Type → row color (hex)

| Type prefix       | Fill      |
|-------------------|-----------|
| `Flight`          | `#DDEBF7` |
| `Train`           | `#FFF2CC` |
| `Lodging`         | `#E2EFDA` |
| `Activity`        | `#FCE4D6` |
| `Luggage storage` | `#EDEDED` |
| `Tips`            | `#E4D7F2` |

Header row: white text on `#305496`. Borders: `#BFBFBF`. Body font: Arial 10
(xlsx) / Helvetica 8 (pdf). Header row frozen in xlsx; `repeatRows=1` in pdf.
PDF page size is landscape A3 so the table fits on a single page.

## Editing workflow

1. Update `travel_info_data/info_notes.txt` (and drop any new attachments
   alongside it).
2. Update the `ROWS` list in `scripts/build_itinerary.py` to match. Keep rows
   in chronological order.
3. Run the script. Commit `info_notes.txt`, the script change, and the
   regenerated `output/` files together so they stay in sync.
