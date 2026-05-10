# Euro 2026 Trip — Project Conventions

This repo holds travel documents and a generated itinerary for a May 2026 Europe trip.

## Source of truth

`travel_info_data/info_notes.txt` is the source of truth for the itinerary. Any
attached files referenced from it (e.g. `Norwegian Air Manolet.pdf`) live in the
same `travel_info_data/` directory and are also authoritative.

When the user changes `info_notes.txt`, treat the new content as the spec and
regenerate the outputs.

### `<...>` markers

Inline placeholders wrapped in angle brackets (e.g.
`<Place here flight to Stockholm. refer to Norwegian Air Manolet.pdf>`) are
**instructions to follow**, not data. Read the referenced file/URL and inline the
resulting facts at that location in the itinerary. Do not put the literal `<...>`
text into the spreadsheet.

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

Writes (overwrites each run):

- `output/euro_2026_itinerary.xlsx`
- `output/euro_2026_itinerary.pdf`

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
| 6 | Location / Details  | Address for lodging; carrier + train/flight number for transit |
| 7 | Departure           | `Station/Airport — HH:MM` for transit; `Check-in: <date>` for lodging |
| 8 | Arrival             | `Station/Airport — HH:MM` for transit; `Check-out: <date>` for lodging |
| 9 | Duration            | `Xh Ym` for transit; `N nights` for lodging                  |
| 10| Notes               | Booking refs, ticket numbers, source URLs                    |

## Type → row color (hex)

| Type prefix       | Fill      |
|-------------------|-----------|
| `Flight`          | `#DDEBF7` |
| `Train`           | `#FFF2CC` |
| `Lodging`         | `#E2EFDA` |
| `Activity`        | `#FCE4D6` |
| `Luggage storage` | `#EDEDED` |

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
