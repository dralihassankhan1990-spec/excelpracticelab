# Excel Practice Lab — Excel for IT Project Management

A professional practice website for the "Excel for IT Project Management — 42 Formula Demos"
course by Muhammad Bilal (Soch & Systems). All frontend, no backend, no database, no sign-up.
Runs entirely in the visitor's browser.

## What visitors can do

- **Pick any of the 42 formulas** — search or filter by category (Lookup, Aggregation,
  Logic, Text, Date, Dynamic array, Statistics, References, Tool).
- **Watch the demo first** — selecting a formula puts its demo video at the top of the
  practice area (played from the YouTube playlist, in order 01–42).
- **Read the formula** — syntax, how to read it, the course example, the common error
  and its fix, and why it matters.
- **Practise by typing** — a built-in Excel-compatible checker evaluates the typed
  formula against the live dataset and says correct / wrong, with the expected answer
  and a one-click model solution. It understands cell references (A1, $A$1, A:A),
  ranges, sheet names, and 60+ functions (XLOOKUP, VLOOKUP, INDEX-MATCH, SUMIFS,
  COUNTIFS, FILTER, UNIQUE, SORT, TEXT, dates, and more).
- **Switch datasets** — the course dataset (120 IT work items) plus three more fictional
  IT project datasets: Bug Tracker (40), Sprint Tasks (36), Project Budget & Costs (30).
  Most exercises work on every dataset; where a dataset lacks a column, the page says
  so and points back to the course dataset.
- **Upload their own data** — CSV file or pasted CSV text. First row = headers. The file
  is read in the browser only; nothing is uploaded anywhere.
- **See the slides** — the 48-slide course deck embedded on the page, with links to the
  GitHub repo (workbook, slides, guide, all 42 demo videos) and the YouTube channel.
- No Drive link anywhere on the page, by design.

## Files

```
index.html            the page
styles.css            glassmorphism styling (dark teal + gold)
app.js                UI + the Excel formula evaluator + exercises
data/formulas.json    the 42 formulas (from units.py)
data/datasets.json    the 4 datasets (course data from the workbook + 3 fictional)
build_data.py         regenerates data/*.json from the workbook + units.py (optional)
```

## Deploy (GitHub → Vercel, free)

1. On GitHub, create a new public repository, e.g. **excel-practice-lab**.
2. Upload everything in this folder (drag-and-drop on the GitHub website, or git push).
3. On Vercel, **Add New → Project → Import** the repository.
4. Framework Preset: **Other**. Build Command: leave empty. Output Directory: leave
   empty (or `.`). Deploy.
5. Share the `https://<name>.vercel.app` link. That is the public link for everyone —
   no account needed to open it.

## Preview locally

```
python3 -m http.server 8000
```
then open http://localhost:8000 (opening index.html directly as a file will not load
the data files — a local server is required).

## Updating the data

Edit `data/formulas.json` / `data/datasets.json` directly, or re-run
`build_data.py` from the Excel training folder (needs the workbook and `units.py`).
Then commit and push — Vercel redeploys automatically.

## Verified

- All 42 exercises check correctly on the course dataset, the Bug Tracker and the
  Sprint Tasks; 30 of 42 on the Budget dataset (it has no hours/story-point columns —
  the page explains this where it applies). Zero evaluator errors.
- Known answers spot-checked: XLOOKUP IT-005 → Fatima Noor, COUNTIF Delayed = 18,
  COUNTBLANK FinishDate = 72, SUM ActualHrs = 1217.8.
- Playlist order verified live: 42 videos, 01 XLOOKUP … 42 SUBSTITUTE/CLEAN, so each
  formula opens its own demo.
