# Organogram Builder

A Streamlit app to enter an institution's positions, draw the organogram, and print it as **PDF** and **Word**.

## Files

| File | Purpose |
|---|---|
| `app.py` | The Streamlit screens (entry forms, table editor, chart, export) |
| `data_utils.py` | Load/save, validation, level calculation, add/edit/delete |
| `chart_utils.py` | Builds the Graphviz chart, including multi-page splitting |
| `export_utils.py` | PDF and Word creation |
| `sample_data.xlsx` | Excel template with sample data and dropdowns |
| `make_template.py` | Script that regenerates the template (optional) |
| `requirements.txt` | Python packages |
| `packages.txt` | **Required on Streamlit Cloud** - installs the Graphviz program |

## Deploy (GitHub + Streamlit Community Cloud)

1. Create a GitHub repository and upload every file in this folder to the top level (not inside a sub-folder).
2. Go to https://share.streamlit.io, sign in with GitHub, choose **Create app**, pick your repository, branch `main`, main file `app.py`, and deploy.
3. If you change `packages.txt` later, reboot the app from the Streamlit Cloud menu.

## Run on your own computer

```
pip install -r requirements.txt
# also install the Graphviz program: https://graphviz.org/download/  (Windows installer, or: sudo apt install graphviz)
streamlit run app.py
```

## Data format (Excel)

**Positions sheet** - one row per position

| Column | Meaning |
|---|---|
| `position_id` | Unique code (P001, P002...). Never reuse or change. |
| `position_title` | Name of the post (required) |
| `department` | Used for box colours (no colour key is printed) |
| `reports_to` | `position_id` of the main manager; **blank for top-level posts** |
| `status` | Filled / Vacant (Vacant posts get a red dashed border) |
| `email`, `phone` | Optional |
| `sort_order` | Optional; left-to-right order under the same manager |

**Extra_Reporting sheet** - only for a second (dotted-line) manager

| Column | Meaning |
|---|---|
| `position_id` | The post with an extra manager |
| `also_reports_to` | `position_id` of the extra manager |
| `relationship` | Optional: Dotted line, Functional, Acting, Project |

Person names are not entered; the chart shows position titles only. (Old files that still have a `name` column will load, and names will then appear.)

The level is never typed; it is calculated from `reports_to`.

## Important notes

* **Streamlit Cloud does not keep your data.** Use *Save as Excel / JSON* in the sidebar before closing and load the file next time.
* **Large charts (50+ posts):** choose *Overview + one page per branch*, or *Left to right*, or A2 paper. One page with 50+ boxes prints with very small text.
* **Malayalam or other non-English position titles:** the on-screen chart uses your browser's fonts. The PDF/Word pictures are drawn on the server, which may not have a suitable font, so such titles could appear as empty boxes. If so, add a suitable font package (for example `fonts-smc` for Malayalam, or `fonts-noto-core`) as a new line in `packages.txt`, and reboot the app. This has not been tested.
* The Word file contains the chart as a picture (not editable) plus an editable table of all positions.
