# 🏕️ Camp Group Sorter

A Streamlit app that sorts camp attendees into groups from a registration CSV or Excel file.

## Features

- Upload any `.csv` or `.xlsx` registration file
- Configure groups by grade, gender, or both
- Two sorting modes:
  - **Size-targeted** — fill groups to a target size, honoring friend pairs where possible
  - **Friend-first** — build groups from friend connection chains, maximizing friend requests honored
- Optional grade filter to include only specific grades
- Download a formatted `.xlsx` file with a Dashboard, Summary, per-bucket tabs, and a Missing Grade tab

## Setup

### Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Deploy to Streamlit Community Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your repo and set `app.py` as the entry point
4. Click **Deploy**

## Expected CSV columns

The app works with any CSV but produces the richest output when these columns are present:

| Column | Used for |
|---|---|
| `First Name` / `Last Name` | Name matching for friend requests |
| `Grade` | Grade-based separation and filtering |
| `Gender` | Gender-based separation |
| `Friend Request` | Resolving friend pairs (comma/semicolon separated names) |
| `T-Shirt Size`, `Mobile Phone Number`, etc. | Passed through to the Excel output |

## Output Excel tabs

| Tab | Contents |
|---|---|
| **Dashboard** | KPI cards, attendee breakdown by bucket, full group roster |
| **Summary** | Every attendee with their group, bucket, and friend satisfaction stats |
| **Per-bucket tabs** | One tab per grade/gender combination |
| **Missing Grade** | Attendees excluded by the grade filter or with unrecognized grades |
