# ⛪ Rock Point Camp Group Sorter

A Streamlit app for Rock Point Church that sorts campers into groups from a registration CSV or Excel file — with full branding for both KIDS and YTH camps.

## Features

- **Two camp modes** — KIDS Camp and YTH Camp, each with its own brand colors, fonts, and workflow
- **KIDS Camp** — sort by grade, gender, and/or friend requests
- **YTH Camp** — sort by grade, gender, community group, previous camp groups, and leaders
- **Two sorting algorithms**
  - *Size-targeted* — fill groups to a target size, honoring friend pairs within the cap
  - *Friend-first* — groups form from friend connection chains, maximizing requests honored
- **Community group support** — upload a community group Excel, leaders CSV, and previous camp groups CSV for YTH
- **Leader assignment** — automatically assigns leaders to groups based on their community group
- **Smart friend matching** — resolves first name, last name, and full name variants; handles all column name variants for friend request fields
- **Formatted Excel output** — Dashboard, Summary, per-bucket tabs, Leaders tab (YTH), and Outside Filter tab

## Setup

### Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Deploy to Streamlit Community Cloud (free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set **Main file path** to `app.py`
5. Click **Deploy** — Streamlit installs dependencies automatically

## Expected CSV columns

### Registration file (required)
| Column | Notes |
|---|---|
| `First Name` / `Last Name` | Required for name matching |
| `Grade` | Enables grade-based sorting and filtering |
| `Gender` | Enables gender-based sorting |
| `Friend Request` | Also accepts: `Friend Requests`, `Youth's Friend Request`, `Youth's Friend Requests` |
| `Community Group` | YTH only — can also be supplied via the Community Group Excel |
| `T-Shirt Size`, `Mobile Phone Number`, etc. | Passed through to Excel output |

### YTH additional files (optional)
| File | Required columns |
|---|---|
| Community Group Excel | `First Name`, `Last Name`, `Community Group` |
| Leaders CSV | `Community Group`, `Leader Name` |
| Previous Camp Groups CSV | `First Name`, `Last Name`, `Group` (or `Camp Group`) |

## Output Excel tabs

| Tab | Contents |
|---|---|
| **Dashboard** | KPI cards, attendee breakdown, full group roster with leader assignments |
| **Summary** | Every attendee with group, bucket, and friend satisfaction stats |
| **Per-bucket tabs** | One tab per grade/gender/community combination |
| **Leaders** | YTH only — group-to-leader assignment table |
| **Outside Filter** | Attendees excluded by grade filter or with unrecognized grades |

## Brand

Built for **Rock Point Church**, Gilbert AZ — using the official KIDS and YTH brand guidelines.
