# Copilot Prompt: Automate Midweek Data Excel Sheet from GraphQL API

## Context & Goal

I have a weekly church midweek data tracking spreadsheet (`Midweek_Data.xlsx`) that I currently populate manually. I want a Python script (hostable and runnable for free, e.g. on GitHub Actions or a free-tier cloud function) that:

1. Calls a GraphQL API **three times** — once for each of three stream IDs
2. Parses the response to extract governorship-level service data
3. Writes the data into the correct sheet and rows of the Excel file, matching the existing format exactly

---

## The GraphQL Query

```graphql
query getStreamGovernorships($id: ID!) {
  streams(where: { id: $id }) {
    id
    name
    leader {
      id
      firstName
      lastName
      fullName
      __typename
    }
    governorships {
      council {
        name
        leader {
          fullName
        }
      }
      name
      id
      stream_name
      bacentaCount
      aggregateServiceRecords(limit: 1, skip: 0) {
        id
        attendance
        income
        numberOfServices
        week
        __typename
      }
      services(limit: 1, skip: 0) {
        id
        createdAt
        attendance
        income
        week
        serviceDate {
          date
          __typename
        }
      }
      leader {
        id
        fullName
        __typename
      }
      __typename
    }
    __typename
  }
}
```

**GraphQL API Endpoint:** `https://api-synago.firstlovecenter.com/graphql`

**Stream IDs to query (run the query once per ID):**

```
STREAM_ID_1 = "2dd77486-5d8d-4231-96e9-6d042500198a"   # Colossians (Council 1 - Isaac Agyeman)
STREAM_ID_2 = "2d0f5804-0462-442f-93cc-25db95912589"   # Galatians
STREAM_ID_3 = "804e3aaf-e868-4772-a9f6-f0de76941d01"   # Jesus Night
```

---

## Sample API Response Structure

Here is an example of a real response for STREAM_ID_1 (Colossians). Use this to understand the data shape:

```json
{
  "data": {
    "streams": [
      {
        "id": "2dd77486-5d8d-4231-96e9-6d042500198a",
        "name": "Colossians",
        "leader": { "fullName": "Isaac Agyeman" },
        "governorships": [
          {
            "council": {
              "name": "Colossians 1",
              "leader": { "fullName": "Isaac Agyeman" }
            },
            "name": "THS Ecomog",
            "bacentaCount": 5,
            "aggregateServiceRecords": [
              {
                "attendance": 6,
                "income": 122,
                "numberOfServices": 2,
                "week": 16
              }
            ],
            "services": [
              {
                "attendance": 25,
                "income": 517,
                "week": 46,
                "serviceDate": { "date": "2025-11-13" }
              }
            ],
            "leader": { "fullName": "Alex Baah-Gyambrah" }
          }
          // ... more governorships
        ]
      }
    ]
  }
}
```

**Key fields to extract per governorship:**

- `name` → Governorship name
- `leader.fullName` → Team Leader
- `bacentaCount` → No. Of Bacentas
- `council.name` → Used to determine which sheet to write to (e.g. "Colossians 1" → sheet "Colossians 1")
- `Income (USD)` → Leave as 0 (not in API)
- `Bacentas on Vacation` → Leave blank (not in API)

### ⚠️ Critical: Data Source Priority Logic (per governorship)

The script must determine the current ISO week number at runtime (e.g. week 16) and apply the following priority logic to resolve `attendance`, `income`, `numberOfServices`, and `comment` for each governorship:

**Step 1 — Check `aggregateServiceRecords[0]`:**

- If it exists AND its `week` field == current week number → use it as the primary source:
  - `attendance` = `aggregateServiceRecords[0].attendance`
  - `income` = `aggregateServiceRecords[0].income`
  - `numberOfServices` = `aggregateServiceRecords[0].numberOfServices`
  - `comment` = _(blank)_

**Step 2 — Fallback: Check `services[0]`:**

- If Step 1 did not match (aggregate is empty, or its `week` != current week) → check `services[0]`
- If `services[0]` exists AND its `week` == current week → this is a joint service held at governorship level:
  - `attendance` = `services[0].attendance`
  - `income` = `services[0].income`
  - `numberOfServices` = `1` _(always 1 for a joint service)_
  - `comment` = `"Joint Service"`

**Step 3 — Final fallback:**

- If neither matched → write all zeros and blank comment:
  - `attendance` = 0, `income` = 0, `numberOfServices` = 0, `comment` = _(blank)_

**Derived fields (apply after resolving the above):**

- `Services/Bacentas` → format as `"{numberOfServices}/{bacentaCount}"`
- `Services Not Held` → `bacentaCount - numberOfServices`

---

## Target Excel File Structure

The file `Midweek_Data.xlsx` has the following sheets relevant to this task:

- **`Colossians 1`** — Governorships under council "Colossians 1" (leader: Isaac Agyeman)
- **`Colossians 2`** — Governorships under council "Colossians 2" (leader: Edwin Ogoe)
- **`Colossians 3`** — Governorships under council "Colossians 3" (leader: Nathan Kudowor)

### Sheet Column Layout (same pattern for all Colossians sheets):

| Col | Header               | Source                                           |
| --- | -------------------- | ------------------------------------------------ |
| A   | Governorship         | `governorship.name`                              |
| B   | Team Leader          | `governorship.leader.fullName`                   |
| C   | No. Of Bacentas      | `bacentaCount`                                   |
| D   | Bacentas on Vacation | leave blank                                      |
| E   | Att                  | resolved `attendance` (see priority logic)       |
| F   | Income(GHS)          | resolved `income` (see priority logic)           |
| G   | Income(USD)          | 0                                                |
| H   | No. Of Services      | resolved `numberOfServices` (see priority logic) |
| I   | Services/Bacentas    | `"{numberOfServices}/{bacentaCount}"`            |
| J   | Services Not Held    | `bacentaCount - numberOfServices`                |
| K   | Comment              | `"Joint Service"` if fallback used, else blank   |

- Row 1: Sheet title (e.g. "GALATIANS" — leave as-is, do not overwrite)
- Row 2: Column headers — do not overwrite
- Row 3 onward: Data rows, one per governorship
- Last row: TOTAL row — use Excel SUM formulas (e.g. `=SUM(C3:C{last_data_row})`)

### Summary Sheet

After updating the detail sheets, also update the **`Summary`** sheet. The relevant rows for Colossians are:

| Overseer       | Oversight Area | Columns to update                           |
| -------------- | -------------- | ------------------------------------------- |
| Isaac Agyeman  | Colossians 1   | Bacentas, Att, Income(GHS), No. of Services |
| Edwin Ogoe     | Colossians 2   | Same                                        |
| Nathan Kudowor | Colossians 3   | Same                                        |

The Summary sheet pulls totals — these should be Excel formulas referencing the detail sheets (e.g. `='Colossians 1'!C{total_row}`), not hardcoded values.

---

## Script Requirements

### Language & Libraries

- **Python 3.10+**
- `requests` — for GraphQL API calls
- `openpyxl` — for Excel read/write
- No other dependencies beyond stdlib + these two

### Authentication — Login Flow

The API uses a JWT Bearer token obtained by logging in first. The script must:

**Step 1 — Login to get the access token:**

```
POST https://ndx3y4sa3znyoxzin6bzmoy6fi0jvruc.lambda-url.eu-west-2.on.aws/auth/login
Content-Type: application/json

{
    "email": "dabick14@gmail.com",
    "password": "<password>"
}
```

Response shape:

```json
{
    "message": "Login successful",
    "tokens": {
        "accessToken": "<jwt_token>",
        "refreshToken": "<refresh_token>"
    },
    "user": { ... }
}
```

Extract `tokens.accessToken` from the response.

**Step 2 — Use the access token for all GraphQL requests:**

```
Authorization: <accessToken>
```

_(No "Bearer" prefix — pass the token value directly as the Authorization header value, as shown in the curl example)_

**Credentials as environment variables:**

- `FLC_EMAIL` — login email
- `FLC_PASSWORD` — login password

Do **not** hardcode credentials in the script. Store them as GitHub Actions secrets and load via `os.environ`.

**Token expiry:** The access token expires after ~30 minutes. Since the script runs three queries in quick succession this is fine — fetch the token once at startup and reuse it for all three queries.

### Script Behaviour

1. Load the existing `Midweek_Data.xlsx` (do not recreate from scratch — preserve all other sheets and formatting)
2. For each of the 3 stream IDs, call the GraphQL API
3. Group governorships by `council.name` → determines target sheet
4. Clear existing data rows (rows 3 to last data row, not headers or totals) in the relevant sheet
5. Write fresh data rows starting at row 3
6. Rewrite the TOTAL row at the end using `SUM` formulas
7. Update the Summary sheet totals (via formula references or direct SUM)
8. Save output as `Midweek_Data_updated.xlsx` (preserve original)

### Data Resolution

- Apply the **priority logic** described above for every governorship — do not assume `aggregateServiceRecords` is always present or current
- Use Python's `datetime.date.today().isocalendar().week` to get the current ISO week number at runtime

### Error Handling

- If the API returns an error or empty data for a stream, log a warning and skip that stream (don't crash)
- Print a summary of how many governorships were written per sheet

### Free Hosting Options

Structure the script so it can be run as:

- A **standalone Python script** (run locally or via cron)
- A **GitHub Actions workflow** (triggered weekly via `schedule: cron`)

Include a sample `github-actions.yml` that:

- Runs every Wednesday at 8am UTC (`cron: '0 8 * * 3'`)
- Checks out the repo
- Installs dependencies (`pip install requests openpyxl`)
- Runs the script with `FLC_EMAIL` and `FLC_PASSWORD` from GitHub Secrets
- Commits and pushes the updated Excel file back to the repo

---

## Example Expected Output (Colossians 1 sheet, first few rows)

```
Row 1: [GALATIANS header]
Row 2: [Governorship | Team Leader | No. Of Bacentas | Bacentas on Vacation | Att | Income(GHS) | Income(USD) | No. Of Services | Services/Bacentas | Services Not Held | Comment]
Row 3: THS Ecomog        | Alex Baah-Gyambrah       | 5 |  | 6  | 122  | 0 | 2 | 2/5 | 3 |               ← aggregate matched week 16
Row 4: THS Agbogba       | Collins Quarcoo          | 7 |  | 17 | 429  | 0 | 4 | 4/7 | 3 |               ← aggregate matched week 16
Row 5: Haatso Mabey      | Malcolm Otchere - Forbih | 3 |  | 38 | 4903 | 0 | 1 | 1/3 | 2 |               ← aggregate matched week 16
Row 6: Pantang Campus    | Princess Wilson-Anderson | 4 |  | 2  | 50   | 0 | 1 | 1/4 | 3 | Joint Service ← fell back to services[0]
Row 7: Wisconsin Fruitful| Thomas Dodoo             | 3 |  | 0  | 0    | 0 | 0 | 0/3 | 3 |               ← no data for current week
...
TOTAL row: [=SUM formulas across all data rows, Comment column left blank]
```

---

## Deliverables Requested

1. **`update_midweek_data.py`** — the main Python script
2. **`.github/workflows/update_midweek.yml`** — GitHub Actions workflow
3. **`README.md`** — brief setup instructions (how to set the API token secret, how to run locally)
