# Drought Relief Data: checker page + Cornell Box upload

Partners check their file on a web page, download the checked copy, and upload it to a
**Cornell Box File Request** link. You then combine everything in the Box folder with
one Python script.

```
Partner's browser                          Cornell Box (Cornell-managed)        Your computer
───────────────────                        ─────────────────────────────        ─────────────
checker page (GitHub Pages)                File Request link                    combine.py
 1 add rows, fix errors      ──────────►    → "Drought relief submissions"  ──►  rows_clean.csv
 2 organization details        upload        folder (only you see it)           files_summary.csv
 3 download checked .xlsx                                                       coverage_ward_month.csv
 4 open Box link, upload
```

**Why this setup is the safest:**
- **Data storage:** data goes straight to Cornell-managed storage, under Cornell's
  agreements.
- **Checker page:** it's static, with no server, database, passwords or access codes.
  Files are read in the partner's browser and never sent anywhere by the page.
- **Box File Request:** partners can upload but can't see the folder or each other's
  files, and they don't need a Box account.
- **Versions:** Box keeps every file, including earlier versions.

## Contents

| Path | Goes to | Purpose |
|---|---|---|
| `docs/index.html` | GitHub Pages | The checker page |
| `docs/config.js` | GitHub Pages | **The one line you edit:** your Box File Request link |
| `docs/xlsx.full.min.js` | GitHub Pages | Spreadsheet reader (SheetJS 0.18.5) used by the page |
| `docs/drought_relief_disbursement_template.xlsx` | GitHub Pages | Blank template partners can download |
| `docs/.nojekyll` | GitHub Pages | Tells GitHub to serve the files as they are |
| `combine/combine.py` | Your computer | Reads the Box folder, re-checks every row, writes the combined dataset |
| `combine/requirements.txt` | Your computer | `openpyxl`, the only package the script needs |
| `ward_county.csv` | Reference | The 125 wards and their county, as used by the template, page and script |

---

## Part A: Box folder and upload link (about 5 minutes)

1. In Cornell Box (https://cornell.box.com), create a folder, e.g.
   **Drought relief submissions**. Keep it private to you, plus any team members you add.
2. On the folder: **⋯ (More options) → File Request**.
3. Optional: click **Edit** to add a short description (e.g. "Upload the checked file
   from the Drought Relief Data Checker") and ask uploaders for their email.
4. Copy the File Request link. It looks like `https://cornell.app.box.com/f/...`.

## Part B: Put the checker page online (about 5 minutes)

1. Open `docs/config.js` and paste the Box link between the quotes:
   ```js
   window.BOX_FILE_REQUEST_URL = "https://cornell.app.box.com/f/abc123...";
   ```
2. On github.com, create a new **public** repository, e.g. `drought-relief-checker`,
   with no README. Public is fine: the repo holds only the page and the blank template,
   never data. GitHub Pages from a private repository needs a paid plan.
3. In a terminal, from this folder:
   ```bash
   git init
   git add .
   git commit -m "Drought relief data checker"
   git branch -M main
   git remote add origin https://github.com/<your-username>/drought-relief-checker.git
   git push -u origin main
   ```
4. On GitHub: **Settings → Pages → Build and deployment**:
   *Source:* **Deploy from a branch**, *Branch:* **main**, folder **/docs** → **Save**.
5. After a minute or two the page is live at
   `https://<your-username>.github.io/drought-relief-checker/`.

## Part C: Test it, then send the link

1. Open the page, upload a filled-in template, fix anything flagged, fill in step 2,
   download the checked file in step 3, and upload it with the step 4 button.
2. Check the file appeared in your Box folder, then run Part D on it.
3. Send partners the page link. Example:

   > Please share your drought relief disbursement records using this page:
   > https://&lt;your-username&gt;.github.io/drought-relief-checker/
   > It checks your file in your browser, shows anything that needs fixing, and then gives you a
   > secure Cornell Box link to upload the checked file. You don't need an account. Partial
   > data is welcome. Questions: sc2875@cornell.edu

Partners who can't use the page can upload the filled-in template straight to the Box
link. The combining script checks those files too.

---

## Part D: Combine the submissions (on your computer)

**Get the files locally,** either way:
- **Box Drive** (Cornell's desktop sync app): the folder appears on your computer, e.g.
  `~/Library/CloudStorage/Box-Box/Drought relief submissions` on a Mac.
- Or, on the Box website, select the folder → **Download**, and unzip it.

**Run:**
```bash
cd combine
python3 -m venv .venv && source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python combine.py "/path/to/Drought relief submissions" --out combined_output --latest-per-org
```

- `--latest-per-org` keeps only the newest file from each organization, for when partners
  re-upload corrections. Leave it out to use every file. Files it skips are listed in
  `run_log.txt`.
- The script reads `.xlsx` and `.csv`, searching subfolders too. Old `.xls` files are listed
  in the log so you can ask for `.xlsx`.

**Outputs** (in `combined_output/`):

| File | One row per | Use |
|---|---|---|
| `rows_clean.csv` | record without errors | **Your analysis dataset** |
| `rows_all.csv` | every record read | Includes rows with errors; `status` and `issues` say what's wrong |
| `files_summary.csv` | uploaded file | Who sent it, contact, rows OK / warning / error, wards, months, households, their message |
| `coverage_ward_month.csv` | ward (all 125) | One column per month since 2024-01: `aid`, `zero` (reported no aid), or blank (no record) |
| `combined.xlsx` | – | All of the above as sheets |
| `run_log.txt` | – | Files read, skipped or not used, and totals |

Row columns: `source_file, organization, file_row` (the row number in the partner's
Excel file, so you can point them to it), then `county, ward, month, aid_type, amount,
amount_unit, households, agency, payment_form, notes, status, issues`.

**Duplicates are flagged, never dropped.** A row gets a warning if the same ward + month
+ aid type + agency appears twice in one file, or in another organization's file. The
second case often means the county and the implementing agency both reported the same
distribution, a double count you'll want to resolve.

---

## Checks (same rules on the page and in `combine.py`)

| Check | Result |
|---|---|
| Ward, County, Month, Type of aid missing | error |
| Month not readable as a month | error (`Mar 2024`, `03/2024`, `2024-3`, Excel dates are converted to `YYYY-MM`) |
| Month in the future | error |
| Amount not a number or `not applicable` | error (`3,000.00` and `KES 1,000` are converted to plain numbers) |
| Amount negative | error |
| Currency / unit or Households missing (unless amount is 0) | error |
| Households not a whole number | error |
| Notes over 4,000 characters | error |
| "No disbursement this month" with an empty amount | amount set to 0 |
| County not one of the 22 counties in scope | error |
| Ward not one of the 125 wards in scope | warning |
| Ward belongs to a different county than the one entered | warning |
| Month before 2024-01 | warning |
| "No disbursement" with a non-zero amount | warning |
| Duplicate ward + month + aid + agency | warning |

Ward and county names are matched ignoring case, spaces and punctuation, then written in
the template's spelling. Every automatic correction is listed: in the checked file's
**Checker notes** sheet, and in the `issues` column as `FIXED: …`.

## Counties and wards

The template, page and script use the **125 wards** in `ward_county.csv`, in **22 ASAL
counties**. Tharaka Nithi has none of these wards, so it isn't offered. Counties come
from NDMA's ward list (June 2026), with two exceptions to keep in mind:

| Ward | County used | Note |
|---|---|---|
| Churo/Amaya | Baringo | Not in NDMA's June 2026 list (last surveyed Aug 2025); please confirm |
| Kapedo/Napeitom | Turkana | Not in NDMA's June 2026 list (last surveyed Aug 2025); please confirm |
| Kapenguria | West Pokot | Not in NDMA's June 2026 list (last surveyed Aug 2025); please confirm |
| Evurore | Embu | NDMA lists it under both Embu and Tharaka Nithi; Embu used |

To change the list, edit `ward_county.csv` and ask for the template, page and script to be
regenerated from it, so all three stay identical.

In the template, the County column comes first. Once a county is chosen, the Ward
drop-down shows only that county's wards. This works in Excel and LibreOffice; Google
Sheets ignores filtered drop-downs, but the checker page catches any mismatch.

## The checked file partners upload

| Sheet | Contents |
|---|---|
| Data | Same layout as the template's Data tab (County first), with corrected values |
| Submission info | Organization, contact, message, source file, date checked, checker version, counts |
| Checker notes | Each warning and automatic correction, by row |
