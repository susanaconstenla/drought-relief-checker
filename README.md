# Drought Relief Data Checker

A web page where partner organizations check their ward-month drought relief
disbursement data and upload it securely to the Cornell University research team.

**Page:** served by GitHub Pages from the `docs/` folder.

## How it works

1. The partner enters their organization details.
2. They upload their filled-in Excel template (or a .csv with the same headings). The
   page checks every row in the browser and shows anything that needs fixing.
3. They download a checked copy of the file. The page then clears their data from the
   browser.
4. They upload the checked file through a Cornell Box File Request link.

The page never sends data anywhere itself: files are read in the partner's browser, and
the only copy that leaves their computer is the one they upload to Box.

## Files

| Path | Purpose |
|---|---|
| `docs/index.html` | The checker page |
| `docs/config.js` | The Cornell Box File Request link used in step 4 |
| `docs/xlsx.full.min.js` | Spreadsheet reader (SheetJS 0.18.5) |
| `docs/drought_relief_disbursement_template.xlsx` | Blank template partners can download |

## Contact

Susana Constenla-Villoslada, Cornell University · sc2875@cornell.edu
