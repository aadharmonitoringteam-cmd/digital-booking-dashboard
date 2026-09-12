# Daily Booking Dashboard

Responsive dashboard for the **Daily Booking Report** Excel file.

### Current sample data
- Daily Booking Report rows: 505
- Divisions: 8
- Sample total transactions: 21,986
- Sample total amount: ₹1,272,141.63

## What it does

- Reads `data/Digital Txn Status.xlsx` directly in the browser.
- Works on desktop and mobile.
- Division filter and office search.
- KPI cards for transactions, amount, offices and divisions.
- Division-wise transaction and amount charts.
- Payment/category chart.
- Top 10 offices.
- Sorts office performance by transaction count.
- Downloads the filtered office table as CSV.
- Has **Upload Daily Excel** so you can test a new daily report immediately without changing the website.
- No server/database is required.

## Daily update on GitHub

1. Keep the Excel file name exactly:
   `data/Digital Txn Status.xlsx`
2. Every day replace that file with the new report.
3. Commit and push the changed Excel file to GitHub.
4. Refresh the dashboard link.

The website fetches the Excel file each time it opens, so the dashboard code does **not** need to be edited for normal daily report replacement.

### Important Excel requirement

The workbook should contain a sheet named **Daily Booking Report** with the same column structure, including:
- Office ID
- Office Name
- Product
- the `(Cnt)` / `(Amt)` booking columns
- Division

The app ignores the summary row where Office ID is blank.

## GitHub Pages

Create a GitHub repository and upload:
- `index.html`
- `app.js`
- `styles.css`
- `data/Digital Txn Status.xlsx`

Then in GitHub:
**Settings → Pages → Deploy from a branch → main → / (root) → Save**

After GitHub publishes it, your dashboard will be available at your GitHub Pages URL.

## Optional: local testing

Because browsers can block local Excel fetching with `file://`, use a small local web server, or simply upload the project to GitHub Pages.

## Privacy

Excel parsing happens in the user's browser. The page does not upload the selected Excel file to an application server.
