# Refreshing dashboard data

The dashboard loads `dashboard-data.json` beside `index.html`; end users do not
upload files. Each month, regenerate the JSON from the updated workbook and
publish both files together:

```powershell
python -m pip install -r requirements.txt
python build_dashboard_data.py "C:\path\to\updated-workbook.xlsx"
```

The converter checks every workbook sheet and ignores `NA`, `N/A`, and `Blank`
values rather than treating them as zero. Country rows are grouped by region;
when a region has an `All` summary row, the summary is used instead of adding
the country rows a second time. It imports populated database, consent,
marketable, campaign, form, eStore, Engagement, and audience data. The
dashboard status tooltip identifies sheets with repeated placeholder values
and sheets with no numeric data. Product columns include `Blank` when the
workbook provides it. Review converter warnings for rows with missing keys and
the dashboard data-status note after each refresh.

Serve the dashboard from a web server; opening `index.html` directly from disk
does not allow the browser to fetch its adjacent JSON file. Keep the hosting
location access-controlled when the workbook contains sensitive data. The
dashboard revalidates the JSON on each visit so monthly updates stay current
without forcing browsers to discard their cached copy. For faster repeat loads,
enable ETag or Last-Modified validation and gzip/Brotli compression on the host.
