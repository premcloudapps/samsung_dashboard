# Refreshing dashboard data

The dashboard loads `dashboard-data.json` beside `index.html`; end users do not
upload files. Each month, regenerate the JSON from the updated workbook and
publish both files together:

```powershell
python -m pip install -r requirements.txt
python build_dashboard_data.py "C:\path\to\updated-workbook.xlsx"
```

The converter checks every workbook sheet. It imports the populated database,
consent, marketable, campaign, form, and Engaged Contacts sheets. Engaged
Contacts values are included in the Executive Overview and flagged for review
because they repeat across rows. The Marketable - I4B sheet is
mapped into the Marketable section; it currently has no metric values, so I4B
is shown as unavailable until the workbook is populated. The Registered, Sale, and
Abandoned eStore sheets and the Global Industry, Employee Range, and Lead Source
audience sheets are included as requested and flagged in the dashboard because
their counts repeat across rows. The Without I4B sheet is also included in the
Marketable section and flagged for repeated values. The Non Estore Contact
sheet is also included in the Marketable section; its values are flagged for
review. The current eStore, audience, and Marketable breakdown product values
do not reconcile to their totals. Other blank or repeated-placeholder sheets
are excluded. Review the converter output and dashboard data-status note after
each refresh.

Serve the dashboard from a web server; opening `index.html` directly from disk
does not allow the browser to fetch its adjacent JSON file. Keep the hosting
location access-controlled when the workbook contains sensitive data.
