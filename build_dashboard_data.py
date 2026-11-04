import argparse
import json
import math
from pathlib import Path

from openpyxl import load_workbook


PRODUCTS = ["Mobile", "Display", "Climate", "Others", "Non Standard"]
REGIONS = ["SEUK SMB", "SEBN", "SEG", "SESA", "SEF", "SEP", "SEH", "SEACE"]
MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]

METRIC_SHEETS = {
    "Total Database": ("db", "Total Database"),
    "Total - Opt-in contacts": ("optIn", "Total - Opt-in Contacts"),
    "Total - Opt-Out contacts": ("optOut", "Total - Opt-Out Contacts"),
    "Total - Opt-in-out Contacts": (
        "neither",
        "Total - Niether Opt-in/Out Contacts",
    ),
    "New Contacts": ("newContacts", "New Contacts"),
    "Marketable Database": ("market", "Marketable Database"),
}
ESTORE_SHEETS = {
    "Estore Contacts - Registered": (
        "registered",
        "Estore Contacts - Registered",
    ),
    "Estore Contacts - SALE ": ("sale", "Estore Contacts - Sale"),
    "Estore Contacts - Abandoned": (
        "abandoned",
        "Estore Contacts - Abandoned",
    ),
}
ENGAGEMENT_SHEETS = {
    "Engaged Contacts": ("engaged", "Engaged Contacts"),
}
MARKETABLE_BREAKDOWN_SHEETS = {
    "Marketable - I4B": (
        "i4b",
        "Marketable Database - I4B",
    ),
    "Marketable - Without I4B": (
        "withoutI4B",
        " Marketable Database - Without I4B",
    ),
    "Marketable - Non Estore Contact": (
        "nonEstore",
        "Marketable - Non Estore Contacts",
    ),
}
AUDIENCE_SHEETS = {
    "Marketable - Global Industry": (
        "industry",
        "Global Industry",
        "Marketable - Global Industry Count",
    ),
    "Marketable -No of employee": (
        "employees",
        "No of employee",
        "Marketable - No of employee Count",
    ),
    "Marketable - Lead Source": (
        "sources",
        "Lead_Source",
        "Lead_Source count",
    ),
}
CAMPAIGN_SHEET = "Campaign Peformance"
FORM_SHEET = "Form submission"


def normalized(value):
    return " ".join(str(value or "").strip().lower().split())


def number(value, location, blank_as_zero=False):
    if value is None or value == "":
        if blank_as_zero:
            return 0
        raise ValueError(f"{location}: expected a number, found a blank cell.")
    if isinstance(value, str):
        value = value.replace(",", "").strip()
    try:
        result = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{location}: expected a number, found {value!r}.") from error
    if not math.isfinite(result) or result < 0 or not result.is_integer():
        raise ValueError(f"{location}: invalid count {value!r}.")
    return int(result)


def rows_for_metric(workbook, sheet_name, metric_name):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    header_indexes = {name: index for index, name in enumerate(headers) if name}
    metric_index = header_indexes.get(normalized(metric_name))
    if metric_index is None:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing the {metric_name!r} column."
        )
    product_indexes = []
    for product in PRODUCTS:
        if normalized(product) not in header_indexes:
            raise ValueError(
                f"Sheet {sheet_name!r} is missing the {product!r} column."
            )
        product_indexes.append(header_indexes[normalized(product)])

    result = {}
    for row_number, row in enumerate(iterator, start=2):
        key_cells = tuple(row[index] if index < len(row) else None for index in range(4))
        value = row[metric_index] if metric_index < len(row) else None
        product_values = [
            row[index] if index < len(row) else None for index in product_indexes
        ]
        if all(cell is None or cell == "" for cell in (value, *product_values)):
            continue
        if any(cell is None or cell == "" for cell in key_cells):
            raise ValueError(f"Sheet {sheet_name!r}, row {row_number}: missing key.")
        year, quarter, month, region = [str(cell).strip() for cell in key_cells]
        if month not in MONTHS or region not in REGIONS:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown month or region."
            )
        if quarter not in {"Q1", "Q2", "Q3", "Q4"}:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown quarter {quarter!r}."
            )
        key = (year, quarter, month, region)
        if key in result:
            raise ValueError(
                f"Sheet {sheet_name!r} contains a duplicate key on row {row_number}."
            )
        result[key] = {
            "value": number(value, f"{sheet_name!r} row {row_number}"),
            "products": [
                number(cell, f"{sheet_name!r} row {row_number}", blank_as_zero=True)
                for cell in product_values
            ],
        }
    return result


def rows_for_dimension(workbook, sheet_name, category_name, metric_name):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    header_indexes = {name: index for index, name in enumerate(headers) if name}
    required = [
        "year",
        "quater",
        "month",
        "region",
        normalized(category_name),
        normalized(metric_name),
        *[normalized(product) for product in PRODUCTS],
    ]
    missing = [name for name in required if name not in header_indexes]
    if missing:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing columns: {', '.join(missing)}."
        )

    result = []
    for row_number, row in enumerate(iterator, start=2):
        indexes = [header_indexes[name] for name in required]
        cells = [row[index] if index < len(row) else None for index in indexes]
        if all(value is None or value == "" for value in cells[5:]):
            continue
        year, quarter, month, region, category, value, *products = cells
        if any(value is None or value == "" for value in (year, quarter, month, region)):
            raise ValueError(f"Sheet {sheet_name!r}, row {row_number}: missing key.")
        year, quarter, month, region = [
            str(value).strip() for value in (year, quarter, month, region)
        ]
        if month not in MONTHS or region not in REGIONS:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown month or region."
            )
        if quarter not in {"Q1", "Q2", "Q3", "Q4"}:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown quarter {quarter!r}."
            )
        result.append(
            {
                "year": year,
                "quarter": quarter,
                "month": month,
                "region": region,
                "category": str(category or "Blank").strip() or "Blank",
                "value": number(value, f"{sheet_name!r} row {row_number}"),
                "products": [
                    number(
                        product,
                        f"{sheet_name!r} row {row_number}",
                        blank_as_zero=True,
                    )
                    for product in products
                ],
            }
        )
    return result


def cell_value(row, indexes, key):
    index = indexes.get(key)
    return row[index] if index is not None and index < len(row) else None


def nullable_number(value, location):
    if value is None or value == "":
        return None
    result = float(value)
    if not math.isfinite(result):
        raise ValueError(f"{location}: invalid number {value!r}.")
    return result


def read_detail_sheet(workbook, sheet_name, kind):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    indexes = {name: index for index, name in enumerate(headers) if name}
    quarter_header = "quarter" if "quarter" in indexes else "quater"
    required = (
        {
            "year",
            quarter_header,
            "month",
            "region",
            "type",
            "campaign name",
            "total_sent",
            "total_delivered",
            "unique_html_opens",
            "total_clicks",
            "unsubscribes",
        }
        if kind == "campaign"
        else {
            "year",
            quarter_header,
            "month",
            "region",
            "form name",
            "type",
            "total submission",
            "unique submission",
            "new contacts",
            "marketable contacts",
        }
    )
    missing = required - indexes.keys()
    if missing:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing columns: {', '.join(sorted(missing))}."
        )

    result = []
    for row_number, row in enumerate(iterator, start=2):
        if not any(value is not None and value != "" for value in row):
            continue
        get = lambda key: cell_value(row, indexes, key)
        if kind == "campaign":
            result.append(
                {
                    "year": str(get("year") or "").strip(),
                    "quarter": str(get(quarter_header) or "").strip(),
                    "month": str(get("month") or "").strip(),
                    "region": str(get("region") or "").strip(),
                    "type": str(get("type") or "").strip(),
                    "campaign": str(get("campaign name") or "").strip(),
                    "sent": nullable_number(get("total_sent"), f"{sheet_name} row {row_number}"),
                    "delivered": nullable_number(get("total_delivered"), f"{sheet_name} row {row_number}"),
                    "uniqueOpens": nullable_number(get("unique_html_opens"), f"{sheet_name} row {row_number}"),
                    "clicks": nullable_number(get("total_clicks"), f"{sheet_name} row {row_number}"),
                    "unsubscribes": nullable_number(get("unsubscribes"), f"{sheet_name} row {row_number}"),
                }
            )
        else:
            result.append(
                {
                    "year": str(get("year") or "").strip(),
                    "quarter": str(get(quarter_header) or "").strip(),
                    "month": str(get("month") or "").strip(),
                    "region": str(get("region") or "").strip(),
                    "name": str(get("form name") or "").strip(),
                    "type": str(get("type") or "").strip(),
                    "submissions": nullable_number(get("total submission"), f"{sheet_name} row {row_number}"),
                    "uniqueSubmissions": nullable_number(get("unique submission"), f"{sheet_name} row {row_number}"),
                    "newContacts": nullable_number(get("new contacts"), f"{sheet_name} row {row_number}"),
                    "marketableContacts": nullable_number(get("marketable contacts"), f"{sheet_name} row {row_number}"),
                }
            )
    return result


def audit_sheets(workbook):
    audit = []
    loaded_sheets = (
        set(METRIC_SHEETS)
        | set(ESTORE_SHEETS)
        | set(ENGAGEMENT_SHEETS)
        | set(MARKETABLE_BREAKDOWN_SHEETS)
        | set(AUDIENCE_SHEETS)
        | {CAMPAIGN_SHEET, FORM_SHEET}
    )
    for worksheet in workbook.worksheets:
        rows = list(worksheet.iter_rows(values_only=True))
        headers = [normalized(value) for value in rows[0]] if rows else []
        populated_rows = [
            row for row in rows[1:] if any(value is not None and value != "" for value in row)
        ]
        if worksheet.title in loaded_sheets:
            if worksheet.title in MARKETABLE_BREAKDOWN_SHEETS:
                metric_name = normalized(
                    MARKETABLE_BREAKDOWN_SHEETS[worksheet.title][1]
                )
                metric_index = headers.index(metric_name) if metric_name in headers else -1
                product_indexes = [
                    headers.index(normalized(product))
                    for product in PRODUCTS
                    if normalized(product) in headers
                ]
                has_values = any(
                    (metric_index >= 0 and metric_index < len(row) and row[metric_index] not in (None, ""))
                    or any(index < len(row) and row[index] not in (None, "") for index in product_indexes)
                    for row in rows[1:]
                )
                status = "included-placeholder" if has_values else "blank"
            else:
                status = (
                    "included-placeholder"
                    if worksheet.title in ESTORE_SHEETS
                    or worksheet.title in ENGAGEMENT_SHEETS
                    or worksheet.title in AUDIENCE_SHEETS
                    else "loaded"
                )
        else:
            metric_index = 5 if len(headers) > 5 else 4
            values = [
                row[metric_index]
                for row in populated_rows
                if metric_index < len(row) and isinstance(row[metric_index], (int, float))
            ]
            if not values:
                status = "blank"
            elif len(set(values)) <= max(12, math.floor(len(values) * 0.1)):
                status = "placeholder"
            else:
                status = "not-integrated"
        audit.append(
            {
                "name": worksheet.title,
                "populatedRows": len(populated_rows),
                "status": status,
            }
        )
    return audit


def build_payload(input_path):
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    metric_rows = {}
    for sheet_name, (metric, column_name) in METRIC_SHEETS.items():
        metric_rows[metric] = rows_for_metric(workbook, sheet_name, column_name)
    estore_rows = {}
    for sheet_name, (metric, column_name) in ESTORE_SHEETS.items():
        estore_rows[metric] = rows_for_metric(workbook, sheet_name, column_name)
    engagement_rows = {}
    for sheet_name, (metric, column_name) in ENGAGEMENT_SHEETS.items():
        engagement_rows[metric] = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    marketable_breakdown = {}
    for sheet_name, (metric, column_name) in MARKETABLE_BREAKDOWN_SHEETS.items():
        marketable_breakdown[metric] = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    audience_rows = {}
    for sheet_name, (dimension, category_column, metric_column) in AUDIENCE_SHEETS.items():
        audience_rows[dimension] = rows_for_dimension(
            workbook,
            sheet_name,
            category_column,
            metric_column,
        )

    records = {}
    for key, values in metric_rows["db"].items():
        records[key] = {
            "year": key[0],
            "q": key[1],
            "m": key[2],
            "r": key[3],
        }
        records[key]["db"] = values["value"]
        records[key]["dbProducts"] = values["products"]
    for metric, rows in metric_rows.items():
        if metric == "db":
            continue
        for key, values in rows.items():
            if key not in records:
                raise ValueError(
                    f"Sheet data contains a record not found in Total Database: {key!r}."
                )
            records[key][metric] = values["value"]
            records[key][f"{metric}Products"] = values["products"]

    for record in records.values():
        if sum(record["dbProducts"]) != record["db"]:
            raise ValueError(
                f"Product totals do not match Total Database for "
                f"{record['year']}, {record['m']}, {record['r']}."
            )

    metrics = sorted(metric_rows.keys())
    estore_records = {}
    for metric, rows in estore_rows.items():
        for key, values in rows.items():
            record = estore_records.setdefault(
                key,
                {
                    "year": key[0],
                    "quarter": key[1],
                    "month": key[2],
                    "region": key[3],
                },
            )
            record[metric] = values["value"]
            record[f"{metric}Products"] = values["products"]
    years = {record["year"] for record in records.values()}
    expected_records = len(years) * len(MONTHS) * len(REGIONS)
    return {
        "sourceWorkbook": input_path.name,
        "metrics": metrics,
        "records": list(records.values()),
        "estoreMetrics": sorted(estore_rows.keys()),
        "estore": list(estore_records.values()),
        "engagementMetrics": sorted(engagement_rows.keys()),
        "engagement": {
            metric: [
                {
                    "year": key[0],
                    "quarter": key[1],
                    "month": key[2],
                    "region": key[3],
                    "value": values["value"],
                    "products": values["products"],
                }
                for key, values in rows.items()
            ]
            for metric, rows in engagement_rows.items()
        },
        "marketableBreakdownMetrics": sorted(marketable_breakdown.keys()),
        "marketableBreakdown": {
            metric: [
                {
                    "year": key[0],
                    "quarter": key[1],
                    "month": key[2],
                    "region": key[3],
                    "value": values["value"],
                    "products": values["products"],
                }
                for key, values in rows.items()
            ]
            for metric, rows in marketable_breakdown.items()
        },
        "audienceDimensions": audience_rows,
        "missingDatabaseRows": max(0, expected_records - len(records)),
        "campaigns": read_detail_sheet(workbook, CAMPAIGN_SHEET, "campaign"),
        "forms": read_detail_sheet(workbook, FORM_SHEET, "form"),
        "sheetAudit": audit_sheets(workbook),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert the Samsung QBR workbook into data the static dashboard can load."
    )
    parser.add_argument("workbook", type=Path, help="Path to the updated .xlsx workbook")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(__file__).with_name("dashboard-data.json"),
        help="Output JSON path (defaults to dashboard-data.json beside this script)",
    )
    args = parser.parse_args()
    if not args.workbook.is_file():
        parser.error(f"Workbook not found: {args.workbook}")
    payload = build_payload(args.workbook)
    args.output.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    included = sum(
        sheet["status"] in {"loaded", "included-placeholder"}
        for sheet in payload["sheetAudit"]
    )
    print(
        f"Wrote {len(payload['records'])} metric rows, "
        f"{len(payload['estore'])} eStore rows, {len(payload['campaigns'])} campaigns, "
        f"{len(payload['forms'])} forms, and "
        f"{sum(len(rows) for rows in payload['audienceDimensions'].values())} audience rows "
        f"from {included}/{len(payload['sheetAudit'])} "
        f"included workbook sheets to {args.output}."
    )
    for sheet in payload["sheetAudit"]:
        if sheet["status"] not in {"loaded", "included-placeholder"}:
            print(f"Skipped {sheet['name']}: {sheet['status']}")
        elif sheet["status"] == "included-placeholder":
            print(f"Included with placeholder warning: {sheet['name']}")


if __name__ == "__main__":
    main()
