"""Reusable Azure Blob image URL mapper. Defaults to the Sanitory worksheet."""
import csv
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlparse, unquote
from dotenv import load_dotenv
from openpyxl import load_workbook
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions

# One reusable mapper for all workbook categories.
PROJECT_DIR = Path(__file__).resolve().parent.parent
WORKBOOK_NAME = "Master product list_with_all_image_urls.xlsx"
AZURE_CONTAINER = "nivora"

# Worksheet name -> Azure folder prefix. Confirm a prefix in Azure if a category has no blobs.
CATEGORY_PREFIXES = {
    "plumbing": "Nivora Products/categories/Plumbing/",
    "electrical": "Nivora Products/categories/Electrical/",
    "Bath & Kitchen Essentials": "Nivora Products/categories/Bath & Kitchen Essentials/",
    "interior designing": "Nivora Products/categories/Interior Designing/",
    "Hardware": "Nivora Products/categories/Hardware/",
    "Sanitory": "Nivora Products/categories/Sanitory/",
}

URL_HEADERS = ("Image Url", "Image URL", "Image url")
# "hyperlink" displays short clickable text; use "full_url" if Nivora imports cell values directly.
URL_DISPLAY_MODE = "hyperlink"  # allowed: "hyperlink", "full_url"
SAS_VALID_HOURS = 24
OUTPUT_DIR = PROJECT_DIR / "output"
ENV_PATH = PROJECT_DIR / ".env"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
MIN_SCORE = 0.62
MIN_SCORE_GAP = 0.12


def norm(value):
    if value is None:
        return ""
    text = str(value).lower().strip().replace("&", " and ")
    text = re.sub(r"\be\.?\s*w\.?\s*c\.?\b", " ewc ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokens(value):
    ignored = {"the", "of", "and", "for", "with", "product", "products"}
    return {t for t in norm(value).split() if t and t not in ignored}


def blob_label(blob_name):
    # Use filename only. The full path repeats the product name and depresses scores.
    return Path(blob_name).stem


def value_for(row, header_map, *names):
    for name in names:
        idx = header_map.get(norm(name))
        if idx is not None:
            return row[idx - 1]
    return None


def product_text(row, header_map):
    # Excludes size and supplier so sizes can share an image.
    fields = [
        value_for(row, header_map, "Sub Category", "Subcategory", "Category"),
        value_for(row, header_map, "Item Type"),
        value_for(row, header_map, "Item Subtype 1"),
        value_for(row, header_map, "Item Subtype 2"),
        value_for(row, header_map, "Item Name"),
        value_for(row, header_map, "Item Description"),
        value_for(row, header_map, "Item Color"),
    ]
    return " ".join(str(v) for v in fields if v not in (None, ""))


def score_candidate(product, blob_name):
    p_norm, b_norm = norm(product), norm(blob_label(blob_name))
    p_tokens, b_tokens = tokens(p_norm), tokens(b_norm)
    if not p_tokens or not b_tokens:
        return 0.0
    overlap = len(p_tokens & b_tokens)
    precision = overlap / len(b_tokens)
    recall = overlap / len(p_tokens)
    dice = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
    bonus = 0.15 if b_norm and b_norm in p_norm else 0.0
    penalty = 0.0
    for group in (
        {"citi", "orrisa", "orissa"},
        {"indian", "western"},
        {"single", "piece"},
        {"corner", "standard", "steel", "stainless"},
    ):
        p_group, b_group = p_tokens & group, b_tokens & group
        if p_group and b_group and not p_group.intersection(b_group):
            penalty += 0.40
        elif p_group and not b_group:
            penalty += 0.12
    return max(0.0, min(1.0, 0.75 * dice + bonus - penalty))


def set_url_cell(cell, full_url):
    """Store a URL as either a short clickable label or the full cell value."""
    if URL_DISPLAY_MODE == "hyperlink":
        cell.value = "Open Image"
        cell.hyperlink = full_url
        cell.style = "Hyperlink"
    elif URL_DISPLAY_MODE == "full_url":
        cell.value = full_url
        cell.hyperlink = None
    else:
        raise ValueError("URL_DISPLAY_MODE must be 'hyperlink' or 'full_url'.")


def get_cell_url(cell):
    """Return the full URL, including hyperlink target when the cell label is short."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    value = cell.value
    return str(value).strip() if value and str(value).strip().lower().startswith(("https://", "http://")) else ""


def blob_name_from_url(url, container_name):
    """Extract the exact Azure blob name from an existing URL, ignoring SAS query parameters."""
    if not url:
        return ""
    parsed = urlparse(url)
    path = unquote(parsed.path).lstrip("/")
    prefix = f"{container_name}/"
    if path.startswith(prefix):
        return path[len(prefix):]
    return ""


def cleanup_old_mapper_outputs(output_xlsx, report_csv, workbook_stem):
    """Remove only legacy timestamped files generated by this mapper, across categories."""
    xlsx_pattern = re.compile(
        rf"^{re.escape(workbook_stem)}_.+_mapped_\d{{8}}_\d{{6}}\.xlsx$"
    )
    csv_pattern = re.compile(
        r"^.+_mapping_report_\d{8}_\d{6}\.csv$"
    )
    removed = 0
    for path in OUTPUT_DIR.iterdir():
        if not path.is_file() or path in (output_xlsx, report_csv):
            continue
        if xlsx_pattern.fullmatch(path.name) or csv_pattern.fullmatch(path.name):
            try:
                path.unlink()
                removed += 1
            except OSError as exc:
                print(f"WARNING: Could not remove old timestamped mapper output {path.name}: {exc}")
    print(f"Old timestamped mapper files removed: {removed}")


def choose_category():
    categories = list(CATEGORY_PREFIXES)
    print("\nAvailable categories:")
    for index, name in enumerate(categories, start=1):
        print(f"  {index}. {name}")
    while True:
        choice = input("Enter category number or exact worksheet name: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(categories):
            name = categories[int(choice) - 1]
            return name, CATEGORY_PREFIXES[name]
        if choice in CATEGORY_PREFIXES:
            return choice, CATEGORY_PREFIXES[choice]
        print("Invalid selection. Choose a listed number or exact worksheet name.")


def make_sas_url(blob_client, account_name, account_key):
    sas = generate_blob_sas(
        account_name=account_name,
        container_name=blob_client.container_name,
        blob_name=blob_client.blob_name,
        account_key=account_key,
        permission=BlobSasPermissions(read=True),
        expiry=datetime.now(timezone.utc) + timedelta(hours=SAS_VALID_HOURS),
        protocol="https",
    )
    return f"{blob_client.url}?{sas}"


def main():
    global WORKSHEET_NAME, AZURE_PREFIX
    WORKSHEET_NAME, AZURE_PREFIX = choose_category()

    if not ENV_PATH.is_file():
        raise FileNotFoundError(f".env not found at {ENV_PATH}")
    load_dotenv(ENV_PATH, override=True)
    connection = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "").strip()
    if not connection:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING is missing or empty.")
    parts = {}
    for segment in connection.split(";"):
        if "=" in segment:
            k, v = segment.split("=", 1)
            parts[k.strip().lower()] = v.strip()
    account_name, account_key = parts.get("accountname"), parts.get("accountkey")
    if not account_name or not account_key:
        raise RuntimeError("Connection string must contain AccountName and AccountKey to create SAS URLs.")
    if not AZURE_PREFIX.endswith("/"):
        raise ValueError("AZURE_PREFIX must end with '/' to restrict mapping to one category.")

    workbook_path = PROJECT_DIR / WORKBOOK_NAME
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    service = BlobServiceClient.from_connection_string(connection)
    container = service.get_container_client(AZURE_CONTAINER)
    print(f"Worksheet: {WORKSHEET_NAME}")
    print(f"Azure prefix: {AZURE_PREFIX}")
    blob_names = [
        b.name for b in container.list_blobs(name_starts_with=AZURE_PREFIX)
        if Path(b.name).suffix.lower() in IMAGE_EXTENSIONS
    ]
    if not blob_names:
        raise RuntimeError("No image blobs found under the configured Azure prefix.")
    print(f"Image blobs found: {len(blob_names)}")

    wb = load_workbook(workbook_path)
    if WORKSHEET_NAME not in wb.sheetnames:
        raise KeyError(f"Worksheet {WORKSHEET_NAME!r} not found. Sheets: {wb.sheetnames}")
    ws = wb[WORKSHEET_NAME]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    header_map = {norm(h): i + 1 for i, h in enumerate(headers) if h is not None}
    url_col = next((header_map[norm(h)] for h in URL_HEADERS if norm(h) in header_map), None)
    if not url_col:
        raise KeyError(f"Image URL column not found. Headers: {headers}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    # Stable per-category names prevent a new duplicate file on every run.
    # Close the previous output workbook/report before rerunning, otherwise Windows may lock it.
    safe_category = re.sub(r"[^A-Za-z0-9_-]+", "_", WORKSHEET_NAME).strip("_")
    output_xlsx = OUTPUT_DIR / f"{workbook_path.stem}_{safe_category}_mapped.xlsx"
    report_csv = OUTPUT_DIR / f"{safe_category}_mapping_report.csv"
    report = []
    counts = {"REFRESHED_SAME_BLOB": 0, "UPDATED_BY_MATCH": 0, "MATCHED": 0, "AMBIGUOUS": 0, "UNMATCHED": 0, "EXISTING_KEPT_UNCERTAIN": 0}
    available_blobs = set(blob_names)

    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        existing_url = ws.cell(r, url_col).value
        item_id = value_for(row, header_map, "Item ID")
        item_name = value_for(row, header_map, "Item Name")
        subtype1 = value_for(row, header_map, "Item Subtype 1")
        subtype2 = value_for(row, header_map, "Item Subtype 2")
        description = value_for(row, header_map, "Item Description")
        record = {
            "worksheet": WORKSHEET_NAME, "excel_row": r, "item_id": item_id,
            "item_name": item_name, "subtype_1": subtype1, "subtype_2": subtype2,
            "description": description, "existing_url": "YES" if existing_url else "NO",
            "status": "", "matched_blob": "", "score": "", "candidate_blobs": "",
            "full_url": "", "notes": "",
        }
        old_full_url = get_cell_url(ws.cell(r, url_col)) if existing_url and str(existing_url).strip() else ""
        had_existing_value = bool(existing_url and str(existing_url).strip())

        # First choice for existing URLs: refresh the exact same blob by extracting its path.
        # This avoids product-name matching ambiguity and does not require the old SAS token to be valid.
        old_blob = blob_name_from_url(old_full_url, AZURE_CONTAINER) if had_existing_value else ""
        if old_blob and old_blob in available_blobs:
            blob_client = container.get_blob_client(old_blob)
            fresh_url = make_sas_url(blob_client, account_name, account_key)
            set_url_cell(ws.cell(r, url_col), fresh_url)
            record["full_url"] = fresh_url
            record["matched_blob"] = old_blob
            record["status"] = "REFRESHED_SAME_BLOB"
            record["notes"] = "Fresh SAS URL generated for the exact blob path from the existing URL; no product-name guess used."
            counts["REFRESHED_SAME_BLOB"] += 1
            report.append(record)
            continue

        # If the old URL does not identify a blob in this category, use product matching.
        # Uncertain matches never overwrite an existing URL with a guess.
        product = product_text(row, header_map)
        # Exact filename or parent-folder label contained in product attributes wins.
        # This avoids repeated folder/filename words and supports "Standard" as a subtype.
        product_norm = norm(product)
        exact = []
        for b in blob_names:
            filename_label = norm(Path(b).stem)
            folder_label = norm(Path(b).parent.name)
            if (
                (filename_label and filename_label in product_norm)
                or (folder_label and folder_label in product_norm)
            ):
                exact.append(b)
        scored = sorted(((score_candidate(product, b), b) for b in blob_names), reverse=True)
        if len(exact) == 1:
            best_score, best_blob = 1.0, exact[0]
            second_score = max((s for s, b in scored if b != best_blob), default=0.0)
        else:
            best_score, best_blob = scored[0]
            second_score = scored[1][0] if len(scored) > 1 else 0.0

        if len(exact) > 1:
            record["status"] = "AMBIGUOUS"
            record["candidate_blobs"] = " || ".join(exact)
            if had_existing_value:
                record["full_url"] = old_full_url
                record["status"] = "EXISTING_KEPT_UNCERTAIN"
                record["notes"] = "Multiple exact candidates; existing cell retained because a safe replacement could not be determined."
                counts["EXISTING_KEPT_UNCERTAIN"] += 1
            else:
                record["notes"] = "Multiple exact candidates; left blank for review."
                counts["AMBIGUOUS"] += 1
        elif best_score >= MIN_SCORE and (best_score - second_score >= MIN_SCORE_GAP or best_score >= 0.99):
            blob_client = container.get_blob_client(best_blob)
            full_url = make_sas_url(blob_client, account_name, account_key)
            set_url_cell(ws.cell(r, url_col), full_url)
            record["full_url"] = full_url
            record["status"] = "UPDATED_BY_MATCH" if had_existing_value else "MATCHED"
            record["matched_blob"] = best_blob
            record["score"] = round(best_score, 3)
            record["candidate_blobs"] = " || ".join(f"{s:.3f}:{b}" for s, b in scored[:3])
            record["notes"] = f"Fresh read-only SAS URL generated; expires in {SAS_VALID_HOURS} hours."
            counts[record["status"]] += 1
        elif best_score >= MIN_SCORE:
            record["status"] = "AMBIGUOUS"
            record["score"] = round(best_score, 3)
            record["candidate_blobs"] = " || ".join(f"{s:.3f}:{b}" for s, b in scored[:5])
            if had_existing_value:
                record["full_url"] = old_full_url
                record["status"] = "EXISTING_KEPT_UNCERTAIN"
                record["notes"] = "Top candidates too close; existing cell retained because a safe replacement could not be determined."
                counts["EXISTING_KEPT_UNCERTAIN"] += 1
            else:
                record["notes"] = "Top candidates too close; URL left blank for review."
                counts["AMBIGUOUS"] += 1
        else:
            record["status"] = "UNMATCHED"
            record["score"] = round(best_score, 3)
            record["candidate_blobs"] = " || ".join(f"{s:.3f}:{b}" for s, b in scored[:5])
            if had_existing_value:
                record["full_url"] = old_full_url
                record["status"] = "EXISTING_KEPT_UNCERTAIN"
                record["notes"] = "No reliable candidate found; existing cell retained rather than replaced with a guess."
                counts["EXISTING_KEPT_UNCERTAIN"] += 1
            else:
                record["notes"] = "No sufficiently reliable match; URL left blank."
                counts["UNMATCHED"] += 1
        report.append(record)

    # CSV always includes full_url, even when Excel displays "Open Image".
    wb.save(output_xlsx)
    fields = list(report[0].keys()) if report else []
    with report_csv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(report)
    # Only clean timestamped files generated by this mapper for this worksheet.
    # Cleanup happens after the new workbook and report are saved successfully.
    cleanup_old_mapper_outputs(output_xlsx, report_csv, workbook_path.stem)
    print("\n========== MAPPING SUMMARY ==========")
    for status, count in counts.items():
        print(f"{status}: {count}")
    print(f"Output workbook: {output_xlsx}")
    print(f"CSV report: {report_csv}")
    print(f"URL display mode: {URL_DISPLAY_MODE}")
    if URL_DISPLAY_MODE == "hyperlink":
        print("Note: Excel shows 'Open Image'; the CSV report retains full URLs. If Nivora imports cell values directly, switch URL_DISPLAY_MODE to 'full_url'.")
    print("Original workbook was not modified. No images were uploaded.")
    print("Existing URLs are refreshed from their exact original blob path when that blob exists; uncertain product matches keep old values or remain blank.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}")
        sys.exit(1)
