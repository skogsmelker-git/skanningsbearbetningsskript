import os
from pathlib import Path
from datetime import datetime
from pypdf import PdfReader

# ✅ Import your existing logic
# (Make sure this file is in same folder or adjust path)
from bearbetningsskript_TUESDAY import extract_personnummer_candidates


# ----------------------------
# CONFIG
# ----------------------------
INPUT_DIR = "input_pdfs"
LOG_FILE = "validation_log.txt"
SUMMARY_FILE = "validation_summary.txt"

MAX_PAGES = 10


# ----------------------------
# TEXT EXTRACTION
# ----------------------------
def extract_text(reader):
    text = ""
    for page in reader.pages:
        try:
            text += (page.extract_text() or "") + "\n"
        except:
            continue
    return text


# ----------------------------
# PERSONNUMMER LOGIC (FROM YOUR SCRIPT)
# ----------------------------
def get_personnummer(text):
    """
    Uses your script's extraction and merges all confidence levels.
    """
    try:
        high, medium, low = extract_personnummer_candidates(text)

        # Combine all levels, deduplicate
        all_ids = set(high + medium + low)
        return sorted(all_ids)

    except Exception as e:
        return [], f"Extraction error: {e}"


# ----------------------------
# VALIDATION
# ----------------------------
def validate_pdf(pdf_path):
    result = {
        "file": pdf_path.name,
        "page_count": 0,
        "personnummer": [],
        "issues": []
    }

    try:
        reader = PdfReader(pdf_path)
        result["page_count"] = len(reader.pages)
    except Exception as e:
        result["issues"].append(f"Unreadable PDF: {e}")
        return result

    # ✅ Page count check
    if result["page_count"] > MAX_PAGES:
        result["issues"].append(
            f"Too many pages ({result['page_count']} > {MAX_PAGES})"
        )

    # ✅ Extract text
    text = extract_text(reader)

    # ✅ Extract PNR using your logic
    pnrs = get_personnummer(text)

    if isinstance(pnrs, tuple):
        # extraction error case
        result["issues"].append(pnrs[1])
        pnrs = []
    else:
        result["personnummer"] = pnrs

    # ✅ Validation rules
    if len(pnrs) == 0:
        result["issues"].append("No personnummer found")

    elif len(pnrs) > 1:
        result["issues"].append(
            f"Multiple personnummer found ({len(pnrs)})"
        )

    return result


# ----------------------------
# LOGGING
# ----------------------------
def write_log(results):
    total = len(results)
    ok_count = 0
    issue_count = 0

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"Validation run: {datetime.now()}\n")
        f.write("=" * 70 + "\n\n")

        for r in results:
            f.write(f"FILE: {r['file']}\n")
            f.write(f"Pages: {r['page_count']}\n")
            f.write(f"Personnummer: {', '.join(r['personnummer']) or 'None'}\n")

            if r["issues"]:
                issue_count += 1
                f.write("ISSUES:\n")
                for issue in r["issues"]:
                    f.write(f"  - {issue}\n")
            else:
                ok_count += 1
                f.write("STATUS: OK\n")

            f.write("\n" + "-" * 70 + "\n\n")

        # ✅ SUMMARY SECTION
        f.write("\n" + "=" * 70 + "\n")
        f.write("SUMMARY\n")
        f.write("=" * 70 + "\n")
        f.write(f"Total files: {total}\n")
        f.write(f"OK files: {ok_count}\n")
        f.write(f"Files with issues: {issue_count}\n")


def write_summary(results):
    """
    Optional clean overview file (easy to scan quickly)
    """
    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        f.write("FILES WITH ISSUES:\n")
        f.write("=" * 50 + "\n")

        for r in results:
            if r["issues"]:
                f.write(f"{r['file']}\n")
                for issue in r["issues"]:
                    f.write(f"  - {issue}\n")
                f.write("\n")


# ----------------------------
# MAIN
# ----------------------------
def main():
    input_path = Path(INPUT_DIR)

    if not input_path.exists():
        print(f"Input folder not found: {INPUT_DIR}")
        return

    pdfs = list(input_path.glob("*.pdf"))

    if not pdfs:
        print("No PDFs found.")
        return

    results = []

    for pdf in pdfs:
        print(f"Checking: {pdf.name}")
        result = validate_pdf(pdf)
        results.append(result)

    write_log(results)
    write_summary(results)

    print("\n✅ Validation complete")
    print(f"Detailed log: {LOG_FILE}")
    print(f"Summary file: {SUMMARY_FILE}")


if __name__ == "__main__":
    main()