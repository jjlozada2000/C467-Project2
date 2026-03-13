"""
i had to redownload excel for the most recent assignments and for some reason it cooked my cpu
"""

import argparse
import csv
import os
import datetime
import openpyxl
from pymongo import MongoClient


def connect_to_mongo():
    client = MongoClient("mongodb://localhost:27017/")
    db = client["proj2_database"]
    return db


def load_excel_file(file_path):
    print(f"\n[INFO] Loading: {file_path}")

    workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet = workbook.active

    rows_as_dicts = []
    headers = []
    first_row = True

    for row in sheet.iter_rows(values_only=True):
        if first_row:
            for cell in row:
                headers.append(str(cell).strip() if cell is not None else "")
            first_row = False
            continue

        if all(cell is None for cell in row):
            continue

        row_dict = {}
        for i in range(len(headers)):
            value = row[i] if i < len(row) else None

            # build # comes in as a datetime object from Excel and converts to readable string
            if headers[i] == "Build #" and isinstance(value, datetime.datetime):
                value = value.strftime("%-m/%-d/%Y")
            elif value is None:
                value = ""
            else:
                value = str(value).strip()

            row_dict[headers[i]] = value

        rows_as_dicts.append(row_dict)

    workbook.close()
    print(f"[INFO] Loaded {len(rows_as_dicts)} rows")
    return rows_as_dicts


def insert_into_collection(db, collection_name, rows):
    collection = db[collection_name]
    collection.drop()
    collection.insert_many(rows)
    print(f"[INFO] Inserted {len(rows)} rows into '{collection_name}'")
    return collection


def write_to_text_file(filename, label, results):
    with open(filename, "a", encoding="utf-8") as f:
        f.write(f"QUERY: {label}\n")

        if len(results) == 0:
            f.write("No results found.\n")
        else:
            f.write(f"Total results: {len(results)}\n\n")
            for i in range(len(results)):
                row = results[i]
                f.write(f"Record {i + 1} \n")
                for key in row:
                    if key != "_id":
                        f.write(f"  {key}: {row[key]}\n")
                f.write("\n")

    print(f"[INFO] Results written to '{filename}'")


def write_to_csv(filename, results):
    if len(results) == 0:
        print(f"[INFO] No results to write to {filename}")
        return

    all_keys = []
    for row in results:
        for key in row.keys():
            if key != "_id" and key not in all_keys:
                all_keys.append(key)

    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_keys)
        writer.writeheader()
        for row in results:
            clean_row = {}
            for key in all_keys:
                clean_row[key] = row.get(key, "")
            writer.writerow(clean_row)

    print(f"[INFO] CSV saved: '{filename}' ({len(results)} rows)")


# test owner
def search_by_testuser(db, collection_names, username, output_text_file):
    print(f"\n[QUERY] Test Owner = '{username}'")
    all_results = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            owner = doc.get("Test Owner", "").lower()
            if username.lower() in owner:
                found.append(doc)
        print(f"  '{name}': {len(found)} results")
        all_results = all_results + found

    write_to_text_file(output_text_file, f"Test Owner = {username}", all_results)

    safe_name = username.replace(" ", "_")
    write_to_csv(f"testuser_{safe_name}.csv", all_results)

    print(f"[INFO] Total: {len(all_results)}")
    return all_results


# repeatable
def search_repeatable(db, collection_names, output_text_file):
    print("\n[QUERY] All Repeatable bugs")
    all_results = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            value = doc.get("Repeatable?", "").lower()
            if value.startswith("yes"):
                found.append(doc)
        print(f"  '{name}': {len(found)} results")
        all_results = all_results + found

    write_to_text_file(output_text_file, "All Repeatable Bugs", all_results)
    print(f"[INFO] Total: {len(all_results)}")
    return all_results


# blocker
def search_blocker(db, collection_names, output_text_file):
    print("\n[QUERY] All Blocker bugs")
    all_results = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            value = doc.get("Blocker?", "").lower()
            if value.startswith("yes"):
                found.append(doc)
        print(f"  '{name}': {len(found)} results")
        all_results = all_results + found

    write_to_text_file(output_text_file, "All Blocker Bugs", all_results)
    print(f"[INFO] Total: {len(all_results)}")
    return all_results


# repeatable and blocker
def search_repeatable_and_blocker(db, collection_names, output_text_file):
    print("\n[QUERY] Repeatable AND Blocker bugs")
    all_results = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            repeatable = doc.get("Repeatable?", "").lower()
            blocker = doc.get("Blocker?", "").lower()
            if repeatable.startswith("yes") and blocker.startswith("yes"):
                found.append(doc)
        print(f"  '{name}': {len(found)} results")
        all_results = all_results + found

    write_to_text_file(output_text_file, "Repeatable AND Blocker Bugs", all_results)
    print(f"[INFO] Total: {len(all_results)}")
    return all_results


# build date
def search_by_builddate(db, collection_names, build_date, output_text_file):
    print(f"\n[QUERY] Build date containing '{build_date}'")
    all_results = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            build = doc.get("Build #", "").lower()
            if build_date.lower() in build:
                found.append(doc)
        print(f"  '{name}': {len(found)} results")
        all_results = all_results + found

    write_to_text_file(output_text_file, f"Build Date: {build_date}", all_results)
    print(f"[INFO] Total: {len(all_results)}")
    return all_results


# iansearch is database logic call
def iansearch(db, collection_names, output_text_file):
    print("\n[DATABASE LOGIC CALL] Running iansearch...")

    all_bugs = []

    for name in collection_names:
        col = db[name]
        all_docs = list(col.find({}))
        found = []
        for doc in all_docs:
            repeatable = doc.get("Repeatable?", "").lower()
            blocker = doc.get("Blocker?", "").lower()
            if repeatable.startswith("no") and blocker.startswith("no"):
                doc["_source"] = name
                found.append(doc)
        print(f"  '{name}': {len(found)} non-critical bugs")
        all_bugs = all_bugs + found

    print(f"  Total to analyze: {len(all_bugs)}")

    def get_word_set(bug):
        text = bug.get("Test Case", "") + " " + bug.get("Expected Result", "") + " " + bug.get("Actual Result", "")
        words = text.lower().split()
        word_set = set()
        for word in words:
            clean = ""
            for char in word:
                if char.isalpha() or char.isdigit():
                    clean = clean + char
            if len(clean) >= 3:
                word_set.add(clean)
        return word_set

    MATCH_THRESHOLD = 2
    bug_groups = []

    for bug in all_bugs:
        bug["_words"] = get_word_set(bug)
        placed = False
        for group in bug_groups:
            shared = group[0]["_words"].intersection(bug["_words"])
            if len(shared) >= MATCH_THRESHOLD:
                group.append(bug)
                placed = True
                break
        if not placed:
            bug_groups.append([bug])

    output_rows = []

    for group in bug_groups:
        build_numbers = set()
        for bug in group:
            build_val = bug.get("Build #", "")
            if build_val != "":
                build_numbers.add(build_val)

        if len(build_numbers) > 1:
            rep = group[0]
            output_rows.append({
                "Test Case":                    rep.get("Test Case", ""),
                "Expected Result":              rep.get("Expected Result", ""),
                "Actual Result":                rep.get("Actual Result", ""),
                "Number of Builds":             len(build_numbers),
                "Build Numbers":                ", ".join(sorted(build_numbers)),
                "Total Matching Bugs in Group": len(group),
                "Collections":                  ", ".join(set(b["_source"] for b in group))
            })

    write_to_csv("iansearch_results.csv", output_rows)
    write_to_text_file(output_text_file, "DATABASE LOGIC CALL: iansearch()", output_rows)

    print(f"[INFO] iansearch done. {len(output_rows)} multi-build bug groups found.")
    return output_rows


def main():
    parser = argparse.ArgumentParser(description="QA Database Tool — proj2")

    parser.add_argument("--files",         nargs="+", required=True, help="Excel file(s) to load")
    parser.add_argument("--testuser",      type=str,  default=None,  help='Search by Test Owner. Ex: --testuser "Kevin Chaja"')
    parser.add_argument("--repeatable",    action="store_true",       help="Find all Repeatable bugs")
    parser.add_argument("--blocker",       action="store_true",       help="Find all Blocker bugs")
    parser.add_argument("--repeatblocker", action="store_true",       help="Find Repeatable AND Blocker bugs")
    parser.add_argument("--builddate",     type=str,  default=None,  help='Find bugs by build date. Ex: --builddate "2/27/2024"')
    parser.add_argument("--iansearch",     action="store_true",       help="Run the proprietary iansearch logic call")

    args = parser.parse_args()

    print("\n[INFO] Connecting to MongoDB...")
    db = connect_to_mongo()

    collection_names = []

    for file_path in args.files:
        if not os.path.exists(file_path):
            print(f"[ERROR] File not found: {file_path}")
            continue

        collection_name = os.path.splitext(os.path.basename(file_path))[0]
        rows = load_excel_file(file_path)
        insert_into_collection(db, collection_name, rows)
        collection_names.append(collection_name)

    if len(collection_names) == 0:
        print("[ERROR] No files loaded. Exiting.")
        return

    print(f"\n[INFO] Collections ready: {collection_names}")

    output_text_file = "db_answers.txt"
    with open(output_text_file, "w", encoding="utf-8") as f:
        f.write("QA DATABASE QUERY RESULTS\n")
        f.write(f"Files: {args.files}\n")

    if args.testuser is not None:
        search_by_testuser(db, collection_names, args.testuser, output_text_file)

    if args.repeatable:
        search_repeatable(db, collection_names, output_text_file)

    if args.blocker:
        search_blocker(db, collection_names, output_text_file)

    if args.repeatblocker:
        search_repeatable_and_blocker(db, collection_names, output_text_file)

    if args.builddate is not None:
        search_by_builddate(db, collection_names, args.builddate, output_text_file)

    if args.iansearch:
        iansearch(db, collection_names, output_text_file)

    no_queries = (
        args.testuser is None
        and not args.repeatable
        and not args.blocker
        and not args.repeatblocker
        and args.builddate is None
        and not args.iansearch
    )

    if no_queries:
        print("\n[INFO] Files loaded. No query flag given. Available flags:")
        print('  --testuser "Kevin Chaja"')
        print("  --repeatable")
        print("  --blocker")
        print("  --repeatblocker")
        print('  --builddate "2/27/2024"')
        print("  --iansearch")

    print(f"\n[DONE] Text output saved to: {output_text_file}")


if __name__ == "__main__":
    main()


"""
OUTPUT:
1. Work by Master Chaja
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --testuser "Kevin Chaja"

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[QUERY] Test Owner = 'Kevin Chaja'
  'EG4-DBDump_Spring2026_c2-1': 28 results
  'EG4-DBDump_spring2026_c1': 34 results
[INFO] Results written to 'db_answers.txt'
[INFO] CSV saved: 'testuser_Kevin_Chaja.csv' (62 rows)
[INFO] Total: 62

[DONE] Text output saved to: db_answers.txt

2. Repeatable Bugs
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --repeatable

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[QUERY] All Repeatable bugs
  'EG4-DBDump_Spring2026_c2-1': 924 results
  'EG4-DBDump_spring2026_c1': 1283 results
[INFO] Results written to 'db_answers.txt'
[INFO] Total: 2207

[DONE] Text output saved to: db_answers.txt

3. Blocker Bugs
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --blocker

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[QUERY] All Blocker bugs
  'EG4-DBDump_Spring2026_c2-1': 413 results
  'EG4-DBDump_spring2026_c1': 524 results
[INFO] Results written to 'db_answers.txt'
[INFO] Total: 937

[DONE] Text output saved to: db_answers.txt

4. Repeatable and blocker
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --repeatblocker

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[QUERY] Repeatable AND Blocker bugs
  'EG4-DBDump_Spring2026_c2-1': 339 results
  'EG4-DBDump_spring2026_c1': 428 results
[INFO] Results written to 'db_answers.txt'
[INFO] Total: 767

[DONE] Text output saved to: db_answers.txt

5. Bugs on 2/27/2024
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --builddate "2/27/2024"

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[QUERY] Build date containing '2/27/2024'
  'EG4-DBDump_Spring2026_c2-1': 170 results
  'EG4-DBDump_spring2026_c1': 170 results
[INFO] Results written to 'db_answers.txt'
[INFO] Total: 340

[DONE] Text output saved to: db_answers.txt

6. My search
ian@Julians-MacBook-Pro-9 Project 2 % python3 main.py --files EG4-DBDump_Spring2026_c2-1.xlsx EG4-DBDump_spring2026_c1.xlsx --iansearch

[INFO] Connecting to MongoDB...

[INFO] Loading: EG4-DBDump_Spring2026_c2-1.xlsx
[INFO] Loaded 1136 rows
[INFO] Inserted 1136 rows into 'EG4-DBDump_Spring2026_c2-1'

[INFO] Loading: EG4-DBDump_spring2026_c1.xlsx
[INFO] Loaded 1586 rows
[INFO] Inserted 1586 rows into 'EG4-DBDump_spring2026_c1'

[INFO] Collections ready: ['EG4-DBDump_Spring2026_c2-1', 'EG4-DBDump_spring2026_c1']

[DATABASE LOGIC CALL] Running iansearch...
  'EG4-DBDump_Spring2026_c2-1': 85 non-critical bugs
  'EG4-DBDump_spring2026_c1': 127 non-critical bugs
  Total to analyze: 212
[INFO] CSV saved: 'iansearch_results.csv' (7 rows)
[INFO] Results written to 'db_answers.txt'
[INFO] iansearch done. 7 multi-build bug groups found.

[DONE] Text output saved to: db_answers.txt
"""