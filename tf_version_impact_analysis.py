import os
import re
import json
import tempfile
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import subprocess
import platform


# ------------------------------------------------------------------------------
# Extract rc.type values from all files in /policies
# ------------------------------------------------------------------------------
def extract_rc_types_from_policies(policies_dir="policies"):
    results = []

    for root, _, files in os.walk(policies_dir):
        for file in files:
            path = os.path.join(root, file)

            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            except Exception:
                print(f"Warning: couldn't read file {path}")
                continue

            matches = re.findall(r'rc\.type\s+is\s+"([^"]+)"', content)

            for m in matches:
                results.append({
                    "filename": file,
                    "resource_type": m.strip()
                })

    return results


# ------------------------------------------------------------------------------
# Load local AzureRM 4.53 resource list (JSON array of strings)
# ------------------------------------------------------------------------------
def load_local_resource_types(path):
    with open(path, "r") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("Registry file must be a JSON array of strings")

    return set(data)


# ------------------------------------------------------------------------------
# Compare rc.type to AzureRM resource list
# ------------------------------------------------------------------------------
def compare_types(resource_entries, registry_set):
    results = []

    for entry in resource_entries:
        rtype = entry["resource_type"]
        match = rtype in registry_set

        results.append({
            "filename": entry["filename"],
            "resource_type": rtype,
            "match": "Match" if match else "Mismatch"
        })

    return results


# ------------------------------------------------------------------------------
# Write color-highlighted Excel to temp directory + autolaunch
# ------------------------------------------------------------------------------
def write_excel(results):
    df = pd.DataFrame(results).sort_values(["filename", "resource_type"])

    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, "terraform_policy_rc_type_comparison.xlsx")

    df.to_excel(excel_path, index=False)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    match_col = None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value == "match":
            match_col = idx
            break

    for row in ws.iter_rows(min_row=2):
        fill = green if row[match_col - 1].value == "Match" else red
        for c in row:
            c.fill = fill

    wb.save(excel_path)

    # Auto-launch Excel
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(excel_path)
        elif system == "Darwin":
            subprocess.run(["open", excel_path])
        else:
            subprocess.run(["xdg-open", excel_path])
    except Exception:
        print("Excel created but could not auto-open.")

    print(f"\nExcel file created at: {excel_path}")
    return excel_path


# ------------------------------------------------------------------------------
# JSON export
# ------------------------------------------------------------------------------
def export_json(results):
    return json.dumps(results, indent=2)


# ------------------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------------------
def main():
    policies_dir = "policies"
    registry_file = "azurerm-4.53.0-resource-types.json"

    resource_entries = extract_rc_types_from_policies(policies_dir)
    registry_set = load_local_resource_types(registry_file)

    results = compare_types(resource_entries, registry_set)

    write_excel(results)

    print("\nJSON Export:\n")
    print(export_json(results))


if __name__ == "__main__":
    main()
