import pandas as pd
import os


# ============================================================
# REWARI PROPERTY INVESTMENT PROJECT
# PHASE 1: FINAL DATA AUDIT
# ============================================================

print("=" * 70)
print("REWARI PROPERTY INVESTMENT PROJECT")
print("PHASE 1 - FINAL DATA AUDIT")
print("=" * 70)


# ============================================================
# 1. FIND EXCEL FILE
# ============================================================

script_folder = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(
    script_folder,
    "data",
    "Rewari_Collector_RatesCLEANED.xlsx"
)

file_path = os.path.abspath(file_path)

print("\nExcel file location:")
print(file_path)


# ============================================================
# 2. CHECK FILE
# ============================================================

if not os.path.exists(file_path):
    print("\nERROR: Excel file not found!")
    exit()

print("\nExcel file found successfully!")


# ============================================================
# 3. OPEN WORKBOOK
# ============================================================

excel_file = pd.ExcelFile(file_path)

print("\n" + "=" * 70)
print("WORKBOOK SHEETS")
print("=" * 70)

for sheet in excel_file.sheet_names:
    print("-", sheet)


# ============================================================
# 4. MAP SHEETS TO YEARS
# ============================================================

year_sheets = {
    "2023": "Collector Rates 2023",
    "2024": "Collector Rates 2024",
    "2025": "Collector Rates 2025",
    "2026": "Collector Rates 2026"
}


# ============================================================
# 5. CHECK ALL REQUIRED SHEETS
# ============================================================

print("\n" + "=" * 70)
print("YEAR SHEET CHECK")
print("=" * 70)

for year, sheet_name in year_sheets.items():

    if sheet_name in excel_file.sheet_names:
        print(f"✓ {year} → {sheet_name}")
    else:
        print(f"✗ {year} → {sheet_name} NOT FOUND")


# ============================================================
# 6. LOAD DATA
# ============================================================

data = {}

print("\n" + "=" * 70)
print("LOADING DATA")
print("=" * 70)

for year, sheet_name in year_sheets.items():

    if sheet_name not in excel_file.sheet_names:
        continue

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name
    )

    data[year] = df

    print(f"\n{year}")
    print("-" * 40)
    print("Rows    :", df.shape[0])
    print("Columns :", df.shape[1])


# ============================================================
# 7. COLUMN NAMES
# ============================================================

print("\n" + "=" * 70)
print("COLUMN NAMES")
print("=" * 70)

for year, df in data.items():

    print(f"\n{year}:")

    for number, column in enumerate(df.columns, start=1):
        print(f"{number}. {column}")


# ============================================================
# 8. DATA TYPES
# ============================================================

print("\n" + "=" * 70)
print("DATA TYPES")
print("=" * 70)

for year, df in data.items():

    print(f"\n{year}:")

    for column in df.columns:
        print(
            f"{column} --> {df[column].dtype}"
        )


# ============================================================
# 9. MISSING VALUES
# ============================================================

print("\n" + "=" * 70)
print("MISSING VALUE CHECK")
print("=" * 70)

for year, df in data.items():

    print(f"\n{year}:")

    missing = df.isnull().sum()

    missing = missing[missing > 0]

    if len(missing) == 0:
        print("✓ No missing values")

    else:

        for column, count in missing.items():
            print(
                f"{column}: {count} missing"
            )


# ============================================================
# 10. DUPLICATE ROW CHECK
# ============================================================

print("\n" + "=" * 70)
print("DUPLICATE ROW CHECK")
print("=" * 70)

for year, df in data.items():

    duplicate_count = df.duplicated().sum()

    print(
        f"{year}: {duplicate_count} duplicate rows"
    )


# ============================================================
# 11. BASIC VALUE CHECK
# ============================================================

print("\n" + "=" * 70)
print("UNIQUE VALUE CHECK")
print("=" * 70)

for year, df in data.items():

    print("\n" + "-" * 60)
    print("YEAR:", year)
    print("-" * 60)

    for column in df.columns:

        # Only display columns with manageable
        # numbers of unique values

        unique_count = df[column].nunique(
            dropna=True
        )

        print(
            f"{column}: {unique_count} unique values"
        )


# ============================================================
# 12. LOCATION CHECK
# ============================================================

print("\n" + "=" * 70)
print("LOCATION CHECK")
print("=" * 70)

for year, df in data.items():

    if "Location" in df.columns:

        print(f"\n{year}:")

        print(
            df["Location"]
            .value_counts(dropna=False)
        )

    else:

        print(
            f"\n{year}: Location column not found"
        )


# ============================================================
# 13. PROPERTY TYPE CHECK
# ============================================================

print("\n" + "=" * 70)
print("PROPERTY TYPE CHECK")
print("=" * 70)

for year, df in data.items():

    if "Property Type" in df.columns:

        print(f"\n{year}:")

        print(
            df["Property Type"]
            .value_counts(dropna=False)
        )

    else:

        print(
            f"\n{year}: Property Type column not found"
        )


# ============================================================
# 14. LOCATION × PROPERTY TYPE
# ============================================================

print("\n" + "=" * 70)
print("LOCATION × PROPERTY TYPE")
print("=" * 70)

for year, df in data.items():

    if (
        "Location" in df.columns
        and
        "Property Type" in df.columns
    ):

        print(f"\n{year}:")

        table = pd.crosstab(
            df["Location"],
            df["Property Type"]
        )

        print(table)


# ============================================================
# 15. AREA / SEGMENT CHECK
# ============================================================

print("\n" + "=" * 70)
print("AREA / SEGMENT CHECK")
print("=" * 70)

for year, df in data.items():

    print(f"\n{year}:")

    if "Area" in df.columns:

        print(
            "Unique Areas:",
            df["Area"].nunique(dropna=True)
        )

    if "Segment" in df.columns:

        print(
            "Unique Segments:",
            df["Segment"].nunique(dropna=True)
        )


# ============================================================
# 16. FINAL SUMMARY
# ============================================================

print("\n" + "=" * 70)
print("FINAL AUDIT SUMMARY")
print("=" * 70)

for year, df in data.items():

    print(
        f"{year}: {len(df)} records"
    )

print("\nOriginal Excel file was NOT modified.")

print("\n" + "=" * 70)
print("PHASE 1 AUDIT COMPLETED")
print("=" * 70)