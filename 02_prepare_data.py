import pandas as pd
import os


# ============================================================
# REWARI PROPERTY INVESTMENT PROJECT
# STEP 2: PREPARE STANDARDIZED MASTER DATASET
# ============================================================

print("=" * 70)
print("REWARI PROPERTY INVESTMENT PROJECT")
print("STEP 2 - PREPARE MASTER DATASET")
print("=" * 70)


# ============================================================
# 1. FILE PATHS
# ============================================================

project_folder = os.path.dirname(
    os.path.abspath(__file__)
)

input_file = os.path.join(
    project_folder,
    "data",
    "Rewari_Collector_RatesCLEANED.xlsx"
)

output_folder = os.path.join(
    project_folder,
    "output"
)

os.makedirs(output_folder, exist_ok=True)


# ============================================================
# 2. CHECK INPUT FILE
# ============================================================

if not os.path.exists(input_file):

    print("\nERROR: Excel file not found.")
    print(input_file)
    exit()

print("\nExcel file found successfully.")


# ============================================================
# 3. YEAR → SHEET MAPPING
# ============================================================

year_sheets = {
    2023: "Collector Rates 2023",
    2024: "Collector Rates 2024",
    2025: "Collector Rates 2025",
    2026: "Collector Rates 2026"
}


# ============================================================
# 4. READ EACH YEAR
# ============================================================

all_data = []

print("\n" + "=" * 70)
print("READING YEARLY DATA")
print("=" * 70)


for year, sheet_name in year_sheets.items():

    print(f"\nReading {year}...")

    df = pd.read_excel(
        input_file,
        sheet_name=sheet_name
    )

    # Add year
    df["Year"] = year

    # Remove accidental spaces from column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    all_data.append(df)

    print(
        f"✓ {year}: {len(df)} records"
    )


# ============================================================
# 5. COMBINE ALL YEARS
# ============================================================

master_df = pd.concat(
    all_data,
    ignore_index=True
)


print("\n" + "=" * 70)
print("COMBINED DATASET")
print("=" * 70)

print(
    "Total records:",
    len(master_df)
)

print(
    "Total columns:",
    len(master_df.columns)
)


# ============================================================
# 6. CLEAN TEXT COLUMNS
# ============================================================

print("\nCleaning text fields...")

text_columns = master_df.select_dtypes(
    include=["object"]
).columns


for column in text_columns:

    master_df[column] = (
        master_df[column]
        .astype(str)
        .str.strip()
    )


# ============================================================
# 7. STANDARDIZE LOCATION
# ============================================================

if "Location" in master_df.columns:

    master_df["Location"] = (
        master_df["Location"]
        .str.strip()
        .str.title()
    )


# ============================================================
# 8. STANDARDIZE PROPERTY TYPE
# ============================================================

if "Property Type" in master_df.columns:

    master_df["Property Type"] = (
        master_df["Property Type"]
        .str.strip()
        .str.title()
    )


# ============================================================
# 9. CREATE GEOGRAPHICAL IDENTIFIER
# ============================================================

print("\nCreating geographical identifiers...")


if (
    "Area" in master_df.columns
    and
    "Segment" in master_df.columns
):

    master_df["Geo_ID"] = (
        master_df["Location"].fillna("")
        + "_"
        + master_df["Area"].fillna("")
        + "_"
        + master_df["Segment"].fillna("")
    )

elif "Area" in master_df.columns:

    master_df["Geo_ID"] = (
        master_df["Location"].fillna("")
        + "_"
        + master_df["Area"].fillna("")
    )

else:

    master_df["Geo_ID"] = (
        master_df["Location"].fillna("")
    )


# ============================================================
# 10. CREATE ANALYSIS ID
# ============================================================

if "Property Type" in master_df.columns:

    master_df["Analysis_ID"] = (
        master_df["Geo_ID"]
        + "_"
        + master_df["Property Type"]
    )

else:

    master_df["Analysis_ID"] = (
        master_df["Geo_ID"]
    )


# ============================================================
# 11. CHECK PROPERTY TYPES
# ============================================================

print("\n" + "=" * 70)
print("PROPERTY TYPES")
print("=" * 70)

if "Property Type" in master_df.columns:

    print(
        master_df["Property Type"]
        .value_counts(dropna=False)
    )


# ============================================================
# 12. CHECK LOCATIONS
# ============================================================

print("\n" + "=" * 70)
print("LOCATIONS")
print("=" * 70)

if "Location" in master_df.columns:

    print(
        master_df["Location"]
        .value_counts(dropna=False)
    )


# ============================================================
# 13. CHECK YEAR COUNTS
# ============================================================

print("\n" + "=" * 70)
print("YEAR COUNTS")
print("=" * 70)

print(
    master_df["Year"]
    .value_counts()
    .sort_index()
)


# ============================================================
# 14. CHECK LOCATION × PROPERTY TYPE
# ============================================================

print("\n" + "=" * 70)
print("LOCATION × PROPERTY TYPE")
print("=" * 70)

if (
    "Location" in master_df.columns
    and
    "Property Type" in master_df.columns
):

    location_property = pd.crosstab(
        master_df["Location"],
        master_df["Property Type"]
    )

    print(location_property)


# ============================================================
# 15. SAVE MASTER DATASET
# ============================================================

output_file = os.path.join(
    output_folder,
    "Rewari_Master_2023_2026.csv"
)

master_df.to_csv(
    output_file,
    index=False
)


# ============================================================
# 16. FINAL SUMMARY
# ============================================================

print("\n" + "=" * 70)
print("MASTER DATASET CREATED")
print("=" * 70)

print("\nFile saved at:")
print(output_file)

print("\nTotal records:", len(master_df))

print("\nColumns:")
for column in master_df.columns:
    print("-", column)

print("\n" + "=" * 70)
print("STEP 2 COMPLETED SUCCESSFULLY")
print("=" * 70)

print("""
IMPORTANT:
The original Excel file was NOT modified.

Next step:
STEP 3 - Identify and standardize geographical
         names across 2023-2026.
""")