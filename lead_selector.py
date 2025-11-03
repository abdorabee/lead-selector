import pandas as pd
import glob
import ast

# ---------------------------
# CONFIGURATION
# ---------------------------

# Folder where your Excel files are stored (same directory by default)
INPUT_PATH = "./"        # Change if files are in another folder
OUTPUT_FILE = "Top_25_Leads.xlsx"

# ---------------------------
# LOAD AND MERGE FILES
# ---------------------------

all_files = glob.glob(INPUT_PATH + "*.xlsx")
dataframes = []

for file in all_files:
    try:
        df = pd.read_excel(file)
        df["Source_File"] = file.split("/")[-1]
        dataframes.append(df)
        print(f"Loaded: {file} ({len(df)} rows)")
    except Exception as e:
        print(f"❌ Error reading {file}: {e}")

if not dataframes:
    raise ValueError("No Excel files found in folder!")

merged = pd.concat(dataframes, ignore_index=True)
print(f"\n✅ Merged {len(dataframes)} files with {len(merged)} total rows.\n")

# ---------------------------
# CLEAN AND NORMALIZE
# ---------------------------

# Keep only relevant columns (some may not exist in all files)
columns_to_keep = [
    "LinkedinUsername",
    "LinkedinUrl",
    "Email",
    "PhoneNumbers",
    "Countries",
    "JobCompanySize",
    "LastJobTitle",
    "Source_File"
]

merged = merged[[c for c in columns_to_keep if c in merged.columns]].copy()

# Normalize 'Countries' to count them properly
def count_countries(value):
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        try:
            value = ast.literal_eval(value)
        except Exception:
            value = [value]
    return len(value)

merged["CountryCount"] = merged["Countries"].apply(count_countries)  # type: ignore[attr-defined]

# Normalize phone and email existence
merged["HasEmail"] = merged["Email"].notna().astype(int)  # type: ignore[attr-defined]
merged["HasPhone"] = merged["PhoneNumbers"].apply(  # type: ignore[attr-defined]
    lambda x: int(pd.notna(x) and str(x).lower() not in ["nan", "[]", "[null]"])
)

# ---------------------------
# SCORING
# ---------------------------

def company_size_score(size):
    if pd.isna(size):
        return 0
    size = str(size).lower()
    if "large" in size or "1001" in size or "5001" in size:
        return 3
    elif "medium" in size or "201" in size or "1000" in size:
        return 2
    elif "small" in size or "1" in size or "50" in size:
        return 1
    return 0

merged["CompanyScore"] = merged["JobCompanySize"].apply(company_size_score)  # type: ignore[attr-defined]
merged["MultiCountry"] = merged["CountryCount"].apply(lambda x: 1 if x > 1 else 0)  # type: ignore[attr-defined]

# Total weighted score
merged["TotalScore"] = (
    merged["CompanyScore"] * 1.0
    + merged["MultiCountry"] * 2.0
    + merged["HasEmail"] * 1.0
    + merged["HasPhone"] * 1.0
)

# ---------------------------
# RANK AND EXPORT
# ---------------------------

# Sort and get top 25
top_25 = merged.sort_values(by="TotalScore", ascending=False).head(25)  # type: ignore[call-overload]

# Export
top_25.to_excel(OUTPUT_FILE, index=False)
print(f"✅ Exported Top 25 leads → {OUTPUT_FILE}")

# Optional: display top 5 preview
print("\n🔝 Top 5 Leads Preview:")
print(top_25[["LinkedinUsername", "JobCompanySize", "Countries", "Email", "PhoneNumbers", "TotalScore"]].head())
