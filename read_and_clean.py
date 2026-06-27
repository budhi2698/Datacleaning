import pandas as pd

# Read the Excel file
df = pd.read_excel('tests.xlsx', header=None)

# Extract metadata
code = df.iloc[1, 1]
date = df.iloc[2, 1]

# Set column names from row 3 (index 3)
df.columns = df.iloc[3]

# Drop the header rows (first 4 rows)
df = df.iloc[4:].reset_index(drop=True)

print("=== Metadata ===")
print(f"Code: {code}")
print(f"Date: {date}")
print("\n=== Original Data ===")
print(df.to_string())
print("\n=== Data Info ===")
print(df.info())
print("\n=== Missing Values ===")
print(df.isnull().sum())
print("\n=== Duplicate Rows ===")
print(df.duplicated().sum())

# Clean the data
# 1. Remove duplicate rows
df_cleaned = df.drop_duplicates()

# 2. Drop completely empty rows
df_cleaned = df_cleaned.dropna(how='all')

# 3. Reset index
df_cleaned = df_cleaned.reset_index(drop=True)

# 4. Forward fill the Serial column where possible
df_cleaned['Serial'] = df_cleaned['Serial'].ffill()

print("\n=== Cleaned Data ===")
print(df_cleaned.to_string())
print("\n=== Cleaned Data Info ===")
print(df_cleaned.info())

# Save cleaned data
df_cleaned.to_excel('cleandata.xlsx', index=False)
print("\nCleaned data saved as cleandata.xlsx")

