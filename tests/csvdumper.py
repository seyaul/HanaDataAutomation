import pandas as pd

# Define your input/output paths
input_excel = input("enter yo path: ").strip().strip('"\'')
output_csv = "brand_map.csv"

# Read the Excel file
df = pd.read_excel(input_excel, dtype={"Item ID": str})

# Keep only the necessary columns
columns_to_keep = ["Item ID", "Brand", "CATEGORY"]
df_subset = df[columns_to_keep]

# Drop duplicates (optional but recommended for mapping)
df_subset = df_subset.drop_duplicates()

# Save to CSV
df_subset.to_csv(output_csv, index=False)

print(f"✅ Saved {len(df_subset)} unique rows to {output_csv}")
