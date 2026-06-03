import pandas as pd
import os

# Load your CSV file
df = pd.read_excel("VAT file 2026_latest.xlsx")

# Output folder for customer files
output_dir = "customer_files_EU"
os.makedirs(output_dir, exist_ok=True)

# Group by Customer Id and create one file per customer
for customer_id, group in df.groupby("Customer ID"):
    file_name = f"{output_dir}/{customer_id}.csv"
    group.to_csv(file_name, index=False)
    print(f"Created {file_name}")
