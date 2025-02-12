import pandas as pd

# Load dataset
file_path = "2_processed_linkedin_data.xlsx"  # Update with the correct local path
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Compute mean engagement per author
author_engagement = df.groupby("TL")[["reactions", "comments", "shares"]].mean()

# Save results to an Excel file
output_xlsm = "3a_author_mean_engagement.xlsx"  # Keep consistent variable naming
with pd.ExcelWriter(output_xlsm, engine='openpyxl') as writer:
    author_engagement.to_excel(writer, sheet_name="Mean Engagement")

# Correct the variable name in the print statement
print(f"Excel file saved as {output_xlsm}")
