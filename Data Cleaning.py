import pandas as pd
import os
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def clean_dataset(file_path, output_file_path):
    # Load the dataset from the correct sheet
    sheet_name = "Technology & Innovation"  # Updated sheet name
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Define numeric columns manually
    numeric_columns = ['reactions', 'comments', 'shares']

    # Identify rows with missing numeric values before filling
    missing_values = df[df[numeric_columns].isna().any(axis=1)]

    # Fill missing numeric values with zero
    df[numeric_columns] = df[numeric_columns].fillna(0)

    # Convert numeric columns to proper format
    for col in numeric_columns:
        df[col] = df[col].astype(str).str.replace(r'[,\.]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    # Identify and remove duplicates based on 'Post URL'
    duplicates = df[df.duplicated(subset=['Post URL'], keep=False)]
    df.drop_duplicates(subset=['Post URL'], inplace=True)

    # Convert 'Nur Repost' column to categorical (1 for 'x', 0 for empty)
    df['Nur Repost'] = df['Nur Repost'].apply(lambda x: 1 if str(x).strip().lower() == 'x' else 0).astype('int8')

    # Trim whitespaces in text columns and fill missing values
    text_columns = ['TL', 'Post content']
    df[text_columns] = df[text_columns].apply(lambda col: col.str.strip().fillna(''))

    # Optimize data types to reduce memory usage
    df = df.astype({col: 'int32' for col in numeric_columns})

    # Generate a summary of changes
    changes_summary = {
        "Duplicates Removed": len(duplicates),
        "Missing Numeric Values Set to 0": len(missing_values)
    }

    logging.info(f"Changes Summary: {changes_summary}")

    # Save cleaned dataset, removed duplicates, and missing numeric values in separate sheets
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        duplicates.to_excel(writer, sheet_name="Removed Duplicates", index=False)
        missing_values.to_excel(writer, sheet_name="Missing Numeric Values", index=False)

    logging.info(f"Cleaned dataset, removed duplicates, and missing numeric values saved to {output_file_path}")


# Define file paths
base_path = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.join(base_path, "data_set.xlsx")
output_file_path = os.path.join(base_path, "data_set_cleaned.xlsx")

clean_dataset(input_file_path, output_file_path)