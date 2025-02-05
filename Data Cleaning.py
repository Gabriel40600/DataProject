import pandas as pd
import os


def clean_dataset(file_path, sheet_name):
    # Load the dataset
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Fill missing numeric values with zero
    numeric_columns = ['reactions', 'comments', 'shares']
    df[numeric_columns] = df[numeric_columns].fillna(0)

    # Convert numeric columns to proper format
    for col in numeric_columns:
        df[col] = df[col].astype(str).str.replace(',', '').str.replace('.', '')
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    # Identify duplicates based on 'Post URL'
    duplicates = df[df.duplicated(subset=['Post URL'], keep=False)]

    # Remove duplicates, keeping only the first occurrence
    df = df.drop_duplicates(subset=['Post URL'])

    # Fill missing values in 'Nur Repost' with 'No' (assuming it's a categorical indicator)
    df['Nur Repost'] = df['Nur Repost'].fillna('No')

    # Trim whitespaces in text columns
    text_columns = ['TL', 'Post content', 'Post URL']
    for col in text_columns:
        df[col] = df[col].astype(str).str.strip()

    # Mark empty spaces as null
    df.replace(r'^\s*$', pd.NA, regex=True, inplace=True)

    # Remove rows with invalid or missing URLs
    df = df[df['Post URL'].str.startswith('http', na=False)]

    # Generate a summary of changes
    changes_summary = {
        "Duplicates Removed": len(duplicates),
        "Missing Numeric Values Set to 0": df[numeric_columns].isna().sum().to_dict()
    }

    print("Changes Summary:", changes_summary)

    return df


# Define file paths
base_path = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.join(base_path, "data_set.xlsx")
output_file_path = os.path.join(base_path, "data_set_cleaned.xlsx")

# Usage example
cleaned_df = clean_dataset(input_file_path, "Technology & Innovation")
# Save cleaned data
cleaned_df.to_excel(output_file_path, index=False)

print(f"Cleaned dataset saved to {output_file_path}")
