import pandas as pd
import os
import logging

# Setup logging
log_file_path = "1_data_cleaning_log.txt"
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def clean_dataset(file_path, output_file_path):
    sheet_name = "Technology & Innovation"
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    numeric_columns = ['reactions', 'comments', 'shares']
    text_columns = ['TL', 'Post content']

    missing_values = df[df[numeric_columns].isna().any(axis=1)]
    df[numeric_columns] = df[numeric_columns].fillna(0)

    for col in numeric_columns:
        df[col] = df[col].astype(str).str.replace(r'[,\.]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    df[text_columns] = df[text_columns].astype(str).apply(lambda col: col.str.strip().fillna(''))

    duplicates = df[df.duplicated(subset=['Post URL'], keep=False)]
    df.drop_duplicates(subset=['Post URL'], inplace=True)

    df['Nur Repost'] = df['Nur Repost'].apply(lambda x: 1 if str(x).strip().lower() == 'x' else 0).astype('int8')
    df = df.astype({col: 'int32' for col in numeric_columns})

    changes_summary = {
        "Duplicates Removed": len(duplicates),
        "Missing Numeric Values Set to 0": len(missing_values)
    }

    with open(log_file_path, "a") as log_file:
        log_file.write("\nCleaning Summary:\n")
        log_file.write(f"Duplicates Removed: {len(duplicates)}\n")
        log_file.write(f"Duplicate Entries:\n{duplicates.to_string(index=False)}\n")
        log_file.write(f"Missing Numeric Values Set to 0: {len(missing_values)}\n")
        for col in numeric_columns:
            missing_count = missing_values[col].isna().sum()
            if missing_count > 0:
                log_file.write(f"Column '{col}' had {missing_count} missing values replaced with 0.\n")
        log_file.write("-------------------------------------\n")

    logging.info(f"Cleaned dataset and logs saved to {output_file_path} and {log_file_path}")

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        duplicates.to_excel(writer, sheet_name="Removed Duplicates", index=False)
        missing_values.to_excel(writer, sheet_name="Missing Numeric Values", index=False)


base_path = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.join(base_path, "0_data_set.xlsx")
output_file_path = os.path.join(base_path, "1_data_set_cleaned.xlsx")
clean_dataset(input_file_path, output_file_path)
