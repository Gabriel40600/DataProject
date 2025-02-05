import pandas as pd


# Load the Excel file
def analyze_linkedin_data(file_path, output_file):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Select only numeric columns
    numeric_columns = ["reactions", "comments", "shares", "Sentiment_Positive", "Sentiment_Neutral",
                       "Sentiment_Negative", "Sentiment_Compound", "Post Length", "Emoji Count"]

    df_numeric = df[numeric_columns]

    # Generate descriptive statistics
    descriptive_stats = df_numeric.describe()

    # Save to Excel
    descriptive_stats.to_excel(output_file, index=False)
    print(f"Processed data saved to {output_file}")


# Provide the file path
file_path = "processed_linkedin_data.xlsx"
output_file = "deskriptive_statistik.xlsx"
analyze_linkedin_data(file_path, output_file)
