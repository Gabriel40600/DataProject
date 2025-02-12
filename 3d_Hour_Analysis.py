import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import scipy.stats as stats
from scikit_posthocs import posthoc_dunn
import seaborn as sns

# Load the LinkedIn data
file_path = "2_processed_linkedin_data.xlsx"
df = pd.read_excel(file_path)

# Remove invalid timestamp entries
df_cleaned = df[df['Post Timestamp (ISO)'] != "Invalid LinkedIn ID"].copy()

# Convert timestamp to datetime and extract the hour of the day
df_cleaned['Post Timestamp (ISO)'] = pd.to_datetime(df_cleaned['Post Timestamp (ISO)'], errors='coerce')
df_cleaned = df_cleaned.dropna(subset=['Post Timestamp (ISO)'])  # Drop rows where conversion failed
df_cleaned['Hour of Day'] = df_cleaned['Post Timestamp (ISO)'].dt.hour

# Count posts per hour
posts_per_hour = df_cleaned.groupby("Hour of Day")["Post Timestamp (ISO)"].count()

# Normalize engagement metrics by the number of posts per hour
metrics = ['reactions', 'comments', 'shares']
df_cleaned[metrics] = df_cleaned[metrics].div(df_cleaned["Hour of Day"].map(posts_per_hour), axis=0)

# Save normalized interactions to an Excel file
output_file = "3d_normalized_interactions_by_hour.xlsx"
df_cleaned.to_excel(output_file, index=False)
print(f"Normalized data saved to {output_file}")

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
kruskal_results = []
dunn_results_list = []

for col in metrics:
    # Prepare data for statistical tests
    groups = [df_cleaned[df_cleaned['Hour of Day'] == hour][col].dropna() for hour in df_cleaned['Hour of Day'].unique()]

    # Check normality using Shapiro-Wilk test
    normality_p_values = [stats.shapiro(group)[1] for group in groups if len(group) > 3]  # Shapiro requires >3 samples
    is_normal = all(p > 0.05 for p in normality_p_values) if normality_p_values else False
    normality_results.append([col, "Normal" if is_normal else "Not Normal"])

    print(f"Normality check for {col}: {'Normal' if is_normal else 'Not Normal'}")

    if not is_normal:
        # Run Kruskal-Wallis test if data is not normal
        h_stat, kw_p_value = stats.kruskal(*groups)
        kruskal_results.append([col, h_stat, kw_p_value, "Significant" if kw_p_value < 0.05 else "Not Significant"])
        print(f"Kruskal-Wallis Test for {col}: H-statistic = {h_stat:.3f}, p-value = {kw_p_value:.5f}")

        # If Kruskal-Wallis is significant, run Dunn's Test for pairwise comparisons
        if kw_p_value < 0.05:
            dunn_results = posthoc_dunn(df_cleaned, val_col=col, group_col="Hour of Day", p_adjust="bonferroni")
            dunn_results.reset_index(inplace=True)
            dunn_results.insert(0, "Metric", col)  # Add metric column for clarity
            dunn_results_list.append(dunn_results)

# Save results to an Excel file
output_file_tests = "3d_interactions_by_hour_tests.xlsx"

with pd.ExcelWriter(output_file_tests, engine='xlsxwriter') as writer:
    pd.DataFrame(normality_results, columns=["Metric", "Normality"]).to_excel(writer, sheet_name="Normality Test",
                                                                              index=False)
    pd.DataFrame(kruskal_results, columns=["Metric", "H-Statistic", "p-value", "Significance"]).to_excel(writer,
                                                                                                         sheet_name="Kruskal-Wallis Test",
                                                                                                         index=False)

    if dunn_results_list:
        dunn_final_results = pd.concat(dunn_results_list, ignore_index=True)
        dunn_final_results.to_excel(writer, sheet_name="Dunn's Test", index=False)

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file_tests}")
