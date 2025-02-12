import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
from scikit_posthocs import posthoc_dunn
import seaborn as sns

# Define file path
file_path = "2_processed_linkedin_data.xlsx"

# Load the dataset
df = pd.read_excel(file_path)

# Define columns of interest
link_col = "Contains Link"  # Ensure this column exists in the dataset
interaction_cols = ["reactions", "comments", "shares"]

# Create a new column for Link Presence
df["Link Presence"] = df[link_col].apply(lambda x: "With Link" if x == 1 else "No Link")

# Group by Link Presence and calculate the average interactions
interaction_summary = df.groupby("Link Presence")[interaction_cols].mean().reset_index()

# Calculate Percentage Change
percentage_change = interaction_summary.set_index("Link Presence").pct_change().iloc[1] * 100
percentage_change = percentage_change.to_frame(name="Percentage Change").reset_index()

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
kruskal_results = []
dunn_results_combined = pd.DataFrame()

# Preserve category names instead of converting to numbers
df["Link Presence"] = df["Link Presence"].astype("category")

for col in interaction_cols:
    stat, p_value = stats.shapiro(df[col].dropna())
    is_normal = "Yes" if p_value > 0.05 else "No"
    normality_results.append((col, stat, p_value, is_normal))
    print(f"{col}: W={stat:.4f}, p={p_value:.4f}, Normal: {is_normal}")

# Perform Kruskal-Wallis Test
print("\nKruskal-Wallis Test:")
for col in interaction_cols:
    group_data = [df[df["Link Presence"] == category][col].dropna().values for category in df["Link Presence"].unique()]
    kruskal_stat, kruskal_p = stats.kruskal(*group_data)
    significance = "Significant" if kruskal_p < 0.05 else "Not Significant"
    kruskal_results.append((col, kruskal_stat, kruskal_p, significance))
    print(f"{col}: H={kruskal_stat:.4f}, p={kruskal_p:.4f}, {significance}")

# Perform Dunn's Test if Kruskal-Wallis is significant
for col in interaction_cols:
    if any(p < 0.05 for _, _, p, _ in kruskal_results):
        print(f"\nDunn's Post Hoc Test for {col}:")
        dunn_results = posthoc_dunn(df, val_col=col, group_col="Link Presence", p_adjust="bonferroni")
        dunn_results.insert(0, "Metric", col)  # Label column for merging
        dunn_results_combined = pd.concat([dunn_results_combined, dunn_results], axis=0)
        print(dunn_results)

# Save all results in a single Excel file
output_file = "3i_linkedin_link_analysis.xlsx"
with pd.ExcelWriter(output_file) as writer:
    interaction_summary.to_excel(writer, sheet_name="Interaction Summary", index=False)
    percentage_change.to_excel(writer, sheet_name="Percentage Change", index=False)
    df_normality = pd.DataFrame(normality_results, columns=["Metric", "W-Statistic", "P-Value", "Normal"])
    df_normality.to_excel(writer, sheet_name="Normality Test", index=False)
    df_kruskal = pd.DataFrame(kruskal_results, columns=["Metric", "H-Statistic", "P-Value", "Significance"])
    df_kruskal.to_excel(writer, sheet_name="Kruskal-Wallis Test", index=False)
    if not dunn_results_combined.empty:
        dunn_results_combined.to_excel(writer, sheet_name="Dunn's Test", index=False)

print("Analysis completed. Results saved.")
