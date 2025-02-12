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
emoji_col = "Contains Emoji"  # Ensure this column exists in the dataset
interaction_cols = ["reactions", "comments", "shares"]

# Create a new column for Emoji Presence
df["Emoji Presence"] = df[emoji_col].apply(lambda x: "with Emoji" if x == 1 else "No Emoji")

# Group by Emoji Presence and calculate the average interactions
interaction_summary = df.groupby("Emoji Presence")[interaction_cols].mean().reset_index()

# Save the results to an Excel file
output_file = "3g_linkedin_emoji_analysis.xlsx"
interaction_summary.to_excel(output_file, index=False)

# Plot the interaction comparison
fig, ax = plt.subplots(figsize=(8, 6))
interaction_summary.set_index("Emoji Presence").plot(kind="bar", ax=ax)
plt.title("Average Interactions by Emoji Presence")
plt.ylabel("Average Count")
plt.xlabel("Emoji Presence")
plt.xticks(rotation=0)
plt.legend(title="Interaction Type")
plt.grid(axis="y", linestyle="--", alpha=0.7)

# Save the plot
plot_file = "3g_linkedin_post_interaction_emoji_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
kruskal_results = []
dunn_results_list = []

# Preserve category names instead of converting to numbers
df["Emoji Presence"] = df["Emoji Presence"].astype("category")

for col in interaction_cols:
    # Prepare data for statistical tests
    groups = [df[df["Emoji Presence"] == emoji_type][col].dropna() for emoji_type in df["Emoji Presence"].unique()]

    # Check normality using Shapiro-Wilk test
    normality_p_values = [stats.shapiro(group)[1] for group in groups if len(group) > 3]  # Shapiro requires >3 samples

    if len(normality_p_values) < len(groups):
        print(f"Skipping Shapiro-Wilk test for {col}: Not enough data in some groups.")

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
            dunn_results = posthoc_dunn(df, val_col=col, group_col="Emoji Presence", p_adjust="bonferroni")
            dunn_results.reset_index(inplace=True)
            dunn_results.insert(0, "Metric", col)  # Add metric column for clarity
            dunn_results_list.append(dunn_results)

# ✅ Compute Percentage Change Between "With Emoji" and "No Emoji"
emoji_group = df.groupby("Emoji Presence")[interaction_cols].mean()

# Calculate percentage changes
percentage_changes = []
for col in interaction_cols:
    no_emoji_value = emoji_group.loc["No Emoji", col]
    with_emoji_value = emoji_group.loc["with Emoji", col]
    percentage_change = ((with_emoji_value - no_emoji_value) / no_emoji_value) * 100
    percentage_changes.append([col, percentage_change])

# Convert to DataFrame
percentage_change_df = pd.DataFrame(percentage_changes, columns=["Metric", "Percentage Change"])

# Save results to an Excel file
output_file_tests = "3g_Emoji_tests.xlsx"

with pd.ExcelWriter(output_file_tests, engine='xlsxwriter') as writer:
    pd.DataFrame(normality_results, columns=["Metric", "Normality"]).to_excel(writer, sheet_name="Normality Test", index=False)
    pd.DataFrame(kruskal_results, columns=["Metric", "H-Statistic", "p-value", "Significance"]).to_excel(writer, sheet_name="Kruskal-Wallis Test", index=False)

    # Ensure we save an empty table if Dunn's Test has no results
    if not dunn_results_list:
        dunn_results_list.append(pd.DataFrame(columns=["Metric", "Post Type", "Compared Post Type"]))

    dunn_final_results = pd.concat(dunn_results_list, ignore_index=True)
    dunn_final_results.to_excel(writer, sheet_name="Dunn's Test", index=False)

    # ✅ Save percentage change results
    percentage_change_df.to_excel(writer, sheet_name="Percentage Change", index=False)

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file_tests}")
