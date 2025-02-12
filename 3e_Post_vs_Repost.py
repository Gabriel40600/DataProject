import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
from scipy.stats import ttest_ind, mannwhitneyu
from scikit_posthocs import posthoc_dunn
import seaborn as sns

# Define file path dynamically
file_path = "2_processed_linkedin_data.xlsx"

# Load the dataset
df = pd.read_excel(file_path)

# Define columns of interest
repost_col = "Nur Repost"
interaction_cols = ["reactions", "comments", "shares"]

# Create a new column for post type
df["Post Type"] = df[repost_col].apply(lambda x: "Repost" if x == 1 else "Original")

# Group by post type and calculate the average interactions
interaction_summary = df.groupby("Post Type")[interaction_cols].mean().reset_index()

# Save the results to an Excel file
output_file = "3e_linkedin_post_analysis.xlsx"
interaction_summary.to_excel(output_file, index=False)

# Plot the interaction comparison
fig, ax = plt.subplots(figsize=(8, 6))
interaction_summary.set_index("Post Type").plot(kind="bar", ax=ax)
plt.title("Average Interactions by Post Type")
plt.ylabel("Average Count")
plt.xlabel("Post Type")
plt.xticks(rotation=0)
plt.legend(title="Interaction Type")
plt.grid(axis="y", linestyle="--", alpha=0.7)

# Save the plot
plot_file = "3e_linkedin_post_interaction_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
test_results = []

for col in interaction_cols:
    # Prepare data for statistical tests
    groups = [df[df["Post Type"] == post_type][col].dropna() for post_type in df["Post Type"].unique()]

    # Check normality using Shapiro-Wilk test
    normality_p_values = [stats.shapiro(group)[1] for group in groups if len(group) > 3]  # Shapiro requires >3 samples
    is_normal = all(p > 0.05 for p in normality_p_values) if normality_p_values else False
    normality_results.append([col, "Normal" if is_normal else "Not Normal"])

    print(f"Normality check for {col}: {'Normal' if is_normal else 'Not Normal'}")

    # Perform t-test or Mann-Whitney U test for binary comparison
    if is_normal:
        t_stat, t_p_value = ttest_ind(*groups, equal_var=False)
        test_results.append([col, "t-test", t_stat, t_p_value, "Significant" if t_p_value < 0.05 else "Not Significant"])
    else:
        u_stat, u_p_value = mannwhitneyu(*groups, alternative="two-sided")
        test_results.append([col, "Mann-Whitney U", u_stat, u_p_value, "Significant" if u_p_value < 0.05 else "Not Significant"])

# ✅ Compute Percentage Change Between "Original" and "Repost"
post_type_group = df.groupby("Post Type")[interaction_cols].mean()

# Calculate percentage changes
percentage_changes = []
for col in interaction_cols:
    original_value = post_type_group.loc["Original", col]
    repost_value = post_type_group.loc["Repost", col]
    percentage_change = ((repost_value - original_value) / original_value) * 100
    percentage_changes.append([col, percentage_change])

# Convert to DataFrame
percentage_change_df = pd.DataFrame(percentage_changes, columns=["Metric", "Percentage Change"])

# Save results to an Excel file
output_file_tests = "3e_post_vs_repost_tests.xlsx"

with pd.ExcelWriter(output_file_tests, engine="xlsxwriter") as writer:
    pd.DataFrame(normality_results, columns=["Metric", "Normality"]).to_excel(writer, sheet_name="Normality Test", index=False)
    pd.DataFrame(test_results, columns=["Metric", "Test", "Statistic", "p-value", "Significance"]).to_excel(writer, sheet_name="Statistical Tests", index=False)
    percentage_change_df.to_excel(writer, sheet_name="Percentage Change", index=False)

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file_tests}")