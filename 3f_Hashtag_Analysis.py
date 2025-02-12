import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
from scipy.stats import ttest_ind, mannwhitneyu
import seaborn as sns

# Define file path dynamically
file_path = "2_processed_linkedin_data.xlsx"

# Load the dataset
df = pd.read_excel(file_path)

# Define columns of interest
hashtag_col = "Contains Hashtag"
interaction_cols = ["reactions", "comments", "shares"]

# Create a new column for Hashtag Presence
df["Hashtag Presence"] = df[hashtag_col].apply(lambda x: "With Hashtag" if x == 1 else "No Hashtag")

# Group by Hashtag Presence and calculate the average interactions
interaction_summary = df.groupby("Hashtag Presence")[interaction_cols].mean().reset_index()

# Plot the interaction comparison
fig, ax = plt.subplots(figsize=(8, 6))
interaction_summary.set_index("Hashtag Presence").plot(kind="bar", ax=ax)
plt.title("Average Interactions by Hashtag Presence")
plt.ylabel("Average Count")
plt.xlabel("Hashtag Presence")
plt.xticks(rotation=0)
plt.legend(title="Interaction Type")
plt.grid(axis="y", linestyle="--", alpha=0.7)

# Save the plot
plot_file = "3f_hashtag_interaction_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# ✅ Perform Normality Test (Shapiro-Wilk) and Statistical Tests
print("Normality Test (Shapiro-Wilk):")
normality_results = []
test_results = []

for col in interaction_cols:
    # Prepare data for statistical tests
    groups = [df[df["Hashtag Presence"] == hashtag_type][col].dropna() for hashtag_type in df["Hashtag Presence"].unique()]

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

# ✅ Compute Percentage Change Between "With Hashtag" and "No Hashtag"
hashtag_group = df.groupby("Hashtag Presence")[interaction_cols].mean()

# Calculate percentage changes
percentage_changes = []
for col in interaction_cols:
    if "No Hashtag" in hashtag_group.index and "With Hashtag" in hashtag_group.index:
        no_hashtag_value = hashtag_group.loc["No Hashtag", col]
        with_hashtag_value = hashtag_group.loc["With Hashtag", col]
        percentage_change = ((with_hashtag_value - no_hashtag_value) / no_hashtag_value) * 100
        percentage_changes.append([col, percentage_change])
    else:
        percentage_changes.append([col, None])  # Handle cases where one category is missing

# Convert to DataFrame
percentage_change_df = pd.DataFrame(percentage_changes, columns=["Metric", "Percentage Change"])

# ✅ Save all results to an Excel file
output_file_tests = "3f_hashtag_tests.xlsx"

with pd.ExcelWriter(output_file_tests, engine="xlsxwriter") as writer:
    interaction_summary.to_excel(writer, sheet_name="Average Interactions", index=False)
    pd.DataFrame(normality_results, columns=["Metric", "Normality"]).to_excel(writer, sheet_name="Normality Test", index=False)
    pd.DataFrame(test_results, columns=["Metric", "Test", "Statistic", "p-value", "Significance"]).to_excel(writer, sheet_name="Statistical Tests", index=False)
    percentage_change_df.to_excel(writer, sheet_name="Percentage Change", index=False)

# ✅ Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file_tests}")
