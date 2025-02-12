import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
from scikit_posthocs import posthoc_dunn
import seaborn as sns

# Define file path (update this based on your actual file location)
file_path = file_path = "2_processed_linkedin_data.xlsx"

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
output_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/3e_linkedin_post_analysis.xlsx"
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
plot_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/3e_linkedin_post_interaction_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
kruskal_results = []
dunn_results_list = []

for col in interaction_cols:
    # Prepare data for statistical tests
    groups = [df[df["Post Type"] == post_type][col].dropna() for post_type in df["Post Type"].unique()]

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
            dunn_results = posthoc_dunn(df, val_col=col, group_col="Post Type", p_adjust="bonferroni")
            dunn_results.reset_index(inplace=True)
            dunn_results.insert(0, "Metric", col)  # Add metric column for clarity
            dunn_results_list.append(dunn_results)

# Save results to an Excel file
output_file_tests = "/Users/gabrielbarrera/Documents/GitHub/DataProject/3e_post_vs_repost_tests.xlsx"

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