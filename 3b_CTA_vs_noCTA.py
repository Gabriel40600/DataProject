import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
import seaborn as sns

# Define file path (update this based on your actual file location)
file_path = "/2_processed_linkedin_data.xlsx"

# Load the dataset
df = pd.read_excel(file_path)

# Define columns of interest
cta_col = "CTA Present"
interaction_cols = ["reactions", "comments", "shares"]

# Ensure CTA column is correctly formatted as 0 or 1
df[cta_col] = df[cta_col].fillna(0).astype(int)

# Create a new column for CTA presence
df["CTA Type"] = df[cta_col].apply(lambda x: "With CTA" if x == 1 else "Without CTA")

# Group by CTA presence and calculate average interactions
cta_analysis = df.groupby("CTA Type")[interaction_cols].mean().reset_index()

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []
cta_group = df[df[cta_col] == 1][interaction_cols]
no_cta_group = df[df[cta_col] == 0][interaction_cols]

for col in interaction_cols:
    stat_cta, p_cta = stats.shapiro(cta_group[col].dropna())
    stat_no_cta, p_no_cta = stats.shapiro(no_cta_group[col].dropna())

    normality_results.append([col, "CTA Group", p_cta, "Normal" if p_cta > 0.05 else "Not Normal"])
    normality_results.append([col, "No CTA Group", p_no_cta, "Normal" if p_no_cta > 0.05 else "Not Normal"])
    print(f"{col}: CTA Group p-value: {p_cta:.5f}, No CTA Group p-value: {p_no_cta:.5f}")

# Perform Mann-Whitney U Test
print("\nMann-Whitney U Test:")
mann_whitney_results = []

for col in interaction_cols:
    u_stat, mw_p = stats.mannwhitneyu(cta_group[col].dropna(), no_cta_group[col].dropna(), alternative="two-sided")
    mann_whitney_results.append([col, u_stat, mw_p, "Significant" if mw_p < 0.05 else "Not Significant"])
    print(f"{col} - U-statistic: {u_stat:.3f}, p-value: {mw_p:.5f}")

# Save results to an Excel file
output_file = "/3b_cta_engagement_analysis.xlsx"

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    cta_analysis.to_excel(writer, sheet_name="CTA Engagement", index=False)
    pd.DataFrame(normality_results, columns=["Metric", "Group", "p-value", "Normality"]).to_excel(writer, sheet_name="Normality Test", index=False)
    pd.DataFrame(mann_whitney_results, columns=["Metric", "U-Statistic", "p-value", "Significance"]).to_excel(writer, sheet_name="Mann-Whitney Test", index=False)

# Plot the interaction comparison
fig, ax = plt.subplots(figsize=(8, 6))
cta_analysis.set_index("CTA Type").plot(kind="bar", ax=ax)
plt.title("Effect of Call-to-Action on Engagement")
plt.ylabel("Average Count")
plt.xlabel("CTA Presence")
plt.xticks(rotation=0)
plt.legend(title="Interaction Type")
plt.grid(axis="y", linestyle="--", alpha=0.7)

# Save the plot
plot_file = "/3b_cta_engagement_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file} \nPlot saved to: {plot_file}")