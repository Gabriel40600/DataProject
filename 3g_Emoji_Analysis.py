import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as stats
from scipy.stats import mannwhitneyu
import seaborn as sns

# Define file path
file_path = "2_processed_linkedin_data.xlsx"

# Load the dataset
df = pd.read_excel(file_path)

# Define columns of interest
emoji_col = "Contains Emoji"
interaction_cols = ["reactions", "comments", "shares"]

# Create a new column for Emoji Presence
df["Emoji Presence"] = df[emoji_col].apply(lambda x: "With Emoji" if x == 1 else "No Emoji")

# Group by Emoji Presence and calculate the average interactions
interaction_summary = df.groupby("Emoji Presence")[interaction_cols].mean().reset_index()

# Perform Normality Test (Shapiro-Wilk)
print("Normality Test (Shapiro-Wilk):")
normality_results = []

df["Emoji Presence"] = df["Emoji Presence"].astype("category")

for col in interaction_cols:
    stat, p_value = stats.shapiro(df[col].dropna())
    is_normal = "Yes" if p_value > 0.05 else "No"
    normality_results.append((col, stat, p_value, is_normal))
    print(f"{col}: W={stat:.4f}, p={p_value:.4f}, Normal: {is_normal}")

# Perform Mann-Whitney U Test
print("\nMann-Whitney U Test:")
mann_whitney_results = []

for col in interaction_cols:
    group_emoji = df[df["Emoji Presence"] == "With Emoji"][col].dropna()
    group_no_emoji = df[df["Emoji Presence"] == "No Emoji"][col].dropna()

    u_stat, p_value = mannwhitneyu(group_emoji, group_no_emoji, alternative="two-sided")
    significance = "Significant" if p_value < 0.05 else "Not Significant"
    mann_whitney_results.append([col, u_stat, p_value, significance])
    print(f"{col}: U={u_stat:.4f}, p={p_value:.4f}, {significance}")

# Save results to an Excel file
output_file = "3g_Emoji_tests.xlsx"

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    interaction_summary.to_excel(writer, sheet_name="Average Interactions", index=False)
    pd.DataFrame(normality_results, columns=["Metric", "W-Statistic", "P-Value", "Normal"]).to_excel(writer, sheet_name="Normality Test", index=False)
    pd.DataFrame(mann_whitney_results, columns=["Metric", "U-Statistic", "P-Value", "Significance"]).to_excel(writer, sheet_name="Mann-Whitney Test", index=False)

# Visualization: Boxplot to show engagement distribution per emoji presence
fig, ax = plt.subplots(figsize=(8, 6))
interaction_summary.set_index("Emoji Presence").plot(kind="bar", ax=ax)
plt.title("Effect of Emojis on Engagement")
plt.ylabel("Average Count")
plt.xlabel("Emoji Presence")
plt.xticks(rotation=0)
plt.legend(title="Interaction Type")
plt.grid(axis="y", linestyle="--", alpha=0.7)
plt.savefig("3g_emoji_interaction_plot.png", bbox_inches="tight")

print("Analysis completed! Results saved to 3g_Emoji_tests.xlsx and visualization as 3g_emoji_interaction_plot.png")

