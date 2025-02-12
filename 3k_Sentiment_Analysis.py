import pandas as pd
import scipy.stats as stats
from scikit_posthocs import posthoc_dunn
import matplotlib.pyplot as plt
import seaborn as sns

# Load the processed LinkedIn data
file_path = "2_processed_linkedin_data.xlsx"
df = pd.read_excel(file_path)

# Define sentiment and engagement columns
sentiment_columns = ["Positive Sentiment", "Neutral Sentiment", "Negative Sentiment", "Compound Sentiment"]
engagement_columns = ["reactions", "comments", "shares"]

# Clean data by removing any NaN values in sentiment and engagement columns
df_cleaned = df.dropna(subset=sentiment_columns + engagement_columns)

# Categorize posts into sentiment groups based on compound sentiment score
bins = [-1, -0.5, 0, 0.5, 1]
labels = ["Very Negative", "Negative", "Neutral", "Positive"]
df_cleaned["Sentiment Category"] = pd.cut(df_cleaned["Compound Sentiment"], bins=bins, labels=labels)

# Compute mean and median engagement per sentiment category
engagement_summary = df_cleaned.groupby("Sentiment Category")[engagement_columns].agg(['mean', 'median'])

# Perform Kruskal-Wallis test for each engagement metric
df_grouped = df_cleaned.groupby("Sentiment Category")[engagement_columns]
kruskal_results = {metric: stats.kruskal(*[group[metric].values for _, group in df_grouped]) for metric in engagement_columns}

# Convert Kruskal-Wallis results to DataFrame with correct column names
kruskal_df = pd.DataFrame(kruskal_results, index=["H Statistic", "p-value"]).T

# Perform Dunn’s test for post-hoc analysis and combine all results into one DataFrame
dunn_results_combined = []
for metric in engagement_columns:
    dunn_df = posthoc_dunn(df_cleaned, val_col=metric, group_col="Sentiment Category", p_adjust='bonferroni')
    dunn_df.insert(0, "Metric", metric)
    dunn_results_combined.append(dunn_df)

dunn_results_df = pd.concat(dunn_results_combined)

# Save results to Excel
with pd.ExcelWriter("3k_Sentiment_Analysis.xlsx") as writer:
    engagement_summary.to_excel(writer, sheet_name="Engagement Summary")
    kruskal_df.to_excel(writer, sheet_name="Kruskal-Wallis Test")
    dunn_results_df.to_excel(writer, sheet_name="Dunn's Test", index=False)

# Visualization: Boxplot to show engagement distribution per sentiment
plt.figure(figsize=(12, 6))
for metric in engagement_columns:
    plt.figure(figsize=(8, 6))
    sns.boxplot(x="Sentiment Category", y=metric, data=df_cleaned, palette="coolwarm")
    plt.title(f"Engagement ({metric}) by Sentiment Category")
    plt.xticks(rotation=45)
    plt.savefig(f"{metric}_sentiment_boxplot.png")
    plt.close()

print("Sentiment impact analysis on engagement completed. Results saved to 3k_Sentiment_Analysis.xlsx")
