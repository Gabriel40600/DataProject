import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset
file_path = "processed_linkedin_data.xlsx"
df_raw = pd.read_excel(file_path)

# Clean invalid timestamp values
df_raw = df_raw[df_raw["Post Timestamp (ISO)"].str.contains("Invalid LinkedIn ID") == False].copy()

# Convert timestamp column to datetime
df_raw["Post Timestamp (ISO)"] = pd.to_datetime(df_raw["Post Timestamp (ISO)"], errors="coerce")

# Drop any rows with missing timestamps after conversion
df_raw = df_raw.dropna(subset=["Post Timestamp (ISO)"])

# Extract day of the week and hour
df_raw["Day of Week"] = df_raw["Post Timestamp (ISO)"].dt.day_name()
df_raw["Hour of Day"] = df_raw["Post Timestamp (ISO)"].dt.hour

# Aggregate engagement metrics by day and hour
df_grouped = df_raw.groupby(["Day of Week", "Hour of Day"])[["reactions", "comments", "shares"]].mean().reset_index()

# Normalize engagement values
df_grouped[['reactions', 'comments', 'shares']] = df_grouped[['reactions', 'comments', 'shares']].apply(
    lambda x: x / x.max()
)

# Save processed data to an Excel file
output_file = "linkedin_posting_analysis.xlsx"
df_grouped.to_excel(output_file, index=False)
print(f"Processed data saved to {output_file}")

# Reorder days for correct weekly sequence
ordered_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
df_grouped["Day of Week"] = pd.Categorical(df_grouped["Day of Week"], categories=ordered_days, ordered=True)

# Pivot data for separate heatmaps
heatmap_reactions = df_grouped.pivot(index="Day of Week", columns="Hour of Day", values="reactions")
heatmap_comments = df_grouped.pivot(index="Day of Week", columns="Hour of Day", values="comments")
heatmap_shares = df_grouped.pivot(index="Day of Week", columns="Hour of Day", values="shares")

# Plot heatmaps
def plot_heatmap(data, title, cmap):
    plt.figure(figsize=(12, 6))
    sns.heatmap(data, cmap=cmap, annot=True, fmt=".2f", linewidths=0.5)
    plt.title(title)
    plt.xlabel("Hour of Day")
    plt.ylabel("Day of Week")
    plt.show()

plot_heatmap(heatmap_reactions, "Best Posting Times for Reactions", "Reds")
plot_heatmap(heatmap_comments, "Best Posting Times for Comments", "Greens")
plot_heatmap(heatmap_shares, "Best Posting Times for Shares", "Blues")
