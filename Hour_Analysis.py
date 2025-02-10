import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Load the LinkedIn data
file_path = "processed_linkedin_data.xlsx"  # Update with the correct path if necessary
df = pd.read_excel(file_path)

# Remove invalid timestamp entries
df_cleaned = df[df['Post Timestamp (ISO)'] != "Invalid LinkedIn ID"].copy()

# Convert timestamp to datetime and extract the hour of the day
df_cleaned['Post Timestamp (ISO)'] = pd.to_datetime(df_cleaned['Post Timestamp (ISO)'], errors='coerce')
df_cleaned = df_cleaned.dropna(subset=['Post Timestamp (ISO)'])  # Drop rows where conversion failed
df_cleaned['Hour of Day'] = df_cleaned['Post Timestamp (ISO)'].dt.hour

# Count posts per hour
posts_per_hour = df_cleaned['Hour of Day'].value_counts().sort_index()

# Aggregate total interactions (reactions, comments, shares) by hour
metrics = ['reactions', 'comments', 'shares']
interaction_data = df_cleaned.groupby('Hour of Day')[metrics].sum()

# Normalize by number of posts per hour
normalized_interactions = interaction_data.div(posts_per_hour, axis=0)

# Save to an Excel file
output_file = "normalized_interactions_by_hour.xlsx"
normalized_interactions.to_excel(output_file)

# Create a grouped bar chart with dual y-axis
bar_width = 0.3  # Width of each bar
x = np.arange(len(normalized_interactions.index))  # X positions for bars

fig, ax1 = plt.subplots(figsize=(12, 6))

# Primary y-axis for Reactions
ax1.bar(x - bar_width, normalized_interactions['reactions'], width=bar_width, label="Reactions", color='tab:blue')
ax1.set_ylabel("Avg Reactions per Post", color='tab:blue')
ax1.tick_params(axis='y', labelcolor='tab:blue')

# Secondary y-axis for Comments & Shares
ax2 = ax1.twinx()
ax2.bar(x, normalized_interactions['comments'], width=bar_width, label="Comments", color='tab:orange')
ax2.bar(x + bar_width, normalized_interactions['shares'], width=bar_width, label="Shares", color='tab:green')
ax2.set_ylabel("Avg Comments & Shares per Post", color='black')

# Formatting
ax1.set_xlabel("Hour of the Day")
ax1.set_title("Average Interactions per Post by Hour with Dual Scale")
ax1.set_xticks(x)
ax1.set_xticklabels(normalized_interactions.index)
ax1.grid(axis='y', linestyle='--', alpha=0.7)

# Legend
fig.legend(loc="upper right", bbox_to_anchor=(1,1), bbox_transform=ax1.transAxes)

plt.tight_layout()
plt.show()

print(f"Data successfully saved to {output_file}")
