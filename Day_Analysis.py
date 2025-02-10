import pandas as pd
import matplotlib.pyplot as plt

# Load the LinkedIn data
file_path = "processed_linkedin_data.xlsx"  # Update with the correct path if necessary
df = pd.read_excel(file_path)

# Remove invalid timestamp entries
df_cleaned = df[df['Post Timestamp (ISO)'] != "Invalid LinkedIn ID"].copy()

# Convert timestamp to datetime and extract the day of the week
df_cleaned['Post Timestamp (ISO)'] = pd.to_datetime(df_cleaned['Post Timestamp (ISO)'], errors='coerce')
df_cleaned = df_cleaned.dropna(subset=['Post Timestamp (ISO)'])  # Drop rows where conversion failed
df_cleaned['Day of Week'] = df_cleaned['Post Timestamp (ISO)'].dt.day_name()

# Count posts per day
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
posts_per_day = df_cleaned['Day of Week'].value_counts().reindex(day_order, fill_value=0)

# Aggregate total interactions (reactions, comments, shares) by day of the week
metrics = ['reactions', 'comments', 'shares']
interaction_data = df_cleaned.groupby('Day of Week')[metrics].sum().reindex(day_order)

# Normalize by number of posts per day
normalized_interactions = interaction_data.div(posts_per_day, axis=0)

# Save to an Excel file
output_file = "interactions_by_day.xlsx"
normalized_interactions.to_excel(output_file)

print(f"Data successfully saved to {output_file}")
