import pandas as pd
import matplotlib.pyplot as plt

# Define file path
file_path = "deskriptive_statistik.xlsx"

# Load the Excel file
xls = pd.ExcelFile(file_path)

# Load "Descriptive Stats" and create a boxplot
descriptive_stats = xls.parse("Descriptive Stats").set_index("Unnamed: 0")
descriptive_stats_filtered = descriptive_stats.loc[["reactions", "comments", "shares"], ["min", "25%", "50%", "75%", "max"]].T

plt.figure(figsize=(12, 6))
plt.boxplot([descriptive_stats_filtered["reactions"],
             descriptive_stats_filtered["comments"],
             descriptive_stats_filtered["shares"]], tick_labels=["Reactions", "Comments", "Shares"])
plt.title("Distribution of Reactions, Comments, and Shares")
plt.ylabel("Count")
plt.yscale("log")
plt.savefig("boxplot_engagement.png")
plt.close()

# Load "Post Analysis" and create a bar chart
post_analysis = xls.parse("Post Analysis")
plt.figure(figsize=(8, 6))
plt.bar(post_analysis["Metric"], post_analysis["Value"], color=['blue', 'green', 'red'])
plt.title("Percentage of Posts Containing Emojis, Hashtags, and Links")
plt.ylabel("Percentage (%)")
plt.savefig("bar_post_analysis.png")
plt.close()

# Function to create grouped bar charts for reactions, comments, and shares
def plot_grouped_bar(sheet_name, title, filename):
    data = xls.parse(sheet_name)
    data.set_index(data.columns[0], inplace=True)
    data.plot(kind='bar', figsize=(10, 6))
    plt.title(title)
    plt.ylabel("Count")
    plt.savefig(filename)
    plt.close()

# Save visualizations for Link Analysis
plot_grouped_bar("Link Reactions", "Engagement Metrics - Reactions for Posts with and without Links", "bar_link_reactions.png")
plot_grouped_bar("Link Comments", "Engagement Metrics - Comments for Posts with and without Links", "bar_link_comments.png")
plot_grouped_bar("Link Shares", "Engagement Metrics - Shares for Posts with and without Links", "bar_link_shares.png")

# Save visualizations for Emoji Analysis
plot_grouped_bar("Emoji Reactions", "Engagement Metrics - Reactions for Posts with and without Emojis", "bar_emoji_reactions.png")
plot_grouped_bar("Emoji Comments", "Engagement Metrics - Comments for Posts with and without Emojis", "bar_emoji_comments.png")
plot_grouped_bar("Emoji Shares", "Engagement Metrics - Shares for Posts with and without Emojis", "bar_emoji_shares.png")

# Save visualizations for Hashtag Analysis
plot_grouped_bar("Hashtag Reactions", "Engagement Metrics - Reactions for Posts with and without Hashtags", "bar_hashtag_reactions.png")
plot_grouped_bar("Hashtag Comments", "Engagement Metrics - Comments for Posts with and without Hashtags", "bar_hashtag_comments.png")
plot_grouped_bar("Hashtag Shares", "Engagement Metrics - Shares for Posts with and without Hashtags", "bar_hashtag_shares.png")

print("Visualizations saved successfully!")
