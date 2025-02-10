import pandas as pd
import matplotlib.pyplot as plt

# Define file path (update this based on your actual file location)
file_path = "/Users/gabrielbarrera/Documents/GitHub/DataProject/processed_linkedin_data.xlsx"

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
output_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/linkedin_post_analysis.xlsx"
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
plot_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/linkedin_post_interaction_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file} \nPlot saved to: {plot_file}")

