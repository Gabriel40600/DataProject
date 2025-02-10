import pandas as pd
import matplotlib.pyplot as plt

# Define file path (update this based on your actual file location)
file_path = "/Users/gabrielbarrera/Documents/GitHub/DataProject/processed_linkedin_data.xlsx"

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

# Save the results to an Excel file
output_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/cta_engagement_analysis.xlsx"
cta_analysis.to_excel(output_file, index=False)

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
plot_file = "/Users/gabrielbarrera/Documents/GitHub/DataProject/cta_engagement_plot.png"
plt.savefig(plot_file, bbox_inches="tight")

# Show confirmation
print(f"Analysis completed! \nResults saved to: {output_file} \nPlot saved to: {plot_file}")
