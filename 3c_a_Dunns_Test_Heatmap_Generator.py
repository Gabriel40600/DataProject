import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load the Dunn's Test results
file_path = "/Users/gabrielbarrera/Documents/GitHub/DataProject/3c_interactions_by_day_tests.xlsx"
dunn_test_results = pd.read_excel(file_path, sheet_name="Dunn's Test")

# Define engagement metrics and days
metrics = ["reactions", "comments", "shares"]
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# Generate Heatmaps for each engagement metric
for metric in metrics:
    plt.figure(figsize=(8, 6))
    
    # Extract subset for the current metric
    subset = dunn_test_results[dunn_test_results["Metric"] == metric].set_index("index")[days]
    
    # Generate heatmap
    sns.heatmap(subset.astype(float), annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
    plt.title(f"Dunn's Test - Pairwise Engagement Differences for {metric.capitalize()}")
    plt.ylabel("Day of the Week")
    plt.xlabel("Compared Day")

    # Save heatmap
    heatmap_file = f"/Users/gabrielbarrera/Documents/GitHub/DataProject/3c_a_dunn_heatmap_{metric}.png"
    plt.savefig(heatmap_file, bbox_inches="tight")
    print(f"Saved heatmap: {heatmap_file}")

plt.show()
