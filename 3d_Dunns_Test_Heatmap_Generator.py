import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Define file path
file_path = "3d_interactions_by_hour_tests.xlsx"

# Load the Dunn's Test results
dunn_test_results = pd.read_excel(file_path, sheet_name="Dunn's Test")

# Define engagement metrics and hours
metrics = ["reactions", "comments", "shares"]
hours = list(range(24))  # 0 to 23 hours

# Generate Heatmaps for each engagement metric
for metric in metrics:
    plt.figure(figsize=(10, 6))

    # Extract subset for the current metric
    subset = dunn_test_results[dunn_test_results["Metric"] == metric].set_index("index")[hours]

    # Convert values to float and replace NaNs for visualization
    subset = subset.astype(float).fillna(1.0)

    # Generate heatmap
    sns.heatmap(subset, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
    plt.title(f"Dunn's Test - Pairwise Engagement Differences for {metric.capitalize()} by Hour")
    plt.ylabel("Hour of the Day")
    plt.xlabel("Compared Hour")

    # Save heatmap
    heatmap_file = f"3d_dunn_heatmap_{metric}.png"
    plt.savefig(heatmap_file, bbox_inches="tight")
    print(f"Saved heatmap: {heatmap_file}")

plt.show()
