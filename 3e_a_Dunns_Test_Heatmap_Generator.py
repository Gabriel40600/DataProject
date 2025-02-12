import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Define file path (only the file name, assuming script runs in the same directory)
file_path = "3e_post_vs_repost_tests.xlsx"

# Load the Dunn's Test results
try:
    dunn_test_results = pd.read_excel(file_path, sheet_name="Dunn's Test")
except ValueError:
    print("Error: 'Dunn's Test' sheet not found. Please check the file.")
    exit()

# Debug: Print available columns in the dataset
print("Available columns in Dunn's Test sheet:", dunn_test_results.columns)

# Define engagement metrics
metrics = ["reactions", "comments", "shares"]
post_types = ["Original", "Repost"]

# Ensure 'index' column exists before proceeding
if "index" not in dunn_test_results.columns:
    print("Error: 'index' column not found in Dunn's Test results. Check dataset format.")
    exit()

# Generate Heatmaps for each engagement metric
for metric in metrics:
    plt.figure(figsize=(8, 6))

    # Extract subset for the current metric
    subset = dunn_test_results[dunn_test_results["Metric"] == metric]

    # Debug: Print subset size before proceeding
    print(f"Processing {metric}: Found {subset.shape[0]} rows.")

    # Ensure the data is not empty
    if subset.empty:
        print(f"Skipping {metric}: No valid data found in the dataset.")
        continue

    subset = subset.set_index("index")[post_types]

    # Convert values to float and replace NaNs for visualization
    subset = subset.astype(float).fillna(1.0)

    # Check if there is any data left
    if subset.shape[0] == 0 or subset.shape[1] == 0:
        print(f"Skipping {metric}: No valid comparison data available.")
        continue

    # Generate heatmap
    sns.heatmap(subset, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
    plt.title(f"Dunn's Test - Pairwise Engagement Differences for {metric.capitalize()} (Original vs. Repost)")
    plt.ylabel("Post Type")
    plt.xlabel("Compared Post Type")

    # Save heatmap (only file name, assuming script runs in the same directory)
    heatmap_file = f"3e_a_dunn_heatmap_{metric}.png"
    plt.savefig(heatmap_file, bbox_inches="tight")
    print(f"Saved heatmap: {heatmap_file}")

plt.show()
