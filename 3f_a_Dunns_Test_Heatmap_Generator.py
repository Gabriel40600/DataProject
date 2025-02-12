import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Define file path
file_path = "3f_Hashtag_tests.xlsx"

# Load the Dunn's Test results
try:
    dunn_test_results = pd.read_excel(file_path, sheet_name="Dunn's Test")
except ValueError:
    print("Error: 'Dunn's Test' sheet not found. Please check the file.")
    exit()

# Debug: Print available columns in the dataset
dunn_test_results.columns = dunn_test_results.columns.str.strip()  # Remove any spaces in column names
print("Available columns in Dunn's Test sheet:", dunn_test_results.columns.tolist())

# Define engagement metrics and ensure correct post types
metrics = ["reactions", "comments", "shares"]
post_types = ["No Hashtag", "with Hashtag"]  # Corrected post type names

# Ensure 'index' column exists before proceeding
if "index" not in dunn_test_results.columns:
    print("Error: 'index' column not found. Available columns:", dunn_test_results.columns.tolist())
    exit()

# Debug: Print available metrics
print("Unique values in 'Metric' column:", dunn_test_results["Metric"].unique())

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

    # ✅ Set "index" as row names and select "No Hashtag" & "with Hashtag"
    subset = subset.set_index("index")[post_types]

    # Convert values to float and replace NaNs for visualization
    subset = subset.astype(float).fillna(1.0)

    # Check if there is any data left
    if subset.shape[0] == 0 or subset.shape[1] == 0:
        print(f"Skipping {metric}: No valid comparison data available.")
        continue

    # Generate heatmap
    sns.heatmap(subset, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
    plt.title(f"Dunn's Test - Pairwise Engagement Differences for {metric.capitalize()} (No Hashtag vs. with Hashtag)")
    plt.ylabel("Post Type")
    plt.xlabel("Compared Post Type")

    # Save heatmap
    heatmap_file = f"3f_a_dunn_heatmap_{metric}.png"
    plt.savefig(heatmap_file, bbox_inches="tight")
    print(f"Saved heatmap: {heatmap_file}")

plt.show()
