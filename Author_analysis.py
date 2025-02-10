import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.preprocessing import LabelEncoder

# Load dataset
file_path = "processed_linkedin_data.xlsx"
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Encode authors
df["TL_encoded"] = LabelEncoder().fit_transform(df["TL"])

# Selecting numeric columns only
numeric_df = df.select_dtypes(include=["number"])

# Selecting author and engagement metrics
engagement_metrics = ["reactions", "comments", "shares"]

# Compute correlation values for TL_encoded vs engagement metrics
correlation_values = numeric_df.corr().loc["TL_encoded", engagement_metrics]

# Save correlation results to an Excel file
output_excel = "author_engagement_correlation.xlsx"
correlation_values.to_frame().T.to_excel(output_excel, index=False)
print(f"Correlation values saved to {output_excel}")

# Display correlation values
print("Correlation between TL_encoded and engagement metrics:")
print(correlation_values)

# Create a bar chart visualization
plt.figure(figsize=(8, 5))
correlation_values.plot(kind='bar', color=['blue', 'orange', 'green'])
plt.title("Correlation Between Authors and Engagement Metrics")
plt.ylabel("Correlation Coefficient")
plt.xticks(rotation=0)
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Save the visualization
output_image = "author_engagement_correlation.png"
plt.savefig(output_image)
print(f"Visualization saved as {output_image}")

# Show the plot
plt.show()
