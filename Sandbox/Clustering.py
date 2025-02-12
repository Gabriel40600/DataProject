import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

# Load the processed LinkedIn dataset
file_path = "../2_processed_linkedin_data.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Extracting the post content for clustering
text_data = df["Post content"].dropna().tolist()

# Convert text data into TF-IDF features
vectorizer = TfidfVectorizer(stop_words="english", max_features=5000)  # Limit features for performance
X = vectorizer.fit_transform(text_data)

# Determine the optimal number of clusters using the Elbow Method
inertia = []
K_range = range(2, 10)

for k in K_range:
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    kmeans.fit(X)
    inertia.append(kmeans.inertia_)

# Plot the Elbow Method graph
plt.figure(figsize=(8, 5))
plt.plot(K_range, inertia, marker='o', linestyle='--')
plt.xlabel("Number of Clusters")
plt.ylabel("Inertia")
plt.title("Elbow Method for Optimal K in K-Means")
plt.show()

# Apply K-Means clustering with the chosen number of clusters (k=6)
optimal_k = 6
kmeans = KMeans(n_clusters=optimal_k, random_state=42, n_init=10)
df_cleaned = df.dropna(subset=["Post content"]).copy()  # Only keep rows where text exists
df_cleaned["Cluster"] = kmeans.fit_predict(X)

# Get the top terms per cluster
order_centroids = kmeans.cluster_centers_.argsort()[:, ::-1]
terms = vectorizer.get_feature_names_out()

cluster_keywords = {}
for i in range(optimal_k):
    top_terms = [terms[ind] for ind in order_centroids[i, :10]]
    cluster_keywords[f"Cluster {i}"] = top_terms

# Count the number of posts in each cluster
cluster_counts = df_cleaned["Cluster"].value_counts().sort_index()

# Plot the distribution of posts per cluster
plt.figure(figsize=(8, 5))
sns.barplot(x=cluster_counts.index, y=cluster_counts.values, palette="viridis")
plt.xlabel("Cluster")
plt.ylabel("Number of Posts")
plt.title("Distribution of Posts Across 6 Clusters")
plt.xticks(ticks=cluster_counts.index, labels=[f"Cluster {i}" for i in cluster_counts.index])
plt.show()

# Analyze engagement (reactions, comments, shares) by cluster
engagement_metrics = df_cleaned.groupby("Cluster")[["reactions", "comments", "shares"]].mean()

# Display engagement metrics for each cluster
print("\nEngagement Metrics by Cluster:")
print(engagement_metrics)

# Display top terms per cluster
print("\nTop Terms per Cluster:")
for cluster, terms in cluster_keywords.items():
    print(f"{cluster}: {', '.join(terms)}")

# Save the clustered data
df_cleaned.to_excel("linkedin_clustered_data.xlsx", index=False)
print("\nClustered data saved to 'linkedin_clustered_data.xlsx'")
