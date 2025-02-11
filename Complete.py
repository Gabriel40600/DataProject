import os
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.cluster import KMeans
from sklearn.feature_extraction.text import TfidfVectorizer
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from nltk.tokenize import word_tokenize
from nltk import pos_tag
from nltk.corpus import stopwords
from datetime import datetime, timezone
import nltk

# Download necessary NLTK resources
nltk.download("vader_lexicon", quiet=True)
nltk.download("punkt", quiet=True)
nltk.download("averaged_perceptron_tagger", quiet=True)
nltk.download("stopwords", quiet=True)


# ======================= Data Cleaning =======================
def clean_dataset(input_path, output_path):
    df = pd.read_excel(input_path, sheet_name="Technology & Innovation")
    df = df.drop_duplicates(subset=['Post URL']).fillna(0)
    df['Nur Repost'] = df['Nur Repost'].apply(lambda x: 1 if str(x).strip().lower() == 'x' else 0)
    df.to_excel(output_path, index=False)
    return output_path


# ======================= Feature Extraction =======================
def extract_features(input_path, output_path):
    df = pd.read_excel(input_path)
    analyzer = SentimentIntensityAnalyzer()

    df["Post content"] = df["Post content"].astype(str).fillna("")
    df["Sentiment_Positive"], df["Sentiment_Neutral"], df["Sentiment_Negative"], df["Sentiment_Compound"] = zip(
        *df["Post content"].apply(lambda text: analyzer.polarity_scores(text).values())
    )
    df.to_excel(output_path, index=False)
    return output_path


# ======================= Regression Analysis =======================
def regression_analysis(input_path):
    df = pd.read_excel(input_path)
    df['Post Length'] = df['Post content'].astype(str).apply(len)
    X = df[['Post Length']]
    y = df['reactions']
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    model = LinearRegression()
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    print(f"Linear Regression R²: {r2_score(y_test, y_pred):.2f}")


# ======================= Clustering Analysis =======================
def clustering_analysis(input_path, output_path):
    df = pd.read_excel(input_path)
    vectorizer = TfidfVectorizer(stop_words="english", max_features=5000)
    X = vectorizer.fit_transform(df["Post content"].fillna(""))
    kmeans = KMeans(n_clusters=6, random_state=42, n_init=10)
    df["Cluster"] = kmeans.fit_predict(X)
    df.to_excel(output_path, index=False)
    return output_path


# ======================= Visualization =======================
def visualize_descriptive_stats(input_path):
    df = pd.read_excel(input_path)
    df[["reactions", "comments", "shares"]].boxplot()
    plt.title("Distribution of Engagement Metrics")
    plt.show()


# ======================= Main Execution =======================
if __name__ == "__main__":
    cleaned_file = clean_dataset("data_set.xlsx", "data_set_cleaned.xlsx")
    processed_file = extract_features(cleaned_file, "processed_linkedin_data.xlsx")
    regression_analysis(processed_file)
    clustered_file = clustering_analysis(processed_file, "linkedin_clustered_data.xlsx")
    visualize_descriptive_stats(clustered_file)