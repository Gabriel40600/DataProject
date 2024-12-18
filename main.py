import pandas as pd
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from nltk.tokenize import word_tokenize
from nltk import pos_tag
from collections import Counter
import requests
import nltk
import re
import urllib.parse
from datetime import datetime, timezone

# Download necessary NLTK resources
nltk.download("vader_lexicon")
nltk.download("stopwords")
nltk.download("punkt")
nltk.download("averaged_perceptron_tagger")

# Regex patterns for extracting statistics
number_pattern = r"\b(?!19[0-9]{2}|20[0-9]{2})\d+\.?\d*\b"  # Exclude years
percent_pattern = r"\d+\.?\d*%"  # Matches percentages
stat_terms_pattern = r"\b(mean|median|average|increase|decrease|percent|growth|std|variance|correlation)\b"

# Function to extract meaningful statistics
def extract_statistics(text):
    percentages = re.findall(percent_pattern, text)
    numbers = re.findall(number_pattern, text)
    stat_terms = re.findall(stat_terms_pattern, text, flags=re.IGNORECASE)

    # Combine numbers and percentages only if statistical terms are present
    if not stat_terms and not percentages:
        numbers = []  # Exclude numbers if no statistical context

    return {
        "numbers": numbers,
        "percentages": percentages,
        "terms": stat_terms
    }

# Function to classify statistic usage
def classify_statistics(stat_data):
    has_terms = bool(stat_data["terms"])
    has_percentages = bool(stat_data["percentages"])
    has_numbers = bool(stat_data["numbers"])

    if has_terms or has_percentages:
        return "YES"
    elif has_numbers:
        return "MAYBE"
    else:
        return "NO"

# Function to extract emojis from text
def extract_emojis(text):
    emoji_pattern = re.compile(
        "[\U0001F600-\U0001F64F"  # Emoticons
        "\U0001F300-\U0001F5FF"  # Symbols & Pictographs
        "\U0001F680-\U0001F6FF"  # Transport & Map
        "\U0001F1E0-\U0001F1FF"  # Flags
        "\U00002600-\U000027BF"  # Miscellaneous Symbols
        "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
        "\U0001FA70-\U0001FAFF"  # Symbols and Pictographs Extended-A
        "\U00002500-\U00002BEF"  # Chinese/Japanese/Korean characters
        "]+", flags=re.UNICODE
    )
    return emoji_pattern.findall(text)

# Function to extract hashtags from text
def extract_hashtags(text):
    hashtag_pattern = re.compile(r"#\w+")
    return hashtag_pattern.findall(text)

# Function to count adjectives in text
def count_adjectives(text):
    tokens = word_tokenize(text)  # Tokenize text into words
    pos_tags = pos_tag(tokens)  # Tag each word with its part of speech
    adjectives = [word for word, pos in pos_tags if pos in ["JJ", "JJR", "JJS"]]  # Filter adjectives
    return len(adjectives)

# Function to extract adjectives from text
def extract_adjectives(text):
    tokens = word_tokenize(text)  # Tokenize text into words
    pos_tags = pos_tag(tokens)  # Tag each word with its part of speech
    adjectives = [word for word, pos in pos_tags if pos in ["JJ", "JJR", "JJS"]]  # Filter adjectives
    return adjectives

# Function to extract post ID from LinkedIn URL
def extract_post_id(url):
    post_id_pattern = r"activity[-:]?(\d+)"
    match = re.search(post_id_pattern, url)
    if match:
        post_id = match.group(1)
        return post_id
    else:
        return None

# Function to extract and format timestamp from LinkedIn activity URL
class LIPostTimestampExtractor:
    @staticmethod
    def format_timestamp(timestamp_s, get_local: bool = False):
        # Format the timestamp to UTC
        if get_local:
            date = datetime.fromtimestamp(timestamp_s)
            return date.strftime('%a, %d %b %Y %H:%M:%S GMT')
        else:
            date = datetime.fromtimestamp(timestamp_s, tz=timezone.utc)
            return date.strftime('%a, %d %b %Y %H:%M:%S GMT (UTC)')

    @classmethod
    def get_date_from_linkedin_activity(cls, post_url: str, get_local: bool = False) -> str:
        try:
            post_url_unquote = urllib.parse.unquote(post_url)
            match = re.search(r'activity-(\d+)', post_url_unquote)
            if not match:
                return 'Invalid LinkedIn ID'

            linkedin_id = match.group(1)

            # Extract the first 41 bits directly
            first_41_bits = bin(int(linkedin_id))[2:43]

            # Convert to timestamp in milliseconds
            timestamp_ms = int(first_41_bits, 2)

            # Convert to seconds
            timestamp_s = timestamp_ms / 1000

            # Format timestamp
            return cls.format_timestamp(timestamp_s, get_local)

        except (ValueError, IndexError):
            return 'Date not available'

# Initialize Sentiment Analyzer
analyzer = SentimentIntensityAnalyzer()

# Load the Excel file and all sheets
file_path = "data_set.xlsx"  # Replace with your file path
all_sheets = pd.read_excel(file_path, sheet_name=None)

# Specify the target sheet and columns
target_sheet = "Technology & Innovation"
post_content_column = "Post content"
linkedin_url_column = "Post URL"  # Updated column name

if target_sheet in all_sheets:
    # Load the target sheet
    target_df = all_sheets[target_sheet]

    # Check if the LinkedIn URL column exists
    if linkedin_url_column in target_df.columns:
        print(f"'{linkedin_url_column}' column found in the dataset.")
    else:
        print(f"'{linkedin_url_column}' column NOT found in the dataset.")

    # 1. Add sentiment analysis for "Post content"
    if post_content_column in target_df.columns:
        target_df["Sentiment_Positive"] = target_df[post_content_column].fillna("").apply(lambda x: analyzer.polarity_scores(x)["pos"])
        target_df["Sentiment_Neutral"] = target_df[post_content_column].fillna("").apply(lambda x: analyzer.polarity_scores(x)["neu"])
        target_df["Sentiment_Negative"] = target_df[post_content_column].fillna("").apply(lambda x: analyzer.polarity_scores(x)["neg"])
        target_df["Sentiment_Compound"] = target_df[post_content_column].fillna("").apply(lambda x: analyzer.polarity_scores(x)["compound"])

    # 2. Extract emojis and count their occurrences
    if post_content_column in target_df.columns:
        target_df["Post content_emojis"] = target_df[post_content_column].fillna("").apply(extract_emojis)
        target_df["Post content_emoji_count"] = target_df["Post content_emojis"].apply(len)

    # 3. Extract hashtags and count their occurrences
    if post_content_column in target_df.columns:
        target_df["Post content_hashtags"] = target_df[post_content_column].fillna("").apply(extract_hashtags)

    # 4. Extract and classify statistics
    if post_content_column in target_df.columns:
        # Extract statistics for each post
        target_df["Statistics"] = target_df[post_content_column].fillna("").apply(extract_statistics)
        target_df["Uses Statistics"] = target_df["Statistics"].apply(classify_statistics)
        target_df["Numbers"] = target_df["Statistics"].apply(lambda x: ", ".join(x["numbers"]))
        target_df["Percentages"] = target_df["Statistics"].apply(lambda x: ", ".join(x["percentages"]))
        target_df["Statistical Terms"] = target_df["Statistics"].apply(lambda x: ", ".join(x["terms"]))

    # 5. Extract adjectives and count them
    if post_content_column in target_df.columns:
        target_df["Adjectives"] = target_df[post_content_column].fillna("").apply(extract_adjectives)
        target_df["Adjective Count"] = target_df["Adjectives"].apply(len)

    # 6. Extract post ID and timestamp
    if linkedin_url_column in target_df.columns:
        target_df['Post ID'] = target_df[linkedin_url_column].apply(extract_post_id)
        target_df['Post Timestamp'] = target_df[linkedin_url_column].apply(lambda x: LIPostTimestampExtractor.get_date_from_linkedin_activity(x))

    # Save the processed data to a new Excel file
    output_file = "processed_linkedin_data.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        target_df.to_excel(writer, sheet_name=target_sheet, index=False)
    print(f"Processed data has been saved to '{output_file}'")
else:
    print(f"Sheet '{target_sheet}' does not exist in the Excel file.")