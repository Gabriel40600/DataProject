import os
import re
import urllib.parse
from datetime import datetime, timezone

import pandas as pd
import nltk
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from nltk.tokenize import word_tokenize
from nltk import pos_tag

# =============================================================================
# Download necessary NLTK resources (only downloads if not already present)
# =============================================================================
nltk.download("vader_lexicon", quiet=True)
nltk.download("stopwords", quiet=True)
nltk.download("punkt", quiet=True)
nltk.download("averaged_perceptron_tagger", quiet=True)

# =============================================================================
# Precompiled Regex Patterns
# =============================================================================
NUMBER_PATTERN = r"\b(?!19[0-9]{2}|20[0-9]{2})\d+\.?\d*\b"  # Exclude years
PERCENT_PATTERN = r"\d+\.?\d*%"
STAT_TERMS_PATTERN = r"\b(mean|median|average|increase|decrease|percent|growth|std|variance|correlation)\b"
EMOJI_PATTERN = re.compile(
    "[\U0001F600-\U0001F64F"
    "\U0001F300-\U0001F5FF"
    "\U0001F680-\U0001F6FF"
    "\U0001F1E0-\U0001F1FF"
    "\U00002600-\U000027BF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA70-\U0001FAFF"
    "\U00002500-\U00002BEF"
    "]+", flags=re.UNICODE
)
HASHTAG_PATTERN = re.compile(r"#\w+")
QUOTE_PATTERN = re.compile(r'["]([^\"]+)["]')
URL_PATTERN = re.compile(r'https?://\S+|www\.\S+', flags=re.IGNORECASE)


# =============================================================================
# Feature Extraction Functions
# =============================================================================
def contains_question(text: str) -> int:
    return 1 if "?" in text else 0


def extract_link(text: str) -> str:
    match = URL_PATTERN.search(text)
    return match.group(0) if match else ""


def contains_link(text: str) -> int:
    return 1 if URL_PATTERN.search(text) else 0


def contains_quote(text: str) -> int:
    return 1 if QUOTE_PATTERN.search(text) else 0


def contains_hashtag(text: str) -> int:
    return 1 if HASHTAG_PATTERN.search(text) else 0


def contains_emoji(text: str) -> int:
    return 1 if EMOJI_PATTERN.search(text) else 0


def extract_emojis(text: str) -> str:
    return ", ".join(EMOJI_PATTERN.findall(text))


def extract_cta(text: str) -> list:
    cta_phrases = [
        "act now", "apply today", "be sure to", "book now", "buy now", "call today",
        "check out", "click here", "discover", "download now", "find out more",
        "follow this", "get a quote", "join today", "learn more", "order now",
        "register", "save big", "save money", "see more", "shop now", "sign up",
        "start now", "try it today", "visit our", "watch for"
    ]
    text_lower = text.lower()
    found_ctas = [phrase for phrase in cta_phrases if phrase in text_lower]
    return [1 if found_ctas else 0, ", ".join(found_ctas)]


def extract_post_id(url: str) -> str:
    match = re.search(r"activity[-:]?(\d+)", url)
    return match.group(1) if match else None


class LIPostTimestampExtractor:
    @staticmethod
    def format_iso_timestamp(timestamp_s: float) -> str:
        date = datetime.fromtimestamp(timestamp_s, tz=timezone.utc)
        return date.strftime('%Y-%m-%d %H:%M:%S')

    @staticmethod
    def get_unix_timestamp(timestamp_s: float) -> int:
        return int(timestamp_s)

    @classmethod
    def get_date_from_linkedin_activity(cls, post_url: str) -> tuple:
        try:
            match = re.search(r'activity-(\d+)', post_url)
            if not match:
                return 'Invalid LinkedIn ID', None
            linkedin_id = match.group(1)
            first_41_bits = bin(int(linkedin_id))[2:43]
            timestamp_ms = int(first_41_bits, 2)
            timestamp_s = timestamp_ms / 1000
            return cls.format_iso_timestamp(timestamp_s), cls.get_unix_timestamp(timestamp_s)
        except (ValueError, IndexError):
            return 'Date not available', None


def process_linkedin_data(file_path: str, output_file: str):
    df = pd.read_excel(file_path)
    analyzer = SentimentIntensityAnalyzer()
    df["Post content"] = df["Post content"].fillna("")

    df["CTA Present"], df["CTA Found"] = zip(*df["Post content"].apply(extract_cta))
    df["Contains Hashtag"] = df["Post content"].apply(contains_hashtag)
    df["Contains Emoji"] = df["Post content"].apply(contains_emoji)
    df["Extracted Emojis"] = df["Post content"].apply(extract_emojis)
    df["Contains Question"] = df["Post content"].apply(contains_question)
    df["Contains Link"] = df["Post content"].apply(contains_link)
    df["Extracted Link"] = df["Post content"].apply(extract_link)
    df["Contains Quote"] = df["Post content"].apply(contains_quote)
    df["Post ID"] = df["Post URL"].apply(extract_post_id)
    df[["Post Timestamp (ISO)", "Post Timestamp (Unix)"]] = df["Post URL"].apply(
        lambda x: pd.Series(LIPostTimestampExtractor.get_date_from_linkedin_activity(x)))

    df.to_excel(output_file, index=False)
    print(f"Processed data saved to {output_file}")


if __name__ == "__main__":
    process_linkedin_data("data_set_cleaned.xlsx", "processed_linkedin_data.xlsx")
