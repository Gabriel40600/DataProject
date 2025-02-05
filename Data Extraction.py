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
QUOTE_PATTERN = re.compile(r'["“](.*?)["”]')
URL_PATTERN = re.compile(r'https?://\S+|www\.\S+', flags=re.IGNORECASE)


# =============================================================================
# CTA Extraction Function
# =============================================================================
def extract_cta(text: str) -> list:
    """Extracts CTA presence (1/0) and found CTA phrases separately."""
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


# =============================================================================
# Feature Extraction Functions
# =============================================================================
def extract_emojis(text: str) -> list:
    """Extracts all emojis from the text."""
    return EMOJI_PATTERN.findall(text)


def emoji_count(text: str) -> int:
    """Counts the number of emojis in the text."""
    return len(EMOJI_PATTERN.findall(text))


def contains_emoji(text: str) -> bool:
    """Returns True if at least one emoji is present in the text."""
    return bool(EMOJI_PATTERN.search(text))


def extract_hashtags(text: str) -> list:
    """Extracts all hashtags and returns them as a list."""
    return HASHTAG_PATTERN.findall(text)


def contains_hashtag(text: str) -> bool:
    """Checks if there is at least one hashtag in the text."""
    return bool(HASHTAG_PATTERN.search(text))


def extract_statistics(text: str) -> dict:
    """Extracts numbers, percentages, and statistical terms from the text."""
    percentages = re.findall(PERCENT_PATTERN, text)
    numbers = re.findall(NUMBER_PATTERN, text)
    stat_terms = re.findall(STAT_TERMS_PATTERN, text, flags=re.IGNORECASE)
    return {"numbers": numbers, "percentages": percentages, "terms": stat_terms}


def contains_question(text: str) -> bool:
    return "?" in text


def contains_link(text: str) -> bool:
    return bool(URL_PATTERN.search(text))


def contains_quote(text: str) -> bool:
    return bool(QUOTE_PATTERN.search(text))


def extract_personal_pronouns(text: str) -> list:
    tokens = word_tokenize(text)
    pos_tags = pos_tag(tokens)
    return [word for word, tag in pos_tags if tag in {"PRP", "PRP$"}]


def get_post_length(text: str) -> int:
    """Returns the length of the post (number of characters)."""
    return len(text)


def extract_post_id(url: str) -> str:
    """Extracts the LinkedIn post ID from the URL."""
    match = re.search(r"activity[-:]?(\d+)", url)
    return match.group(1) if match else None


class LIPostTimestampExtractor:
    """Extracts and formats a timestamp from a LinkedIn post URL."""

    @staticmethod
    def format_timestamp(timestamp_s: float) -> str:
        """Formats the timestamp (in seconds) to a human-readable date string."""
        date = datetime.fromtimestamp(timestamp_s, tz=timezone.utc)
        return date.strftime('%a, %d %b %Y %H:%M:%S GMT (UTC)')

    @classmethod
    def get_date_from_linkedin_activity(cls, post_url: str) -> str:
        """Extracts the timestamp from a LinkedIn activity URL."""
        try:
            match = re.search(r'activity-(\d+)', post_url)
            if not match:
                return 'Invalid LinkedIn ID'
            linkedin_id = match.group(1)

            first_41_bits = bin(int(linkedin_id))[2:43]
            timestamp_ms = int(first_41_bits, 2)
            timestamp_s = timestamp_ms / 1000
            return cls.format_timestamp(timestamp_s)
        except (ValueError, IndexError):
            return 'Date not available'


# =============================================================================
# Main Processing Function
# =============================================================================
def process_linkedin_data(file_path: str, output_file: str):
    df = pd.read_excel(file_path)

    analyzer = SentimentIntensityAnalyzer()
    df["Post content"] = df["Post content"].fillna("")

    # CTA Extraction
    cta_results = df["Post content"].apply(extract_cta).tolist()
    df["CTA Present"] = [item[0] for item in cta_results]
    df["CTA Found"] = [item[1] for item in cta_results]

    # Sentiment Analysis
    df["Sentiment_Positive"] = df["Post content"].apply(lambda x: analyzer.polarity_scores(x)["pos"])
    df["Sentiment_Neutral"] = df["Post content"].apply(lambda x: analyzer.polarity_scores(x)["neu"])
    df["Sentiment_Negative"] = df["Post content"].apply(lambda x: analyzer.polarity_scores(x)["neg"])
    df["Sentiment_Compound"] = df["Post content"].apply(lambda x: analyzer.polarity_scores(x)["compound"])

    # Feature Extraction
    df["Post Length"] = df["Post content"].apply(get_post_length)
    df["Hashtags"] = df["Post content"].apply(extract_hashtags)
    df["Contains Hashtag"] = df["Post content"].apply(contains_hashtag)
    df["Emojis"] = df["Post content"].apply(extract_emojis)
    df["Emoji Count"] = df["Post content"].apply(emoji_count)  # NEW COLUMN: Number of emojis
    df["Contains Emoji"] = df["Post content"].apply(contains_emoji)
    df["Contains Question"] = df["Post content"].apply(contains_question)
    df["Contains Link"] = df["Post content"].apply(contains_link)
    df["Contains Quote"] = df["Post content"].apply(contains_quote)
    df["Personal Pronouns"] = df["Post content"].apply(extract_personal_pronouns)

    # Extract Post ID & Timestamp
    df["Post ID"] = df["Post URL"].apply(extract_post_id)
    df["Post Timestamp"] = df["Post URL"].apply(LIPostTimestampExtractor.get_date_from_linkedin_activity)

    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Processed data saved to {output_file}")


# =============================================================================
# Main Execution
# =============================================================================
if __name__ == "__main__":
    process_linkedin_data("data_set_cleaned.xlsx", "processed_linkedin_data.xlsx")
