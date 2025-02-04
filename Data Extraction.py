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
# Precompiled Regex Patterns for Performance
# =============================================================================
NUMBER_PATTERN = r"\b(?!19[0-9]{2}|20[0-9]{2})\d+\.?\d*\b"  # Exclude years
PERCENT_PATTERN = r"\d+\.?\d*%"
STAT_TERMS_PATTERN = r"\b(mean|median|average|increase|decrease|percent|growth|std|variance|correlation)\b"
EMOJI_PATTERN = re.compile(
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
HASHTAG_PATTERN = re.compile(r"#\w+")


# =============================================================================
# Helper Functions for Data Extraction and Analysis
# =============================================================================

def extract_statistics(text: str) -> dict:
    """
    Extracts statistical information from the text, including:
        - Numbers (excluding likely years)
        - Percentages
        - Key statistical terms (e.g., mean, median, growth, etc.)
    Returns a dictionary containing the extracted elements.
    """
    percentages = re.findall(PERCENT_PATTERN, text)
    numbers = re.findall(NUMBER_PATTERN, text)
    stat_terms = re.findall(STAT_TERMS_PATTERN, text, flags=re.IGNORECASE)

    # Only consider numbers if there is a statistical context (terms or percentages)
    if not stat_terms and not percentages:
        numbers = []

    return {
        "numbers": numbers,
        "percentages": percentages,
        "terms": stat_terms
    }


def classify_statistics(stat_data: dict) -> str:
    """
    Classifies the statistical usage in text based on the extracted data.
    Returns:
        - "YES" if statistical terms or percentages are found.
        - "MAYBE" if only numbers are found.
        - "NO" otherwise.
    """
    if stat_data["terms"] or stat_data["percentages"]:
        return "YES"
    elif stat_data["numbers"]:
        return "MAYBE"
    else:
        return "NO"


def extract_emojis(text: str) -> list:
    """
    Extracts all emojis present in the text.
    """
    return EMOJI_PATTERN.findall(text)


def extract_hashtags(text: str) -> list:
    """
    Extracts all hashtags from the text.
    """
    return HASHTAG_PATTERN.findall(text)


def contains_hashtag(text: str) -> bool:
    """
    Checks if the text contains any hashtags.
    Returns True if at least one hashtag is found, otherwise False.
    """
    return bool(extract_hashtags(text))


def extract_adjectives(text: str) -> list:
    """
    Extracts adjectives from the text using NLTK's POS tagger.
    Only tokens that are fully alphabetic are considered to avoid including emojis or other symbols.
    """
    tokens = word_tokenize(text)
    pos_tags = pos_tag(tokens)
    adjectives = [word for word, pos in pos_tags if pos in {"JJ", "JJR", "JJS"} and word.isalpha()]
    return adjectives


def extract_post_id(url: str) -> str:
    """
    Extracts the LinkedIn post ID from the URL.
    """
    match = re.search(r"activity[-:]?(\d+)", url)
    return match.group(1) if match else None


class LIPostTimestampExtractor:
    """
    Class to extract and format a timestamp from a LinkedIn post URL.
    """

    @staticmethod
    def format_timestamp(timestamp_s: float, get_local: bool = False) -> str:
        """
        Formats the timestamp (in seconds) to a human-readable date string.
        """
        if get_local:
            date = datetime.fromtimestamp(timestamp_s)
            return date.strftime('%a, %d %b %Y %H:%M:%S GMT')
        else:
            date = datetime.fromtimestamp(timestamp_s, tz=timezone.utc)
            return date.strftime('%a, %d %b %Y %H:%M:%S GMT (UTC)')

    @classmethod
    def get_date_from_linkedin_activity(cls, post_url: str, get_local: bool = False) -> str:
        """
        Extracts the timestamp from a LinkedIn activity URL.
        Returns a formatted date string or an error message if extraction fails.
        """
        try:
            post_url_unquote = urllib.parse.unquote(post_url)
            match = re.search(r'activity-(\d+)', post_url_unquote)
            if not match:
                return 'Invalid LinkedIn ID'
            linkedin_id = match.group(1)

            # Extract the first 41 bits from the binary representation
            first_41_bits = bin(int(linkedin_id))[2:43]
            timestamp_ms = int(first_41_bits, 2)
            timestamp_s = timestamp_ms / 1000
            return cls.format_timestamp(timestamp_s, get_local)
        except (ValueError, IndexError):
            return 'Date not available'


def contains_question(text: str) -> bool:
    """
    Checks if the text contains a question mark.
    Returns True if found, False otherwise.
    """
    return "?" in text


def get_post_length(text: str) -> int:
    """
    Returns the length of the post (number of characters).
    """
    return len(text)


def contains_link(text: str) -> bool:
    """
    Checks if the text contains any hyperlinks.
    Uses a regex pattern to search for common URL formats like http://, https://, or www.
    """
    url_pattern = re.compile(r'https?://\S+|www\.\S+', flags=re.IGNORECASE)
    return bool(url_pattern.search(text))


def contains_quote(text: str) -> bool:
    """
    Checks if the text contains quoted content.
    Searches for text enclosed in standard double quotes (") or smart quotes (“ and ”).
    """
    pattern = re.compile(r'["“](.*?)["”]')
    return bool(pattern.search(text))


def extract_personal_pronouns(text: str) -> list:
    """
    Extracts personal pronouns from the text using NLTK's POS tagger.
    Personal pronouns (PRP) and possessive pronouns (PRP$) are identified.
    Returns a list of personal pronouns found in the text.
    """
    tokens = word_tokenize(text)
    pos_tags = pos_tag(tokens)
    return [word for word, tag in pos_tags if tag in {"PRP", "PRP$"}]


# =============================================================================
# Main Data Processing Function
# =============================================================================

def process_linkedin_data(file_path: str,
                          target_sheet: str = "Technology & Innovation",
                          post_content_column: str = "Post content",
                          linkedin_url_column: str = "Post URL",
                          output_file: str = "processed_linkedin_data.xlsx"):
    """
    Processes LinkedIn data from an Excel file by applying several transformations:

    Step-by-Step Explanation:
    # 1. Load Data: Read the specified sheet from the Excel file.
    # 2. Sentiment Analysis: Compute sentiment scores (positive, neutral, negative, compound)
    #    using NLTK's SentimentIntensityAnalyzer.
    # 3. Emoji Extraction: Extract all emojis from the post content.
    # 4. Hashtag Extraction: Extract all hashtags from the post content.
    # 5. Statistical Data Extraction and Classification:
    #    - Extract numbers, percentages, and key statistical terms.
    #    - Classify the statistical usage in the text.
    # 6. Adjective Extraction: Extract adjectives from the post content.
    # 7. Question Mark Detection: Check if the post contains a question mark.
    # 8. Post Length Calculation: Determine the length of the post in characters.
    # 9. Link Presence Detection: Identify if the post contains hyperlinks.
    # 10. Quote Detection: Identify if the post contains any quoted text.
    # 11. Hashtag Presence Detection: Identify if the post contains hashtags.
    # 12. Personal Pronoun Extraction: Identify personal pronouns and count them.
    # 13. Post ID and Timestamp Extraction: If available, extract the LinkedIn post ID and
    #     convert the timestamp to a human-readable format.
    # 14. Save the Processed Data: Write the transformed DataFrame to a new Excel file.
    """
    # 1. Load Data
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file '{file_path}' not found.")

    all_sheets = pd.read_excel(file_path, sheet_name=None)
    if target_sheet not in all_sheets:
        raise ValueError(f"Sheet '{target_sheet}' does not exist in the Excel file.")

    df = all_sheets[target_sheet]

    # Validate required columns
    if post_content_column not in df.columns:
        raise ValueError(f"Column '{post_content_column}' not found in sheet '{target_sheet}'.")
    if linkedin_url_column not in df.columns:
        print(
            f"Warning: Column '{linkedin_url_column}' not found in sheet '{target_sheet}'. Some features will be skipped.")

    # Initialize NLTK Sentiment Analyzer
    analyzer = SentimentIntensityAnalyzer()

    # Ensure post content is a string (fill missing with empty strings)
    df[post_content_column] = df[post_content_column].fillna("")

    # 2. Sentiment Analysis
    df["Sentiment_Positive"] = df[post_content_column].apply(lambda x: analyzer.polarity_scores(x)["pos"])
    df["Sentiment_Neutral"] = df[post_content_column].apply(lambda x: analyzer.polarity_scores(x)["neu"])
    df["Sentiment_Negative"] = df[post_content_column].apply(lambda x: analyzer.polarity_scores(x)["neg"])
    df["Sentiment_Compound"] = df[post_content_column].apply(lambda x: analyzer.polarity_scores(x)["compound"])

    # 3. Emoji Extraction
    df["Post content_emojis"] = df[post_content_column].apply(extract_emojis)
    df["Post content_emoji_count"] = df["Post content_emojis"].apply(len)

    # 4. Hashtag Extraction
    df["Post content_hashtags"] = df[post_content_column].apply(extract_hashtags)

    # 5. Statistical Data Extraction and Classification
    df["Statistics"] = df[post_content_column].apply(extract_statistics)
    df["Uses Statistics"] = df["Statistics"].apply(classify_statistics)
    df["Numbers"] = df["Statistics"].apply(lambda x: ", ".join(x["numbers"]))
    df["Percentages"] = df["Statistics"].apply(lambda x: ", ".join(x["percentages"]))
    df["Statistical Terms"] = df["Statistics"].apply(lambda x: ", ".join(x["terms"]))

    # 6. Adjective Extraction
    df["Adjectives"] = df[post_content_column].apply(extract_adjectives)
    df["Adjective Count"] = df["Adjectives"].apply(len)

    # 7. Question Mark Detection
    df["Contains Question"] = df[post_content_column].apply(contains_question)

    # 8. Post Length Calculation
    df["Post Length"] = df[post_content_column].apply(get_post_length)

    # 9. Link Presence Detection
    df["Contains Link"] = df[post_content_column].apply(contains_link)

    # 10. Quote Detection
    df["Contains Quote"] = df[post_content_column].apply(contains_quote)

    # 11. Hashtag Presence Detection
    df["Contains Hashtag"] = df[post_content_column].apply(contains_hashtag)

    # 12. Personal Pronoun Extraction
    df["Personal Pronouns"] = df[post_content_column].apply(extract_personal_pronouns)
    df["Personal Pronoun Count"] = df["Personal Pronouns"].apply(len)

    # 13. Post ID and Timestamp Extraction (if URL column is available)
    if linkedin_url_column in df.columns:
        df['Post ID'] = df[linkedin_url_column].apply(lambda url: extract_post_id(str(url)))
        df['Post Timestamp'] = df[linkedin_url_column].apply(
            lambda url: LIPostTimestampExtractor.get_date_from_linkedin_activity(str(url))
        )

    # 14. Save the Processed Data to a New Excel File
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=target_sheet, index=False)
    print(f"Processed data has been saved to '{output_file}'")


# =============================================================================
# Main Execution
# =============================================================================

if __name__ == "__main__":
    # Define input and output file paths
    input_file = "data_set.xlsx"  # Replace with the path to your Excel file
    output_file = "processed_linkedin_data.xlsx"

    try:
        process_linkedin_data(
            file_path=input_file,
            target_sheet="Technology & Innovation",
            post_content_column="Post content",
            linkedin_url_column="Post URL",
            output_file=output_file
        )
    except Exception as e:
        print(f"Error processing data: {e}")