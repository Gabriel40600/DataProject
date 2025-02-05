import pandas as pd


# Load the Excel file
def analyze_linkedin_data(file_path, output_file):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Select only numeric columns
    numeric_columns = ["reactions", "comments", "shares", "Sentiment_Positive", "Sentiment_Neutral",
                       "Sentiment_Negative", "Sentiment_Compound", "Post Length", "Emoji Count"]

    df_numeric = df[numeric_columns]

    # Generate descriptive statistics
    descriptive_stats = df_numeric.describe().transpose()

    # Calculate percentage of posts containing emojis, hashtags, or links
    emoji_percentage = (df["Contains Emoji"].sum() / len(df)) * 100
    hashtag_percentage = (df["Contains Hashtag"].sum() / len(df)) * 100
    link_percentage = (df["Contains Link"].sum() / len(df)) * 100

    # Calculate engagement separately for reactions, comments, and shares
    link_reactions = df.groupby("Contains Link")["reactions"].mean()
    link_comments = df.groupby("Contains Link")["comments"].mean()
    link_shares = df.groupby("Contains Link")["shares"].mean()

    emoji_reactions = df.groupby("Contains Emoji")["reactions"].mean()
    emoji_comments = df.groupby("Contains Emoji")["comments"].mean()
    emoji_shares = df.groupby("Contains Emoji")["shares"].mean()

    hashtag_reactions = df.groupby("Contains Hashtag")["reactions"].mean()
    hashtag_comments = df.groupby("Contains Hashtag")["comments"].mean()
    hashtag_shares = df.groupby("Contains Hashtag")["shares"].mean()

    # Save results to Excel
    with pd.ExcelWriter(output_file) as writer:
        descriptive_stats.to_excel(writer, sheet_name="Descriptive Stats")
        pd.DataFrame({"Metric": ["Emoji %", "Hashtag %", "Link %"],
                      "Value": [emoji_percentage, hashtag_percentage, link_percentage]}).to_excel(writer,
                                                                                                  sheet_name="Post Analysis",
                                                                                                  index=False)
        link_reactions.to_excel(writer, sheet_name="Link Reactions")
        link_comments.to_excel(writer, sheet_name="Link Comments")
        link_shares.to_excel(writer, sheet_name="Link Shares")
        emoji_reactions.to_excel(writer, sheet_name="Emoji Reactions")
        emoji_comments.to_excel(writer, sheet_name="Emoji Comments")
        emoji_shares.to_excel(writer, sheet_name="Emoji Shares")
        hashtag_reactions.to_excel(writer, sheet_name="Hashtag Reactions")
        hashtag_comments.to_excel(writer, sheet_name="Hashtag Comments")
        hashtag_shares.to_excel(writer, sheet_name="Hashtag Shares")

    print(f"Processed data saved to {output_file}")


# Provide the file path
file_path = "processed_linkedin_data.xlsx"
output_file = "deskriptive_statistik.xlsx"
analyze_linkedin_data(file_path, output_file)