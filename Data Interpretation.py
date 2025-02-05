import os
import logging
import pandas as pd
import matplotlib.pyplot as plt


def load_data(file_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Loads the Excel file and returns the DataFrame from the specified sheet.

    Parameters:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to load.

    Returns:
        pd.DataFrame: The loaded DataFrame.

    Raises:
        FileNotFoundError: If the specified file does not exist.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    return pd.read_excel(file_path, sheet_name=sheet_name)


def filter_and_convert_timestamps(df: pd.DataFrame, timestamp_col: str = 'Post Timestamp') -> pd.DataFrame:
    """
    Filters rows with valid timestamps and converts them to datetime objects.

    Parameters:
        df (pd.DataFrame): DataFrame containing raw data.
        timestamp_col (str): Name of the column with timestamps.

    Returns:
        pd.DataFrame: DataFrame with converted and filtered timestamps.
    """
    # Convert values to strings and filter rows containing "GMT (UTC)"
    df_filtered = df[df[timestamp_col].astype(str).str.contains(r'GMT \(UTC\)', na=False)].copy()

    # Convert timestamps to datetime objects
    df_filtered[timestamp_col] = pd.to_datetime(
        df_filtered[timestamp_col],
        errors='coerce',
        format='%a, %d %b %Y %H:%M:%S GMT (UTC)'
    )

    # Drop rows where conversion failed
    return df_filtered.dropna(subset=[timestamp_col])


def plot_hourly_activity(df: pd.DataFrame, timestamp_col: str = 'Post Timestamp',
                         output_file: str = 'hourly_posting_activity.jpg'):
    """
    Creates and saves a line plot showing the hourly posting activity.

    Parameters:
        df (pd.DataFrame): DataFrame with converted timestamps.
        timestamp_col (str): Name of the column with datetime objects.
        output_file (str): Filename for the saved plot.
    """
    # Extract hour from the timestamps
    df['Post Hour'] = df[timestamp_col].dt.hour
    hourly_counts = df['Post Hour'].value_counts().sort_index()

    plt.figure(figsize=(10, 6))
    plt.plot(hourly_counts.index, hourly_counts.values, marker='o', linestyle='-')
    plt.xlabel('Hour of the Day (UTC)')
    plt.ylabel('Number of Posts')
    plt.title('LinkedIn Posting Activity by Hour of the Day')
    plt.xticks(range(0, 24))
    plt.grid(True)
    total_timestamps = len(df)
    plt.text(0.95, 0.01, f'Total Timestamps Analyzed: {total_timestamps}',
             verticalalignment='bottom', horizontalalignment='right',
             transform=plt.gca().transAxes, color='gray', fontsize=10)
    plt.tight_layout()
    plt.savefig(output_file, dpi=300)
    plt.close()


def plot_daily_activity(df: pd.DataFrame, timestamp_col: str = 'Post Timestamp',
                        output_file: str = 'daily_posting_activity.jpg'):
    """
    Creates and saves a bar plot showing posting activity by day of the week.

    Parameters:
        df (pd.DataFrame): DataFrame with converted timestamps.
        timestamp_col (str): Name of the column with datetime objects.
        output_file (str): Filename for the saved plot.
    """
    # Extract day of the week from the timestamps
    df['Post Day'] = df[timestamp_col].dt.day_name()
    # Define the order of days
    days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    daily_counts = df['Post Day'].value_counts().reindex(days_order, fill_value=0)

    plt.figure(figsize=(10, 6))
    plt.bar(daily_counts.index, daily_counts.values, color='skyblue')
    plt.xlabel('Day of the Week')
    plt.ylabel('Number of Posts')
    plt.title('LinkedIn Posting Activity by Day of the Week')
    plt.xticks(rotation=45)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    total_timestamps = len(df)
    plt.text(0.95, 0.01, f'Total Timestamps Analyzed: {total_timestamps}',
             verticalalignment='bottom', horizontalalignment='right',
             transform=plt.gca().transAxes, color='gray', fontsize=10)
    plt.tight_layout()
    plt.savefig(output_file, dpi=300)
    plt.close()


def main():
    """
    Executes the data processing and plotting workflow.

    Process:
      1. Configures logging.
      2. Loads the Excel file and the specified sheet.
      3. Filters and converts the timestamps.
      4. Creates plots for hourly and daily posting activity.
      5. Saves the plots as JPEG files.
    """
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    file_path = "processed_linkedin_data.xlsx"  # Path to the processed Excel file
    sheet_name = ("Sheet1")
    try:
        # Step 1: Load Excel data
        df = load_data(file_path, sheet_name)
        logging.info("Data loaded successfully.")

        # Step 2: Filter and convert timestamps
        df = filter_and_convert_timestamps(df, 'Post Timestamp')
        total_timestamps = len(df)
        logging.info(f"Number of valid timestamps: {total_timestamps}")

        # Step 3: Plot hourly activity
        plot_hourly_activity(df, 'Post Timestamp', 'hourly_posting_activity.jpg')
        logging.info("Hourly activity plot saved as 'hourly_posting_activity.jpg'.")

        # Step 4: Plot daily activity
        plot_daily_activity(df, 'Post Timestamp', 'daily_posting_activity.jpg')
        logging.info("Daily activity plot saved as 'daily_posting_activity.jpg'.")

    except Exception as e:
        logging.error(f"Error processing data: {e}")


if __name__ == "__main__":
    main()
