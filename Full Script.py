import pandas as pd
import matplotlib.pyplot as plt

# Load the processed Excel file
file_path = "processed_linkedin_data.xlsx"  # Replace with your file path
sheet_name = "Technology & Innovation"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Filter rows where Post Timestamp is available and valid
df = df[df['Post Timestamp'].str.contains(r'GMT \(UTC\)', na=False)]

# Convert Post Timestamp to datetime format
df['Post Timestamp'] = pd.to_datetime(df['Post Timestamp'], errors='coerce', format='%a, %d %b %Y %H:%M:%S GMT (UTC)')

# Drop rows where the conversion failed
df = df.dropna(subset=['Post Timestamp'])

# Total number of valid timestamps
total_timestamps = len(df)

# Extract hour from the timestamp
df['Post Hour'] = df['Post Timestamp'].dt.hour

# Group by hour to count the number of posts per hour
hourly_counts = df['Post Hour'].value_counts().sort_index()

# Plot the hourly posting activity and save as jpg
plt.figure(figsize=(10, 6))
plt.plot(hourly_counts.index, hourly_counts.values, marker='o')
plt.xlabel('Hour of the Day (UTC)')
plt.ylabel('Number of Posts')
plt.title('LinkedIn Posting Activity by Hour of the Day')
plt.xticks(range(0, 24))
plt.grid(True)
plt.text(0.95, 0.01, f'Total Timestamps Analyzed: {total_timestamps}', verticalalignment='bottom', horizontalalignment='right',
         transform=plt.gca().transAxes, color='gray', fontsize=10)
plt.savefig('hourly_posting_activity.jpg')
plt.show()

# Group by day of the week to count the number of posts per day
df['Post Day'] = df['Post Timestamp'].dt.day_name()
daily_counts = df['Post Day'].value_counts()[['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']]

# Plot the daily posting activity and save as jpg
plt.figure(figsize=(10, 6))
plt.bar(daily_counts.index, daily_counts.values, color='skyblue')
plt.xlabel('Day of the Week')
plt.ylabel('Number of Posts')
plt.title('LinkedIn Posting Activity by Day of the Week')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.text(0.95, 0.01, f'Total Timestamps Analyzed: {total_timestamps}', verticalalignment='bottom', horizontalalignment='right',
         transform=plt.gca().transAxes, color='gray', fontsize=10)
plt.savefig('daily_posting_activity.jpg')
plt.show()