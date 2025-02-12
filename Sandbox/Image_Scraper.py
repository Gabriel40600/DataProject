import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import requests
import os
import time
from selenium import webdriver
# Set up Selenium WebDriver
def setup_driver():
    options = Options()
    options.add_argument("--headless")  # Run in headless mode
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# Scroll to bottom for lazy-loaded content
def scroll_to_bottom(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.1)  # Allow time for lazy loading


# Extract the main post-specific image
def extract_post_specific_image(driver, post_url):
    try:
        driver.get(post_url)
        scroll_to_bottom(driver)

        # Locate all images broadly
        image_elements = driver.find_elements(By.XPATH, "//img[contains(@src, 'media.licdn.com')]")
        print(f"Found {len(image_elements)} images on the page.")

        # Filter images containing 'feedshare' in their URLs
        feedshare_images = []
        for img in image_elements:
            try:
                src = img.get_attribute("src")
                width = img.get_attribute("width")
                height = img.get_attribute("height")

                if src and "feedshare" in src:
                    feedshare_images.append({"src": src, "width": int(width or 0), "height": int(height or 0)})
            except Exception as e:
                print(f"Error processing an image: {e}")

        if not feedshare_images:
            print(f"No 'feedshare' images found for post: {post_url}")
            return None

        # Select the largest 'feedshare' image
        largest_image = max(feedshare_images, key=lambda x: x["width"] * x["height"])
        print(f"Selected image: {largest_image['src']} (Area: {largest_image['width'] * largest_image['height']})")
        return largest_image["src"]
    except Exception as e:
        print(f"Error extracting image from post {post_url}: {e}")
        return None


# Download the image
def download_image(url, folder="images", filename="latest_image"):
    try:
        response = requests.get(url, stream=True, timeout=10)
        if response.status_code == 200:
            # Get file extension
            content_type = response.headers.get("Content-Type")
            ext = content_type.split("/")[-1] if content_type else "jpg"

            # Save the image
            if not os.path.exists(folder):
                os.makedirs(folder)
            filepath = os.path.join(folder, f"{filename}.{ext}")
            with open(filepath, "wb") as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            print(f"Downloaded: {filepath}")
        else:
            print(f"Failed to download {url} (HTTP {response.status_code})")
    except Exception as e:
        print(f"Error downloading image: {e}")


# Process multiple LinkedIn URLs
def process_multiple_links(file_path, sheet_name, url_column, indices):
    # Load the dataset
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    if sheet_name not in all_sheets:
        print(f"Sheet '{sheet_name}' not found.")
        return

    # Select specific rows
    data = all_sheets[sheet_name]
    urls = data.loc[indices, url_column].dropna()

    # Initialize WebDriver
    driver = setup_driver()

    for idx, post_url in zip(indices, urls):
        print(f"\nProcessing post {idx}: {post_url}")
        feedshare_image_url = extract_post_specific_image(driver, post_url)
        if feedshare_image_url:
            download_image(feedshare_image_url, filename=f"post_{idx}")
        else:
            print(f"No valid images found for post {idx}")

    driver.quit()


# Parameters for the small test sample
file_path = "../0_data_set.xlsx"  # Replace with your file path
sheet_name = "Technology & Innovation"  # Replace with your sheet name
url_column = "Post URL"  # Replace with your column name
indices_to_test = [1, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]  # Specific rows to process

# Run the broader test
process_multiple_links(file_path, sheet_name, url_column, indices_to_test)
