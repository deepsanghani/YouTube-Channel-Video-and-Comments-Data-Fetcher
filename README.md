# YouTube Channel Data Fetcher

This project fetches data from a YouTube channel, including video details and comments, using the YouTube Data API. The data is processed and saved into an Excel file with separate sheets for video details and comments.

## Features
- Fetch channel details by channel name.
- Retrieve video details such as:
  - Title
  - Description
  - Published date (in IST)
  - Views
  - Likes
  - Comments count
  - Video duration
- Extract up to 100 comments (including replies) per video.
- Save data to an Excel file with formatted column widths.

## Prerequisites
- Python 3.7 or higher
- YouTube Data API v3 enabled on Google Cloud Console
- API Key for YouTube Data API

## Installation
1. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Run the Script
1. Open the terminal and run the following command:
    ```bash
    python script_for_fetching_comments.py
    ```
2. Enter the YouTube channel name when prompted.
3. The script will fetch the data and save it to `YouTube_Data.xlsx`.

## Output Example
### Video Data
The `Video Data` sheet in the Excel file will look like this:

| video_id      | title                    | description       | published_date     | view_count | like_count | comment_count | duration      | thumbnail_url       |
|---------------|--------------------------|-------------------|--------------------|------------|------------|---------------|---------------|---------------------|
| abcd1234efgh  | Sample Video Title 1     | This is a sample. | 21-11-2024 10:30:00| 12345      | 678        | 45            | 10 minutes    | http://thumbnail1   |
| ijkl5678mnop  | Sample Video Title 2     | Another sample.   | 20-11-2024 15:00:00| 56789      | 1234       | 89            | 15 minutes    | http://thumbnail2   |

### Comments Data
The `Comments Data` sheet will look like this:

| video_id      | comment_id  | comment_text             | author_name       | published_date     | like_count | reply_to    |
|---------------|-------------|--------------------------|-------------------|--------------------|------------|-------------|
| abcd1234efgh  | comment123  | Great video!             | User1             | 21-11-2024 11:00:00| 12         | None        |
| abcd1234efgh  | reply123    | Thank you!               | Creator           | 21-11-2024 11:05:00| 5          | User1       |
| ijkl5678mnop  | comment456  | Very informative.        | User2             | 20-11-2024 16:00:00| 20         | None        |

## Notes
- The script fetches up to 50 videos and limits comments to 100 per video.
- Ensure that your API key is active and has sufficient quota to avoid errors.

---
