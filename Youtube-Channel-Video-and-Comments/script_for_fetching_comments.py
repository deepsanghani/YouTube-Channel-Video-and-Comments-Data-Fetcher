import googleapiclient.discovery
import pandas as pd
from datetime import datetime
import pytz
import re

API_KEY = 'API_KEY'

def get_youtube_client():
    youtube = googleapiclient.discovery.build("youtube", "v3", developerKey=API_KEY)
    return youtube

def convert_utc_to_ist(utc_date_str):
    utc_date = datetime.strptime(utc_date_str, "%Y-%m-%dT%H:%M:%SZ")
    utc_date = pytz.utc.localize(utc_date)
    ist_timezone = pytz.timezone('Asia/Kolkata')
    ist_date = utc_date.astimezone(ist_timezone)
    return ist_date.strftime("%d-%m-%Y %H:%M:%S")

def convert_duration(iso_duration):
    duration = iso_duration[2:]
    time_parts = re.findall(r'(\d+)([HMS])', duration)
    time_dict = {'H': 0, 'M': 0, 'S': 0}
    for value, unit in time_parts:
        time_dict[unit] = int(value)
    duration_str = ""
    
    if time_dict['H'] > 0:
        duration_str += f"{time_dict['H']} hours "
    if time_dict['M'] > 0:
        duration_str += f"{time_dict['M']} minutes "
    if time_dict['S'] > 0 or (time_dict['H'] == 0 and time_dict['M'] == 0):
        duration_str += f"{time_dict['S']} seconds"
    return duration_str.strip()

def get_channel_id(channel_name):
    youtube = get_youtube_client()
    try:
        request = youtube.search().list(
            part="snippet",
            q=channel_name, 
            type="channel",
            maxResults=1
        )
        response = request.execute()
        if 'items' in response:
            return response['items'][0]['snippet']['channelId']
        else:
            raise ValueError(f"Channel not found with the name: {channel_name}")
    except Exception as e:
        print(f"Error fetching channel ID: {e}")
        return None

def get_videos(channel_id):
    youtube = get_youtube_client()
    videos = []
    try:
        request = youtube.search().list(
            part="snippet",
            channelId=channel_id,
            maxResults=50,
            type="video"
        )
        response = request.execute()
        videos = response['items']
    except Exception as e:
        print(f"Error fetching videos: {e}")
    return videos

def get_video_details(video_id):
    youtube = get_youtube_client()
    try:
        request = youtube.videos().list(
            part="snippet,statistics,contentDetails",
            id=video_id
        )
        response = request.execute()
        if 'items' in response and len(response['items']) > 0:
            video_details = response['items'][0]
            video_data = {
                'video_id': video_id,
                'title': video_details['snippet']['title'],
                'description': video_details['snippet'].get('description', 'N/A'),
                'published_date': convert_utc_to_ist(video_details['snippet']['publishedAt']),
                'view_count': video_details['statistics'].get('viewCount', 'N/A'),
                'like_count': video_details['statistics'].get('likeCount', 'N/A'),
                'comment_count': video_details['statistics'].get('commentCount', 'N/A'),
                'duration': convert_duration(video_details['contentDetails'].get('duration', 'N/A')),
                'thumbnail_url': video_details['snippet']['thumbnails']['high'].get('url', 'N/A')
            }
            return video_data
        else:
            return None
    except Exception as e:
        print(f"Error fetching video details for {video_id}: {e}")
        return None

def get_comments(video_id):
    youtube = get_youtube_client()
    comments = []
    try:
        request = youtube.commentThreads().list(
            part="snippet,replies",
            videoId=video_id,
            textFormat="plainText",
            maxResults=50  # Fetching 50 comments per API call
        )
        comment_counter = 0  # Initialize the counter to track the number of comments fetched

        while request and comment_counter < 100:
            response = request.execute()
            for item in response['items']:
                if comment_counter >= 100:
                    break
                comment = item['snippet']['topLevelComment']['snippet']
                comment_data = {
                    'video_id': video_id,
                    'comment_id': item['id'],
                    'comment_text': comment['textDisplay'],
                    'author_name': comment['authorDisplayName'],
                    'published_date': convert_utc_to_ist(comment['publishedAt']),
                    'like_count': comment['likeCount'],
                    'reply_to': None
                }
                comments.append(comment_data)
                comment_counter += 1
                if 'replies' in item:
                    for reply_item in item['replies']['comments']:
                        if comment_counter >= 100:
                            break
                        reply = reply_item['snippet']
                        reply_data = {
                            'video_id': video_id,
                            'comment_id': reply_item['id'],
                            'comment_text': reply['textDisplay'],
                            'author_name': reply['authorDisplayName'],
                            'published_date': convert_utc_to_ist(reply['publishedAt']),
                            'like_count': reply['likeCount'],
                            'reply_to': comment['authorDisplayName']
                        }
                        comments.append(reply_data)
                        comment_counter += 1

            if comment_counter >= 100:
                break
            request = youtube.commentThreads().list_next(request, response)
    except Exception as e:
        print(f"Error fetching comments for video {video_id}: {e}")
    return comments

def save_to_excel(video_data, comment_data, output_path):
    try:
        video_df = pd.DataFrame(video_data)
        comment_df = pd.DataFrame(comment_data)
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            video_df.to_excel(writer, sheet_name='Video Data', index=False)
            comment_df.to_excel(writer, sheet_name='Comments Data', index=False)
            video_worksheet = writer.sheets['Video Data']
            comment_worksheet = writer.sheets['Comments Data']
            for col_num, column in enumerate(video_df.columns.values):
                video_worksheet.set_column(col_num, col_num, 40)
            
            for col_num, column in enumerate(comment_df.columns.values):
                comment_worksheet.set_column(col_num, col_num,40) 
        print(f"Data saved successfully to {output_path}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

channel_name = input("Enter YouTube Channel Name: ")
try:
    print("Fetching channel data...")
    channel_id = get_channel_id(channel_name)
    if not channel_id:
        print("Channel not found, exiting.")
        exit()
    print(f"Channel ID: {channel_id}")
    print("Fetching video data...")
    videos = get_videos(channel_id)
    video_data = []
    comment_data = []
    for video in videos:
        print(f"Processing video: {video['snippet']['title']}")
        try:
            video_id = video['id']['videoId']
            video_details = get_video_details(video_id)
            if video_details:
                video_data.append(video_details)
            comments = get_comments(video_id)
            comment_data.extend(comments)
        except Exception as e:
            print(f"Error processing video {video['id']['videoId']}: {e}")
    
    output_path = 'C:\\Users\\Deep Sanghani\\Desktop\\Youtube-Channel-Video-and-Comments\\YouTube_Data.xlsx'
    print("Saving to Excel...")
    save_to_excel(video_data, comment_data, output_path)
except Exception as e:
    print(f"An error occurred: {str(e)}")
