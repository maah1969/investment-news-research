import os
import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from groq import Groq
from docx import Document
import smtplib
from email.message import EmailMessage
from googleapiclient.discovery import build
from youtube_transcript_api import YouTubeTranscriptApi

# Load environment variables
load_dotenv()

YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

def get_groq_client():
    """Get the Groq API client."""
    if not GROQ_API_KEY:
        print("Error: GROQ_API_KEY environment variable not set.")
        return None
    return Groq(api_key=GROQ_API_KEY)
        
def fetch_top_trending_investment_videos():
    """
    Search YouTube for the top trending English investment videos 
    published in the last 48 hours.
    """
    if not YOUTUBE_API_KEY:
        print("Error: YOUTUBE_API_KEY environment variable not set.")
        return None

    youtube = build('youtube', 'v3', developerKey=YOUTUBE_API_KEY)
    
    # Target videos from the last 48 hours. YouTube API requires RFC 3339 formatted date-time value (1970-01-01T00:00:00Z)
    published_after = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=48)).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
    
    try:
        search_response = youtube.search().list(
            q="investing OR stock market OR finance news",
            part="id,snippet",
            maxResults=50,
            order="viewCount",
            publishedAfter=published_after,
            relevanceLanguage="en",
            type="video"
        ).execute()

        if not search_response.get("items"):
            print("No trending investment videos found.")
            return None
            
        return search_response["items"]
        
    except Exception as e:
        print(f"Error fetching YouTube data: {e}")
        return None

def get_video_transcript(video_id):
    """Retrieve the English transcript for a given YouTube video ID."""
    try:
        # Try to get English transcript
        transcript_list = YouTubeTranscriptApi().list(video_id)
        transcript = transcript_list.find_transcript(['en'])
        fetched_transcript = transcript.fetch()
        
        # Combine text from all transcript chunks
        full_text = " ".join([snippet.text for snippet in fetched_transcript.snippets])
        
        # Limit the transcript length strictly to avoid exceeding Groq Llama 3 API token limits when batching multiple videos
        max_chars = 3000 
        if len(full_text) > max_chars:
            print(f"Transcript too long ({len(full_text)} chars). Truncating to {max_chars} chars.")
            full_text = full_text[:max_chars] + "\n...[Transcript Truncated]..."
            
        return full_text
    except Exception as e:
        print(f"Error fetching transcript for video {video_id}: {e}")
        return None

def generate_top10_report(client, videos_context):
    """Generate a top 10 YouTube video summary report using Groq."""
    if not videos_context:
        return "動画データが取得できなかったため、レポート生成をスキップしました。"
        
    prompt = f"""
あなたは優秀な証券アナリストです。以下の海外の最新投資YouTube動画について、アクセス数上位から情報を取得できたトップ10のトランスクリプト（要約対象）を読み、それぞれ日本語で詳細に要約したレポートを作成してください。

【重要な指示】
1. 各動画の内容を丁寧に拾い上げ、要約の文章量を1動画あたり300〜400文字程度と詳細に膨らませてください。「具体的な数値」「登場した銘柄名」「市場の背景」「配信者の主張の根拠」などをしっかり盛り込み、読み応えのある構成にしてください。
2. 「風が吹けば桶屋が儲かる」のような独自の将来予測や市場への波及効果の分析は絶対に含めず、あくまで動画内で語られている内容のみを詳細に要約してください。

【動画データ（TOP 10）】
{videos_context}

【レポートのフォーマット】
# 注目トピック　TOP１０

## 注目トピック 第〇位: [動画タイトル] (チャンネル名: [チャンネル名])
（動画の詳細な要約：約300〜400文字のボリュームで充実させて記述）

## 注目トピック 第〇位: [動画タイトル] (チャンネル名: [チャンネル名])
（動画の詳細な要約...同様に記載）

...（以下、第10位までリストアップしてください）
"""
    
    try:
        # Use Llama 3 70B model for fast generation
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            model="llama-3.3-70b-versatile",
            temperature=0.4,
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        print(f"Error generating report with Groq: {e}")
        return f"レポート生成中にエラーが発生しました。\nエラー詳細: {e}"

def generate_filename_summary(client, text_content):
    """Generate a 15-character summary for the filename."""
    prompt = f"""
以下のレポートの内容を一言で表す、15文字以内の簡潔なタイトル（要約）を作成してください。
記号や改行、空白はすべて除き、Windowsのファイル名として使える文字列のみを直接出力してください。
回答には余計な挨拶や「〜です」などは一切含めないでください。
例: 「イラン原油高とAI波及」「エネルギーと技術の波及効果」など

【レポート内容】
{text_content[:1500]}
"""
    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",
            temperature=0.3,
            max_tokens=30,
        )
        summary = chat_completion.choices[0].message.content.strip()
        # Clean up the output string to be safe for filenames
        bad_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '「', '」', ' ', '　', '\n']
        for c in bad_chars:
            summary = summary.replace(c, '')
        return summary[:15]
    except Exception as e:
        print(f"Error generating summary: {e}")
        return "海外YouTubeトップ10"

def save_as_docx(report_text, filepath):
    """Save the markdown-like text to a nice Word document."""
    doc = Document()
    for line in report_text.split('\n'):
        if line.startswith('# '):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith('### '):
            doc.add_heading(line[4:].strip(), level=3)
        else:
            if line.strip():
                doc.add_paragraph(line.strip())
    doc.save(filepath)

def send_email_with_attachment(filepath, filename):
    """Send the generated report via email."""
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        print("Warning: GMAIL_ADDRESS or GMAIL_APP_PASSWORD not set. Skipping email.")
        return

    msg = EmailMessage()
    msg['Subject'] = f"【自動リサーチ】海外投資YouTubeトップ10要約 ({filename})"
    msg['From'] = GMAIL_ADDRESS
    msg['To'] = GMAIL_ADDRESS
    msg.set_content("本日の海外投資ニュースのリサーチ結果を添付します。\n\n※このメールはGitHub Actionsまたはローカル環境から自動送信されています。")

    try:
        with open(filepath, 'rb') as f:
            file_data = f.read()
            msg.add_attachment(
                file_data, 
                maintype='application', 
                subtype='vnd.openxmlformats-officedocument.wordprocessingml.document', 
                filename=filename
            )
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
            smtp.send_message(msg)
            
        print("-> Successfully sent email with attached report!")
    except Exception as e:
        print(f"Error sending email: {e}")

def fetch_and_process_news():
    """
    Fetch trending international finance YouTube video and process it for retail investors using Groq.
    """
    jst_time = datetime.datetime.now(ZoneInfo("Asia/Tokyo"))
    date_str = jst_time.strftime("%Y%m%d")
    
    # Target path as requested for output
    target_dir = os.path.join("G:\\", "マイドライブ", "Antigravity", "海外投資記事", "リサーチ")
    if not os.path.exists(target_dir):
        # Using a relative/local fallback if the exact GDrive path doesn't exist in CI environment
        try:
            os.makedirs(target_dir, exist_ok=True)
            print(f"Created target directory: {target_dir}")
        except Exception as e:
            print(f"Could not create Google Drive directory. Using local 'output' folder instead.")
            target_dir = "output"
            os.makedirs(target_dir, exist_ok=True)

    print(f"[{jst_time}] Starting YouTube video research and plot generation...")
    
    client = get_groq_client()
    if not client:
        return
        
    print("1. Fetching top trending English investment YouTube videos...")
    videos = fetch_top_trending_investment_videos()
    
    if not videos:
        print("Stopping process: Could not find any trending videos.")
        return
        
    print("2. Fetching transcripts and filtering for exactly 10 valid videos...")
    videos_context = ""
    valid_count = 0
    
    for idx, item in enumerate(videos, 1):
        video_id = item["id"]["videoId"]
        title = item["snippet"]["title"]
        channel = item["snippet"]["channelTitle"]
        
        safe_title = title.encode('cp932', 'replace').decode('cp932')
        print(f"-> Checking original Rank {idx}: {safe_title}")
        
        transcript_text = get_video_transcript(video_id)
        if not transcript_text:
            print("   -> No transcript found. Skipping.")
            continue
            
        valid_count += 1
        videos_context += f"【抽出第{valid_count}位 (全体アクセス順位: {idx}位)】\n"
        videos_context += f"タイトル: {title}\n"
        videos_context += f"チャンネル名: {channel}\n"
        videos_context += f"内容/トランスクリプト: {transcript_text}\n\n"
        
        if valid_count >= 10:
            print("-> Successfully collected 10 videos with transcripts.")
            break
            
    if valid_count == 0:
        print("Stopping process: Could not find any videos with valid transcripts.")
        return
        
    print("3. Generating top 10 summary report using Groq...")
    report_text = generate_top10_report(client, videos_context)
    
    print("4. Generating 15-character filename summary...")
    summary = generate_filename_summary(client, report_text)
    
    # Generate filename (e.g., 20260315_01_イラン原油高.docx)
    index = 1
    while True:
        filename = f"{date_str}_{index:02d}_{summary}.docx"
        filepath = os.path.join(target_dir, filename)
        if not os.path.exists(filepath):
            break
        index += 1
        
    print(f"4. Saving report to {filepath} ...")
    try:
        save_as_docx(report_text, filepath)
        print("-> Successfully saved report!")
        
        print("5. Sending report via email...")
        send_email_with_attachment(filepath, filename)
    except Exception as e:
        print(f"Error saving file: {e}")

if __name__ == "__main__":
    fetch_and_process_news()
