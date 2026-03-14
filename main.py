import os
import datetime
import requests
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from groq import Groq
from docx import Document
import smtplib
from email.message import EmailMessage

# Load environment variables
load_dotenv()

NEWS_API_KEY = os.getenv("NEWS_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

def get_groq_client():
    """Get the Groq API client."""
    if not GROQ_API_KEY:
        print("Error: GROQ_API_KEY environment variable not set.")
        return None
    return Groq(api_key=GROQ_API_KEY)

def fetch_top_business_news():
    """Fetch top business headlines from the US using News API."""
    if not NEWS_API_KEY:
        print("Error: NEWS_API_KEY environment variable not set.")
        return None
        
    url = f"https://newsapi.org/v2/top-headlines?country=us&category=business&apiKey={NEWS_API_KEY}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        if data.get("status") == "ok" and data.get("articles"):
            # Return the top 3 articles to avoid overwhelming the prompt
            return data["articles"][:3]
        else:
            print("No articles found or API error:", data)
            return None
    except Exception as e:
        print(f"Error fetching news: {e}")
        return None

def generate_ripple_effect_report(client, articles):
    """Generate a ripple effect analysis report using Groq (Llama 3)."""
    if not articles:
        return "No news articles to analyze."
        
    # Prepare the prompt with the news articles
    news_text = ""
    for i, article in enumerate(articles, 1):
        title = article.get('title', 'No Title')
        description = article.get('description', 'No Description')
        news_text += f"ニュース{i}:\nタイトル: {title}\n概要: {description}\n\n"
        
    prompt = f"""
あなたは優秀な証券アナリストです。以下の米国の最新ビジネスニュースを読み、
「風が吹けば桶屋が儲かる」という視点で、これらのニュースが日本のどの業界や銘柄に予期せぬ恩恵（波及効果）をもたらすか、具体的なシナリオを日本語で論理的に解説するリサーチレポートを作成してください。

要約を作成する際は、ニュースから「具体的な数値」「企業のアクション」「専門家の見解」を必ず抽出したうえで、それらを含めた内容にしてください。

【取得したニュース】
{news_text}

【レポートのフォーマット】
# グローバルニュース・波及効果分析レポート

## 本日のハイライト（ニュース要約）
（取得したニュースの要約。「具体的な数値」「企業のアクション」「専門家の見解」を必ず含めること）

## 波及効果シナリオ（風が吹けば桶屋が儲かる）
（ニュースの事象が、どのような連鎖反応を経て日本の特定の業界・企業に影響するかをステップ・バイ・ステップで解説。最低2つのシナリオを提示してください。）

## 注目すべき日本株セクター
（恩恵を受ける可能性のある具体的な日本の業種やテーマ）
"""
    
    try:
        # Use Llama 3 8B or 70B model for fast generation
        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            model="llama-3.3-70b-versatile",
            temperature=0.7,
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
        return "ニュース波及効果まとめ"

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
    msg['Subject'] = f"【自動リサーチ】本日のニュース分析 ({filename})"
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
    Fetch international finance news and process it for retail investors using Groq.
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

    print(f"[{jst_time}] Starting news research and plot generation...")
    
    client = get_groq_client()
    if not client:
        return
        
    print("1. Fetching US business news...")
    articles = fetch_top_business_news()
    
    if not articles:
        print("Stopping process: Could not fetch news.")
        return
        
    print(f"-> Successfully fetched {len(articles)} articles.")
    print("2. Generating ripple effect report using Groq...")
    report_text = generate_ripple_effect_report(client, articles)
    
    print("3. Generating 15-character filename summary...")
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
