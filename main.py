import os
import re
import time
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from google import genai
from google.genai import types

def main():
    # .envファイルから環境変数を読み込む
    load_dotenv()
    
    print("===== 図書館データ取得・スプレッドシート連携スクリプト =====")
    
    # 環境変数から機密情報を取得
    library_id = os.environ.get("LIBRARY_LOGIN_ID")
    library_pw = os.environ.get("LIBRARY_PASSWORD")
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    
    if not library_id or not library_pw:
        print("エラー: .envファイルに LIBRARY_LOGIN_ID または LIBRARY_PASSWORD が設定されていません。")
        return
        
    res_list = []
    
    print("-> 1. Playwrightによってブラウザを起動し、データ取得を開始します")
    with sync_playwright() as p:
        # headless=Trueでバックグラウンド実行
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 800})
        page = context.new_page()
        
        # --- ログインプロセス ---
        try:
            page.goto("https://www.library.city.kita.lg.jp/opw/OPW/OPWLOGINTIME.CSP?HPFLG=1&NEXT=OPWUSERINFO&DB=LIB")
            page.locator('input[type="text"]').first.fill(library_id)
            page.locator('input[type="password"]').first.fill(library_pw)
            page.locator('input[type="submit"], input[type="image"], button[type="submit"]').first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            # --- メニュー遷移 ---
            page.get_by_text("各種一覧・その他", exact=False).first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            page.get_by_text("新着案内", exact=False).first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            # --- 検索条件指定（一般書、前日、集計） ---
            print("-> 資料の種類「一般書（すべて）」と「前日」を選択し、集計を実行します...")
            page.locator('select[name="SK"]').select_option('1')
            page.locator('select[name="SPAN"]').select_option('1')
            page.locator('input[name="syukei"]').click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
            
            # --- 新着図書の有無を確認 ---
            if page.locator('select[name="WRTCOUNT"]').count() > 0:
                # --- 表示件数(100)の変更 ---
                # JavaScriptのonchangeで自動submitされるため
                print("-> 表示件数を100件に変更します...")
                page.locator('select[name="WRTCOUNT"]').first.select_option('100')
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(2000)
                
                # --- 出版年の降順(SORT=-3)に変更 ---
                # 出版年横の「↓」リンクをクリック
                print("-> 出版年降順で並び替えます...")
                if page.locator('a[href*="SORT=-3"]').count() > 0:
                    page.locator('a[href*="SORT=-3"]').first.click()
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(2000)
                
                # --- データ抽出ループ ---
                print("-> データの抽出を開始します...")
                current_page = 1
                while True:
                    print(f"  - ページ {current_page} を処理中...")
                    
                    rows = page.locator('tr.lightcolor, tr.basecolor').all()
                    for row in rows:
                        tds = row.locator('td').all()
                        if len(tds) >= 6:
                            date_text = tds[0].inner_text().strip()
                            title = tds[2].inner_text().strip()
                            author = tds[3].inner_text().strip()
                            publisher = tds[4].inner_text().strip()
                            year = tds[5].inner_text().strip()
                            
                            # 整形（改行や余分な空白の削除）
                            date_text = re.sub(r'\s+', ' ', date_text)
                            title = re.sub(r'\s+', ' ', title)
                            author = re.sub(r'\s+', ' ', author)
                            publisher = re.sub(r'\s+', ' ', publisher)
                            year = re.sub(r'\s+', ' ', year)
                            
                            if title: # タイトルが空でない場合のみ
                                res_list.append([date_text, title, author, publisher, year])
                    
                    # 「次」ページへのリンクがあるか確認
                    next_links = page.locator('a:has-text("次")')
                    if next_links.count() > 0:
                        try:
                            next_links.first.click()
                            page.wait_for_load_state("networkidle")
                            page.wait_for_timeout(1500)
                            current_page += 1
                        except Exception as e:
                            print(f"次のページへの遷移中にエラーが発生しました、ループを終了します。: {e}")
                            break
                    else:
                        break
            else:
                print("-> 新着図書が存在しません。データの抽出をスキップします。")
                    
            # --- 読書記録の抽出 ---
            rec_list = []
            if len(res_list) > 0:
                print("-> 読書記録データの抽出を開始します...")
                page.get_by_text("利用状況ページ", exact=False).first.click()
                page.wait_for_timeout(1000)
                page.get_by_text("利用状況一覧", exact=False).first.click()
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(2000)

                page.locator('a[href="#ContentRec"]').first.click()
                page.wait_for_timeout(1000)
                
                current_page_rec = 1
                while True:
                    print(f"  - 読書記録 ページ {current_page_rec} を処理中...")
                    form_rec = page.locator('form[name="FormREC"]')
                    rows = form_rec.locator('tr.lightcolor, tr.basecolor').all()
                    for row in rows:
                        tds = row.locator('td').all()
                        if len(tds) >= 7:
                            no = tds[1].inner_text().strip()
                            title = tds[3].inner_text().strip()
                            author = tds[4].inner_text().strip()
                            date_val = tds[6].inner_text().strip()
                            
                            no = re.sub(r'\s+', ' ', no)
                            title = re.sub(r'\s+', ' ', title)
                            author = re.sub(r'\s+', ' ', author)
                            date_val = re.sub(r'\s+', ' ', date_val)
                            
                            if no and title:
                                rec_list.append([no, title, author, date_val])
                    
                    next_links = form_rec.locator('a:has-text("次")')
                    if next_links.count() > 0:
                        try:
                            next_links.first.click()
                            page.wait_for_load_state("networkidle")
                            page.wait_for_timeout(1500)
                            current_page_rec += 1
                        except Exception as e:
                            print(f"読書記録の次のページへの遷移中にエラーが発生しました: {e}")
                            break
                    else:
                        break
            else:
                print("-> 新着図書が0件のため、読書記録のデータ抽出もスキップします。")

        except Exception as e:
            print(f"ブラウザ操作中にエラーが発生しました: {e}")
            page.screenshot(path="main_error.png")
            return
            
        browser.close()
        
    print(f"-> スクレイピング完了: 計 {len(res_list)} 件の図書データを取得しました。")

    # --- Spreadsheet 書き込み・クリア ---
    print("-> 2. Google SpreadSheet への接続を開始します")
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_path = 'credentials.json'
        
        if not os.path.exists(creds_path):
            print(f"エラー: {creds_path} が見つかりません。ファイルが作業フォルダ(d:\\AG\\library_check)に配置されているか確認してください。")
            return

        credentials = Credentials.from_service_account_file(creds_path, scopes=scopes)
        gc = gspread.authorize(credentials)
        
        spreadsheet_id = os.environ.get("SPREADSHEET_ID_NEW_BOOKS")
        if not spreadsheet_id:
            print("エラー: .envファイルに SPREADSHEET_ID_NEW_BOOKS が設定されていません。")
            return
            
        sh = gc.open_by_key(spreadsheet_id)
        
        try:
            worksheet = sh.worksheet("シート1")
        except gspread.exceptions.WorksheetNotFound:
            print("エラー: 'シート1' という名前のシートが見つかりません。")
            return
            
        if not res_list:
            print("-> 新着図書が0件のため、スプレッドシートの中身を全て削除してプログラムを終了します。")
            worksheet.clear()
            print("★ 完了: 新着案内のスプレッドシートのデータを削除しました。")
            return
            
        print("-> 古いデータを削除しています...")
        worksheet.clear()
        
        print("-> 新しいデータを書き込み中...")
        # 列の見出し（ヘッダー）を追加
        header = ["受入日", "タイトル", "著者名", "出版者", "出版年"]
        res_list.insert(0, header)
        
        # A1セルから一括で書き込み
        worksheet.append_rows(res_list, value_input_option='USER_ENTERED')
        
        print(f"★ 完了: 新着案内のスプレッドシートのデータを新しく書き換えました！")
        
        # --- 読書記録のスプレッドシートへの書き込み ---
        if not rec_list:
            print("-> 読書記録の取得データが0件のため、読書記録の書き込み処理をスキップします。")
        else:
            print(f"-> 読書記録完了: 計 {len(rec_list)} 件の読書記録データを取得しました。")
            print("-> 3. 読書記録のスプレッドシートへの書き込みを開始します")
            rec_spreadsheet_id = os.environ.get("SPREADSHEET_ID_READING_HISTORY")
            if not rec_spreadsheet_id:
                print("エラー: .envファイルに SPREADSHEET_ID_READING_HISTORY が設定されていません。")
                return
                
            rec_sh = gc.open_by_key(rec_spreadsheet_id)
            
            try:
                rec_worksheet = rec_sh.worksheet("シート1")
            except gspread.exceptions.WorksheetNotFound:
                print("エラー: 読書記録のスプレッドシートに 'シート1' が見つかりません。")
                return
                
            print("-> 読書記録の古いデータを削除しています...")
            rec_worksheet.clear()
            
            print("-> 読書記録の新しいデータを書き込み中...")
            rec_header = ["No", "タイトル", "著者名", "貸出日"]
            rec_list.insert(0, rec_header)
            
            rec_worksheet.append_rows(rec_list, value_input_option='USER_ENTERED')
            print(f"★ 完了: 読書記録のスプレッドシートのデータを新しく書き換えました！")
            
        # --- Gemini APIによる好みの解析とハイライト ---
        if gemini_api_key and len(res_list) > 1 and len(rec_list) > 1:
            print("-> 4. Geminiを利用して好みの本を解析し、新着図書をハイライトします...")
            try:
                client = genai.Client(api_key=gemini_api_key)
                
                # 読書記録の文字列化 (ヘッダー除く)
                reading_history_text = "\n".join([f"- {row[1]} (著: {row[2]})" for row in rec_list[1:]])
                
                # 新着図書の文字列化 (ヘッダー除く)
                new_books_text = "\n".join([f"- {row[1]} (著: {row[2]})" for row in res_list[1:]])
                
                prompt = f"""
以下の「読書記録」から、ユーザーがどのようなジャンルやテーマ、著者の本を好むかを分析してください。
その後、その分析結果に基づいて、以下の「新着図書」の中からユーザーの好みに合うおすすめの本を抽出してください。
抽出した本の「タイトル（文字列の完全一致）」のリストをJSON形式で返してください。

【読書記録】
{reading_history_text}

【新着図書】
{new_books_text}

出力は以下のJSONフォーマットのみにしてください。
{{
  "recommended_titles": [
    "おすすめの本のタイトル1",
    "おすすめの本のタイトル2"
  ]
}}
"""
                response = client.models.generate_content(
                    model='gemini-3.1-flash-lite-preview',
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                    ),
                )
                
                result_json = response.text
                result_data = json.loads(result_json)
                recommended_titles = result_data.get("recommended_titles", [])
                
                if recommended_titles:
                    print(f"  - おすすめ本として {len(recommended_titles)} 冊が抽出されました。ハイライト処理を行います。")
                    # 新着図書スプレッドシートのB列のタイトルと照合
                    highlight_count = 0
                    for i, row in enumerate(res_list[1:], start=2): # headerは1行目、データは2行目から
                        title = row[1]
                        if title in recommended_titles:
                            cell_name = f"B{i}"
                            worksheet.format(cell_name, {
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 0.0
                                }
                            })
                            print(f"    - ハイライトしました: {title}")
                            highlight_count += 1
                    
                    if highlight_count > 0:
                        print("★ 完了: おすすめ新着図書のハイライトを行いました！")
                    else:
                        print("  - 抽出されたタイトルと新着図書のリストが完全に一致しませんでした。")
                else:
                    print("  - おすすめの本は見つかりませんでした。")
                    
            except Exception as e:
                print(f"Gemini APIによる処理中にエラーが発生しました: {e}")
        elif not gemini_api_key:
            print("-> GEMINI_API_KEYが設定されていないため、おすすめ本のハイライト処理をスキップします。")

        # --- 5. 新着図書のHTMLメール送信 ---
        gmail_address = os.environ.get("GMAIL_ADDRESS")
        gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
        
        if gmail_address and gmail_password and len(res_list) > 1:
            print("-> 5. 新着図書のリストをGmailで送信します...")
            try:
                msg = MIMEMultipart("alternative")
                msg["Subject"] = "【お知らせ】図書館の新着図書リスト"
                msg["From"] = gmail_address
                msg["To"] = gmail_address
                
                # 推奨タイトルが未定義の場合の回避
                if 'recommended_titles' not in locals():
                    recommended_titles = []
                
                # HTMLテーブルの構築
                html_content = """
                <html>
                <head>
                <style>
                  table { border-collapse: collapse; width: 100%; font-family: sans-serif; }
                  th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
                  th { background-color: #f2f2f2; }
                  .highlight { background-color: #ffff99; } /* おすすめ本は黄色 */
                </style>
                </head>
                <body>
                  <h2>本日の新着図書リスト</h2>
                  <p>※黄色でハイライトされているのは、あなたの好みに合わせたおすすめの本です。</p>
                  <table>
                    <tr>
                      <th>受入日</th>
                      <th>タイトル</th>
                      <th>著者名</th>
                      <th>出版者</th>
                      <th>出版年</th>
                    </tr>
                """
                
                for row in res_list[1:]:
                    title = row[1]
                    tr_class = ' class="highlight"' if title in recommended_titles else ''
                    
                    html_content += f"""
                    <tr{tr_class}>
                      <td>{row[0]}</td>
                      <td>{row[1]}</td>
                      <td>{row[2]}</td>
                      <td>{row[3]}</td>
                      <td>{row[4]}</td>
                    </tr>
                    """
                
                html_content += """
                  </table>
                </body>
                </html>
                """
                
                part = MIMEText(html_content, "html")
                msg.attach(part)
                
                print(f"  - {gmail_address} 宛てに送信中...")
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                    server.login(gmail_address, gmail_password)
                    server.send_message(msg)
                    
                print("★ 完了: 新着図書のメールを送信しました！")
            except Exception as e:
                print(f"メールの送信中にエラーが発生しました: {e}")
        elif not gmail_address or not gmail_password:
            print("-> Gmailのアドレスまたはアプリパスワードが設定されていないため、メール送信をスキップします。")
            
    except gspread.exceptions.APIError as api_err:
            print(f"スプレッドシートAPIエラー: APIの権限や共有設定が正しく行われているか確認してください。\n詳細: {api_err}")
    except Exception as e:
        print(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
