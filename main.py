import os
import re
import time
from playwright.sync_api import sync_playwright
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

def main():
    # .envファイルから環境変数を読み込む
    load_dotenv()
    
    print("===== 図書館データ取得・スプレッドシート連携スクリプト =====")
    
    # 環境変数から機密情報を取得
    library_id = os.environ.get("LIBRARY_LOGIN_ID")
    library_pw = os.environ.get("LIBRARY_PASSWORD")
    
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
            
            # --- 表示件数(100)の変更 ---
            # JavaScriptのonchangeで自動submitされるため
            print("-> 表示件数を100件に変更します...")
            page.locator('select[name="WRTCOUNT"]').first.select_option('100')
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
            
            # --- 出版年の降順(SORT=-3)に変更 ---
            # 出版年横の「↓」リンクをクリック
            print("-> 出版年降順で並び替えます...")
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
                    
            # --- 読書記録の抽出 ---
            print("-> 読書記録データの抽出を開始します...")
            rec_list = []
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

        except Exception as e:
            print(f"ブラウザ操作中にエラーが発生しました: {e}")
            page.screenshot(path="main_error.png")
            return
            
        browser.close()
        
    print(f"-> スクレイピング完了: 計 {len(res_list)} 件の図書データを取得しました。")
    if not res_list:
        print("-> 取得データが0件のため、スプレッドシートへの書き込み処理をスキップします。")
        return

    # --- Spreadsheet 書き込み ---
    print("-> 2. Google SpreadSheet への書き込みを開始します")
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
        
    except gspread.exceptions.APIError as api_err:
            print(f"スプレッドシートAPIエラー: APIの権限や共有設定が正しく行われているか確認してください。\n詳細: {api_err}")
    except Exception as e:
        print(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

if __name__ == "__main__":
    main()
