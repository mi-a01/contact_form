from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/run-script', methods=['POST'])
def run_script():
    # 実行したいpython処理
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.common.exceptions import NoSuchElementException
    from bs4 import BeautifulSoup
    import time
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials

    # --- 設定 ---
    name1 = "山田"
    name2 = "太郎"
    フリガナ1 = "ヤマダ"
    フリガナ2 = "タロウ"
    ふりがな1 = "やまだ"
    ふりがな2 = "たろう"
    電話番号1 = "000"
    電話番号2 = "1234"
    電話番号3 = "5678"
    mail = "a@a.com"



    # Google Sheets認証
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    # 対象シート
    sheet = client.open_by_key("1mvFRPB0lotw9Lh1wJlYk2EldqIeT97cbWIDzTW5LqD0").get_worksheet_by_id(2142414190)

    data = sheet.get_all_values()[1:]

    url_content_list = []
    for idx, row in enumerate(data):
        if len(row) > 4 and row[4].strip():  # E列が空でない
            url = row[4].strip()
            content = row[7].strip() if len(row) > 7 else ""
            url_content_list.append((idx + 2, url, content)) 
            # print(url_content_list)

    # Chrome起動オプション
    options = Options()
    # options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    # 各URLに対して処理
    for row_num, url, content in url_content_list:
        log_lines = []
        driver.get(url)
        print("ページを開きました")
        time.sleep(5)

        # --- 「お問い合わせ」リンクを探す ---
        links = driver.find_elements(By.TAG_NAME, "a")
        found_contact = False
        for link in links:
            if "お問い合わせ" in link.text or "contact" in link.text or "お問合" in link.text or "CONTACT" in link.text or "contact" in link.text or "Contact" in link.text or "お問合せ" in link.text.lower():
                href = link.get_attribute("href")
                if href:
                    if href.startswith("/"):
                        from urllib.parse import urljoin
                        href = urljoin(url, href)

                    print(f"[{url}] → お問い合わせページ: {href}")
                    log_lines.append(f"[{url}] → お問い合わせページ: {href}")
                    driver.get(href)
                    time.sleep(2)
                    found_contact = True
                    break

        if not found_contact:
            print(f"[{url}] → お問い合わせリンクが見つかりません")
            log_lines.append(f"[{url}] → お問い合わせリンクが見つかりません")
            print("確認お願いします。")
            message = "確認お願いします。送信結果 を入力して次へ進んでください。\n\n"
            message += "\n*****************************************************************************\n"
            message += "\n".join(log_lines)

            sheet.update_cell(row_num, 11, message)
            # I列（index 8）に入力されるまで待機
            while True:
                user_input = sheet.cell(row_num, 10).value  # I列
                if user_input and user_input.strip():
                    print(f"送信結果が入力されました: {user_input}")
                    break
                print("送信結果入力待ち...")
                time.sleep(5)  # 5秒ごとにチェック

            # J列メッセージ削除（ログ以外のメッセージを消す）
            lines = message.splitlines()
            cleaned_lines = [line for line in lines if not line.strip().startswith("確認お願いします") and line.strip() != ""]
            cleaned = "\n".join(cleaned_lines)
            sheet.update_cell(row_num, 11, cleaned)
            continue
    driver.quit()

    result = {"message": "Pythonスクリプトが実行されました"}
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
