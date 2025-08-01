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
    try:
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

        soup = BeautifulSoup(driver.page_source, "html.parser")

        # --- 内容入力 ---
        textarea_found = False
        for label_text in ["お問い合わせ内容", "内容"]:
            targets = soup.find_all(string=lambda s: s and label_text in s)
            for target in targets:
                parent = target.parent
                while parent:
                    textarea = parent.find("textarea")
                    if textarea:
                        try:
                            textarea_name = textarea.get("name")
                            textarea_id = textarea.get("id")
                            if textarea_name:
                                element = driver.find_element(By.NAME, textarea_name)
                            elif textarea_id:
                                element = driver.find_element(By.ID, textarea_id)
                            else:
                                break
                            element.clear()
                            element.send_keys(content)
                            print(f"[{url}] → '{label_text}' に対応するテキストエリアに入力しました。")
                            log_lines.append(f"[{url}] → '{label_text}' に対応するテキストエリアに入力しました。")
                            textarea_found = True
                            break
                        except NoSuchElementException:
                            break
                    parent = parent.parent
                if textarea_found:
                    break
            if textarea_found:
                break

        if not textarea_found:
            textareas = driver.find_elements(By.TAG_NAME, "textarea")
            for ta in textareas:
                if ta.is_displayed() and ta.is_enabled():
                    ta.send_keys(content)
                    print(f"[{url}] → 明示ラベルなしtextareaに入力成功")
                    log_lines.append(f"[{url}] → 明示ラベルなしtextareaに入力成功")
                    textarea_found = True
                    break

        if not textarea_found:
            print(f"[{url}] → 【手作業が必要です】お問い合わせの種別を選んでください")
            log_lines.append(f"[{url}] → 【手作業が必要です】お問い合わせの種別を選んでください")
    

        # --- 名前入力 ---
        def find_input_near_label(label_elem):
            inputs = label_elem.find_all("input", {"type": "text"})
            if inputs:
                return inputs
            for sibling in label_elem.find_next_siblings():
                inputs = sibling.find_all("input", {"type": "text"})
                if inputs:
                    return inputs
            parent = label_elem.parent
            while parent:
                for sibling in parent.find_next_siblings():
                    inputs = sibling.find_all("input", {"type": "text"})
                    if inputs:
                        return inputs
                parent = parent.parent
            return []

        name_inputs_found = False
        soup = BeautifulSoup(driver.page_source, "html.parser")
        used_names = set()
        name_keywords = ["名前", "お名前", "name", "氏名"]

        # --- 1) inputタグを直接探索（name/idベース） ---
        candidate_inputs = soup.find_all("input", {"type": "text"})
        matched_inputs = []
        for input_tag in candidate_inputs:
            input_name = (input_tag.get("name") or input_tag.get("id") or "").lower()
            if any(kw in input_name for kw in name_keywords):
                matched_inputs.append(input_tag)

        # 該当inputが1つなら全名を入力、2つなら姓・名で分割
        if matched_inputs:
            try:
                if len(matched_inputs) == 1:
                    input_tag = matched_inputs[0]
                    input_name = input_tag.get("name") or input_tag.get("id")
                    if input_name:
                        element = (
                            driver.find_element(By.NAME, input_tag.get("name"))
                            if input_tag.get("name")
                            else driver.find_element(By.ID, input_tag.get("id"))
                        )
                        element.clear()
                        element.send_keys(name1 + name2)
                        print(f"[{url}] → 名前入力（直接input探索）: {name1 + name2}")
                        log_lines.append(f"[{url}] → 名前入力（直接input探索）: {name1 + name2}")
                        name_inputs_found = True
                        used_names.add(input_name)
                elif len(matched_inputs) >= 2:
                    for i, val in enumerate([name1, name2]):
                        input_tag = matched_inputs[i]
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name:
                            element = (
                                driver.find_element(By.NAME, input_tag.get("name"))
                                if input_tag.get("name")
                                else driver.find_element(By.ID, input_tag.get("id"))
                            )
                            element.clear()
                            element.send_keys(val)
                            used_names.add(input_name)
                    print(f"[{url}] → 名前入力（直接input探索・分割）: {name1}, {name2}")
                    log_lines.append(f"[{url}] → 名前入力（直接input探索・分割）: {name1}, {name2}")
                    name_inputs_found = True
            except Exception as e:
                print(f"[{url}] → 名前入力エラー（直接探索）: {e}")
                log_lines.append(f"[{url}] → 名前入力エラー（直接探索）: {e}")

        # --- 2) テキストラベル探索（従来処理） ---
        if not name_inputs_found:
            name_targets = soup.find_all(string=lambda s: s and ("名前" in s or "お名前" in s))

            if name_targets:

                for target in name_targets:
                    label_elem = target.find_parent(['label', 'div', 'span', 'td', 'p', 'dt']) or target.find_parent()
                    if not label_elem:
                        continue

                    parent = label_elem.parent
                    inputs_found = False

                    # 2-1) 親要素内探索
                    if parent:
                        inputs = parent.find_all("input", {"type": "text"})
                        for input_tag in inputs:
                            input_name = input_tag.get("name") or input_tag.get("id")
                            if input_name and input_name not in used_names:
                                try:
                                    element = (
                                        driver.find_element(By.NAME, input_name)
                                        if input_tag.get("name")
                                        else driver.find_element(By.ID, input_name)
                                    )
                                    element.clear()
                                    element.send_keys(name1 + name2)
                                    print(f"[{url}] → 名前入力（親要素内）: {name1 + name2}")
                                    log_lines.append(f"[{url}] → 名前入力（親要素内）: {name1 + name2}")
                                    inputs_found = True
                                    used_names.add(input_name)
                                    break
                                except Exception as e:
                                    print(f"[{url}] → 名前入力エラー（親要素内）: {e}")
                                    log_lines.append(f"[{url}] → 名前入力エラー（親要素内）: {e}")

                    if inputs_found:
                        name_inputs_found = True
                        continue

                    # 2-2) 兄弟要素探索
                    next_elem = label_elem.find_next_sibling()
                    while next_elem:
                        inputs = next_elem.find_all("input", {"type": "text"})
                        if inputs:
                            try:
                                if len(inputs) == 1:
                                    input_name = inputs[0].get("name") or inputs[0].get("id")
                                    if input_name and input_name not in used_names:
                                        element = (
                                            driver.find_element(By.NAME, input_name)
                                            if inputs[0].get("name")
                                            else driver.find_element(By.ID, input_name)
                                        )
                                        element.clear()
                                        element.send_keys(name1 + name2)
                                        print(f"[{url}] → 名前を1つの欄に入力: {name1 + name2}")
                                        log_lines.append(f"[{url}] → 名前を1つの欄に入力: {name1 + name2}")
                                        name_inputs_found = True
                                        used_names.add(input_name)
                                        break
                                elif len(inputs) >= 2:
                                    for i, val in enumerate([name1, name2]):
                                        input_name = inputs[i].get("name") or inputs[i].get("id")
                                        if input_name and input_name not in used_names:
                                            element = (
                                                driver.find_element(By.NAME, input_name)
                                                if inputs[i].get("name")
                                                else driver.find_element(By.ID, input_name)
                                            )
                                            element.clear()
                                            element.send_keys(val)
                                            used_names.add(input_name)
                                    print(f"[{url}] → 名前を分割して入力: {name1}, {name2}")
                                    log_lines.append(f"[{url}] → 名前を分割して入力: {name1}, {name2}")
                                    name_inputs_found = True
                                    break
                            except Exception as e:
                                print(f"[{url}] → 名前入力エラー（兄弟探索）: {e}")
                                log_lines.append(f"[{url}] → 名前入力エラー（兄弟探索）: {e}")
                        if name_inputs_found:
                            break
                        next_elem = next_elem.find_next_sibling()
                    if name_inputs_found:
                        continue

                    # 2-3) 階層探索
                    inputs = find_input_near_label(label_elem)
                    for input_tag in inputs:
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name and input_name not in used_names:
                            try:
                                element = (
                                    driver.find_element(By.NAME, input_name)
                                    if input_tag.get("name")
                                    else driver.find_element(By.ID, input_name)
                                )
                                element.clear()
                                element.send_keys(name1 + name2)
                                print(f"[{url}] → 名前入力（追加探索）: {name1 + name2}")
                                log_lines.append(f"[{url}] → 名前入力（追加探索）: {name1 + name2}")
                                name_inputs_found = True
                                used_names.add(input_name)
                                break
                            except Exception as e:
                                print(f"[{url}] → 名前入力エラー（追加探索）: {e}")
                                log_lines.append(f"[{url}] → 名前入力エラー（追加探索）: {e}")
                    if name_inputs_found:
                        continue

        # --- 最後にログ出力 ---
        if not name_inputs_found:
            print(f"[{url}] → 名前欄が見つかりません")
            log_lines.append(f"[{url}] → 名前欄が見つかりません")

        # ふりがな入力
        def try_fill(inputs, driver, used_names, furi_text, url, log_lines):
            for input_tag in inputs:
                input_name = input_tag.get("name") or input_tag.get("id")
                if input_name and input_name not in used_names:
                    if input_tag.get("value", "").strip():
                        continue
                    try:
                        element = (
                            driver.find_element(By.NAME, input_tag.get("name"))
                            if input_tag.get("name")
                            else driver.find_element(By.ID, input_tag.get("id"))
                        )
                        if element.get_attribute("value").strip():
                            continue
                        element.clear()
                        element.send_keys(furi_text)
                        print(f"[{url}] → ふりがな入力: {furi_text}")
                        log_lines.append(f"[{url}] → ふりがな入力: {furi_text}")
                        used_names.add(input_name)
                        return True
                    except Exception as e:
                        print(f"[{url}] → ふりがな入力エラー: {e}")
                        log_lines.append(f"[{url}] → ふりがな入力エラー: {e}")
            return False

        # --- ふりがな入力処理 ---
        furi_hiragana_inputs_found = False
        soup = BeautifulSoup(driver.page_source, "html.parser")
        used_names = set()
        furi_keywords = ["ふりがな", "furigana", "yomi", "読み"]

        # --- 1) name/id 直接探索 ---
        candidate_inputs = soup.find_all("input", {"type": "text"})
        matched_inputs = []

        for input_tag in candidate_inputs:
            input_name = (input_tag.get("name") or input_tag.get("id") or "").lower()
            if any(kw in input_name for kw in furi_keywords):
                if input_name not in used_names and not input_tag.get("value", "").strip():
                    matched_inputs.append(input_tag)

        try:
            if matched_inputs:
                if len(matched_inputs) == 1:
                    input_tag = matched_inputs[0]
                    input_name = input_tag.get("name") or input_tag.get("id")
                    if input_name:
                        element = (
                            driver.find_element(By.NAME, input_tag.get("name"))
                            if input_tag.get("name")
                            else driver.find_element(By.ID, input_tag.get("id"))
                        )
                        if element.get_attribute("value").strip():
                            pass
                        else:
                            element.clear()
                            element.send_keys(ふりがな1 + ふりがな2)
                            print(f"[{url}] → ふりがな入力（直接input探索）: {ふりがな1 + ふりがな2}")
                            log_lines.append(f"[{url}] → ふりがな入力（直接input探索）: {ふりがな1 + ふりがな2}")
                            furi_hiragana_inputs_found = True
                            used_names.add(input_name)
                elif len(matched_inputs) >= 2:
                    for i, val in enumerate([ふりがな1, ふりがな2]):
                        input_tag = matched_inputs[i]
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name:
                            element = (
                                driver.find_element(By.NAME, input_tag.get("name"))
                                if input_tag.get("name")
                                else driver.find_element(By.ID, input_tag.get("id"))
                            )
                            if element.get_attribute("value").strip():
                                continue
                            element.clear()
                            element.send_keys(val)
                            used_names.add(input_name)
                    print(f"[{url}] → ふりがなを分割して入力: {ふりがな1}, {ふりがな2}")
                    log_lines.append(f"[{url}] → ふりがなを分割して入力: {ふりがな1}, {ふりがな2}")
                    furi_hiragana_inputs_found = True
        except Exception as e:
            print(f"[{url}] → ふりがな入力エラー（直接探索）: {e}")
            log_lines.append(f"[{url}] → ふりがな入力エラー（直接探索）: {e}")

        # --- 2) ラベルテキスト探索 ---
        if not furi_hiragana_inputs_found:
            furi_targets = soup.find_all(string=lambda s: s and any(kw in s for kw in furi_keywords))

            for target in furi_targets:
                label_elem = target.find_parent(['label', 'div', 'span', 'td', 'p', 'dt']) or target.find_parent()
                if not label_elem:
                    continue

                # 2-1) 親要素内探索
                parent = label_elem.parent
                if parent:
                    if try_fill(parent.find_all("input", {"type": "text"})):
                        break

                # 2-2) 兄弟要素探索
                next_elem = label_elem.find_next_sibling()
                while next_elem:
                    if try_fill(next_elem.find_all("input", {"type": "text"})):
                        break
                    next_elem = next_elem.find_next_sibling()
                if furi_hiragana_inputs_found:
                    break

                # 2-3) 階層的追加探索
                if try_fill(find_input_near_label(label_elem)):
                    break

        if not furi_hiragana_inputs_found:
            print(f"[{url}] → ふりがな欄が見つかりません")
            log_lines.append(f"[{url}] → ふりがな欄が見つかりません")

        

        # フリガナ入力
        if not furi_hiragana_inputs_found:
            #　フリガナ入力
            # --- ヘルパー関数 ---
            def find_input_near_label(label_elem):
                inputs = label_elem.find_all("input", {"type": "text"})
                if inputs:
                    return inputs
                for sibling in label_elem.find_next_siblings():
                    inputs = sibling.find_all("input", {"type": "text"})
                    if inputs:
                        return inputs
                parent = label_elem.parent
                while parent:
                    for sibling in parent.find_next_siblings():
                        inputs = sibling.find_all("input", {"type": "text"})
                        if inputs:
                            return inputs
                    parent = parent.parent
                return []

            # --- フリガナ入力処理 ---
            furi_katakana_inputs_found = False
            soup = BeautifulSoup(driver.page_source, "html.parser")
            used_names = set()
            furi_katakana_keywords = ["フリガナ", "ヨミ"]

            # --- 1) inputタグを直接探索（name/idベース） ---
            candidate_inputs = soup.find_all("input", {"type": "text"})
            matched_inputs = []
            for input_tag in candidate_inputs:
                input_name = (input_tag.get("name") or input_tag.get("id") or "").lower()
                if any(kw in input_name for kw in furi_katakana_keywords):
                    matched_inputs.append(input_tag)

            try:
                if matched_inputs:
                    if len(matched_inputs) == 1:
                        input_tag = matched_inputs[0]
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name:
                            element = (
                                driver.find_element(By.NAME, input_tag.get("name"))
                                if input_tag.get("name")
                                else driver.find_element(By.ID, input_tag.get("id"))
                            )
                            element.clear()
                            element.send_keys(フリガナ1 + フリガナ2)
                            print(f"[{url}] → フリガナ入力（直接input探索）: {フリガナ1 + フリガナ2}")
                            log_lines.append(f"[{url}] → フリガナ入力（直接input探索）: {フリガナ1 + フリガナ2}")
                            furi_katakana_inputs_found = True
                            used_names.add(input_name)
                    elif len(matched_inputs) >= 2:
                        for i, val in enumerate([フリガナ1, フリガナ2]):
                            input_tag = matched_inputs[i]
                            input_name = input_tag.get("name") or input_tag.get("id")
                            if input_name:
                                element = (
                                    driver.find_element(By.NAME, input_tag.get("name"))
                                    if input_tag.get("name")
                                    else driver.find_element(By.ID, input_tag.get("id"))
                                )
                                element.clear()
                                element.send_keys(val)
                                used_names.add(input_name)
                        print(f"[{url}] → フリガナを分割して入力: {フリガナ1}, {フリガナ2}")
                        log_lines.append(f"[{url}] → フリガナを分割して入力: {フリガナ1}, {フリガナ2}")
                        furi_katakana_inputs_found = True
            except Exception as e:
                print(f"[{url}] → フリガナ入力エラー（直接探索）: {e}")
                log_lines.append(f"[{url}] → フリガナ入力エラー（直接探索）: {e}")

            # --- 2) ラベルテキストから探索 ---
            if not furi_katakana_inputs_found:
                furi_katakana_targets = soup.find_all(string=lambda s: s and any(kw in s for kw in furi_katakana_keywords))

                if furi_katakana_targets:
                    for target in furi_katakana_targets:
                        label_elem = target.find_parent(['label', 'div', 'span', 'td', 'p', 'dt']) or target.find_parent()
                        if not label_elem:
                            continue

                        parent = label_elem.parent
                        inputs_found = False

                        # 2-1) 親要素内探索
                        if parent:
                            inputs = parent.find_all("input", {"type": "text"})
                            for input_tag in inputs:
                                input_name = input_tag.get("name") or input_tag.get("id")
                                if input_name and input_name not in used_names:
                                    try:
                                        element = (
                                            driver.find_element(By.NAME, input_name)
                                            if input_tag.get("name")
                                            else driver.find_element(By.ID, input_name)
                                        )
                                        element.clear()
                                        element.send_keys(フリガナ1 + フリガナ2)
                                        print(f"[{url}] → フリガナ入力（親要素内）: {フリガナ1 + フリガナ2}")
                                        log_lines.append(f"[{url}] → フリガナ入力（親要素内）: {フリガナ1 + フリガナ2}")
                                        inputs_found = True
                                        used_names.add(input_name)
                                        break
                                    except Exception as e:
                                        print(f"[{url}] → フリガナ入力エラー（親要素内）: {e}")
                                        log_lines.append(f"[{url}] → フリガナ入力エラー（親要素内）: {e}")

                        if inputs_found:
                            furi_katakana_inputs_found = True
                            continue

                        # 2-2) ラベルの兄弟要素探索
                        next_elem = label_elem.find_next_sibling()
                        while next_elem:
                            inputs = next_elem.find_all("input", {"type": "text"})
                            if inputs:
                                try:
                                    if len(inputs) == 1:
                                        input_name = inputs[0].get("name") or inputs[0].get("id")
                                        if input_name and input_name not in used_names:
                                            element = (
                                                driver.find_element(By.NAME, input_name)
                                                if inputs[0].get("name")
                                                else driver.find_element(By.ID, input_name)
                                            )
                                            element.clear()
                                            element.send_keys(フリガナ1 + フリガナ2)
                                            print(f"[{url}] → フリガナを1つの欄に入力: {フリガナ1 + フリガナ2}")
                                            log_lines.append(f"[{url}] → フリガナを1つの欄に入力: {フリガナ1 + フリガナ2}")
                                            furi_katakana_inputs_found = True
                                            used_names.add(input_name)
                                            break
                                    elif len(inputs) >= 2:
                                        for i, val in enumerate([フリガナ1, フリガナ2]):
                                            input_name = inputs[i].get("name") or inputs[i].get("id")
                                            if input_name and input_name not in used_names:
                                                element = (
                                                    driver.find_element(By.NAME, input_name)
                                                    if inputs[i].get("name")
                                                    else driver.find_element(By.ID, input_name)
                                                )
                                                element.clear()
                                                element.send_keys(val)
                                                used_names.add(input_name)
                                        print(f"[{url}] → フリガナを分割して入力: {フリガナ1}, {フリガナ2}")
                                        log_lines.append(f"[{url}] → フリガナを分割して入力: {フリガナ1}, {フリガナ2}")
                                        furi_katakana_inputs_found = True
                                        break
                                except Exception as e:
                                    print(f"[{url}] → フリガナ入力エラー（兄弟探索）: {e}")
                                    log_lines.append(f"[{url}] → フリガナ入力エラー（兄弟探索）: {e}")
                            if furi_katakana_inputs_found:
                                break
                            next_elem = next_elem.find_next_sibling()
                        if furi_katakana_inputs_found:
                            continue

                        # 2-3) 階層的な追加探索
                        inputs = find_input_near_label(label_elem)
                        for input_tag in inputs:
                            input_name = input_tag.get("name") or input_tag.get("id")
                            if input_name and input_name not in used_names:
                                try:
                                    element = (
                                        driver.find_element(By.NAME, input_name)
                                        if input_tag.get("name")
                                        else driver.find_element(By.ID, input_name)
                                    )
                                    element.clear()
                                    element.send_keys(フリガナ1 + フリガナ2)
                                    print(f"[{url}] → フリガナ入力（追加探索）: {フリガナ1 + フリガナ2}")
                                    log_lines.append(f"[{url}] → フリガナ入力（追加探索）: {フリガナ1 + フリガナ2}")
                                    furi_katakana_inputs_found = True
                                    used_names.add(input_name)
                                    break
                                except Exception as e:
                                    print(f"[{url}] → フリガナ入力エラー（追加探索）: {e}")
                                    log_lines.append(f"[{url}] → フリガナ入力エラー（追加探索）: {e}")
                        if furi_katakana_inputs_found:
                            continue

            # --- ログ出力 ---
            if not furi_katakana_inputs_found:
                print(f"[{url}] → フリガナ欄が見つかりません")
                log_lines.append(f"[{url}] → フリガナ欄が見つかりません")
        
        # メール入力
        # ヘルパー関数: ラベルの近くから input を探索する
        def find_input_near_label(label_elem):
            inputs = label_elem.find_all("input", {"type": ["text", "email"]})
            if inputs:
                return inputs
            for sibling in label_elem.find_next_siblings():
                inputs = sibling.find_all("input", {"type": ["text", "email"]})
                if inputs:
                    return inputs
            parent = label_elem.parent
            while parent:
                for sibling in parent.find_next_siblings():
                    inputs = sibling.find_all("input", {"type": ["text", "email"]})
                    if inputs:
                        return inputs
                parent = parent.parent
            return []

        # メールアドレス入力
        mail_inputs_found = False
        soup = BeautifulSoup(driver.page_source, "html.parser")
        mail_targets = soup.find_all(string=lambda s: s and any(kw in s for kw in ["メールアドレス", "メール", "送付先", "mail", "メアド"]))
        used_names = set()

        if mail_targets:
            for target in mail_targets:
                label_elem = target.find_parent(['label', 'div', 'span', 'td', 'p', 'dt']) or target.find_parent()
                if not label_elem:
                    continue

                parent = label_elem.parent
                next_elem = label_elem.find_next_sibling()
                inputs_found = False

                # name/id 属性で直接探索（メール関連）
                # --- メールアドレス入力処理 ---
                candidate_inputs = soup.find_all("input", {"type": ["text", "email"]})

                for input_tag in candidate_inputs:
                    # name または id を取得して小文字に統一
                    input_name = (input_tag.get("name") or input_tag.get("id") or "").lower()

                    # メール関連キーワードを含む name/id でなければスキップ
                    if not any(kw in input_name for kw in ["メールアドレス", "メール", "mail", "送付先", "メアド", "email", "e_mail", "e-mail"]):
                        continue

                    # 同じ name/id を使い回さないようにする
                    if input_name in used_names:
                        continue

                    # value がすでに入力されていればスキップ
                    if input_tag.get("value", "").strip():
                        continue

                    try:
                        # name 優先で要素を取得
                        if input_tag.get("name"):
                            element = driver.find_element(By.NAME, input_tag.get("name"))
                        elif input_tag.get("id"):
                            element = driver.find_element(By.ID, input_tag.get("id"))
                        else:
                            continue  # name も id もない場合はスキップ

                        # 入力済みならスキップ
                        if element.get_attribute("value").strip():
                            continue

                        # 入力
                        element.clear()
                        element.send_keys(mail)
                        print(f"[{url}] → メール入力: {mail}")
                        log_lines.append(f"[{url}] → メール入力: {mail}")
                        inputs_found = True
                        mail_inputs_found = True
                        used_names.add(input_name)
                        break

                    except Exception as e:
                        print(f"[{url}] → メール入力エラー: {e}")
                        log_lines.append(f"[{url}] → メール入力エラー: {e}")
                if inputs_found:
                    mail_inputs_found = True
                    continue

                # ラベル兄弟要素を順に探索
                while next_elem:
                    inputs = next_elem.find_all("input", {"type": ["text", "email"]})
                    for input_tag in inputs:
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name and input_name not in used_names:
                            try:
                                if input_tag.get("value", "").strip():
                                    continue
                                element = (
                                    driver.find_element(By.NAME, input_name)
                                    if input_tag.get("name")
                                    else driver.find_element(By.ID, input_name)
                                )
                                if element.get_attribute("value").strip():
                                    continue
                                element.clear()
                                element.send_keys(mail)
                                print(f"[{url}] → メール入力: {mail}")
                                log_lines.append(f"[{url}] → メール入力: {mail}")
                                inputs_found = True
                                used_names.add(input_name)
                                break
                            except Exception as e:
                                print(f"[{url}] → メール入力エラー: {e}")
                                log_lines.append(f"[{url}] → メール入力エラー: {e}")
                    if inputs_found:
                        mail_inputs_found = True
                        break
                    next_elem = next_elem.find_next_sibling()

                if inputs_found:
                    continue

                # 階層を上がって追加探索
                inputs = find_input_near_label(label_elem)
                for input_tag in inputs:
                    input_name = input_tag.get("name") or input_tag.get("id")
                    if input_name and input_name not in used_names:
                        try:
                            if input_tag.get("value", "").strip():
                                continue
                            element = (
                                driver.find_element(By.NAME, input_name)
                                if input_tag.get("name")
                                else driver.find_element(By.ID, input_name)
                            )
                            if element.get_attribute("value").strip():
                                continue
                            element.clear()
                            element.send_keys(mail)
                            print(f"[{url}] → メール入力（追加探索）: {mail}")
                            log_lines.append(f"[{url}] → メール入力（追加探索）: {mail}")
                            inputs_found = True
                            used_names.add(input_name)
                            break
                        except Exception as e:
                            print(f"[{url}] → メール入力エラー: {e}")
                            log_lines.append(f"[{url}] → メール入力エラー: {e}")
                if inputs_found:
                    mail_inputs_found = True
                    continue

                # 親要素内のinput探索
                if parent:
                    inputs = parent.find_all("input", {"type": ["text", "email"]})
                    for input_tag in inputs:
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name and input_name not in used_names:
                            try:
                                if input_tag.get("value", "").strip():
                                    continue
                                element = (
                                    driver.find_element(By.NAME, input_name)
                                    if input_tag.get("name")
                                    else driver.find_element(By.ID, input_name)
                                )
                                if element.get_attribute("value").strip():
                                    continue
                                element.clear()
                                element.send_keys(mail)
                                print(f"[{url}] → メール入力（親要素内）: {mail}")
                                log_lines.append(f"[{url}] → メール入力（親要素内）: {mail}")
                                inputs_found = True
                                used_names.add(input_name)
                                break
                            except Exception as e:
                                print(f"[{url}] → メール入力エラー（親要素内）: {e}")
                                log_lines.append(f"[{url}] → メール入力エラー（親要素内）: {e}")
                    if inputs_found:
                        mail_inputs_found = True
                        continue

        if not mail_inputs_found:
            print(f"[{url}] → メールアドレス欄が見つかりません")
            log_lines.append(f"[{url}] → メールアドレス欄が見つかりません")




        # 電話番号入力
        def find_input_near_label(label_elem):
            inputs = label_elem.find_all("input", {"type": "text"})
            if inputs:
                return inputs
            for sibling in label_elem.find_next_siblings():
                inputs = sibling.find_all("input", {"type": "text"})
                if inputs:
                    return inputs
            parent = label_elem.parent
            while parent:
                for sibling in parent.find_next_siblings():
                    inputs = sibling.find_all("input", {"type": "text"})
                    if inputs:
                        return inputs
                parent = parent.parent
            return []

        電話番号_inputs_found = False
        soup = BeautifulSoup(driver.page_source, "html.parser")
        used_phones = set()
        電話番号_keywords = ["電話番号", "tel", "phone", "number"]

        # --- 1) inputタグを直接探索（name/idベース） ---
        candidate_inputs = soup.find_all("input", {"type": ["text", "tel"]})
        matched_inputs = []
        for input_tag in candidate_inputs:
            input_name = (input_tag.get("name") or input_tag.get("id") or "").lower()
            value = input_tag.get("value") or ""
            if any(kw in input_name for kw in 電話番号_keywords) and not value.strip():
                matched_inputs.append(input_tag)

        try:
            if matched_inputs:
                if len(matched_inputs) == 1:
                    input_tag = matched_inputs[0]
                    input_name = input_tag.get("name") or input_tag.get("id")
                    if input_name:
                        element = (
                            driver.find_element(By.NAME, input_tag.get("name"))
                            if input_tag.get("name")
                            else driver.find_element(By.ID, input_tag.get("id"))
                        )
                        element.clear()
                        element.send_keys(電話番号1 + 電話番号2 + 電話番号3)
                        print(f"[{url}] → 電話番号入力（直接input探索）: {電話番号1 + 電話番号2 + 電話番号3}")
                        log_lines.append(f"[{url}] → 電話番号入力（直接input探索）: {電話番号1 + 電話番号2 + 電話番号3}")
                        電話番号_inputs_found = True
                        used_phones.add(input_name)
                elif len(matched_inputs) >= 3:
                    for i, val in enumerate([電話番号1, 電話番号2, 電話番号3]):
                        input_tag = matched_inputs[i]
                        input_name = input_tag.get("name") or input_tag.get("id")
                        if input_name:
                            element = (
                                driver.find_element(By.NAME, input_tag.get("name"))
                                if input_tag.get("name")
                                else driver.find_element(By.ID, input_tag.get("id"))
                            )
                            element.clear()
                            element.send_keys(val)
                            used_phones.add(input_name)
                    print(f"[{url}] → 電話番号入力（直接input探索・分割）: {電話番号1}, {電話番号2}, {電話番号3}")
                    log_lines.append(f"[{url}] → 電話番号入力（直接input探索・分割）: {電話番号1}, {電話番号2}, {電話番号3}")
                    電話番号_inputs_found = True
        except Exception as e:
            print(f"[{url}] → 電話番号入力エラー（直接探索）: {e}")
            log_lines.append(f"[{url}] → 電話番号入力エラー（直接探索）: {e}")

        # --- 2) テキストラベル探索 ---
        if not 電話番号_inputs_found:
            電話番号_targets = soup.find_all(string=lambda s: s and any(kw in s for kw in ["電話番号"]))
            if 電話番号_targets:
                for target in 電話番号_targets:
                    label_elem = target.find_parent(['label', 'div', 'span', 'td', 'p', 'dt']) or target.find_parent()
                    if not label_elem:
                        continue

                    parent = label_elem.parent
                    inputs_found = False

                    # 2-1) 親要素内
                    if parent:
                        inputs = parent.find_all("input", {"type": "text"})
                        for input_tag in inputs:
                            input_name = input_tag.get("name") or input_tag.get("id")
                            value = input_tag.get("value") or ""
                            if input_name and input_name not in used_phones and not value.strip():
                                try:
                                    element = (
                                        driver.find_element(By.NAME, input_name)
                                        if input_tag.get("name")
                                        else driver.find_element(By.ID, input_name)
                                    )
                                    element.clear()
                                    element.send_keys(電話番号1 + 電話番号2 + 電話番号3)
                                    print(f"[{url}] → 電話番号入力（親要素内）: {電話番号1 + 電話番号2 + 電話番号3}")
                                    log_lines.append(f"[{url}] → 電話番号入力（親要素内）: {電話番号1 + 電話番号2 + 電話番号3}")
                                    電話番号_inputs_found = True
                                    used_phones.add(input_name)
                                    break
                                except Exception as e:
                                    print(f"[{url}] → 電話番号入力エラー（親要素内）: {e}")
                                    log_lines.append(f"[{url}] → 電話番号入力エラー（親要素内）: {e}")
                    if 電話番号_inputs_found:
                        continue

                    # 2-2) 兄弟要素
                    next_elem = label_elem.find_next_sibling()
                    while next_elem:
                        inputs = next_elem.find_all("input", {"type": "text"})
                        if inputs:
                            try:
                                if len(inputs) == 1:
                                    input_name = inputs[0].get("name") or inputs[0].get("id")
                                    value = inputs[0].get("value") or ""
                                    if input_name and input_name not in used_phones and not value.strip():
                                        element = (
                                            driver.find_element(By.NAME, input_name)
                                            if inputs[0].get("name")
                                            else driver.find_element(By.ID, input_name)
                                        )
                                        element.clear()
                                        element.send_keys(電話番号1 + 電話番号2 + 電話番号3)
                                        print(f"[{url}] → 電話番号を1つの欄に入力: {電話番号1 + 電話番号2 + 電話番号3}")
                                        log_lines.append(f"[{url}] → 電話番号を1つの欄に入力: {電話番号1 + 電話番号2 + 電話番号3}")
                                        電話番号_inputs_found = True
                                        used_phones.add(input_name)
                                        break
                                elif len(inputs) >= 3:
                                    for i, val in enumerate([電話番号1, 電話番号2, 電話番号3]):
                                        input_name = inputs[i].get("name") or inputs[i].get("id")
                                        value = inputs[i].get("value") or ""
                                        if input_name and input_name not in used_phones and not value.strip():
                                            element = (
                                                driver.find_element(By.NAME, input_name)
                                                if inputs[i].get("name")
                                                else driver.find_element(By.ID, input_name)
                                            )
                                            element.clear()
                                            element.send_keys(val)
                                            used_phones.add(input_name)
                                    print(f"[{url}] → 電話番号を分割して入力: {電話番号1}, {電話番号2}, {電話番号3}")
                                    log_lines.append(f"[{url}] → 電話番号を分割して入力: {電話番号1}, {電話番号2}, {電話番号3}")
                                    電話番号_inputs_found = True
                                    break
                            except Exception as e:
                                print(f"[{url}] → 電話番号入力エラー（兄弟探索）: {e}")
                                log_lines.append(f"[{url}] → 電話番号入力エラー（兄弟探索）: {e}")
                        if 電話番号_inputs_found:
                            break
                        next_elem = next_elem.find_next_sibling()
                    if 電話番号_inputs_found:
                        continue

                    # 2-3) 階層探索
                    inputs = find_input_near_label(label_elem)
                    for input_tag in inputs:
                        input_name = input_tag.get("name") or input_tag.get("id")
                        value = input_tag.get("value") or ""
                        if input_name and input_name not in used_phones and not value.strip():
                            try:
                                element = (
                                    driver.find_element(By.NAME, input_name)
                                    if input_tag.get("name")
                                    else driver.find_element(By.ID, input_name)
                                )
                                element.clear()
                                element.send_keys(電話番号1 + 電話番号2 + 電話番号3)
                                print(f"[{url}] → 電話番号入力（追加探索）: {電話番号1 + 電話番号2 + 電話番号3}")
                                log_lines.append(f"[{url}] → 電話番号入力（追加探索）: {電話番号1 + 電話番号2 + 電話番号3}")
                                電話番号_inputs_found = True
                                used_phones.add(input_name)
                                break
                            except Exception as e:
                                print(f"[{url}] → 電話番号入力エラー（追加探索）: {e}")
                                log_lines.append(f"[{url}] → 電話番号入力エラー（追加探索）: {e}")
                    if 電話番号_inputs_found:
                        continue

        # --- 最後にログ出力 ---
        if not 電話番号_inputs_found:
            print(f"[{url}] → 電話番号欄が見つかりません")
            log_lines.append(f"[{url}] → 電話番号欄が見つかりません")


    except Exception as e:
        print(f"[{url}] → エラー発生: {e}")

    # --- 処理完了待機 ---
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

driver.quit()
