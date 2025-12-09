"""
University Syllabus Scraper (Interactive CLI Version)

概要:
    大学のポータルサイトからシラバスデータを収集するスクレイピングツールです。
    対話形式で年度・キャンパス・学部などの条件を指定し、
    動的なDOM更新（Ajax）を検知して安全にデータを取得します。

作成者: ryujp1
作成日: 2025-12-09
"""

import os
import time
import getpass  # パスワードを隠して入力するためのライブラリ
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# 定数設定 (Configuration)
# ==========================================
# 大学のログインURL（）
URL_LOGIN = ""

# 学部の選択肢リスト（システム上の表記と一致させる）
DEPT_OPTIONS = [
    "指示なし", "コンピュータサイエンス学部",""
]

# キャンパスの選択肢リスト
CAMPUS_OPTIONS = ["指示なし", ""]


# ==========================================
# ユーティリティ関数 (Utility Functions)
# ==========================================

def install_japanese_font():
    """
    Linux環境（Google Colab等）で日本語フォントをインストールします。
    ローカル環境で実行する場合は不要な場合があります。
    """
    if not os.path.exists("/usr/share/fonts/opentype/ipafont-gothic"):
        print("Installing Japanese fonts... (approx. 30 sec)")
        os.system("apt-get -y install fonts-ipafont-gothic")


def safe_send_keys(driver, element_id, text, max_retries=5):
    """
    要素に対して安全にテキストを入力します。
    StaleElementReferenceExceptionが発生した場合、リトライを行います。

    Args:
        driver: Selenium WebDriverのインスタンス
        element_id (str): 対象要素のID属性
        text (str): 入力するテキスト
        max_retries (int): 最大リトライ回数

    Returns:
        bool: 入力に成功した場合はTrue、失敗した場合はFalse
    """
    print(f" 入力試行中...")
    for i in range(max_retries):
        try:
            elem = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.ID, element_id))
            )
            elem.clear()
            elem.send_keys(text)
            print(f"  -> 入力成功")
            return True
        except (StaleElementReferenceException, TimeoutException):
            time.sleep(2)
        except Exception as e:
            print(f"  入力エラー: {e}")
            time.sleep(1)

    print(f"  エラー: {element_id} に入力できませんでした。")
    return False


def safe_select_by_text(driver, element_id, target_text, max_retries=5):
    """
    プルダウンメニューから指定されたテキストを含む項目を選択します。
    画面更新によるStaleエラーに対応しています。

    Args:
        driver: Selenium WebDriverインスタンス
        element_id (str): select要素のID
        target_text (str): 選択したい項目のテキスト（部分一致）
        max_retries (int): 最大リトライ回数

    Returns:
        bool: 選択成功ならTrue
    """
    for i in range(max_retries):
        try:
            elem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            select = Select(elem)
            found_text = None

            # 空白を除去してマッチングを行う
            for opt in select.options:
                if target_text in opt.text.replace(" ", "").replace("　", ""):
                    found_text = opt.text
                    break

            if found_text:
                select.select_by_visible_text(found_text)
                print(f"  -> '{found_text}' を選択成功")
                return True
            else:
                # 選択肢がロードされるのを待つ
                time.sleep(1)
        except Exception:
            time.sleep(1)

    return False


def ask_user_selection(label_name, options_list):
    """
    CLI上でユーザーに選択肢を提示し、入力を受け付けます。

    Args:
        label_name (str): 項目名（例：学部）
        options_list (list): 選択肢のリスト

    Returns:
        str: 選択された項目のテキスト
    """
    print(f"\n【{label_name}】を選択してください:")
    for i, opt in enumerate(options_list):
        print(f"  {i}: {opt}")

    while True:
        try:
            val = input(f">> 番号を入力 (0-{len(options_list)-1}, Enterで0): ")
            if val == "":
                return options_list[0]
            idx = int(val)
            if 0 <= idx < len(options_list):
                return options_list[idx]
        except ValueError:
            pass
        print("エラー: 正しい番号を入力してください。")


def set_dropdown_field(driver, label_text, wait):
    """
    ラベル名から隣接するプルダウンを探し、動的に選択肢を取得してユーザーに入力させます。
    Ajaxによる選択肢の遅延ロードに対応しています。
    """
    xpath = f"//td[contains(text(), '{label_text}')]/following-sibling::td//select"

    # 1. 要素の出現待ち
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        print(f"  (スキップ: {label_text} の項目が見つかりませんでした)")
        return

    # 2. 選択肢のロード待ち（最大10秒）
    print(f"  '{label_text}' の選択肢を読み込んでいます...")
    options_text = []
    for _ in range(10):
        try:
            elem = driver.find_element(By.XPATH, xpath)
            select = Select(elem)
            options_text = [opt.text for opt in select.options]
            if len(options_text) > 1:  # 「指示なし」以外が読み込まれたらOK
                break
            time.sleep(1)
        except Exception:
            time.sleep(1)

    # 3. ユーザー入力と設定
    if not options_text:
        options_text = ["指示なし"]

    selected_text = ask_user_selection(label_text, options_text)

    try:
        elem = driver.find_element(By.XPATH, xpath)
        Select(elem).select_by_visible_text(selected_text)
        print(f"  -> {label_text}: {selected_text} をセットしました")
    except Exception as e:
        print(f"  セットエラー: {e}")


# ==========================================
# メイン処理 (Main Execution)
# ==========================================
def main():
    # --- 0. 環境設定 ---
    # Colab等で必要な場合のみ実行
    # install_japanese_font()

    # --- 1. ユーザー入力（認証情報） ---
    # GitHub公開用に、コード内のハードコーディングを避けて入力させる
    print("\n" + "="*40 + "\n 認証情報の入力\n" + "="*40)
    user_id = input("学籍番号を入力してください: ")
    password = getpass.getpass("パスワードを入力してください (非表示): ")

    # --- 2. 基本条件の設定 ---
    print("\n" + "="*40 + "\n STEP 1: 基本条件の設定\n" + "="*40)
    while True:
        in_year = input(">> 【年度】を入力 (例: 2025): ")
        if in_year.isdigit() and len(in_year) == 4:
            break
        if in_year == "":
            in_year = "2025"
            break

    target_campus = ask_user_selection("キャンパス", CAMPUS_OPTIONS)
    target_dept = ask_user_selection("学部（時間割所属）", DEPT_OPTIONS)

    # --- 3. ブラウザ起動 ---
    options = Options()
    options.add_argument('--headless')  # ヘッドレスモード
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')

    print("\nブラウザを起動中...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- ログイン処理 ---
        driver.get(URL_LOGIN)
        wait.until(EC.visibility_of_element_located((By.NAME, "userName"))).send_keys(user_id)
        driver.find_element(By.NAME, "password").send_keys(password)

        try:
            driver.find_element(By.XPATH, "//button[contains(text(), 'ログイン')]").click()
        except:
            driver.find_element(By.CSS_SELECTOR, ".btn.waves-effect.waves-light").click()

        print("ログイン成功。メニュー移動中...")
        time.sleep(10)

        # --- メニュー移動 ---
        try:
            wait.until(EC.element_to_be_clickable((By.ID, "menu-link-mt-kyomu"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@title='シラバス']"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable((By.ID, "menu-link-mf-156037"))).click()
            print("シラバス検索画面へ移動しました。")
            time.sleep(5)
        except Exception as e:
            print(f"メニュー操作エラー: {e}")
            return

        # フレーム切り替え処理
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
        iframes = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
        if len(iframes) > 0:
            driver.switch_to.frame(0)

        # --- 基本条件の適用 ---
        print("\n基本条件を適用します...")
        if not safe_send_keys(driver, "nendo", in_year):
            print("年度入力失敗のため中断")
            return

        if target_campus != "指示なし":
            safe_select_by_text(driver, "campusCd", target_campus)
            time.sleep(2)

        if target_dept != "指示なし":
            success = safe_select_by_text(driver, "jikanwariShozokuSelect", target_dept)
            if success:
                print("  (画面項目の追加を待機中...)")
                time.sleep(5)

        # --- STEP 2: 詳細条件の入力 ---
        print("\n" + "="*40 + "\n STEP 2: 詳細条件の入力\n" + "="*40)

        # 動的に選択肢を取得する項目
        extra_fields = ["学年", "学期", "曜日", "時限"]
        for label in extra_fields:
            set_dropdown_field(driver, label, wait)

        # 自由入力項目
        sub_name = input("\n【開講科目名】を入力 (入力なしでスキップ): ")
        if sub_name:
            try:
                xpath = f"//td[contains(text(), '開講科目名')]/following-sibling::td//input[@type='text']"
                driver.find_element(By.XPATH, xpath).send_keys(sub_name)
                print(f"  -> 科目名: {sub_name}")
            except Exception:
                pass

        # --- 検索実行 ---
        print("\n検索を実行します...")
        search_btn = driver.find_element(By.XPATH, "//input[contains(@value, '検索開始')]")
        driver.execute_script("arguments[0].click();", search_btn)
        time.sleep(5)

        # --- データ抽出・保存 ---
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        rows = soup.select("table.normal tr")
        if len(rows) == 0:
            rows = soup.select("table tbody tr")

        current_data = []
        for row in rows:
            cells = row.find_all("td")
            if len(cells) < 7:
                continue
            try:
                period = cells[3].text.strip().replace("\n", "")
                subject = cells[5].text.strip()
                teacher = cells[6].text.strip()
                if subject == "":
                    continue

                current_data.append({
                    "年度": in_year,
                    "学部": target_dept,
                    "科目名": subject,
                    "教員名": teacher,
                    "曜日時限": period
                })
            except Exception:
                continue

        print(f"\n検索結果: {len(current_data)} 件")

        if current_data:
            # ファイル名生成（使用できない文字を除去）
            safe_name = target_dept.replace("/", "").replace(" ", "")
            filename = f"syllabus_{in_year}_{safe_name}_search.csv"

            df = pd.DataFrame(current_data)
            df.to_csv(filename, index=False, encoding="utf-8-sig")
            print(f"★保存完了: {filename}")
        else:
            print("データが見つかりませんでした。")

    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        driver.save_screenshot("fatal_error.png")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()