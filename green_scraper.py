"""
Green Japan お気に入りページスクレイピングツール

このスクリプトはGreen Japanのお気に入りページにログインし、
保存した求人情報を取得してExcelファイルに出力します。

取得データ:
    企業名、給与、勤務地、時間、働き方、平均年齢、みなし残業、平均残業、
    休日日数、実務経験、利用言語、掲載ページ、社員数、設立年数など
"""

import os
import time
import datetime
import pandas as pd
import subprocess  # 追加
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import logging
import getpass
from selenium.webdriver.common.keys import Keys
import requests

try:
    import config  # 設定ファイルをインポート
    HAS_CONFIG = True
except ImportError:
    HAS_CONFIG = False

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("scraping.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class GreenScraper:
    """Green Japanのスクレイピングを行うクラス"""
    
    def __init__(self):
        """初期化メソッド - WebDriverの設定とURLの定義"""
        self.base_url = "https://www.green-japan.com"
        self.login_url = f"{self.base_url}/login"
        self.favorites_url = f"{self.base_url}/favorites/sent"
        
        # driver属性を明示的に初期化
        self.driver = None
        
        # Chromeオプションの設定
        self.chrome_options = Options()
        
        # config.pyからヘッドレスモード設定を読み込む
        if HAS_CONFIG and hasattr(config, 'HEADLESS_MODE') and config.HEADLESS_MODE:
            self.chrome_options.add_argument("--headless")
            logger.info("ヘッドレスモードが有効化されました")
        
        # Chromeプロファイルの設定
        self.using_profile = False
        if HAS_CONFIG and hasattr(config, 'USE_CHROME_PROFILE') and config.USE_CHROME_PROFILE:
            if hasattr(config, 'CHROME_PROFILE_PATH') and config.CHROME_PROFILE_PATH:
                profile_path = config.CHROME_PROFILE_PATH
                profile_name = "Default"
                if hasattr(config, 'CHROME_PROFILE_NAME') and config.CHROME_PROFILE_NAME:
                    profile_name = config.CHROME_PROFILE_NAME
                
                # プロファイルパスの設定
                user_data_dir = profile_path
                self.chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
                self.chrome_options.add_argument(f"--profile-directory={profile_name}")
                logger.info(f"Chromeプロファイルを使用: {user_data_dir}\\{profile_name}")
                self.using_profile = True
        
        # 共通のChrome設定
        self.chrome_options.add_argument("--window-size=1920,1080")
        self.chrome_options.add_argument("--disable-gpu")
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        # この設定でGoogleログインのセッション維持を改善
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_options.add_experimental_option("useAutomationExtension", False)
        
        # WebDriverの初期化
        try:
            # ChromeDriverManager.install() がフォルダを返す場合に対応
            driver_path = ChromeDriverManager().install()
            logger.info(f"ChromeDriverManagerが検出したパス: {driver_path}")
            
            # パスがTHIRD_PARTY_NOTICESを含む場合、親ディレクトリを検索
            if "THIRD_PARTY_NOTICES" in driver_path:
                driver_dir = os.path.dirname(driver_path)
                logger.info(f"検索するディレクトリ: {driver_dir}")
                
                # ディレクトリ内でchromedriver.exeを検索
                for root, dirs, files in os.walk(driver_dir):
                    for file in files:
                        if file.endswith("chromedriver.exe"):
                            driver_path = os.path.join(root, file)
                            logger.info(f"見つかったchromedriver.exe: {driver_path}")
                            break
                    if "chromedriver.exe" in files:
                        break
            
            # パスが/を含む場合、Windowsの\に変換
            driver_path = driver_path.replace("/", "\\")
            
            logger.info(f"使用するドライバーパス: {driver_path}")
            
            try:
                self.driver = webdriver.Chrome(
                    service=Service(driver_path),
                    options=self.chrome_options
                )
                # Selenium検出を回避するためのJavaScriptを実行
                self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            except WebDriverException as e:
                error_msg = str(e)
                # プロファイルが使用中の場合は、リモートデバッグ接続を試みる
                if "user data directory is already in use" in error_msg and HAS_CONFIG and getattr(config, 'USE_CHROME_PROFILE', False):
                    logger.info("プロファイルロック検出: リモートデバッグ接続を試みます")
                    
                    # リモートデバッグポート
                    debug_port = getattr(config, 'REMOTE_DEBUGGING_PORT', 9222)
                    
                    # Chrome実行ファイルパスの取得（エラーメッセージ用）
                    chrome_path = getattr(config, 'CHROME_EXECUTABLE_PATH', None)
                    if not chrome_path or not os.path.exists(chrome_path):
                        # 代替パスを試す
                        for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                                  r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]:
                            if os.path.exists(p):
                                chrome_path = p
                                break
                    
                    if not chrome_path:
                        chrome_path = "chrome.exe"  # パスが見つからない場合は単にchrome.exeとする
                    
                    # プロファイル情報の取得（エラーメッセージ用）
                    profile_path = getattr(config, 'CHROME_PROFILE_PATH', '')
                    profile_name = getattr(config, 'CHROME_PROFILE_NAME', 'Default')
                    
                    # 日本語パスを含むプロファイルの場合のフォールバック
                    if "ゆうと" in profile_path or any(ord(c) > 127 for c in profile_path):
                        temp_profile_dir = r"C:\Temp_Chrome_Debug"
                    else:
                        # 元のプロファイルパスを使用
                        temp_profile_dir = profile_path
                    
                    # まず既存のChromeへの接続を試みる
                    try:
                        # 既存の接続を確認
                        try:
                            logger.info(f"既存のChromeプロセスを確認中: http://127.0.0.1:{debug_port}/json")
                            response = requests.get(f"http://127.0.0.1:{debug_port}/json", timeout=3)
                            if response.status_code == 200:
                                logger.info("既存のChromeプロセスに接続できます。新規ウィンドウは開きません。")
                                
                                # リモートデバッグ接続用のオプション
                                fallback_options = Options()
                                fallback_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
                                
                                # リモートデバッグ接続
                                logger.info(f"既存Chromeへリモートデバッグ接続: 127.0.0.1:{debug_port}")
                                self.driver = webdriver.Chrome(
                                    service=Service(driver_path),
                                    options=fallback_options
                                )
                                logger.info("既存Chromeへの接続に成功しました")
                            else:
                                logger.error(f"既存Chromeへの接続失敗: ステータスコード {response.status_code}")
                                raise ConnectionError(f"既存Chromeへの接続に失敗しました。接続先: http://127.0.0.1:{debug_port}/json")
                        except requests.exceptions.ConnectionError:
                            # 自動でリモートデバッグ用Chromeを起動（共通オプションを追加）
                            logger.info("既存のChromeリモートデバッグが有効ではありません。デバッグ用Chromeを自動起動します。")
                            chrome_command = [
                                chrome_path,
                                f"--remote-debugging-port={debug_port}",
                                f"--user-data-dir={temp_profile_dir}",
                                f"--profile-directory={profile_name}",
                                "--window-size=1920,1080",
                                "--disable-gpu",
                                "--no-sandbox",
                                "--disable-dev-shm-usage",
                                "--disable-blink-features=AutomationControlled",
                                "--exclude-switches=enable-automation",
                                "--disable-extensions",
                                "--no-first-run",
                                "--no-default-browser-check"
                            ]
                            logger.info(f"起動コマンド: {' '.join(chrome_command)}")
                            subprocess.Popen(chrome_command)
                            # プロセス起動を待機
                            time.sleep(5)
                            # 自動起動したChromeへの接続
                            fallback_options = Options()
                            fallback_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
                            fallback_options.add_argument("--disable-dev-shm-usage")
                            fallback_options.add_argument("--no-sandbox")
                            self.driver = webdriver.Chrome(
                                service=Service(driver_path),
                                options=fallback_options
                            )
                            logger.info("自動起動したChromeへの接続に成功しました")
                        except Exception as check_error:
                            logger.error(f"既存Chrome確認中にエラー: {str(check_error)}")
                            raise
                        
                    except Exception as connect_error:
                        # エラーがConnectionErrorの場合はそのまま伝播させる（すでにメッセージを表示済み）
                        if isinstance(connect_error, ConnectionError):
                            raise
                        
                        # その他のエラーの場合は追加情報を表示
                        logger.error(f"Chrome接続エラー: {str(connect_error)}")
                        logger.error("自動的にChromeを起動せず、エラーで終了します。")
                        raise
                else:
                    logger.error(f"ChromeDriverの初期化中にエラーが発生しました: {error_msg}")
                    raise
            
        except Exception as e:
            logger.error(f"ChromeDriverの初期化中にエラーが発生しました: {str(e)}")
            raise

        # ドライバーの初期化確認
        if self.driver is None:
            logger.error("WebDriverの初期化に失敗しました")
            raise Exception("WebDriverの初期化に失敗しました")
            
        # タイムアウト時間を延長（30秒）
        self.wait = WebDriverWait(self.driver, 30)
        
        # データ保存用のディレクトリ作成
        today = datetime.datetime.now().strftime("%Y%m%d")
        self.output_dir = f"output_{today}"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
    def login(self, use_google=False):
        """Green Japanにログインする"""
        # ヘッダー要素でログイン状態を確認
        self.driver.get(self.base_url)
        time.sleep(3)
        try:
            header_elem = self.driver.find_element(
                By.CSS_SELECTOR,
                "#js-react-header > header > div > nav > div.js-header-menu-target.mdl-navigation__link > div"
            )
            if "阪" in header_elem.text:
                logger.info("ヘッダー要素に '阪' が含まれており、ログイン済みと判断します。")
                return True
        except Exception as e:
            logger.debug(f"ヘッダー要素確認失敗: {str(e)}")
        
        # プロファイルを使用している場合は、ログイン済みの可能性が高いのでチェック
        if self.using_profile:
            logger.info("Chromeプロファイルを使用しているため、ログイン状態を確認します...")
            # まずホームページを開く
            self.driver.get(self.base_url)
            time.sleep(3)
            
            # マイページなどのリンクがあるかチェック
            try:
                my_page_link = self.driver.find_elements(By.XPATH, "//a[contains(@href, '/mypage')]")
                if my_page_link:
                    logger.info("すでにログインしています。ログイン処理をスキップします。")
                    return True
                
                # 他のログイン済み要素の確認
                logged_in_elements = self.driver.find_elements(By.CSS_SELECTOR, ".header-utility__login-status")
                if logged_in_elements:
                    logger.info("すでにログインしています。ログイン処理をスキップします。")
                    return True
                    
                logger.info("ログインしていないようです。ログイン処理を実行します。")
            except Exception as e:
                logger.warning(f"ログイン状態の確認中にエラー: {str(e)}")
                logger.info("ログイン処理を実行します。")
        
        # 通常のログイン処理を実行
        if use_google:
            return self.login_with_google()
        else:
            # 既存のメールアドレスとパスワードでのログイン処理
            return self.login_with_email_password()
    
    def login_with_google(self):
        """Google アカウントでログインする"""
        try:
            logger.info("Google アカウントでログインを試みています...")
            self.driver.get(self.login_url)
            time.sleep(3)  # ページの読み込みを待機

            # Google アカウントでログインボタンをクリック
            logger.info("Googleログインボタンを探しています...")
            
            # より正確なセレクタを使用（CSSセレクタ）
            try:
                # CSSセレクタを使ってボタンを特定
                google_login_button = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#content_cont > div.wrap640 > div > form > div > div.mt30 > a.social-login-button.google-button"))
                )
                logger.info("CSSセレクタでGoogleログインボタンを見つけました")
            except Exception as e:
                logger.warning(f"CSSセレクタでのボタン特定に失敗: {str(e)}")
                # フォールバック: クラス名で検索
                try:
                    google_login_button = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.social-login-button.google-button"))
                    )
                    logger.info("クラス名でGoogleログインボタンを見つけました")
                except Exception as e:
                    logger.warning(f"クラス名でのボタン特定に失敗: {str(e)}")
                    # 最後の手段: テキスト内容で検索
                    google_login_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Google')]"))
                    )
                    logger.info("テキスト内容でGoogleログインボタンを見つけました")
            
            logger.info("Googleログインボタンをクリック...")
            google_login_button.click()
            time.sleep(3)  # リダイレクトを待機

            # 新しいウィンドウが開いたか確認
            if len(self.driver.window_handles) > 1:
                logger.info("新しいウィンドウに切り替えます...")
                self.driver.switch_to.window(self.driver.window_handles[-1])
            
            # Googleのログインページに遷移
            current_url = self.driver.current_url
            logger.info(f"現在のURL: {current_url}")
            
            if "accounts.google.com" in current_url:
                # Googleアカウントのメールアドレスを入力
                logger.info("メールアドレス入力フィールドを待機しています...")
                email_field = self.wait.until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
                )
                
                # configファイルからGoogleメールアドレスを取得
                email = ""
                if HAS_CONFIG and hasattr(config, 'GOOGLE_EMAIL') and config.GOOGLE_EMAIL:
                    email = config.GOOGLE_EMAIL
                    logger.info("設定ファイルからGoogleメールアドレスを読み込みました")
                
                # 設定が無い場合はユーザー入力を求める
                if not email:
                    email = input("Googleアカウントのメールアドレスを入力してください: ")
                
                email_field.clear()
                email_field.send_keys(email)
                email_field.send_keys(Keys.RETURN)
                time.sleep(3)
                
                # パスワードを入力
                logger.info("パスワード入力フィールドを待機しています...")
                password_field = self.wait.until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
                )
                
                # configファイルからGoogleパスワードを取得
                password = ""
                if HAS_CONFIG and hasattr(config, 'GOOGLE_PASSWORD') and config.GOOGLE_PASSWORD:
                    password = config.GOOGLE_PASSWORD
                    logger.info("設定ファイルからGoogleパスワードを読み込みました")
                
                # 設定が無い場合はユーザー入力を求める
                if not password:
                    password = getpass.getpass("Googleアカウントのパスワードを入力してください: ")
                
                password_field.clear()
                password_field.send_keys(password)
                password_field.send_keys(Keys.RETURN)
                time.sleep(5)
                
                # 認証が完了するまで待機
                logger.info("認証が完了するまで待機しています...")
                max_wait_time = 60  # 最大待機時間（秒）
                wait_interval = 3    # 確認間隔（秒）
                elapsed_time = 0
                
                while elapsed_time < max_wait_time:
                    current_url = self.driver.current_url
                    if self.base_url in current_url:
                        logger.info("認証が完了し、Green Japanに戻りました")
                        return True
                    
                    time.sleep(wait_interval)
                    elapsed_time += wait_interval
                    logger.info(f"認証待機中... 経過時間: {elapsed_time}秒")
                
                # タイムアウト
                logger.error("認証のタイムアウトが発生しました")
                return False
            else:
                logger.error(f"Googleアカウント認証ページに遷移できませんでした: {current_url}")
                return False

        except Exception as e:
            logger.error(f"Google アカウントでのログイン中にエラーが発生しました: {str(e)}")
            # エラー発生時にスクリーンショットを保存
            try:
                screenshot_path = os.path.join(self.output_dir, "error_screenshot.png")
                self.driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
            except:
                pass
            return False
    
    def login_with_email_password(self):
        """
        Green Japanにログインする
        
        Returns:
            bool: ログイン成功ならTrue、失敗ならFalse
        """
        try:
            logger.info("ログインページにアクセスしています...")
            self.driver.get(self.login_url)
            
            # config.pyからログイン情報を読み込む
            email = ""
            password = ""
            
            if HAS_CONFIG and hasattr(config, 'EMAIL') and config.EMAIL:
                email = config.EMAIL
                
            if HAS_CONFIG and hasattr(config, 'PASSWORD') and config.PASSWORD:
                password = config.PASSWORD
            
            # メールアドレスとパスワードの入力を要求（未指定の場合）
            if not email:
                email = input("Green Japanのメールアドレスを入力してください: ")
            if not password:
                password = getpass.getpass("Green Japanのパスワードを入力してください: ")
            
            # メールアドレス入力
            email_field = self.wait.until(EC.presence_of_element_located((By.ID, "user_email")))
            email_field.clear()
            email_field.send_keys(email)
            
            # パスワード入力
            password_field = self.driver.find_element(By.ID, "user_password")
            password_field.clear()
            password_field.send_keys(password)
            
            # ログインボタンをクリック
            login_button = self.driver.find_element(By.NAME, "commit")
            login_button.click()
            
            # ログイン成功の確認
            self.wait.until(lambda driver: self.base_url in driver.current_url)
            logger.info("ログインに成功しました")
            return True
            
        except Exception as e:
            logger.error(f"ログイン中にエラーが発生しました: {str(e)}")
            return False
    
    def scrape_favorites(self, max_retries=None, retry_delay=None):
        """
        お気に入りページから求人情報をスクレイピングする
        
        Args:
            max_retries (int): スクレイピング失敗時の最大リトライ回数
            retry_delay (int): リトライまでの待機時間（秒）
            
        Returns:
            pd.DataFrame: スクレイピングしたデータのデータフレーム
        """
        # config.pyから設定を読み込む
        if max_retries is None and HAS_CONFIG and hasattr(config, 'MAX_RETRIES'):
            max_retries = config.MAX_RETRIES
        elif max_retries is None:
            max_retries = 2
            
        if retry_delay is None and HAS_CONFIG and hasattr(config, 'RETRY_DELAY'):
            retry_delay = config.RETRY_DELAY
        elif retry_delay is None:
            retry_delay = 5
        
        all_job_data = []
        retry_count = 0
        
        while retry_count <= max_retries:
            try:
                logger.info("お気に入りページにアクセスしています...")
                self.driver.get(self.favorites_url)
                # 動的ロード対応: ページ最下部までスクロールして全件読み込む
                self.infinite_scroll(scroll_pause_time=1.0, max_scrolls=100)
                
                # (更新) 新しい DOM 構造に合わせてリンク要素を取得
                job_links = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                        "#__next > div.MuiBox-root.css-0 > div > div.css-1t1ayi5 > div.MuiBox-root.css-13vg3tq > div > a"
                    ))
                )
                logger.info(f"{len(job_links)}件の求人リンクが見つかりました")
                
                for i, link in enumerate(job_links):
                    try:
                        logger.info(f"求人 {i+1}/{len(job_links)} の情報を取得中...")
                        
                        # リンク要素から URL とテキストを取得
                        job_url = link.get_attribute("href")
                        company_name = link.text.strip()
                        
                        # （必要に応じて）リンク内の詳細情報を取得
                        detail_items = link.find_elements(By.CSS_SELECTOR, ".card-info__detail-item")
                        
                        # 各項目の初期化
                        job_data = {
                            "企業名": company_name,
                            "給与": "",
                            "勤務地": "",
                            "時間": "",
                            "働き方": "",
                            "平均年齢": "",
                            "みなし残業": "",
                            "平均残業": "",
                            "休日日数": "",
                            "実務経験": "",
                            "利用言語": "",
                            "掲載ページ": job_url,
                            "社員数": "",
                            "設立年数": "",
                            "採用予定": "",
                            "希望度": "",
                            "応募": "",
                            "結果": "",
                            "HPの作りこみ": "",
                            "転職会議の点数": "",
                            "ライトハウス": "",
                            "▼働き方特徴": ""
                        }
                        
                        # 各項目の情報を取得
                        for item in detail_items:
                            try:
                                item_text = item.text
                                
                                if "給与" in item_text:
                                    job_data["給与"] = item_text.replace("給与：", "").strip()
                                elif "勤務地" in item_text:
                                    job_data["勤務地"] = item_text.replace("勤務地：", "").strip()
                                elif "時間" in item_text:
                                    job_data["時間"] = item_text.replace("時間：", "").strip()
                                elif "働き方" in item_text:
                                    job_data["働き方"] = item_text.replace("働き方：", "").strip()
                                
                                # 言語情報の取得（タグから）
                                language_tags = card.find_elements(By.CSS_SELECTOR, ".card-tag__item")
                                languages = [tag.text for tag in language_tags]
                                job_data["利用言語"] = ", ".join(languages)
                                
                            except Exception as e:
                                logger.warning(f"項目の取得中にエラー: {str(e)}")
                        
                        # 詳細ページへアクセスして追加情報を取得
                        self.get_detailed_info(job_url, job_data)
                        
                        all_job_data.append(job_data)
                        logger.info(f"求人 {i+1} の情報取得が完了しました: {company_name}")
                        
                    except Exception as e:
                        logger.error(f"求人 {i+1} の処理中にエラーが発生しました: {str(e)}")
                
                # 全ての求人の処理が完了したらループを抜ける
                break
                
            except Exception as e:
                logger.error(f"スクレイピング中にエラーが発生しました: {str(e)}")
                retry_count += 1
                
                if retry_count <= max_retries:
                    logger.info(f"{retry_delay}秒後にリトライします ({retry_count}/{max_retries})...")
                    time.sleep(retry_delay)
                else:
                    logger.error("最大リトライ回数に達しました。処理を終了します。")
                    break
        
        # DataFrameに変換
        return pd.DataFrame(all_job_data)
    
    def get_detailed_info(self, job_url, job_data):
        """
        求人詳細ページにアクセスして追加情報を取得する
        
        Args:
            job_url (str): 求人詳細ページのURL
            job_data (dict): 更新する求人データの辞書
        """
        try:
            # 現在のウィンドウハンドルを保存
            current_window = self.driver.current_window_handle
            
            # 新しいタブで詳細ページを開く
            self.driver.execute_script(f"window.open('{job_url}');")
            
            # 新しいタブに切り替え
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
            # ページの読み込みを待機
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            
            try:
                # 会社情報セクションを取得
                company_info = self.driver.find_elements(By.CSS_SELECTOR, ".job-offer-company-details__list-item")
                
                for info in company_info:
                    try:
                        info_text = info.text
                        
                        if "社員数" in info_text:
                            job_data["社員数"] = info_text.replace("社員数", "").strip()
                        elif "設立年" in info_text or "創業" in info_text:
                            job_data["設立年数"] = info_text.strip()
                        elif "平均年齢" in info_text:
                            job_data["平均年齢"] = info_text.replace("平均年齢", "").strip()
                        elif "残業時間" in info_text:
                            job_data["平均残業"] = info_text.replace("平均残業時間", "").strip()
                        elif "休日日数" in info_text:
                            job_data["休日日数"] = info_text.replace("年間休日日数", "").strip()
                        elif "みなし残業" in info_text:
                            job_data["みなし残業"] = info_text.replace("みなし残業", "").strip()
                    except Exception as e:
                        logger.warning(f"会社情報項目の処理中にエラー: {str(e)}")
                
                # 求人要件情報の取得
                requirements = self.driver.find_elements(By.CSS_SELECTOR, ".job-offer-requirements__box")
                
                for req in requirements:
                    try:
                        req_title = req.find_element(By.CSS_SELECTOR, ".job-offer-requirements__label").text
                        req_content = req.find_element(By.CSS_SELECTOR, ".job-offer-requirements__content").text
                        
                        if "必須経験" in req_title or "必要経験" in req_title:
                            job_data["実務経験"] = req_content.strip()
                    except Exception as e:
                        logger.warning(f"求人要件項目の処理中にエラー: {str(e)}")
                
            except Exception as e:
                logger.warning(f"詳細情報取得中にエラー: {str(e)}")

            # ―――― 追加: 詳細ページから企業名、給与、勤務地を取得 ――――
            try:
                # 企業名取得（新しいDOM構造）
                detail_company_name = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "#__next > div.MuiBox-root.css-0 > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters.css-2hiy9a > div > div > aside > div.MuiPaper-root.MuiPaper-outlined.MuiPaper-rounded.MuiCard-root.css-1sbkbfv > a > div.MuiCardContent-root.css-1qw96cp > h6"
                ).text.strip()
                job_data["企業名"] = detail_company_name
            except Exception as e:
                logger.warning(f"詳細ページの企業名取得に失敗: {e}")

            try:
                # 年収ラベル確認
                label_salary = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "#__next > div.MuiBox-root.css-0 > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters.css-2hiy9a > div > div > div > div.css-78jar2 > div:nth-child(4) > p.MuiTypography-root.MuiTypography-body2.css-1r6p42g"
                )
                if "年収" in label_salary.text:
                    salary_elem = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#__next > div.MuiBox-root.css-0 > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters.css-2hiy9a > div > div > div > div.css-78jar2 > div:nth-child(8) > p.MuiTypography-root.MuiTypography-body2.MuiTypography-alignJustify.css-abo7h2"
                    )
                    job_data["給与"] = salary_elem.text.strip()
            except Exception as e:
                logger.warning(f"給与取得中にエラー: {e}")

            try:
                # 勤務地ラベル確認
                label_location = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "#__next > div.MuiBox-root.css-0 > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters.css-2hiy9a > div > div > div > div.css-78jar2 > div:nth-child(11) > p.MuiTypography-root.MuiTypography-body2.css-1r6p42g"
                )
                if "勤務地" in label_location.text:
                    loc_elem = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#__next > div.MuiBox-root.css-0 > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters.css-2hiy9a > div > div > div > div.css-78jar2 > div:nth-child(11) > p.MuiTypography-root.MuiTypography-body2.MuiTypography-alignJustify.css-abo7h2"
                    )
                    job_data["勤務地"] = loc_elem.text.strip()
            except Exception as e:
                logger.warning(f"勤務地取得中にエラー: {e}")
            # ―――― 追加ここまで ――――

            # 元のタブに戻る
            self.driver.close()
            self.driver.switch_to.window(current_window)
            
        except Exception as e:
            logger.error(f"詳細ページのアクセス中にエラーが発生しました: {str(e)}")
            # エラーが発生しても元のタブに戻るように試みる
            try:
                if len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(current_window)
            except:
                pass
    
    def save_to_excel(self, data):
        """
        スクレイピングしたデータをExcelに保存する
        
        Args:
            data (pd.DataFrame): 保存するデータフレーム
            
        Returns:
            str: 保存したファイルのパス
        """
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(self.output_dir, f"green_jobs_{timestamp}.xlsx")
        
        try:
            # Excelファイルを作成
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            
            # DataFrameをExcelに書き込む（B2セルから開始）
            data.to_excel(writer, sheet_name='求人情報', startrow=1, startcol=1, index=False)
            
            writer.close()
            logger.info(f"データを {file_path} に保存しました")
            return file_path
            
        except Exception as e:
            logger.error(f"Excelへの保存中にエラーが発生しました: {str(e)}")
            return None
    
    def close(self):
        """WebDriverを閉じる"""
        self.driver.quit()
        logger.info("WebDriverを閉じました")

    # ―――――― 無限スクロールメソッドの追加 ――――――
    def infinite_scroll(self, scroll_pause_time: float = 1.0, max_scrolls: int = 100):
        """
        ページ下部までスクロールし、動的ロードされるコンテンツを全件読み込む

        @param scroll_pause_time: スクロール後の待機時間（秒）
        @param max_scrolls: 最大スクロール回数
        """
        last_height = self.driver.execute_script("return document.body.scrollHeight")
        for i in range(max_scrolls):
            # ページ最下部までスクロール
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(scroll_pause_time)
            new_height = self.driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                logger.info(f"スクロール完了: {i+1}回")
                return
            last_height = new_height
        logger.warning(f"最大スクロール回数({max_scrolls})に到達しました")

def main():
    """メイン実行関数"""
    scraper = GreenScraper()
    
    try:
        # ログイン
        use_google = input("Google アカウントでログインしますか？ (y/n): ").strip().lower() == 'y'
        if scraper.login(use_google=use_google):
            # お気に入りページのスクレイピング
            job_data = scraper.scrape_favorites()
            
            # 結果の保存
            if not job_data.empty:
                file_path = scraper.save_to_excel(job_data)
                if file_path:
                    print(f"\n処理が完了しました。データは {file_path} に保存されています。")
            else:
                print("\nスクレイピングされたデータがありません。")
        else:
            print("\nログインに失敗しました。")
    finally:
        # ブラウザを閉じる
        scraper.close()

if __name__ == "__main__":
    main() 