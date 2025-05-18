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
                self.infinite_scroll(scroll_pause_time=2.0, max_scrolls=100)
                
                # (更新) 新しい DOM 構造に合わせてリンク要素を取得
                job_links = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                        "#__next > div.MuiBox-root[class*='css-'] > div > div[class*='css-'] > div.MuiBox-root[class*='css-'] > div > a"
                    ))
                )
                logger.info(f"{len(job_links)}件の求人リンクが見つかりました")
                
                # 各求人リンクのURLを事前に取得して保存
                job_urls = []
                for link in job_links:
                    try:
                        job_url = link.get_attribute("href")
                        job_urls.append(job_url)
                    except Exception as e:
                        logger.warning(f"求人URLの取得中にエラー: {str(e)}")
                
                logger.info(f"取得したURL数: {len(job_urls)}")
                logger.info(f"取得したURL: {job_urls}")
                
                
                # 各リンクに対応する給与情報を取得する準備
                job_salaries = []
                try:
                    # XPathを使って各求人カードの給与情報を取得
                    for i in range(1, len(job_urls) + 1):
                        try:
                            # 提供されたXPathパターンを使用
                            xpath = f"/html/body/div[1]/div[1]/div/div[1]/div[{i}]/div/a/div[2]/div[2]/div[1]/span"
                            
                            # 要素の取得を試みる
                            salary_elems = self.driver.find_elements(By.XPATH, xpath)
                            
                            if salary_elems and len(salary_elems) > 0:
                                salary_text = salary_elems[0].text.strip()
                                # 「円」が含まれる値のみを追加
                                if "円" in salary_text:
                                    job_salaries.append(salary_text)
                                    logger.info(f"求人 {i} の給与情報: {salary_text}")
                                else:
                                    job_salaries.append("")
                                    logger.info(f"求人 {i} の給与情報に「円」が含まれていないためスキップ")
                            else:
                                job_salaries.append("")
                                logger.warning(f"求人 {i} の給与情報要素が見つかりませんでした")
                        except Exception as e:
                            job_salaries.append("")
                            logger.warning(f"求人 {i} の給与情報取得中にエラー: {str(e)}")
                    
                    logger.info(f"取得した給与情報数: {len(job_salaries)}")
                except Exception as e:
                    logger.warning(f"給与情報の取得に失敗: {str(e)}")
                
                # URLごとに詳細ページにアクセスして情報を取得
                for i, job_url in enumerate(job_urls):
                    try:
                        if i >= 2:
                            break  
                        logger.info(f"求人 {i+1}/{len(job_urls)} の情報を取得中...")
                        
                        # 求人詳細ページに遷移
                        self.driver.get(job_url)
                        # ページ読み込みのために3秒待機
                        time.sleep(3)

                        # 各項目の初期化
                        job_data = {
                            "企業名": "",
                            "給与": job_salaries[i] if i < len(job_salaries) else "",  # 給与情報を事前に取得した値から割り当て
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
                            "採用人数": "",
                            "希望度": "個別で記入",
                            "応募資格": "個別で記入",
                            "結果": "個別で記入",
                            "HPの作りこみ": "個別で記入",
                            "転職会議の点数": "個別で記入",
                            "ライトハウス": "個別で記入",
                        }
                        
                        # 詳細情報を取得するロジックを試行
                        try:
                            # 詳細項目を取得
                            detail_items = self.driver.find_elements(By.CSS_SELECTOR, 
                                "#__next > div.MuiBox-root[class*='css-'] > div > div.MuiContainer-root[class*='css-'] > div > div > div > div[class*='css-'] > div")
                            
                            # DOM構造を確認し、該当する詳細情報を取得
                            for item in detail_items:
                                try:
                                    item_text = item.text
                                    
                                    if "勤務地" in item_text:
                                        job_data["勤務地"] = item_text.replace("勤務地：", "").strip()
                                    elif "時間" in item_text:
                                        job_data["時間"] = item_text.replace("時間：", "").strip()
                                    elif "働き方" in item_text:
                                        job_data["働き方"] = item_text.replace("働き方：", "").strip()
                                    
                                    # 言語情報の取得（タグから）
                                    language_tags = self.driver.find_elements(By.CSS_SELECTOR, ".card-tag__item")
                                    if language_tags:
                                        languages = [tag.text for tag in language_tags]
                                        job_data["利用言語"] = ", ".join(languages)
                                except Exception as e:
                                    logger.warning(f"詳細項目の取得中にエラー: {str(e)}")
                        except Exception as e:
                            logger.warning(f"詳細項目の全体取得に失敗: {str(e)}")
                        
                        # 詳細ページへアクセスして追加情報を取得
                        self.get_detailed_info(job_url, job_data)
                        
                        all_job_data.append(job_data)
                        
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
            # 現在のURLを保存
            current_url = self.driver.current_url
            
            # 同じタブで詳細ページにアクセス（新しいタブを開かない）
            logger.info(f"詳細ページにアクセス: {job_url}")
            self.driver.get(job_url)
            
            # ページ読み込みのために待機
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            
            # try:
            #     # 会社情報セクションを取得
            #     company_info = self.driver.find_elements(By.CSS_SELECTOR, ".job-offer-company-details__list-item")
            #     logger.info(f"会社情報セクション: {company_info}")
            #     for info in company_info:
            #         try:
            #             info_text = info.text
                        
            #             if "社員数" in info_text:
            #                 job_data["社員数"] = info_text.replace("社員数", "").strip()
            #             elif "設立年" in info_text or "創業" in info_text:
            #                 job_data["設立年数"] = info_text.strip()
            #             elif "平均年齢" in info_text:
            #                 job_data["平均年齢"] = info_text.replace("平均年齢", "").strip()
            #             elif "残業時間" in info_text:
            #                 job_data["平均残業"] = info_text.replace("平均残業時間", "").strip()
            #             elif "休日日数" in info_text:
            #                 job_data["休日日数"] = info_text.replace("年間休日日数", "").strip()
            #             elif "みなし残業" in info_text:
            #                 job_data["みなし残業"] = info_text.replace("みなし残業", "").strip()
            #         except Exception as e:
            #             logger.warning(f"会社情報項目の処理中にエラー: {str(e)}")
                
            #     # 求人要件情報の取得
            #     requirements = self.driver.find_elements(By.CSS_SELECTOR, ".job-offer-requirements__box")
                
            #     for req in requirements:
            #         try:
            #             req_title = req.find_element(By.CSS_SELECTOR, ".job-offer-requirements__label").text
            #             req_content = req.find_element(By.CSS_SELECTOR, ".job-offer-requirements__content").text
                        
            #             if "必須経験" in req_title or "必要経験" in req_title:
            #                 job_data["実務経験"] = req_content.strip()
            #         except Exception as e:
            #             logger.warning(f"求人要件項目の処理中にエラー: {str(e)}")
                

            # except Exception as e:
            #     logger.warning(f"詳細情報取得中にエラー: {str(e)}")

            # 詳細ページから必要な情報を取得（get_field_valueメソッドを使用）
            try:
                # 企業名取得（詳細ページから取得するとより正確）
                company_name = self.get_field_value("企業名")
                if company_name:
                    job_data["企業名"] = company_name
                else:
                    # 複数の方法で企業名を取得（バックアップ）
                    try:
                        company_elem = self.driver.find_element(
                            By.CSS_SELECTOR,
                            "#__next > div.MuiBox-root[class*='css-'] div[class*='MuiContainer-root'] aside div[class*='MuiCard-root'] a div[class*='MuiCardContent-root'] h6"
                        )
                        job_data["企業名"] = company_elem.text.strip()
                    except Exception as e:
                        logger.warning(f"企業名要素の取得に失敗: {e}")
                
                # 各情報の取得
                if not job_data["給与"]:
                    salary = self.get_field_value("年収")
                    if salary and "円" in salary:
                        job_data["給与"] = salary
                
                if not job_data["勤務地"]:
                    location = self.get_field_value("勤務地")
                    if location:
                        job_data["勤務地"] = location
                
                work_time = self.get_field_value("勤務時間")
                if work_time:
                    job_data["時間"] = work_time

                holiday = self.get_field_value("休日・休暇")
                if holiday:
                    job_data["休日日数"] = holiday

                benefits = self.get_field_value("待遇・福利厚生")
                if benefits:
                    job_data["待遇・福利厚生"] = benefits
                
                work_style = self.get_field_value("働き方")
                if work_style:
                    job_data["働き方"] = work_style

                number_of_employees = self.get_field_value("採用人数")
                if number_of_employees:
                    job_data["採用人数"] = number_of_employees
                
                application_qualification = self.get_field_value("応募資格")
                if application_qualification:
                    job_data["応募資格"] = application_qualification
                
                try:
                    # 指定されたXPathを使用して利用言語を取得
                    xpath = "//*[@id=\"__next\"]/div[1]/div/div[1]/div/div/div/div[1]/div[3]/div[4]/span"
                    language_elems = self.driver.find_elements(By.XPATH, xpath)
                    
                    if language_elems and len(language_elems) > 0:
                        # 複数の言語要素がある場合は結合
                        languages = [elem.text.strip() for elem in language_elems if elem.text.strip()]
                        if languages:
                            job_data["利用言語"] = ", ".join(languages)
                            logger.info(f"XPathで取得した利用言語: {job_data['利用言語']}")
                except Exception as e:
                    logger.warning(f"XPathによる利用言語取得中にエラー: {str(e)}")
                
            except Exception as e:
                logger.warning(f"get_field_valueによる詳細情報取得中にエラー: {str(e)}")
                
                # フォールバック：旧手法で情報を取得
                try:
                    # 企業名取得（新しいDOM構造）
                    detail_company_name = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#__next > div.MuiBox-root[class*='css-'] div[class*='MuiContainer-root'] aside div[class*='MuiCard-root'] a div[class*='MuiCardContent-root'] h6"
                    ).text.strip()
                    if detail_company_name and not job_data["企業名"]:
                        job_data["企業名"] = detail_company_name
                except Exception as e:
                    logger.warning(f"詳細ページの企業名取得に失敗: {e}")

                # try:
                #     # 年収ラベル確認
                #     label_salary = self.driver.find_element(
                #         By.CSS_SELECTOR,
                #         "#__next > div.MuiBox-root[class*='css-'] > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters[class*='css-'] > div > div > div > div[class*='css-'] > div:nth-child(4) > p.MuiTypography-root.MuiTypography-body2[class*='css-']"
                #     )
                #     if "年収" in label_salary.text and not job_data["給与"]:
                #         salary_elem = self.driver.find_element(
                #             By.CSS_SELECTOR,
                #             "#__next > div.MuiBox-root[class*='css-'] > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters[class*='css-'] > div > div > div > div[class*='css-'] > div:nth-child(8) > p.MuiTypography-root.MuiTypography-body2.MuiTypography-alignJustify[class*='css-']"
                #         )
                #         job_data["給与"] = salary_elem.text.strip()
                # except Exception as e:
                #     logger.warning(f"給与取得中にエラー: {e}")

                # try:
                #     # 勤務地ラベル確認
                #     label_location = self.driver.find_element(
                #         By.CSS_SELECTOR,
                #         "#__next > div.MuiBox-root[class*='css-'] > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters[class*='css-'] > div > div > div > div[class*='css-'] > div:nth-child(11) > p.MuiTypography-root.MuiTypography-body2.MuiTypography-alignJustify[class*='css-']"
                #     )
                #     if "勤務地" in label_location.text and not job_data["勤務地"]:
                #         loc_elem = self.driver.find_element(
                #             By.CSS_SELECTOR,
                #             "#__next > div.MuiBox-root[class*='css-'] > div > div.MuiContainer-root.MuiContainer-maxWidthMd.MuiContainer-disableGutters[class*='css-'] > div > div > div > div[class*='css-'] > div:nth-child(11) > p.MuiTypography-root.MuiTypography-body2.MuiTypography-alignJustify[class*='css-']"
                #         )
                #         job_data["勤務地"] = loc_elem.text.strip()
                # except Exception as e:
                #     logger.warning(f"勤務地取得中にエラー: {e}")

            # 会社情報の取得
            # 会社情報のリンクを取得して遷移
            try:
                # 指定されたXPathを持つaリンクを探す
                company_link = self.driver.find_element(By.XPATH, "/html/body/div[1]/header/div[3]/div[2]/nav/div/div/a[1]")
                
                # リンクのテキストを取得
                link_text = company_link.text.strip()
                logger.info(f"取得したリンクのテキスト: {link_text}")
                
                # リンクをクリックして遷移
                company_link.click()
                
                # ページ遷移後に待機
                time.sleep(2)
                
                logger.info("会社情報ページに遷移しました")

                # 会社情報ページのデータを取得
            #     try:
            #         # 設立年数の取得
            #         establishment_years_elem = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[2]/div/div[6]/div/p")
            #         job_data["設立年数"] = establishment_years_elem.text.strip()
            #         logger.info(f"設立年数: {job_data['設立年数']}")
            #     except Exception as e:
            #         logger.warning(f"設立年数の取得中にエラーが発生しました: {str(e)}")
            #         job_data["設立年数"] = ""

            #     # 社員数の取得
            #     try:
            #         # 社員数の要素を取得
            #         employee_count_elem = self.driver.find_element(By.XPATH, "//*[@id='__next']/div[1]/div/div[2]/div[2]/div/div[11]/div/p")
            #         job_data["社員数"] = employee_count_elem.text.strip()
            #         logger.info(f"社員数: {job_data['社員数']}")
            #     except Exception as e:
            #         logger.warning(f"社員数の取得中にエラーが発生しました: {str(e)}")
            #         job_data["社員数"] = ""
                
            #     # 平均年齢の取得
            #     try:
            #         # 平均年齢の要素を取得
            #         average_age_elem = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div[2]/div/div[12]/div/p")
            #         job_data["平均年齢"] = average_age_elem.text.strip()
            #         logger.info(f"平均年齢: {job_data['平均年齢']}")
            #     except Exception as e:
            #         logger.warning(f"平均年齢の取得中にエラーが発生しました: {str(e)}")
            #         job_data["平均年齢"] = ""
                job_data = self.get_company_info(job_data)
                logger.info(f"会社情報: {job_data}")
            except Exception as e:
                logger.warning(f"会社情報ページへの遷移中にエラーが発生しました: {str(e)}")
            
            # 元のページに戻る
            logger.info(f"元のURLに戻ります: {current_url}")
            self.driver.get(current_url)
            time.sleep(2)  # ページ遷移のための待機
            
        except Exception as e:
            logger.error(f"詳細ページのアクセス中にエラーが発生しました: {str(e)}")
            try:
                # エラー回復：お気に入りページに戻る
                logger.info("エラー回復：お気に入りページに戻ります")
                self.driver.get(self.favorites_url)
                time.sleep(3)  # ページ読み込みのために待機
            except:
                logger.error("回復失敗：ブラウザセッションが無効です")
    
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

    def get_field_value(self, field_name):
        """フィールド名から値を柔軟に取得する"""
        try:
            # フィールド名を含むラベル要素を検索
            labels = self.driver.find_elements(By.CSS_SELECTOR, "p[class*='css-']")
            for label in labels:
                if field_name in label.text:
                    # 親要素を取得
                    parent = label.find_element(By.XPATH, "./..")
                    # 親要素内の値を持つ要素を取得（通常2番目のp要素）
                    value_elem = parent.find_elements(By.TAG_NAME, "p")[1]
                    return value_elem.text.strip()
            return ""
        except Exception as e:
            logger.warning(f"{field_name}の取得中にエラー: {e}")
            return ""

    def get_company_info(self, job_data):
        """
        会社情報ページから情報を柔軟に取得する
        固定XPathではなく、ラベルテキストを元に情報を特定
        
        Args:
            job_data (dict): 更新する求人データ辞書
        
        Returns:
            dict: 更新された求人データ辞書
        """
        try:
            # 基本コンテナXPath
            base_container = "/html/body/div[1]/div[1]/div/div[2]/div[2]/div"
            
            # コンテナ内の全div要素を取得
            container_divs = self.driver.find_elements(By.XPATH, f"{base_container}/div")
            logger.info(f"コンテナ内のdiv要素数: {len(container_divs)}")
            
            # すべてのdiv要素を処理 （ラベル／値を改行で分割して抽出）
            for i, div in enumerate(container_divs):
                div_text = div.text.strip()
                logger.info(f"div[{i}] テキスト: {div_text}")
                
                # 改行で分割
                lines = div_text.split('\n')
                if len(lines) < 2:
                    continue
                
                label = lines[0].strip()
                value = '\n'.join(lines[1:]).strip()
                
                # ラベルに応じて job_data を更新
                if "設立" in label:
                    job_data["設立年数"] = value
                    logger.info(f"設立年数を格納: {value}")
                    continue
                if "社員数" in label or "従業員数" in label:
                    job_data["社員数"] = value
                    logger.info(f"社員数を格納: {value}")
                    continue
                if "平均年齢" in label:
                    job_data["平均年齢"] = value
                    logger.info(f"平均年齢を格納: {value}")
                    continue
            
            # （必要であればこれまでのSVGアイコン検索やバックアップ処理を後段に残します）
            
            logger.info(f"最終取得情報: 設立年数={job_data.get('設立年数','未取得')}, 社員数={job_data.get('社員数','未取得')}, 平均年齢={job_data.get('平均年齢','未取得')}")
            return job_data
            
        except Exception as e:
            logger.error(f"会社情報の取得中にエラー: {e}")
            return job_data
def main():
    """メイン実行関数"""
    scraper = GreenScraper()
    
    try:
        # ログイン方法の設定
        use_google = False
        
        # 標準入力からの入力が不要な場合はconfigファイルから自動判定
        if HAS_CONFIG:
            # Google認証情報の確認
            has_google_auth = (
                hasattr(config, 'GOOGLE_EMAIL') and config.GOOGLE_EMAIL and 
                hasattr(config, 'GOOGLE_PASSWORD') and config.GOOGLE_PASSWORD
            )
            # Green Japan認証情報の確認
            has_green_auth = (
                hasattr(config, 'EMAIL') and config.EMAIL and 
                hasattr(config, 'PASSWORD') and config.PASSWORD
            )
            
            # 認証情報の優先度に基づいて判断
            if has_google_auth:
                logger.info("Google認証情報が設定されているため、Googleログインを使用します")
                use_google = True
            elif has_green_auth:
                logger.info("Green Japan認証情報が設定されているため、通常ログインを使用します")
            else:
                # 両方未設定の場合は入力を求める
                use_google = input("Google アカウントでログインしますか？ (y/n): ").strip().lower() == 'y'
        else:
            # configファイルがない場合は入力を求める
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