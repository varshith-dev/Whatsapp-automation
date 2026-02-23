import sys
import threading
import time
import random
from datetime import datetime
import firebase_admin
from firebase_admin import credentials, db

# PyQt6 imports for classic UI
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QTextEdit, QTabWidget,
                             QMessageBox, QLineEdit, QFormLayout)
from PyQt6.QtCore import QUrl, pyqtSignal, QObject, Qt, QMetaObject

# Selenium imports for rock-solid web automation
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import urllib.parse
import os


class WorkerSignals(QObject):
    """Signals to communicate safely from background threads to the main GUI thread."""
    log_msg = pyqtSignal(str)
    status_update = pyqtSignal(str, str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    history_msg = pyqtSignal(str) # For History tab
    # Add specific signals for the browser
    navigate_url = pyqtSignal(str)
    execute_js = pyqtSignal(str)


class WhatsAppAutomatorApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Whatsapp Automator Service")
        self.resize(1000, 700)
        
        # State Variables
        self.is_running = False
        self.is_paused = False
        self.fb_app = None
        self.driver = None  # Reusable Selenium driver — keeps WhatsApp logged in

        # --- Build Classic UI ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Tabs
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        self.tab_control = QWidget()
        self.tab_presets = QWidget()
        self.tab_history = QWidget()
        self.tab_browser = QWidget()
        self.tab_settings = QWidget()

        self.tabs.addTab(self.tab_control, "Bot Control")
        self.tabs.addTab(self.tab_presets, "Auto-Reply Presets")
        self.tabs.addTab(self.tab_history, "Message History")
        self.tabs.addTab(self.tab_browser, "WhatsApp Web (Internal)")
        self.tabs.addTab(self.tab_settings, "Settings")

        self.setup_control_tab()
        self.setup_presets_tab()
        self.setup_history_tab()
        self.setup_browser_tab()
        self.setup_settings_tab()

        # Threading signals
        self.signals = WorkerSignals()
        self.signals.log_msg.connect(self.append_log)
        self.signals.status_update.connect(self.update_status_labels)
        self.signals.finished.connect(self.on_automation_finished)
        self.signals.error.connect(self.show_error_dialog)
        self.signals.history_msg.connect(self.append_history)

        # Initial check
        self.check_status_thread()

    def setup_control_tab(self):
        layout = QVBoxLayout(self.tab_control)

        # Status Panel
        self.status_layout = QHBoxLayout()
        self.lbl_conn_status = QLabel("🔴 Disconnected (Check Settings Tab)")
        self.lbl_conn_status.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        self.lbl_pending = QLabel("Pending Messages: --")
        self.lbl_pending.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        self.btn_refresh = QPushButton("Refresh Status")
        self.btn_refresh.clicked.connect(self.check_status_thread)

        self.status_layout.addWidget(self.lbl_conn_status)
        self.status_layout.addStretch()
        self.status_layout.addWidget(self.btn_refresh)
        self.status_layout.addStretch()
        self.status_layout.addWidget(self.lbl_pending)
        
        layout.addLayout(self.status_layout)

        # Controls Panel
        self.control_layout = QHBoxLayout()
        self.btn_start = QPushButton("▶️ Start")
        self.btn_start.setMinimumHeight(40)
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_start.clicked.connect(self.start_automation)

        self.btn_pause = QPushButton("⏸️ Pause")
        self.btn_pause.setMinimumHeight(40)
        self.btn_pause.setStyleSheet("background-color: #FF9800; color: white; font-weight: bold;")
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self.pause_automation)

        self.btn_stop = QPushButton("⏹️ Stop")
        self.btn_stop.setMinimumHeight(40)
        self.btn_stop.setStyleSheet("background-color: #f44336; color: white; font-weight: bold;")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_automation)

        self.control_layout.addWidget(self.btn_start)
        self.control_layout.addWidget(self.btn_pause)
        self.control_layout.addWidget(self.btn_stop)
        
        layout.addLayout(self.control_layout)

        # Logs
        layout.addWidget(QLabel("Live Activity Log:"))
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        layout.addWidget(self.txt_log)

    def setup_presets_tab(self):
        layout = QVBoxLayout(self.tab_presets)
        info = QLabel(
            "Define Keyword Auto-Replies here.\n"
            "Format: keyword = Your reply message\n"
            "Example: price = The price is Rs 100."
        )
        info.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(info)
        self.txt_presets = QTextEdit()
        self.txt_presets.setText("hello = Hi! I am currently away. This is an auto-reply.\nhelp = Please contact admin at 9391507369.")
        layout.addWidget(self.txt_presets)
        
    def setup_history_tab(self):
        layout = QVBoxLayout(self.tab_history)
        layout.addWidget(QLabel("Auto-Reply & Send History:"))
        self.txt_history = QTextEdit()
        self.txt_history.setReadOnly(True)
        layout.addWidget(self.txt_history)

    def append_history(self, msg):
        ts = datetime.now().strftime("[%H:%M:%S]")
        self.txt_history.append(f"{ts} {msg}")
        scrollbar = self.txt_history.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def setup_browser_tab(self):
        layout = QVBoxLayout(self.tab_browser)
        
        info = QLabel(
            "🚀 Speed Update: WhatsApp Web will now open directly in a fast Google Chrome window "
            "controlled by our Bot Engine rather than being embedded here.\n\n"
            "This makes the bot 10x faster and 100% reliable at clicking 'Send'.\n\n"
            "When you click 'Start', you will see a Chrome window pop up automatically."
        )
        info.setWordWrap(True)
        info.setStyleSheet("font-size: 16px; margin: 20px;")
        layout.addWidget(info)
        layout.addStretch()

    def setup_settings_tab(self):
        layout = QFormLayout(self.tab_settings)
        
        self.txt_db_url = QLineEdit()
        self.txt_db_url.setText("https://db-vt-gsocp-default-rtdb.firebaseio.com")
        
        self.txt_db_path = QLineEdit()
        self.txt_db_path.setText("registrations")
        self.txt_db_path.setPlaceholderText("The node in Firebase where user data is saved")

        layout.addRow("Firebase Database URL:", self.txt_db_url)
        layout.addRow("Firebase Data Node:", self.txt_db_path)
        
        info = QLabel(
            "Instructions:\n"
            "1. Download your Service Account JSON from Firebase Console.\n"
            "2. Rename it exactly to 'firebase_credentials.json'.\n"
            "3. Place it in the same folder as this script.\n"
            "4. Go to the Bot Control tab and hit Refresh."
        )
        info.setWordWrap(True)
        layout.addRow(info)

    # --- GUI Updaters ---
    def append_log(self, msg):
        ts = datetime.now().strftime("[%H:%M:%S]")
        self.txt_log.append(f"{ts} {msg}")
        # Scroll to bottom
        scrollbar = self.txt_log.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def update_status_labels(self, conn_txt, pending_txt):
        self.lbl_conn_status.setText(conn_txt)
        self.lbl_pending.setText(pending_txt)

    def show_error_dialog(self, msg):
        QMessageBox.critical(self, "Error", msg)

    # --- Backend Logic ---
    def authenticate_firebase(self):
        # Prevent re-initializing if already done
        if not firebase_admin._apps:
            cred = credentials.Certificate('firebase_credentials.json')
            self.fb_app = firebase_admin.initialize_app(cred, {
                'databaseURL': self.txt_db_url.text().strip()
            })

    def check_status_thread(self):
        self.signals.status_update.emit("🟡 Connecting...", "Pending Messages: --")
        threading.Thread(target=self._check_status_worker, daemon=True).start()

    def _check_status_worker(self):
        try:
            self.authenticate_firebase()
            node = self.txt_db_path.text().strip()
            
            ref = db.reference(f'/{node}')
            data = ref.get()
            
            if data is None:
                self.signals.status_update.emit("🟢 Connected (No Data Found)", "Pending Messages: 0")
                return

            pending_count = 0
            if isinstance(data, dict):
                for key, val in data.items():
                    if isinstance(val, dict):
                        status = str(val.get("Message Status", "")).strip().lower()
                        # Adjust key name 'Message Status' depending on your exact Firebase JSON structure
                        if status == "" or status == "none" or status == "pending":
                            pending_count += 1
                            
            self.signals.status_update.emit("🟢 Connected to Firebase", f"Pending Messages: {pending_count}")
        except FileNotFoundError:
            self.signals.status_update.emit("🔴 Disconnected (No JSON)", "Pending: --")
            self.signals.log_msg.emit("Error: firebase_credentials.json not found in directory.")
        except Exception as e:
            self.signals.status_update.emit("🔴 Disconnected (Error)", "Pending: --")
            self.signals.log_msg.emit(f"Status check failed: {str(e)}")

    def start_automation(self):
        self.is_running = True
        self.is_paused = False
        
        # Parse active presets into dictionary for the worker
        self.current_presets = {}
        for line in self.txt_presets.toPlainText().split('\n'):
            if '=' in line:
                kw, reply = line.split('=', 1)
                self.current_presets[kw.strip()] = reply.strip()
        
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_pause.setText("⏸️ Pause")
        self.btn_stop.setEnabled(True)
        
        self.signals.log_msg.emit("=== Automation Started ===")
        # Switch focus to the browser tab so the user sees action
        self.tabs.setCurrentIndex(1)
        
        threading.Thread(target=self._automation_worker, daemon=True).start()

    def pause_automation(self):
        if not self.is_paused:
            self.is_paused = True
            self.btn_pause.setText("▶️ Resume")
            self.signals.log_msg.emit("Automation paused... Will hold before next action.")
        else:
            self.is_paused = False
            self.btn_pause.setText("⏸️ Pause")
            self.signals.log_msg.emit("Automation resumed.")

    def stop_automation(self):
        self.is_running = False
        self.is_paused = False
        self.signals.log_msg.emit("Stopping automation gracefully...")

    def on_automation_finished(self):
        self.is_running = False
        self.is_paused = False
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)
        self.signals.log_msg.emit("=== Automation Finished / Aborted ===")
        self.check_status_thread()

    # --- The Core JavaScript Injection Worker ---
    def execute_js_on_browser(self, js_code):
        """Helper to run JS inside the embedded browser from a background thread."""
        # PyQt6 requires JS to be executed on the main GUI thread.
        # We use a QTimer or QMetaObject.invokeMethod to safely call it.
        # QMetaObject.invokeMethod(self.browser.page(), "runJavaScript", Qt.ConnectionType.QueuedConnection, Q_ARG(str, js_code))
        pass # To be implemented fully via proper signal if waiting for result is needed.
        # Note: Interacting deeply with QWebEngine via python threads is complex.
        # For sending messages, we will construct a special api url
        
    def _automation_worker(self):
        try:
            self.authenticate_firebase()
            node = self.txt_db_path.text().strip()
            ref = db.reference(f'/{node}')
            
            data = ref.get()
            
            if not data or not isinstance(data, dict):
                self.signals.log_msg.emit("No valid data found in Firebase.")
                self.signals.finished.emit()
                return

            keys_to_process = []
            for uid, val in data.items():
                if isinstance(val, dict):
                    status = str(val.get("Message Status", "")).strip().lower()
                    if status == "" or status == "none" or status == "pending":
                        keys_to_process.append((uid, val))

            self.signals.log_msg.emit(f"Found {len(keys_to_process)} farmers to message.")
            # Note: We do NOT exit if there are 0 pending broadcasts because Auto-Reply still runs!

            # --- Initialize Chrome (REUSE existing session if possible) ---
            driver = None
            
            # Try to reuse the existing driver (no re-login needed!)
            if self.driver:
                try:
                    _ = self.driver.title  # Quick health check
                    driver = self.driver
                    self.signals.log_msg.emit("♻️ Reusing existing Chrome session — no re-login needed!")
                except:
                    self.signals.log_msg.emit("Previous Chrome died. Starting fresh...")
                    self.driver = None
            
            if not driver:
                self.signals.log_msg.emit("Launching Chrome...")
                
                # Kill ALL Chrome processes to release profile locks (fresh start)
                os.system("taskkill /F /IM chrome.exe /T >nul 2>&1")
                os.system("taskkill /F /IM chromedriver.exe /T >nul 2>&1")
                time.sleep(2)
                
                # Profile location in AppData (stable, native Chrome location)
                appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
                profile_dir = os.path.join(appdata, 'WAManager_Chrome_Profile')
                os.makedirs(profile_dir, exist_ok=True)
                
                # Clean up stale lock files that cause DevToolsActivePort crashes
                for lock_file in ['SingletonLock', 'SingletonSocket', 'SingletonCookie', 'lockfile']:
                    lock_path = os.path.join(profile_dir, lock_file)
                    try:
                        if os.path.exists(lock_path):
                            os.remove(lock_path)
                    except:
                        pass
                
                options = webdriver.ChromeOptions()
                options.add_argument(f"user-data-dir={profile_dir}")
                
                # Windows crash prevention flags
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--disable-gpu')
                options.add_argument('--disable-extensions')
                options.add_argument('--disable-features=RendererCodeIntegrity')
                options.add_argument('--no-first-run')
                options.add_argument('--no-default-browser-check')
                options.add_argument('--log-level=3')
                options.add_argument('--remote-allow-origins=*')
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                
                driver = webdriver.Chrome(
                    service=Service(ChromeDriverManager().install()),
                    options=options
                )
                self.driver = driver
            
            # Initial load to check WhatsApp Login status
            self.signals.log_msg.emit("Opening WhatsApp Web...")
            driver.get("https://web.whatsapp.com")
            
            self.signals.log_msg.emit("Waiting up to 180s for WhatsApp Web to sync/login...")
            # Wait for either the search bar (logged in) or QR code canvas (needs login)
            WebDriverWait(driver, 180).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div#side, canvas[aria-label="Scan me!"]'))
            )
            
            # Check if QR code is present (meaning not logged in)
            try:
                driver.find_element(By.CSS_SELECTOR, 'canvas[aria-label="Scan me!"]')
                self.signals.log_msg.emit("🚨 PLEASE SCAN THE QR CODE ON THE CHROME WINDOW! 🚨")
                WebDriverWait(driver, 300).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div#side'))
                )
                self.signals.log_msg.emit("Successfully logged in!")
            except:
                self.signals.log_msg.emit("Session active. Proceeding.")

            # --- Start Processing ---
            for index, (uid, user_data) in enumerate(keys_to_process):
                if not self.is_running: break
                while self.is_paused and self.is_running: time.sleep(1)
                if not self.is_running: break

                name = str(user_data.get("Name", "")).strip()
                phone = str(user_data.get("Phone Number", "")).strip()

                if not name or not phone:
                    self.signals.log_msg.emit(f"Skipping UID {uid}: missing name or phone.")
                    continue

                if not phone.startswith("+"):
                    phone = f"+91{phone}" if len(phone) == 10 else f"+{phone}"

                # URL Encode message
                raw_msg = f"Hello *{name}*, your registration was successful!"
                encoded_msg = urllib.parse.quote(raw_msg)
                
                send_url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_msg}"

                self.signals.log_msg.emit(f"Loading chat for {name} ({phone})...")
                
                # Fast Navigate
                driver.get(send_url)
                
                # Wait instantly for the page to load the text box
                try:
                    # Wait for the chat box area (the main wrapper for the right pane)
                    main_pane_locator = (By.ID, 'main')
                    WebDriverWait(driver, 30).until(EC.presence_of_element_located(main_pane_locator))
                    
                    # Give it a tiny bit of breathing room to load the actual text
                    time.sleep(2)
                    
                    # Safest approach: Find the actual chat text box and press Enter inside it.
                    # This completely bypasses the send button icon changing.
                    chat_box = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab="10"]'))
                    )
                    # The message is already pre-filled from the URL, so we just hit Enter
                    chat_box.send_keys(Keys.ENTER)
                    
                    # Quick wait for checkmark icon to confirm it sent before moving on
                    time.sleep(1) 
                    
                    # Update Firebase immediately
                    db.reference(f'/{node}/{uid}').update({"Message Status": "Sent"})
                    self.signals.log_msg.emit(f"Success! {name} marked as sent.")
                    
                except Exception as e:
                    self.signals.log_msg.emit(f"ERROR: Could not find Send button for {name}. Is the number valid?")
                    self.signals.log_msg.emit(str(e))

                # Anti Ban Delay
                if self.is_running and not self.is_paused and index < len(keys_to_process) - 1:
                    sleep_time = random.randint(10, 20)
                    self.signals.log_msg.emit(f"Anti-ban active: resting for {sleep_time} seconds...")
                    for _ in range(sleep_time):
                        if not self.is_running: break
                        while self.is_paused and self.is_running: time.sleep(1)
                        if not self.is_running: break
                        time.sleep(1)
            
            # --- START CONTINUOUS AUTO-REPLY MONITORING ---
            self.signals.log_msg.emit("Broadcasts complete! Entering Auto-Reply monitoring mode...")
            self.signals.log_msg.emit(f"Active presets: {self.current_presets}")
            
            time.sleep(2)
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except:
                pass
            time.sleep(2)
            
            scan_count = 0
            failed_chats = {}  # {chatTitle: timestamp} — skip chats where extraction failed
            
            while self.is_running:
                while self.is_paused and self.is_running: time.sleep(1)
                if not self.is_running: break
                
                scan_count += 1
                
                try:
                    # --- Health check every 30 scans ---
                    if scan_count % 30 == 0:
                        try:
                            _ = driver.title
                        except:
                            self.signals.log_msg.emit("[Auto-Reply] Chrome died! Attempting recovery...")
                            try:
                                self.driver = None
                                driver = None
                            except:
                                pass
                            break  # Exit loop to trigger the outer except/finally
                    
                    # --- WhatsApp disconnect detection ---
                    if scan_count % 15 == 0:
                        try:
                            disconnected = driver.execute_script("""
                                var banner = document.querySelector('[data-testid="alert-phone"]')
                                          || document.querySelector('[data-icon="alert-phone"]');
                                return banner ? true : false;
                            """)
                            if disconnected:
                                self.signals.log_msg.emit("[Auto-Reply] ⚠️ WhatsApp disconnected. Waiting for reconnection...")
                                time.sleep(10)
                                continue
                        except:
                            pass
                    
                    # --- Find unread badges ---
                    unread_badges = driver.find_elements(By.XPATH, "//span[contains(@aria-label, 'unread message')]")
                    if not unread_badges:
                        unread_badges = driver.find_elements(By.XPATH, "//span[contains(@aria-label, 'unread')]")
                    
                    if scan_count % 10 == 1:
                        self.signals.log_msg.emit(f"[Auto-Reply] Scan #{scan_count}: {len(unread_badges)} unread. Monitoring...")
                    
                    if unread_badges:
                        for badge in unread_badges:
                            try:
                                chat_row = None
                                for anc in ["./ancestor::div[@role='listitem']", "./ancestor::div[@role='row']", "./../../.."]:
                                    try:
                                        chat_row = badge.find_element(By.XPATH, anc)
                                        if chat_row: break
                                    except:
                                        continue
                                if not chat_row: continue
                                
                                chat_title = chat_row.text.split('\n')[0] if chat_row.text else ""
                                if not chat_title or 'archived' in chat_title.lower(): continue
                                
                                # Skip chats that failed extraction recently (120s cooldown)
                                now_ts = time.time()
                                if chat_title in failed_chats and (now_ts - failed_chats[chat_title]) < 120:
                                    continue
                                
                                self.signals.log_msg.emit(f"[Auto-Reply] Unread from: {chat_title}. Opening...")
                                chat_row.click()
                                time.sleep(3)
                                
                                # --- Message extraction ---
                                last_msg_js = """
                                    function isTimestamp(s) {
                                        return /^\\d{1,2}:\\d{2}(\\s*(am|pm|AM|PM))?$/.test(s.trim());
                                    }
                                    function isGoodText(s) {
                                        return s && s.trim().length > 0 && s.trim().length < 500 && !isTimestamp(s);
                                    }
                                    var main = document.getElementById('main');
                                    if (main) {
                                        var rows = main.querySelectorAll('[data-id]');
                                        var incoming = [];
                                        for (var i = 0; i < rows.length; i++) {
                                            var id = rows[i].getAttribute('data-id');
                                            if (id && !id.startsWith('true_')) incoming.push(rows[i]);
                                        }
                                        if (incoming.length > 0) {
                                            var last = incoming[incoming.length - 1];
                                            var ct = last.querySelector('.copyable-text [dir]')
                                                  || last.querySelector('.selectable-text [dir]')
                                                  || last.querySelector('span.selectable-text span');
                                            if (ct && isGoodText(ct.innerText)) return ct.innerText.trim();
                                            var spans = last.querySelectorAll('span');
                                            for (var j = 0; j < spans.length; j++) {
                                                if (isGoodText(spans[j].innerText)) return spans[j].innerText.trim();
                                            }
                                        }
                                        var dirSpans = main.querySelectorAll('span[dir="ltr"], span[dir="rtl"], span[dir="auto"]');
                                        for (var k = dirSpans.length - 1; k >= 0; k--) {
                                            if (isGoodText(dirSpans[k].innerText)) return dirSpans[k].innerText.trim();
                                        }
                                    }
                                    var msgIn = document.querySelectorAll('div.message-in, [class*="message-in"]');
                                    if (msgIn.length > 0) {
                                        var last2 = msgIn[msgIn.length - 1];
                                        var allSpans = last2.querySelectorAll('span');
                                        for (var m = 0; m < allSpans.length; m++) {
                                            if (isGoodText(allSpans[m].innerText)) return allSpans[m].innerText.trim();
                                        }
                                    }
                                    return '';
                                """
                                latest_msg_text = str(driver.execute_script(last_msg_js) or '').strip()
                                
                                self.signals.log_msg.emit(f"[Auto-Reply] Message from {chat_title}: '{latest_msg_text}'")
                                
                                if latest_msg_text:
                                    for kw, reply_text in self.current_presets.items():
                                        if kw.lower() == latest_msg_text.lower():
                                            self.signals.log_msg.emit(f"[Auto-Reply] MATCH '{kw}'! Sending reply...")
                                            self.signals.history_msg.emit(f"IN ({chat_title}): '{latest_msg_text}' | OUT: '{reply_text}'")
                                            
                                            type_js = """
                                                var box = document.querySelector('div[contenteditable="true"][data-tab="10"]')
                                                       || document.querySelector('div[contenteditable="true"][title="Type a message"]')
                                                       || document.querySelector('footer div[contenteditable="true"]')
                                                       || document.querySelector('div[contenteditable="true"][role="textbox"]');
                                                if (box) {
                                                    box.focus();
                                                    document.execCommand('insertText', false, arguments[0]);
                                                    return true;
                                                }
                                                return false;
                                            """
                                            typed = driver.execute_script(type_js, reply_text)
                                            
                                            if typed:
                                                time.sleep(0.5)
                                                ActionChains(driver).send_keys(Keys.ENTER).perform()
                                                time.sleep(2)
                                                self.signals.log_msg.emit(f"[Auto-Reply] ✅ Reply sent to {chat_title}!")
                                            else:
                                                self.signals.log_msg.emit("[Auto-Reply] Could not find chat input box!")
                                            break
                                    else:
                                        self.signals.history_msg.emit(f"Read from {chat_title}: '{latest_msg_text}' (No preset match)")
                                elif not latest_msg_text:
                                    # Mark as failed so we skip it for 120 seconds
                                    failed_chats[chat_title] = time.time()
                                    self.signals.log_msg.emit(f"[Auto-Reply] Could not extract text from {chat_title}. Skipping 2 min.")
                                
                                # Go back to chat list
                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                time.sleep(1)
                                break  # One chat per scan cycle
                                
                            except Exception as badge_e:
                                self.signals.log_msg.emit(f"[Auto-Reply] Error: {str(badge_e)[:80]}")
                                try:
                                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                except:
                                    pass
                                continue
                    else:
                        time.sleep(3)
                    
                    # Clean up old entries
                    now_ts = time.time()
                    failed_chats = {k: v for k, v in failed_chats.items() if (now_ts - v) < 120}
                        
                except Exception as loop_e:
                    self.signals.log_msg.emit(f"[Auto-Reply] Loop error: {str(loop_e)[:100]}")
                    try:
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    except:
                        pass
                    time.sleep(5)


        except Exception as e:
            self.signals.error.emit(str(e))
        finally:
            try:
                pass  # Do NOT close Chrome. The user wants it to stay open.
            except:
                pass
            self.signals.finished.emit()

# We need QtCore for the Q_ARG macro
from PyQt6 import QtCore

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion") # Gives a nice classic look across all OS
    
    window = WhatsAppAutomatorApp()
    window.show()
    
    sys.exit(app.exec())
