#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MatterID - Manager v2.5
Professional Model United Nations Management System
Powered by MatterID Authentication
"""

import sys
import csv
import json
import threading
import urllib.request
import ssl
import certifi
import re
import logging
import traceback
import webbrowser
import os
from urllib.parse import urlparse, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
import random

import firebase_admin
from firebase_admin import credentials, firestore, auth

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit,
    QHBoxLayout, QMessageBox, QComboBox, QFileDialog, QSplashScreen,
    QDialog, QProgressBar, QProgressDialog, QStatusBar, QMenu,
    QAbstractItemView, QTabWidget, QTextEdit, QScrollArea, QGridLayout,
    QFrame, QSplitter, QGroupBox, QFormLayout, QSpacerItem, QSizePolicy,
    QListWidget, QListWidgetItem, QInputDialog, QCheckBox
)
from PyQt6.QtGui import QPixmap, QKeySequence, QColor, QBrush, QAction, QFont
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QDateTime, QSettings

# Logging Setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Global Variables
db = None
LOGIN_TOKEN = None

# MatterID Color Scheme
MATTERID_COLORS = {
    'primary': '#4169E1',        # Royal Blue
    'accent': '#007BFF',         # Electric Blue
    'hover': '#365bb8',          # Darker Blue
    'success': '#4CAF50',        # Material Green
    'warning': '#FFC107',        # Material Amber
    'error': '#F44336',          # Material Red
    'background': '#2e2e2e',     # Dark Slate
    'card_bg': '#3a3a3a',        # Carbon Gray
    'border': '#555555',         # Steel Gray
    'text_primary': '#ffffff',   # Pure White
    'text_secondary': '#cccccc', # Light Gray
    'text_muted': '#aaaaaa',     # Muted Gray
    'cyan_accent': '#00D4FF'     # Cyan Accent
}

# Configuration Manager
class ConfigManager:
    def __init__(self):
        self.settings = QSettings("MatterID", "Manager")
        self.default_config = {
            "key_url": "https://api.eliomatters.com/sangammun.json",
            "collection_name": "registrations",
            "table_columns": [
                {"display": "Document ID", "field": None, "editable": False},
                {"display": "First Name", "field": "name", "editable": True},
                {"display": "Email", "field": "email", "editable": True},
                {"display": "Phone No.", "field": "phone", "editable": True},
                {"display": "School", "field": "school", "editable": True},
                {"display": "School Info : If other than dps", "field": "customSchool", "editable": True},
                {"display": "Committee Preferences", "field": "committeePreferences", "editable": True},
                {"display": "Portfolio Preferences", "field": "portfolioPreferences", "editable": True},
                {"display": "DOB", "field": "dob", "editable": True},
                {"display": "Final Committee", "field": "finalCommittee", "editable": True},
                {"display": "Final Portfolio", "field": "finalPortfolio", "editable": True},
                {"display": "Payment SS URL", "field": "screenshotURL", "editable": True}
            ],
            "recent_configs": []
        }
    
    def get_config(self):
        config = {}
        for key, default_value in self.default_config.items():
            if key == "table_columns":
                saved_columns = self.settings.value(key, default_value)
                if isinstance(saved_columns, str):
                    try:
                        config[key] = json.loads(saved_columns)
                    except json.JSONDecodeError:
                        config[key] = default_value
                else:
                    config[key] = saved_columns or default_value
            else:
                config[key] = self.settings.value(key, default_value)
        return config
    
    def save_config(self, config):
        for key, value in config.items():
            if key == "table_columns":
                self.settings.setValue(key, json.dumps(value))
            else:
                self.settings.setValue(key, value)
        self.settings.sync()
    
    def add_recent_config(self, config_name, config):
        recent = self.get_config().get("recent_configs", [])
        recent = [r for r in recent if r.get("name") != config_name]
        recent.insert(0, {"name": config_name, "config": config})
        recent = recent[:10]
        
        current_config = self.get_config()
        current_config["recent_configs"] = recent
        self.save_config(current_config)
    
    def export_config(self, filepath):
        config = self.get_config()
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    
    def import_config(self, filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            config = json.load(f)
        self.save_config(config)
        return config

# Demo Data Generator
class DemoDataGenerator:
    @staticmethod
    def generate_demo_delegates():
        """Generate 10 sample delegates for demonstration"""
        delegates = {}
        
        committees = ["UNHRC", "Lok Sabha", "UNGA-Disec", "UNCSW", "Continuous Crisis Committee", "International Press"]
        schools = ["DPS RK Puram", "DPS Mathura Road", "Ryan International", "Modern School", "Delhi Public School", "Sanskriti School"]
        
        sample_names = [
            "Arjun Sharma", "Priya Patel", "Rahul Gupta", "Ananya Singh", 
            "Karthik Iyer", "Sneha Reddy", "Vikram Malhotra", "Riya Kapoor",
            "Aditya Jain", "Kavya Nair"
        ]
        
        for i, name in enumerate(sample_names):
            doc_id = f"demo_{i+1:03d}"
            delegates[doc_id] = {
                "name": name,
                "email": f"{name.lower().replace(' ', '.')}@example.com",
                "phone": f"+91 9{random.randint(100000000, 999999999)}",
                "school": random.choice(schools),
                "customSchool": "",
                "committeePreferences": random.choice(committees),
                "portfolioPreferences": f"Delegate of {random.choice(['India', 'USA', 'China', 'France', 'UK', 'Germany'])}",
                "dob": f"{random.randint(2005, 2008)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
                "finalCommittee": random.choice(committees),
                "finalPortfolio": f"Delegate of {random.choice(['India', 'USA', 'China', 'France', 'UK', 'Germany'])}",
                "screenshotURL": "https://example.com/payment_screenshot.jpg",
                "updatedAt": datetime.now()
            }
        
        return delegates
    
    @staticmethod
    def generate_demo_attendance():
        """Generate sample attendance data"""
        attendance = {}
        
        # Generate attendance patterns
        patterns = [
            [True, True, True],    # Perfect attendance
            [True, True, False],   # Missed day 3
            [True, False, True],   # Missed day 2
            [False, True, True],   # Missed day 1
            [True, False, False],  # Only day 1
            [False, False, True],  # Only day 3
        ]
        
        for i in range(1, 11):
            doc_id = f"demo_{i:03d}"
            pattern = random.choice(patterns)
            attendance[doc_id] = {
                "day1": pattern[0],
                "day2": pattern[1],
                "day3": pattern[2],
                "updatedAt": datetime.now(),
                "recordedBy": "demo_user"
            }
        
        return attendance

# Import exceptions
try:
    from firebase_admin.auth import ExpiredIdTokenError, InvalidIdTokenError
except ImportError:
    class ExpiredIdTokenError(Exception):
        pass
    class InvalidIdTokenError(Exception):
        pass

# Helper Functions
EMAIL_REGEX = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

def is_valid_email(email):
    return re.match(EMAIL_REGEX, email) is not None

def format_timestamp(ts):
    if isinstance(ts, datetime):
        try:
            local_dt = ts.astimezone()
            return local_dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ts.strftime("%Y-%m-%d %H:%M:%S (UTC?)")
    elif isinstance(ts, QDateTime):
        return ts.toString("yyyy-MM-dd HH:mm:ss")
    elif ts:
        return str(ts)
    return ""

def get_initials(name):
    if not name:
        return "??"
    words = str(name).strip().split()
    if len(words) == 0:
        return "??"
    elif len(words) == 1:
        return words[0][:2].upper()
    else:
        return (words[0][0] + words[-1][0]).upper()

# Callback Handler
class _CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global LOGIN_TOKEN
        parsed = urlparse(self.path)
        logging.info(f"Received HTTP GET: {self.path}")

        if parsed.path == "/callback":
            params = parse_qs(parsed.query)
            token_list = params.get("token", [])
            if token_list:
                LOGIN_TOKEN = token_list[0]
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.end_headers()
                html_response = """<html><body style='background-color: #2e2e2e; color: white;'>
                    <div style='text-align: center; padding: 50px;'>
                    <h1 style='color: #4169E1;'>🆔 MatterID</h1>
                    <h2 style='color: #4CAF50;'>Login Successful!</h2>
                    <p>You may close this window and return to MatterID - Manager.</p>
                    </div></body></html>"""
                self.wfile.write(html_response.encode('utf-8'))
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return

        self.send_response(404)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(b"<html><body><h2>Not Found</h2></body></html>")

# WebLoginDialog
class WebLoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("MatterID - Manager Login")
        self.setModal(True)
        self.setFixedSize(400, 200)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {MATTERID_COLORS['background']};
                color: {MATTERID_COLORS['text_primary']};
            }}
            QPushButton {{
                background-color: {MATTERID_COLORS['primary']};
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background-color: {MATTERID_COLORS['hover']};
            }}
            QLabel {{
                color: {MATTERID_COLORS['text_secondary']};
                font-size: 14px;
            }}
        """)

        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("🆔 MatterID - Manager v2.5")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        subtitle_label = QLabel("Professional MUN Management • Powered by MatterID")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_label.setStyleSheet(f"font-size: 12px; color: {MATTERID_COLORS['text_muted']}; margin-bottom: 20px;")
        layout.addWidget(subtitle_label)
        
        layout.addWidget(QLabel("Click below to authenticate with MatterID:"))

        self.login_btn = QPushButton("🆔 Login with MatterID")
        self.login_btn.clicked.connect(self.start_web_login)
        layout.addWidget(self.login_btn)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.cancel_login)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MATTERID_COLORS['error']};
                color: white;
            }}
            QPushButton:hover {{
                background-color: #d32f2f;
            }}
        """)
        layout.addWidget(self.cancel_btn)

        self.setLayout(layout)

        self.server = None
        self.server_thread = None

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(500)
        self.poll_timer.timeout.connect(self.check_token)

    def start_web_login(self):
        global LOGIN_TOKEN
        LOGIN_TOKEN = None

        try:
            self.server = HTTPServer(("127.0.0.1", 5000), _CallbackHandler)
        except OSError:
            QMessageBox.critical(self, "Error", "Could not start local server on port 5000. Port may be in use.")
            return

        self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.server_thread.start()

        webbrowser.open("https://eliomatters.com/auth.html")

        self.login_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.status_label.setText("🔄 Waiting for MatterID authentication…")
        self.status_label.setStyleSheet(f"color: {MATTERID_COLORS['accent']};")
        self.poll_timer.start()

    def check_token(self):
        global LOGIN_TOKEN
        if LOGIN_TOKEN:
            self.poll_timer.stop()
            self.status_label.setText("✅ Token received. Verifying…")
            self.status_label.setStyleSheet(f"color: {MATTERID_COLORS['success']};")
            QApplication.processEvents()
            QTimer.singleShot(100, self.accept)

    def cancel_login(self):
        self.poll_timer.stop()
        if self.server:
            try:
                self.server.shutdown()
                self.server.server_close()
            except Exception:
                pass
        self.reject()

    def closeEvent(self, event):
        self.poll_timer.stop()
        if self.server:
            try:
                self.server.shutdown()
                self.server.server_close()
            except Exception:
                pass
        super().closeEvent(event)

# KeyDownloadThread
class KeyDownloadThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(dict, dict)
    error = pyqtSignal(str)

    def __init__(self, key_url, auth_url):
        super().__init__()
        self.key_url = key_url
        self.auth_url = auth_url or key_url  # Use same key if auth_url not provided

    def run(self):
        logging.info("Starting MatterID key download from API…")
        try:
            ctx = ssl.create_default_context(cafile=certifi.where())
            
            with urllib.request.urlopen(self.key_url, context=ctx, timeout=20) as response:
                service_data = json.loads(response.read().decode('utf-8'))
            
            self.progress.emit(50)
            
            # Try to get auth data from separate URL, fallback to service data
            try:
                with urllib.request.urlopen(self.auth_url, context=ctx, timeout=20) as response:
                    auth_data = json.loads(response.read().decode('utf-8'))
            except:
                # Use service data as fallback
                auth_data = service_data
            
            self.progress.emit(100)
            
            logging.info("MatterID keys downloaded successfully.")
            self.finished.emit(service_data, auth_data)

        except Exception as e:
            logging.error(f"Error downloading keys: {e}\n{traceback.format_exc()}")
            self.error.emit(f"Error downloading keys: {e}")

# DownloadSplashScreen
class DownloadSplashScreen(QDialog):
    def __init__(self, config_manager):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Dialog)
        self.setModal(True)
        self.config_manager = config_manager
        self.service_key_data = None
        self.auth_key_data = None
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {MATTERID_COLORS['background']};
                border: 2px solid {MATTERID_COLORS['primary']};
                border-radius: 10px;
            }}
            QLabel {{
                color: {MATTERID_COLORS['text_primary']};
            }}
            QProgressBar {{
                border: 2px solid {MATTERID_COLORS['border']};
                border-radius: 5px;
                text-align: center;
                background-color: {MATTERID_COLORS['card_bg']};
            }}
            QProgressBar::chunk {{
                background-color: {MATTERID_COLORS['primary']};
                border-radius: 3px;
            }}
        """)

        layout = QVBoxLayout()
        layout.setSpacing(15)

        # Logo/Header
        self.logo_label = QLabel("🆔")
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_label.setStyleSheet(f"font-size: 48px; color: {MATTERID_COLORS['primary']};")
        
        title_label = QLabel("MatterID - Manager v2.5")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {MATTERID_COLORS['primary']};")
        
        subtitle_label = QLabel("Professional MUN Management • Powered by MatterID")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_label.setStyleSheet(f"font-size: 12px; color: {MATTERID_COLORS['text_muted']};")

        self.message_label = QLabel("🔄 Downloading MatterID configuration…")
        self.message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.message_label.setStyleSheet(f"font-size: 14px; color: {MATTERID_COLORS['text_secondary']};")

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        layout.addWidget(self.logo_label)
        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        layout.addWidget(self.message_label)
        layout.addWidget(self.progress_bar)
        self.setLayout(layout)
        self.resize(400, 250)

        config = self.config_manager.get_config()
        key_url = config.get("key_url", "https://api.eliomatters.com/sangammun.json")
        auth_url = "https://api.eliomatters.com/eliomatter.json"  # Use specific auth URL

        self.thread = KeyDownloadThread(key_url, auth_url)
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.download_finished)
        self.thread.error.connect(self.download_error)
        self.thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def download_finished(self, service_key_data, auth_key_data):
        self.service_key_data = service_key_data
        self.auth_key_data = auth_key_data
        self.message_label.setText("✅ Configuration loaded. Initializing Firebase…")
        logging.info("Key download finished. Initializing Firebase…")
        try:
            service_cred = credentials.Certificate(self.service_key_data)
            auth_cred = credentials.Certificate(self.auth_key_data)
            
            if not firebase_admin._apps:
                default_app = firebase_admin.initialize_app(service_cred)
                # Initialize auth app for token verification
                auth_app = firebase_admin.initialize_app(auth_cred, name='auth')
                logging.info("Firebase apps initialized successfully.")
            else:
                logging.warning("Firebase app already initialized.")

            global db
            db = firestore.client()

            self.accept()

        except ValueError as e:
            logging.error(f"Invalid MatterID key format: {e}")
            QMessageBox.critical(self, "MatterID Error", f"Invalid key format: {e}")
            self.reject()
        except Exception as e:
            logging.error(f"Firebase initialization error: {e}\n{traceback.format_exc()}")
            QMessageBox.critical(self, "MatterID Error", f"Firebase init error: {e}")
            self.reject()

    def download_error(self, error_message):
        logging.error(f"DownloadSplashScreen received error: {error_message}")
        self.message_label.setText("⚠️ Connection failed. Starting in demo mode…")
        self.message_label.setStyleSheet(f"color: {MATTERID_COLORS['warning']};")
        
        # Auto-accept to continue in demo mode
        QTimer.singleShot(2000, self.accept)

# Configuration Tab Widget
class ConfigTab(QWidget):
    config_changed = pyqtSignal()
    
    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.init_ui()
        self.load_config()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("⚙️ MatterID Configuration")
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        # Basic Configuration Group
        basic_group = QGroupBox("Basic Configuration")
        basic_layout = QFormLayout()
        
        self.key_url_edit = QLineEdit()
        self.key_url_edit.setPlaceholderText("Enter the MatterID key URL")
        basic_layout.addRow("MatterID Key URL:", self.key_url_edit)
        
        self.collection_edit = QLineEdit()
        self.collection_edit.setPlaceholderText("Firestore collection name")
        basic_layout.addRow("Collection Name:", self.collection_edit)
        
        basic_group.setLayout(basic_layout)
        layout.addWidget(basic_group)
        
        # Table Columns Configuration
        columns_group = QGroupBox("Table Columns Configuration")
        columns_layout = QVBoxLayout()
        
        # Buttons for column management
        column_buttons = QHBoxLayout()
        self.add_column_btn = QPushButton("➕ Add Column")
        self.remove_column_btn = QPushButton("➖ Remove Selected")
        self.move_up_btn = QPushButton("⬆️ Move Up")
        self.move_down_btn = QPushButton("⬇️ Move Down")
        
        self.add_column_btn.clicked.connect(self.add_column)
        self.remove_column_btn.clicked.connect(self.remove_column)
        self.move_up_btn.clicked.connect(self.move_column_up)
        self.move_down_btn.clicked.connect(self.move_column_down)
        
        column_buttons.addWidget(self.add_column_btn)
        column_buttons.addWidget(self.remove_column_btn)
        column_buttons.addWidget(self.move_up_btn)
        column_buttons.addWidget(self.move_down_btn)
        column_buttons.addStretch()
        
        columns_layout.addLayout(column_buttons)
        
        # Columns list
        self.columns_list = QListWidget()
        self.columns_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        columns_layout.addWidget(self.columns_list)
        
        # Column edit form
        edit_group = QGroupBox("Edit Selected Column")
        edit_layout = QFormLayout()
        
        self.display_name_edit = QLineEdit()
        self.field_name_edit = QLineEdit()
        self.editable_combo = QComboBox()
        self.editable_combo.addItems(["True", "False"])
        
        edit_layout.addRow("Display Name:", self.display_name_edit)
        edit_layout.addRow("Field Name:", self.field_name_edit)
        edit_layout.addRow("Editable:", self.editable_combo)
        
        self.update_column_btn = QPushButton("💾 Update Column")
        self.update_column_btn.clicked.connect(self.update_selected_column)
        edit_layout.addRow(self.update_column_btn)
        
        edit_group.setLayout(edit_layout)
        columns_layout.addWidget(edit_group)
        
        columns_group.setLayout(columns_layout)
        layout.addWidget(columns_group)
        
        # Action buttons
        action_layout = QHBoxLayout()
        
        self.save_config_btn = QPushButton("💾 Save Configuration")
        self.export_btn = QPushButton("📤 Export Config")
        self.import_btn = QPushButton("📥 Import Config")
        self.reset_btn = QPushButton("🔄 Reset to Defaults")
        
        self.save_config_btn.clicked.connect(self.save_config)
        self.export_btn.clicked.connect(self.export_config)
        self.import_btn.clicked.connect(self.import_config)
        self.reset_btn.clicked.connect(self.reset_to_defaults)
        
        action_layout.addWidget(self.save_config_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addWidget(self.import_btn)
        action_layout.addStretch()
        action_layout.addWidget(self.reset_btn)
        
        layout.addLayout(action_layout)
        layout.addStretch()
        
        self.setLayout(layout)
        
        # Connect events
        self.columns_list.itemClicked.connect(self.on_column_selected)
    
    def load_config(self):
        config = self.config_manager.get_config()
        
        self.key_url_edit.setText(config.get("key_url", ""))
        self.collection_edit.setText(config.get("collection_name", ""))
        
        # Load columns
        self.columns_list.clear()
        for column in config.get("table_columns", []):
            item = QListWidgetItem(f"{column['display']} ({column.get('field', 'None')})")
            item.setData(Qt.ItemDataRole.UserRole, column)
            self.columns_list.addItem(item)
    
    def add_column(self):
        column = {
            "display": "New Column",
            "field": "new_field",
            "editable": True
        }
        item = QListWidgetItem(f"{column['display']} ({column['field']})")
        item.setData(Qt.ItemDataRole.UserRole, column)
        self.columns_list.addItem(item)
    
    def remove_column(self):
        current_row = self.columns_list.currentRow()
        if current_row >= 0:
            self.columns_list.takeItem(current_row)
    
    def move_column_up(self):
        current_row = self.columns_list.currentRow()
        if current_row > 0:
            item = self.columns_list.takeItem(current_row)
            self.columns_list.insertItem(current_row - 1, item)
            self.columns_list.setCurrentRow(current_row - 1)
    
    def move_column_down(self):
        current_row = self.columns_list.currentRow()
        if current_row < self.columns_list.count() - 1:
            item = self.columns_list.takeItem(current_row)
            self.columns_list.insertItem(current_row + 1, item)
            self.columns_list.setCurrentRow(current_row + 1)
    
    def on_column_selected(self, item):
        column_data = item.data(Qt.ItemDataRole.UserRole)
        if column_data:
            self.display_name_edit.setText(column_data.get("display", ""))
            self.field_name_edit.setText(column_data.get("field", ""))
            self.editable_combo.setCurrentText("True" if column_data.get("editable", True) else "False")
    
    def update_selected_column(self):
        current_item = self.columns_list.currentItem()
        if current_item:
            column_data = {
                "display": self.display_name_edit.text(),
                "field": self.field_name_edit.text() or None,
                "editable": self.editable_combo.currentText() == "True"
            }
            current_item.setData(Qt.ItemDataRole.UserRole, column_data)
            current_item.setText(f"{column_data['display']} ({column_data.get('field', 'None')})")
    
    def get_current_config(self):
        columns = []
        for i in range(self.columns_list.count()):
            item = self.columns_list.item(i)
            column_data = item.data(Qt.ItemDataRole.UserRole)
            if column_data:
                columns.append(column_data)
        
        return {
            "key_url": self.key_url_edit.text(),
            "collection_name": self.collection_edit.text(),
            "table_columns": columns
        }
    
    def save_config(self):
        config = self.get_current_config()
        full_config = self.config_manager.get_config()
        full_config.update(config)
        self.config_manager.save_config(full_config)
        self.config_changed.emit()
        QMessageBox.information(self, "Success", "Configuration saved successfully!")
    
    def export_config(self):
        filepath, _ = QFileDialog.getSaveFileName(self, "Export Configuration", "config.json", "JSON Files (*.json)")
        if filepath:
            try:
                self.config_manager.export_config(filepath)
                QMessageBox.information(self, "Success", f"Configuration exported to {filepath}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export configuration: {e}")
    
    def import_config(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Import Configuration", "", "JSON Files (*.json)")
        if filepath:
            try:
                self.config_manager.import_config(filepath)
                self.load_config()
                self.config_changed.emit()
                QMessageBox.information(self, "Success", "Configuration imported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import configuration: {e}")
    
    def reset_to_defaults(self):
        reply = QMessageBox.question(self, "Reset Configuration", 
                                   "Are you sure you want to reset to default configuration?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                   QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.config_manager.save_config(self.config_manager.default_config)
            self.load_config()
            self.config_changed.emit()

# Attendance Card Widget
class AttendanceCard(QFrame):
    attendance_changed = pyqtSignal(str, str, bool)  # doc_id, day, present
    
    def __init__(self, doc_id, user_data, attendance_data=None):
        super().__init__()
        self.doc_id = doc_id
        self.user_data = user_data
        self.attendance_data = attendance_data or {}
        self.day_checkboxes = {}
        self.init_ui()
    
    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.Box)
        self.setStyleSheet(f"""
            QFrame {{
                border: 2px solid {MATTERID_COLORS['primary']};
                border-radius: 8px;
                background-color: {MATTERID_COLORS['card_bg']};
                margin: 5px;
            }}
            QFrame:hover {{
                border-color: {MATTERID_COLORS['hover']};
                background-color: #404040;
            }}
            QLabel {{
                border: none;
                color: {MATTERID_COLORS['text_primary']};
            }}
            QCheckBox {{
                color: {MATTERID_COLORS['text_primary']};
                font-weight: bold;
            }}
            QCheckBox::indicator:checked {{
                background-color: {MATTERID_COLORS['success']};
                border: 2px solid {MATTERID_COLORS['success']};
            }}
            QCheckBox::indicator:unchecked {{
                background-color: {MATTERID_COLORS['card_bg']};
                border: 2px solid {MATTERID_COLORS['border']};
            }}
        """)
        self.setFixedSize(250, 160)
        
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Initials circle
        initials = get_initials(self.user_data.get("name", ""))
        initials_label = QLabel(initials)
        initials_label.setStyleSheet(f"""
            QLabel {{
                background-color: {MATTERID_COLORS['primary']};
                color: white;
                border-radius: 20px;
                font-size: 16px;
                font-weight: bold;
                border: none;
            }}
        """)
        initials_label.setFixedSize(40, 40)
        initials_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Name
        name_label = QLabel(self.user_data.get("name", "Unknown"))
        name_label.setStyleSheet("font-weight: bold; font-size: 12px; color: white; border: none;")
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setWordWrap(True)
        
        # Committee
        committee = self.user_data.get("finalCommittee", "Not Assigned")
        committee_label = QLabel(f"📋 {committee}")
        committee_label.setStyleSheet("font-size: 10px; color: #cccccc; border: none;")
        committee_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        committee_label.setWordWrap(True)
        
        # Attendance checkboxes
        attendance_layout = QHBoxLayout()
        attendance_layout.setSpacing(10)
        
        for day in ["day1", "day2", "day3"]:
            day_num = day[-1]
            checkbox = QCheckBox(f"Day {day_num}")
            checkbox.setChecked(self.attendance_data.get(day, False))
            checkbox.stateChanged.connect(
                lambda state, d=day: self.on_attendance_changed(d, state == Qt.CheckState.Checked.value)
            )
            self.day_checkboxes[day] = checkbox
            attendance_layout.addWidget(checkbox)
        
        # Header layout
        header_layout = QHBoxLayout()
        header_layout.addStretch()
        header_layout.addWidget(initials_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        layout.addWidget(name_label)
        layout.addWidget(committee_label)
        layout.addLayout(attendance_layout)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def on_attendance_changed(self, day, present):
        self.attendance_data[day] = present
        self.attendance_changed.emit(self.doc_id, day, present)
    
    def update_attendance(self, day, present):
        """Update attendance from external source"""
        if day in self.day_checkboxes:
            self.day_checkboxes[day].setChecked(present)
            self.attendance_data[day] = present

# Attendance View Widget
class AttendanceView(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.attendance_cards = {}
        self.attendance_data = {}
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("✅ Attendance Management")
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        # Controls
        controls_layout = QHBoxLayout()
        
        # Search
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Search by name or committee...")
        self.search_edit.textChanged.connect(self.filter_attendance)
        
        # Day selector
        self.day_combo = QComboBox()
        self.day_combo.addItems(["All Days", "Day 1", "Day 2", "Day 3"])
        
        # Quick actions
        self.mark_present_btn = QPushButton("✅ Mark All Present")
        self.mark_absent_btn = QPushButton("❌ Mark All Absent")
        self.save_attendance_btn = QPushButton("💾 Save to Database")
        self.test_db_btn = QPushButton("🔍 Test Database")
        self.export_btn = QPushButton("📊 Export Attendance")
        
        self.mark_present_btn.clicked.connect(self.mark_all_present)
        self.mark_absent_btn.clicked.connect(self.mark_all_absent)
        self.save_attendance_btn.clicked.connect(self.save_all_attendance)
        self.test_db_btn.clicked.connect(self.test_database_connection)
        self.export_btn.clicked.connect(self.export_attendance)
        
        controls_layout.addWidget(QLabel("Search:"))
        controls_layout.addWidget(self.search_edit)
        controls_layout.addWidget(QLabel("Day:"))
        controls_layout.addWidget(self.day_combo)
        controls_layout.addWidget(self.mark_present_btn)
        controls_layout.addWidget(self.mark_absent_btn)
        controls_layout.addWidget(self.save_attendance_btn)
        controls_layout.addWidget(self.test_db_btn)
        controls_layout.addStretch()
        controls_layout.addWidget(self.export_btn)
        
        layout.addLayout(controls_layout)
        
        # Statistics bar
        self.stats_label = QLabel("Total: 0 | Present: 0 | Rate: 0%")
        self.stats_label.setStyleSheet(f"""
            QLabel {{
                background-color: {MATTERID_COLORS['card_bg']};
                color: {MATTERID_COLORS['text_primary']};
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
                border: 1px solid {MATTERID_COLORS['border']};
            }}
        """)
        layout.addWidget(self.stats_label)
        
        # Scroll area for attendance cards
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        self.cards_widget = QWidget()
        self.cards_layout = QGridLayout()
        self.cards_layout.setSpacing(10)
        self.cards_widget.setLayout(self.cards_layout)
        
        scroll_area.setWidget(self.cards_widget)
        layout.addWidget(scroll_area)
        
        self.setLayout(layout)
    
    def update_attendance_data(self, users_data, attendance_data=None):
        # Clear existing cards
        for i in reversed(range(self.cards_layout.count())):
            child = self.cards_layout.itemAt(i).widget()
            if child:
                child.setParent(None)
        
        self.attendance_cards.clear()
        
        # Use demo data if no attendance data provided
        if attendance_data is None:
            self.attendance_data = DemoDataGenerator.generate_demo_attendance()
        else:
            self.attendance_data = attendance_data
        
        # Add attendance cards
        row = 0
        col = 0
        cols_per_row = 5
        
        for doc_id, user_data in users_data.items():
            if user_data:
                attendance = self.attendance_data.get(doc_id, {})
                card = AttendanceCard(doc_id, user_data, attendance)
                card.attendance_changed.connect(self.on_attendance_changed)
                self.attendance_cards[doc_id] = card
                
                self.cards_layout.addWidget(card, row, col)
                col += 1
                if col >= cols_per_row:
                    col = 0
                    row += 1
        
        # Add stretch to fill remaining space
        self.cards_layout.setRowStretch(row + 1, 1)
        self.cards_layout.setColumnStretch(cols_per_row, 1)
        
        self.update_statistics()
    
    def on_attendance_changed(self, doc_id, day, present):
        """Handle attendance change from card"""
        if doc_id not in self.attendance_data:
            self.attendance_data[doc_id] = {}
        
        self.attendance_data[doc_id][day] = present
        self.attendance_data[doc_id]["updatedAt"] = firestore.SERVER_TIMESTAMP if db else datetime.now()
        self.attendance_data[doc_id]["recordedBy"] = "matterid_user"  # TODO: Get actual user ID
        
        # Save to database if connected
        if db and not self.main_window.demo_mode:
            try:
                logging.info(f"Attempting to save attendance for {doc_id}: {day} = {present}")
                
                # Create the attendance document data
                attendance_doc = {
                    day: present,
                    "updatedAt": firestore.SERVER_TIMESTAMP,
                    "recordedBy": "matterid_user"
                }
                
                # Save to Firestore with merge=True to update existing or create new
                result = db.collection("attendance").document(doc_id).set(
                    attendance_doc, 
                    merge=True
                )
                
                logging.info(f"✅ Attendance saved successfully for {doc_id}: {day} = {present}")
                
                # Update status bar to show save confirmation
                self.main_window.update_status(f"Attendance saved: {doc_id} - {day}")
                
            except Exception as e:
                logging.error(f"❌ Error saving attendance for {doc_id}: {e}")
                logging.error(f"Full error: {traceback.format_exc()}")
                
                # Show error message to user
                QMessageBox.warning(
                    self.main_window,
                    "Attendance Save Error", 
                    f"Could not save attendance for {doc_id}:\n{str(e)}"
                )
        else:
            logging.info(f"Demo mode: Attendance for {doc_id} saved locally only")
        
        self.update_statistics()
    
    def filter_attendance(self):
        search_text = self.search_edit.text().lower()
        
        for doc_id, card in self.attendance_cards.items():
            user_data = card.user_data
            should_show = True
            
            if search_text:
                name = str(user_data.get("name", "")).lower()
                committee = str(user_data.get("finalCommittee", "")).lower()
                
                should_show = (search_text in name or search_text in committee)
            
            card.setVisible(should_show)
        
        self.update_statistics()
    
    def mark_all_present(self):
        selected_day = self.day_combo.currentText()
        if selected_day == "All Days":
            days = ["day1", "day2", "day3"]
        else:
            day_num = selected_day.split()[-1]
            days = [f"day{day_num}"]
        
        visible_cards = [card for card in self.attendance_cards.values() if card.isVisible()]
        
        for card in visible_cards:
            for day in days:
                card.update_attendance(day, True)
        
        self.update_statistics()
    
    def mark_all_absent(self):
        selected_day = self.day_combo.currentText()
        if selected_day == "All Days":
            days = ["day1", "day2", "day3"]
        else:
            day_num = selected_day.split()[-1]
            days = [f"day{day_num}"]
        
        visible_cards = [card for card in self.attendance_cards.values() if card.isVisible()]
        
        for card in visible_cards:
            for day in days:
                card.update_attendance(day, False)
        
        self.update_statistics()
    
    def test_database_connection(self):
        """Test database connection and permissions"""
        if not db:
            QMessageBox.warning(self, "No Database", "Database connection not available. Running in demo mode.")
            return
        
        try:
            # Test basic connection
            logging.info("Testing database connection...")
            
            # Try to read from registrations collection
            config = self.main_window.config_manager.get_config()
            collection_name = config.get("collection_name", "registrations")
            test_query = db.collection(collection_name).limit(1)
            docs = list(test_query.stream())
            
            # Try to read from attendance collection
            attendance_query = db.collection("attendance").limit(1)
            attendance_docs = list(attendance_query.stream())
            
            # Try to write a test document to attendance
            test_doc_id = "test_connection_" + str(int(datetime.now().timestamp()))
            test_data = {
                "day1": True,
                "day2": False,
                "day3": True,
                "updatedAt": firestore.SERVER_TIMESTAMP,
                "recordedBy": "connection_test"
            }
            
            db.collection("attendance").document(test_doc_id).set(test_data)
            
            # Clean up test document
            db.collection("attendance").document(test_doc_id).delete()
            
            # Show success message
            QMessageBox.information(
                self, 
                "Database Connection Test",
                f"✅ Database connection successful!\n\n"
                f"📊 Registrations collection: {len(docs)} documents found\n"
                f"✅ Attendance collection: {len(attendance_docs)} documents found\n"
                f"✅ Write permissions: Working\n"
                f"✅ Read permissions: Working"
            )
            
            logging.info("✅ Database connection test passed!")
            
        except Exception as e:
            error_msg = f"❌ Database connection test failed:\n\n{str(e)}"
            logging.error(f"Database test error: {e}")
            logging.error(f"Full traceback: {traceback.format_exc()}")
            
            QMessageBox.critical(self, "Database Connection Error", error_msg)
    
    def save_all_attendance(self):
        """Manually save all attendance data to Firebase"""
        if not db or self.main_window.demo_mode:
            QMessageBox.information(self, "Demo Mode", "Running in demo mode. Attendance data is stored locally only.")
            return
        
        if not self.attendance_data:
            QMessageBox.information(self, "No Data", "No attendance data to save.")
            return
        
        # Show progress dialog
        progress = QProgressDialog("Saving attendance data to Firebase...", "Cancel", 0, len(self.attendance_data), self)
        progress.setWindowTitle("Saving Attendance")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)
        
        saved_count = 0
        error_count = 0
        
        for i, (doc_id, attendance) in enumerate(self.attendance_data.items()):
            if progress.wasCanceled():
                break
                
            try:
                # Prepare the data for Firebase
                firebase_data = {
                    "day1": attendance.get("day1", False),
                    "day2": attendance.get("day2", False), 
                    "day3": attendance.get("day3", False),
                    "updatedAt": firestore.SERVER_TIMESTAMP,
                    "recordedBy": "matterid_user"
                }
                
                # Save to Firebase
                db.collection("attendance").document(doc_id).set(firebase_data, merge=True)
                saved_count += 1
                logging.info(f"Saved attendance for {doc_id}")
                
            except Exception as e:
                error_count += 1
                logging.error(f"Error saving attendance for {doc_id}: {e}")
            
            progress.setValue(i + 1)
            QApplication.processEvents()
        
        progress.close()
        
        # Show results
        if error_count == 0:
            QMessageBox.information(self, "Success", f"Successfully saved attendance for {saved_count} delegates!")
            self.main_window.update_status(f"Attendance saved: {saved_count} delegates")
        else:
            QMessageBox.warning(self, "Partial Success", 
                               f"Saved: {saved_count}\nErrors: {error_count}\nCheck console for details.")
    
    def update_statistics(self):
        visible_cards = [card for card in self.attendance_cards.values() if card.isVisible()]
        total = len(visible_cards)
        
        if total == 0:
            self.stats_label.setText("Total: 0 | Present: 0 | Rate: 0%")
            return
        
        # Calculate statistics for visible cards
        day1_present = sum(1 for card in visible_cards if card.attendance_data.get("day1", False))
        day2_present = sum(1 for card in visible_cards if card.attendance_data.get("day2", False))
        day3_present = sum(1 for card in visible_cards if card.attendance_data.get("day3", False))
        
        avg_present = (day1_present + day2_present + day3_present) / (total * 3) * 100
        
        self.stats_label.setText(
            f"Total: {total} | Day 1: {day1_present} | Day 2: {day2_present} | "
            f"Day 3: {day3_present} | Overall Rate: {avg_present:.1f}%"
        )
    
    def export_attendance(self):
        if not self.attendance_data:
            QMessageBox.information(self, "Export Error", "No attendance data to export.")
            return
        
        default_filename = f"attendance_export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Attendance", default_filename, "CSV Files (*.csv)")
        if not file_path:
            return
        
        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file)
                
                # Header
                header = ["Document ID", "Name", "Committee", "Day 1", "Day 2", "Day 3", "Pattern", "Days Present"]
                writer.writerow(header)
                
                # Data rows
                for doc_id, card in self.attendance_cards.items():
                    if card.isVisible():
                        user_data = card.user_data
                        attendance = self.attendance_data.get(doc_id, {})
                        
                        day1 = "✅" if attendance.get("day1", False) else "❌"
                        day2 = "✅" if attendance.get("day2", False) else "❌"
                        day3 = "✅" if attendance.get("day3", False) else "❌"
                        
                        # Generate pattern
                        pattern_chars = []
                        pattern_chars.append("P" if attendance.get("day1", False) else "A")
                        pattern_chars.append("P" if attendance.get("day2", False) else "A")
                        pattern_chars.append("P" if attendance.get("day3", False) else "A")
                        pattern = "".join(pattern_chars)
                        
                        days_present = sum([attendance.get("day1", False), 
                                          attendance.get("day2", False), 
                                          attendance.get("day3", False)])
                        
                        row = [
                            doc_id,
                            user_data.get("name", ""),
                            user_data.get("finalCommittee", ""),
                            day1, day2, day3,
                            pattern,
                            days_present
                        ]
                        writer.writerow(row)
            
            QMessageBox.information(self, "Success", f"Attendance data exported to:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Error exporting attendance data:\n{e}")

# Analytics View Widget
class AnalyticsView(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.users_data = {}
        self.attendance_data = {}
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel("📈 MatterID Analytics Dashboard")
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        
        self.refresh_btn = QPushButton("🔄 Refresh Data")
        self.export_report_btn = QPushButton("📊 Export Report")
        
        self.refresh_btn.clicked.connect(self.refresh_analytics)
        self.export_report_btn.clicked.connect(self.export_comprehensive_report)
        
        header_layout.addWidget(header_label)
        header_layout.addStretch()
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.export_report_btn)
        
        layout.addLayout(header_layout)
        
        # Scroll area for analytics content
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        self.analytics_widget = QWidget()
        self.analytics_layout = QVBoxLayout()
        self.analytics_widget.setLayout(self.analytics_layout)
        
        scroll_area.setWidget(self.analytics_widget)
        layout.addWidget(scroll_area)
        
        self.setLayout(layout)
        
        # Initialize with demo data
        self.refresh_analytics()
    
    def refresh_analytics(self):
        """Refresh all analytics data"""
        # Clear existing analytics
        for i in reversed(range(self.analytics_layout.count())):
            child = self.analytics_layout.itemAt(i).widget()
            if child:
                child.setParent(None)
        
        # Get data (use demo data if no real data available)
        if not self.users_data:
            self.users_data = DemoDataGenerator.generate_demo_delegates()
        if not self.attendance_data:
            self.attendance_data = DemoDataGenerator.generate_demo_attendance()
        
        # Generate analytics sections
        self.create_key_statistics()
        self.create_committee_distribution()
        self.create_school_analysis()
        self.create_attendance_analytics()
        self.create_registration_timeline()
    
    def create_key_statistics(self):
        """Create key statistics section"""
        stats_group = QGroupBox("📊 Key Statistics")
        stats_layout = QGridLayout()
        
        total_registrations = len(self.users_data)
        committees = set(data.get("finalCommittee", "") for data in self.users_data.values() if data.get("finalCommittee"))
        active_committees = len(committees) if committees else 0
        schools = set(data.get("school", "") for data in self.users_data.values() if data.get("school"))
        participating_schools = len(schools) if schools else 0
        
        # Calculate attendance statistics
        total_possible_days = total_registrations * 3
        total_present_days = 0
        for attendance in self.attendance_data.values():
            total_present_days += sum([
                attendance.get("day1", False),
                attendance.get("day2", False),
                attendance.get("day3", False)
            ])
        
        overall_attendance_rate = (total_present_days / total_possible_days * 100) if total_possible_days > 0 else 0
        
        # Create statistic cards
        stats = [
            ("👥 Total Registrations", str(total_registrations), MATTERID_COLORS['primary']),
            ("🏛️ Active Committees", str(active_committees), MATTERID_COLORS['accent']),
            ("🏫 Participating Schools", str(participating_schools), MATTERID_COLORS['success']),
            ("📈 Overall Attendance", f"{overall_attendance_rate:.1f}%", MATTERID_COLORS['warning'])
        ]
        
        for i, (title, value, color) in enumerate(stats):
            card = self.create_stat_card(title, value, color)
            row = i // 2
            col = i % 2
            stats_layout.addWidget(card, row, col)
        
        stats_group.setLayout(stats_layout)
        self.analytics_layout.addWidget(stats_group)
    
    def create_stat_card(self, title, value, color):
        """Create a statistic card widget"""
        card = QFrame()
        card.setFrameStyle(QFrame.Shape.Box)
        card.setStyleSheet(f"""
            QFrame {{
                border: 2px solid {color};
                border-radius: 8px;
                background-color: {MATTERID_COLORS['card_bg']};
                padding: 10px;
            }}
            QLabel {{
                border: none;
                color: {MATTERID_COLORS['text_primary']};
            }}
        """)
        
        layout = QVBoxLayout()
        
        title_label = QLabel(title)
        title_label.setStyleSheet(f"font-size: 12px; color: {MATTERID_COLORS['text_secondary']};")
        
        value_label = QLabel(value)
        value_label.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {color};")
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        card.setLayout(layout)
        
        return card
    
    def create_committee_distribution(self):
        """Create committee distribution chart"""
        committee_group = QGroupBox("🏛️ Committee Distribution")
        committee_layout = QVBoxLayout()
        
        # Count delegates per committee
        committee_counts = {}
        for data in self.users_data.values():
            committee = data.get("finalCommittee", "Unassigned")
            committee_counts[committee] = committee_counts.get(committee, 0) + 1
        
        # Create text-based bar chart
        if committee_counts:
            max_count = max(committee_counts.values())
            
            chart_text = QTextEdit()
            chart_text.setReadOnly(True)
            chart_text.setMaximumHeight(200)
            chart_text.setStyleSheet(f"""
                QTextEdit {{
                    background-color: {MATTERID_COLORS['card_bg']};
                    color: {MATTERID_COLORS['text_primary']};
                    border: 1px solid {MATTERID_COLORS['border']};
                    font-family: 'Courier New', monospace;
                    font-size: 11px;
                }}
            """)
            
            chart_content = "Committee Distribution:\n\n"
            for committee, count in sorted(committee_counts.items(), key=lambda x: x[1], reverse=True):
                bar_length = int((count / max_count) * 30) if max_count > 0 else 0
                bar = "█" * bar_length
                percentage = (count / len(self.users_data) * 100) if self.users_data else 0
                chart_content += f"{committee:<25} {bar:<30} {count:>3} ({percentage:>5.1f}%)\n"
            
            chart_text.setPlainText(chart_content)
            committee_layout.addWidget(chart_text)
        
        committee_group.setLayout(committee_layout)
        self.analytics_layout.addWidget(committee_group)
    
    def create_school_analysis(self):
        """Create school participation analysis"""
        school_group = QGroupBox("🏫 School Participation Analysis")
        school_layout = QVBoxLayout()
        
        # Count delegates per school
        school_counts = {}
        for data in self.users_data.values():
            school = data.get("school", "Unknown")
            school_counts[school] = school_counts.get(school, 0) + 1
        
        if school_counts:
            max_count = max(school_counts.values())
            
            chart_text = QTextEdit()
            chart_text.setReadOnly(True)
            chart_text.setMaximumHeight(200)
            chart_text.setStyleSheet(f"""
                QTextEdit {{
                    background-color: {MATTERID_COLORS['card_bg']};
                    color: {MATTERID_COLORS['text_primary']};
                    border: 1px solid {MATTERID_COLORS['border']};
                    font-family: 'Courier New', monospace;
                    font-size: 11px;
                }}
            """)
            
            chart_content = "School Participation:\n\n"
            for school, count in sorted(school_counts.items(), key=lambda x: x[1], reverse=True):
                bar_length = int((count / max_count) * 25) if max_count > 0 else 0
                bar = "█" * bar_length
                percentage = (count / len(self.users_data) * 100) if self.users_data else 0
                chart_content += f"{school:<30} {bar:<25} {count:>3} ({percentage:>5.1f}%)\n"
            
            chart_text.setPlainText(chart_content)
            school_layout.addWidget(chart_text)
        
        school_group.setLayout(school_layout)
        self.analytics_layout.addWidget(school_group)
    
    def create_attendance_analytics(self):
        """Create attendance pattern analysis"""
        attendance_group = QGroupBox("📅 Attendance Analytics")
        attendance_layout = QVBoxLayout()
        
        # Analyze attendance patterns
        patterns = {}
        day_stats = {"day1": 0, "day2": 0, "day3": 0}
        
        for doc_id, attendance in self.attendance_data.items():
            # Generate pattern
            pattern_chars = []
            for day in ["day1", "day2", "day3"]:
                is_present = attendance.get(day, False)
                pattern_chars.append("P" if is_present else "A")
                if is_present:
                    day_stats[day] += 1
            
            pattern = "".join(pattern_chars)
            patterns[pattern] = patterns.get(pattern, 0) + 1
        
        # Pattern descriptions
        pattern_descriptions = {
            "PPP": "Perfect Attendance",
            "PPA": "Missed Day 3",
            "PAP": "Missed Day 2", 
            "APP": "Missed Day 1",
            "PPO": "Only Days 1-2",
            "PAA": "Only Day 1",
            "APP": "Only Day 1",
            "AAP": "Only Day 3",
            "AAA": "Absent All Days"
        }
        
        chart_text = QTextEdit()
        chart_text.setReadOnly(True)
        chart_text.setMaximumHeight(250)
        chart_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {MATTERID_COLORS['card_bg']};
                color: {MATTERID_COLORS['text_primary']};
                border: 1px solid {MATTERID_COLORS['border']};
                font-family: 'Courier New', monospace;
                font-size: 11px;
            }}
        """)
        
        total_delegates = len(self.users_data)
        chart_content = "Attendance Patterns:\n\n"
        
        # Daily statistics
        chart_content += "Daily Attendance:\n"
        for day, count in day_stats.items():
            day_num = day[-1]
            percentage = (count / total_delegates * 100) if total_delegates > 0 else 0
            bar_length = int((count / total_delegates * 20)) if total_delegates > 0 else 0
            bar = "█" * bar_length
            chart_content += f"Day {day_num}: {bar:<20} {count:>3}/{total_delegates} ({percentage:>5.1f}%)\n"
        
        chart_content += "\nAttendance Patterns:\n"
        if patterns:
            max_pattern_count = max(patterns.values())
            for pattern, count in sorted(patterns.items(), key=lambda x: x[1], reverse=True):
                description = pattern_descriptions.get(pattern, "Custom Pattern")
                percentage = (count / total_delegates * 100) if total_delegates > 0 else 0
                bar_length = int((count / max_pattern_count * 15)) if max_pattern_count > 0 else 0
                bar = "█" * bar_length
                chart_content += f"{pattern} ({description:<18}) {bar:<15} {count:>3} ({percentage:>5.1f}%)\n"
        
        chart_text.setPlainText(chart_content)
        attendance_layout.addWidget(chart_text)
        
        attendance_group.setLayout(attendance_layout)
        self.analytics_layout.addWidget(attendance_group)
    
    def create_registration_timeline(self):
        """Create registration timeline analysis"""
        timeline_group = QGroupBox("📅 Registration Timeline")
        timeline_layout = QVBoxLayout()
        
        info_label = QLabel("Registration timeline analysis would show when delegates registered over time.")
        info_label.setStyleSheet(f"color: {MATTERID_COLORS['text_secondary']}; font-style: italic; padding: 10px;")
        timeline_layout.addWidget(info_label)
        
        # Demo timeline data
        demo_text = QTextEdit()
        demo_text.setReadOnly(True)
        demo_text.setMaximumHeight(120)
        demo_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {MATTERID_COLORS['card_bg']};
                color: {MATTERID_COLORS['text_primary']};
                border: 1px solid {MATTERID_COLORS['border']};
                font-family: 'Courier New', monospace;
                font-size: 11px;
            }}
        """)
        
        demo_content = """Registration Timeline (Demo Data):

Week 1: ████████████████████ 8 registrations (80%)
Week 2: ████                 2 registrations (20%)
Week 3: ██                   0 registrations (0%)

Peak registration period: Week 1
Average daily registrations: 1.4"""
        
        demo_text.setPlainText(demo_content)
        timeline_layout.addWidget(demo_text)
        
        timeline_group.setLayout(timeline_layout)
        self.analytics_layout.addWidget(timeline_group)
    
    def export_comprehensive_report(self):
        """Export comprehensive analytics report"""
        default_filename = f"matterid_analytics_report_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Analytics Report", default_filename, "CSV Files (*.csv)")
        if not file_path:
            return
        
        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file)
                
                # Report header
                writer.writerow([f"MatterID - Manager v2.5 Analytics Report"])
                writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
                writer.writerow([])
                
                # Key statistics
                writer.writerow(["KEY STATISTICS"])
                writer.writerow(["Metric", "Value"])
                writer.writerow(["Total Registrations", len(self.users_data)])
                
                committees = set(data.get("finalCommittee", "") for data in self.users_data.values() if data.get("finalCommittee"))
                writer.writerow(["Active Committees", len(committees)])
                
                schools = set(data.get("school", "") for data in self.users_data.values() if data.get("school"))
                writer.writerow(["Participating Schools", len(schools)])
                writer.writerow([])
                
                # Committee distribution
                writer.writerow(["COMMITTEE DISTRIBUTION"])
                writer.writerow(["Committee", "Count", "Percentage"])
                committee_counts = {}
                for data in self.users_data.values():
                    committee = data.get("finalCommittee", "Unassigned")
                    committee_counts[committee] = committee_counts.get(committee, 0) + 1
                
                for committee, count in sorted(committee_counts.items(), key=lambda x: x[1], reverse=True):
                    percentage = (count / len(self.users_data) * 100) if self.users_data else 0
                    writer.writerow([committee, count, f"{percentage:.1f}%"])
                writer.writerow([])
                
                # School analysis
                writer.writerow(["SCHOOL PARTICIPATION"])
                writer.writerow(["School", "Count", "Percentage"])
                school_counts = {}
                for data in self.users_data.values():
                    school = data.get("school", "Unknown")
                    school_counts[school] = school_counts.get(school, 0) + 1
                
                for school, count in sorted(school_counts.items(), key=lambda x: x[1], reverse=True):
                    percentage = (count / len(self.users_data) * 100) if self.users_data else 0
                    writer.writerow([school, count, f"{percentage:.1f}%"])
                writer.writerow([])
                
                # Attendance analysis
                writer.writerow(["ATTENDANCE ANALYSIS"])
                writer.writerow(["Pattern", "Description", "Count", "Percentage"])
                
                patterns = {}
                for attendance in self.attendance_data.values():
                    pattern_chars = []
                    for day in ["day1", "day2", "day3"]:
                        pattern_chars.append("P" if attendance.get(day, False) else "A")
                    pattern = "".join(pattern_chars)
                    patterns[pattern] = patterns.get(pattern, 0) + 1
                
                pattern_descriptions = {
                    "PPP": "Perfect Attendance",
                    "PPA": "Missed Day 3",
                    "PAP": "Missed Day 2",
                    "APP": "Missed Day 1",
                    "PAA": "Only Day 1",
                    "AAP": "Only Day 3",
                    "AAA": "Absent All Days"
                }
                
                for pattern, count in sorted(patterns.items(), key=lambda x: x[1], reverse=True):
                    description = pattern_descriptions.get(pattern, "Custom Pattern")
                    percentage = (count / len(self.users_data) * 100) if self.users_data else 0
                    writer.writerow([pattern, description, count, f"{percentage:.1f}%"])
            
            QMessageBox.information(self, "Success", f"Comprehensive analytics report exported to:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Error exporting analytics report:\n{e}")
    
    def update_data(self, users_data, attendance_data=None):
        """Update analytics with new data"""
        self.users_data = users_data
        if attendance_data:
            self.attendance_data = attendance_data
        self.refresh_analytics()

# User Card Widget
class UserCard(QFrame):
    edit_requested = pyqtSignal(str)
    
    def __init__(self, doc_id, user_data):
        super().__init__()
        self.doc_id = doc_id
        self.user_data = user_data
        self.init_ui()
    
    def init_ui(self):
        self.setFrameStyle(QFrame.Shape.Box)
        self.setStyleSheet(f"""
            QFrame {{
                border: 2px solid {MATTERID_COLORS['primary']};
                border-radius: 8px;
                background-color: {MATTERID_COLORS['card_bg']};
                margin: 5px;
            }}
            QFrame:hover {{
                border-color: {MATTERID_COLORS['hover']};
                background-color: #404040;
            }}
        """)
        self.setFixedSize(250, 180)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Initials circle
        initials = get_initials(self.user_data.get("name", ""))
        initials_label = QLabel(initials)
        initials_label.setStyleSheet(f"""
            QLabel {{
                background-color: {MATTERID_COLORS['primary']};
                color: white;
                border-radius: 25px;
                font-size: 18px;
                font-weight: bold;
                border: none;
            }}
        """)
        initials_label.setFixedSize(50, 50)
        initials_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Name
        name_label = QLabel(self.user_data.get("name", "Unknown"))
        name_label.setStyleSheet("font-weight: bold; font-size: 14px; color: white; border: none;")
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setWordWrap(True)
        
        # Email
        email_label = QLabel(self.user_data.get("email", ""))
        email_label.setStyleSheet("font-size: 10px; color: #cccccc; border: none;")
        email_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        email_label.setWordWrap(True)
        
        # Phone
        phone_label = QLabel(self.user_data.get("phone", ""))
        phone_label.setStyleSheet("font-size: 10px; color: #cccccc; border: none;")
        phone_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Committee & Portfolio
        committee = self.user_data.get("finalCommittee", "Not Assigned")
        portfolio = self.user_data.get("finalPortfolio", "Not Assigned")
        
        committee_label = QLabel(f"📋 {committee}")
        committee_label.setStyleSheet("font-size: 9px; color: #aaaaaa; border: none;")
        committee_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        committee_label.setWordWrap(True)
        
        portfolio_label = QLabel(f"🎯 {portfolio}")
        portfolio_label.setStyleSheet("font-size: 9px; color: #aaaaaa; border: none;")
        portfolio_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        portfolio_label.setWordWrap(True)
        
        # Layout
        header_layout = QHBoxLayout()
        header_layout.addStretch()
        header_layout.addWidget(initials_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        layout.addWidget(name_label)
        layout.addWidget(email_label)
        layout.addWidget(phone_label)
        layout.addWidget(committee_label)
        layout.addWidget(portfolio_label)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.edit_requested.emit(self.doc_id)

# User View Widget
class UserView(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.user_cards = {}
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("👥 Delegate Profiles")
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        # Search and filter for user view
        search_layout = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Search delegates by name, email, or phone...")
        self.search_edit.textChanged.connect(self.filter_users)
        
        search_layout.addWidget(QLabel("Search:"))
        search_layout.addWidget(self.search_edit)
        
        layout.addLayout(search_layout)
        
        # Scroll area for user cards
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        self.cards_widget = QWidget()
        self.cards_layout = QGridLayout()
        self.cards_layout.setSpacing(10)
        self.cards_widget.setLayout(self.cards_layout)
        
        scroll_area.setWidget(self.cards_widget)
        layout.addWidget(scroll_area)
        
        self.setLayout(layout)
    
    def update_users(self, users_data):
        # Clear existing cards
        for i in reversed(range(self.cards_layout.count())):
            child = self.cards_layout.itemAt(i).widget()
            if child:
                child.setParent(None)
        
        self.user_cards.clear()
        
        # Add new cards
        row = 0
        col = 0
        cols_per_row = 4
        
        for doc_id, user_data in users_data.items():
            if user_data:
                card = UserCard(doc_id, user_data)
                card.edit_requested.connect(self.main_window.edit_user)
                self.user_cards[doc_id] = card
                
                self.cards_layout.addWidget(card, row, col)
                col += 1
                if col >= cols_per_row:
                    col = 0
                    row += 1
        
        # Add stretch to fill remaining space
        self.cards_layout.setRowStretch(row + 1, 1)
        self.cards_layout.setColumnStretch(cols_per_row, 1)
    
    def filter_users(self):
        search_text = self.search_edit.text().lower()
        
        for doc_id, card in self.user_cards.items():
            user_data = card.user_data
            should_show = True
            
            if search_text:
                name = str(user_data.get("name", "")).lower()
                email = str(user_data.get("email", "")).lower()
                phone = str(user_data.get("phone", "")).lower()
                
                should_show = (search_text in name or 
                             search_text in email or 
                             search_text in phone)
            
            card.setVisible(should_show)

# User Edit Dialog
class UserEditDialog(QDialog):
    def __init__(self, doc_id, user_data, config_manager, parent=None):
        super().__init__(parent)
        self.doc_id = doc_id
        self.user_data = user_data.copy()
        self.config_manager = config_manager
        self.field_widgets = {}
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle(f"Edit Delegate: {self.user_data.get('name', 'Unknown')}")
        self.setModal(True)
        self.resize(450, 650)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {MATTERID_COLORS['background']};
                color: {MATTERID_COLORS['text_primary']};
            }}
            QGroupBox {{
                border: 2px solid {MATTERID_COLORS['primary']};
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: {MATTERID_COLORS['primary']};
            }}
        """)
        
        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("✏️ Edit Delegate")
        header_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        # User info header
        header_layout = QHBoxLayout()
        
        initials = get_initials(self.user_data.get("name", ""))
        initials_label = QLabel(initials)
        initials_label.setStyleSheet(f"""
            QLabel {{
                background-color: {MATTERID_COLORS['primary']};
                color: white;
                border-radius: 30px;
                font-size: 20px;
                font-weight: bold;
            }}
        """)
        initials_label.setFixedSize(60, 60)
        initials_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        user_info_layout = QVBoxLayout()
        name_label = QLabel(self.user_data.get("name", "Unknown"))
        name_label.setStyleSheet("font-size: 16px; font-weight: bold; color: white;")
        doc_id_label = QLabel(f"ID: {self.doc_id}")
        doc_id_label.setStyleSheet("font-size: 12px; color: #cccccc;")
        
        user_info_layout.addWidget(name_label)
        user_info_layout.addWidget(doc_id_label)
        
        header_layout.addWidget(initials_label)
        header_layout.addLayout(user_info_layout)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        
        # Form fields
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        form_widget = QWidget()
        form_layout = QFormLayout()
        
        # Get editable fields from config
        config = self.config_manager.get_config()
        for column in config.get("table_columns", []):
            field_name = column.get("field")
            display_name = column.get("display")
            editable = column.get("editable", True)
            
            if field_name and editable and field_name != "updatedAt":
                field_widget = QLineEdit()
                field_widget.setText(str(self.user_data.get(field_name, "")))
                self.field_widgets[field_name] = field_widget
                form_layout.addRow(f"{display_name}:", field_widget)
        
        form_widget.setLayout(form_layout)
        scroll_area.setWidget(form_widget)
        layout.addWidget(scroll_area)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.save_btn = QPushButton("💾 Save Changes")
        self.cancel_btn = QPushButton("❌ Cancel")
        
        self.save_btn.clicked.connect(self.save_changes)
        self.cancel_btn.clicked.connect(self.reject)
        
        self.save_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MATTERID_COLORS['success']};
                color: white;
                padding: 10px 20px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #45a049;
            }}
        """)
        
        self.cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MATTERID_COLORS['error']};
                color: white;
                padding: 10px 20px;
            }}
            QPushButton:hover {{
                background-color: #d32f2f;
            }}
        """)
        
        button_layout.addStretch()
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def save_changes(self):
        for field_name, widget in self.field_widgets.items():
            self.user_data[field_name] = widget.text().strip()
        self.accept()
    
    def get_updated_data(self):
        return self.user_data

# Initial MatterID Splash Screen
def show_matterid_splash_screen(app):
    dialog = QDialog()
    dialog.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Dialog)
    dialog.setStyleSheet(f"""
        QDialog {{
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                stop:0 {MATTERID_COLORS['background']}, 
                stop:1 {MATTERID_COLORS['card_bg']});
            border: 3px solid {MATTERID_COLORS['primary']};
            border-radius: 15px;
        }}
        QLabel {{
            color: {MATTERID_COLORS['text_primary']};
        }}
    """)
    
    layout = QVBoxLayout()
    layout.setSpacing(20)
    
    # MatterID Logo
    logo_label = QLabel("🆔")
    logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    logo_label.setStyleSheet(f"font-size: 64px; color: {MATTERID_COLORS['primary']};")
    
    # Title
    title_label = QLabel("MatterID - Manager v2.5")
    title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    title_label.setStyleSheet(f"font-weight: bold; font-size: 28px; color: {MATTERID_COLORS['primary']};")
    
    # Subtitle
    subtitle_label = QLabel("Professional Model United Nations Management System")
    subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    subtitle_label.setStyleSheet(f"font-size: 14px; color: {MATTERID_COLORS['text_secondary']};")
    
    # Powered by
    powered_label = QLabel("Powered by MatterID Authentication")
    powered_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    powered_label.setStyleSheet(f"font-size: 12px; color: {MATTERID_COLORS['text_muted']}; font-style: italic;")
    
    layout.addWidget(logo_label)
    layout.addWidget(title_label)
    layout.addWidget(subtitle_label)
    layout.addWidget(powered_label)
    
    dialog.setLayout(layout)
    dialog.resize(500, 300)
    dialog.show()
    
    QTimer.singleShot(3000, dialog.accept)
    app.processEvents()
    dialog.exec()

# MainWindow
DEBOUNCE_TIME_MS = 350
SAVE_FEEDBACK_DURATION_MS = 1500
UNSAVED_COLOR = QColor(255, 255, 204)
SAVE_SUCCESS_COLOR = QColor(204, 255, 204)
SAVE_ERROR_COLOR = QColor(255, 204, 204)

class MainWindow(QMainWindow):
    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        logging.info("Initializing MatterID - Manager v2.5…")
        self.setWindowTitle("MatterID - Manager v2.5 • Configuration")
        self.resize(1400, 900)
        self.unsaved_changes = set()
        self.all_loaded_data = {}
        self.attendance_data = {}
        self.demo_mode = False

        self.init_ui()
        self.load_data()

    def init_ui(self):
        self.tab_widget = QTabWidget()
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        self.setCentralWidget(self.tab_widget)
        
        # Configuration Tab
        self.config_tab = ConfigTab(self.config_manager)
        self.config_tab.config_changed.connect(self.on_config_changed)
        self.tab_widget.addTab(self.config_tab, "⚙️ Configuration")
        
        # Spreadsheet Tab
        self.spreadsheet_tab = QWidget()
        self.init_spreadsheet_tab()
        self.tab_widget.addTab(self.spreadsheet_tab, "📊 Spreadsheet View")
        
        # User View Tab
        self.user_view = UserView(self)
        self.tab_widget.addTab(self.user_view, "👥 Delegate View")
        
        # Attendance Tab (NEW)
        self.attendance_view = AttendanceView(self)
        self.tab_widget.addTab(self.attendance_view, "✅ Attendance")
        
        # Analytics Tab (NEW)
        self.analytics_view = AnalyticsView(self)
        self.tab_widget.addTab(self.analytics_view, "📈 Analytics")
        
        # Status Bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.status_row_count_label = QLabel("Rows: 0 / 0")
        self.statusBar.addPermanentWidget(self.status_row_count_label)
        
        # Version label
        version_label = QLabel("MatterID - Manager v2.5")
        version_label.setStyleSheet(f"color: {MATTERID_COLORS['text_muted']}; font-style: italic;")
        self.statusBar.addPermanentWidget(version_label)
        
        self.update_status("Ready • MatterID - Manager v2.5")

    def on_tab_changed(self, index):
        """Update window title when tab changes"""
        tab_names = ["Configuration", "Spreadsheet View", "Delegate View", "Attendance", "Analytics"]
        if 0 <= index < len(tab_names):
            self.setWindowTitle(f"MatterID - Manager v2.5 • {tab_names[index]}")

    def init_spreadsheet_tab(self):
        layout = QVBoxLayout()
        
        # Header
        header_label = QLabel("📊 Data Management")
        header_label.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {MATTERID_COLORS['primary']}; margin: 10px;")
        layout.addWidget(header_label)
        
        # Search and filter controls
        self.search_field_combo = QComboBox()
        self.search_field_combo.addItems(["Document ID", "name", "email", "phone", "school"])
        self.search_text_edit = QLineEdit()
        self.search_text_edit.setPlaceholderText("🔍 Enter search query (Press / to focus)…")
        self.search_button = QPushButton("🔍 Search")
        self.reset_button = QPushButton("🔄 Reset")

        self.search_debounce_timer = QTimer(self)
        self.search_debounce_timer.setSingleShot(True)
        self.search_debounce_timer.timeout.connect(self.trigger_search)

        self.search_button.clicked.connect(self.trigger_search)
        self.reset_button.clicked.connect(self.reset_view)
        self.search_text_edit.textChanged.connect(self.on_search_text_changed)

        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search Field:"))
        search_layout.addWidget(self.search_field_combo)
        search_layout.addWidget(self.search_text_edit)
        search_layout.addWidget(self.search_button)
        search_layout.addWidget(self.reset_button)

        # Filter controls
        self.filter_field_combo = QComboBox()
        self.filter_field_combo.addItems(["committeePreferences", "finalCommittee", "paymentStatus"])
        self.filter_text_edit = QLineEdit()
        self.filter_text_edit.setPlaceholderText("Enter exact filter value")
        self.filter_button = QPushButton("🎯 Apply Filter")
        self.download_button = QPushButton("💾 Download Filtered")

        self.filter_debounce_timer = QTimer(self)
        self.filter_debounce_timer.setSingleShot(True)
        self.filter_debounce_timer.timeout.connect(self.trigger_filter)

        self.filter_button.clicked.connect(self.trigger_filter)
        self.download_button.clicked.connect(self.download_filtered_data)
        self.filter_text_edit.textChanged.connect(self.on_filter_text_changed)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter by:"))
        filter_layout.addWidget(self.filter_field_combo)
        filter_layout.addWidget(self.filter_text_edit)
        filter_layout.addWidget(self.filter_button)
        filter_layout.addWidget(self.download_button)

        # Table
        self.table = QTableWidget()
        self.update_table_structure()
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked |
            QTableWidget.EditTrigger.EditKeyPressed |
            QTableWidget.EditTrigger.AnyKeyPressed
        )
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)
        self.table.itemSelectionChanged.connect(self.update_button_states)

        # Action buttons
        self.refresh_button = QPushButton("🔄 Refresh")
        self.refresh_button.setToolTip("Reload all data from Firestore")
        self.refresh_button.clicked.connect(self.refresh_data)

        self.save_all_button = QPushButton("💾 Save All Changes")
        self.save_all_button.setToolTip("Save all rows with unsaved changes")
        self.save_all_button.clicked.connect(self.autosave_all_rows)

        self.export_all_button = QPushButton("📤 Export All")
        self.export_all_button.setToolTip("Export entire dataset from Firestore as CSV")
        self.export_all_button.clicked.connect(self.export_all_data)

        self.delete_button = QPushButton("🗑️ Delete Selected")
        self.delete_button.setToolTip("Delete selected row(s) from Firestore")
        self.delete_button.clicked.connect(self.delete_selected_documents)
        self.delete_button.setObjectName("delete_button")
        self.delete_button.setEnabled(False)

        action_layout = QHBoxLayout()
        action_layout.addWidget(self.refresh_button)
        action_layout.addWidget(self.save_all_button)
        action_layout.addWidget(self.export_all_button)
        action_layout.addStretch()
        action_layout.addWidget(self.delete_button)

        layout.addLayout(search_layout)
        layout.addLayout(filter_layout)
        layout.addWidget(self.table)
        layout.addLayout(action_layout)
        
        self.spreadsheet_tab.setLayout(layout)
        self.table.itemChanged.connect(self.handle_cell_change)

    def update_table_structure(self):
        config = self.config_manager.get_config()
        table_columns = config.get("table_columns", [])
        
        self.table.setColumnCount(len(table_columns))
        self.table.setHorizontalHeaderLabels([col.get("display", "") for col in table_columns])
        
        self.field_mapping = {
            i: col.get("field")
            for i, col in enumerate(table_columns)
            if col.get("field") is not None and col.get("editable", True)
        }

    def on_config_changed(self):
        self.update_table_structure()
        self.load_data(reload_all=False)

    def edit_user(self, doc_id):
        user_data = self.all_loaded_data.get(doc_id)
        if user_data:
            dialog = UserEditDialog(doc_id, user_data, self.config_manager, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                updated_data = dialog.get_updated_data()
                try:
                    updated_data["updatedAt"] = firestore.SERVER_TIMESTAMP if db else datetime.now()
                    
                    if db:
                        config = self.config_manager.get_config()
                        collection_name = config.get("collection_name", "registrations")
                        db.collection(collection_name).document(doc_id).update(updated_data)
                    
                    self.all_loaded_data[doc_id] = updated_data
                    self.populate_table()
                    self.user_view.update_users(self.all_loaded_data)
                    
                    QMessageBox.information(self, "Success", "Delegate updated successfully!")
                    
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to update delegate: {e}")

    def load_data(self, reload_all=True):
        if reload_all:
            self.update_status("Loading data…")
            QApplication.processEvents()
            
            try:
                if db is None:
                    # Demo mode
                    self.demo_mode = True
                    self.all_loaded_data = DemoDataGenerator.generate_demo_delegates()
                    self.attendance_data = DemoDataGenerator.generate_demo_attendance()
                    logging.info(f"Demo mode: Loaded {len(self.all_loaded_data)} demo delegates.")
                    self.update_status("Ready • Demo Mode • MatterID - Manager v2.5")
                else:
                    # Production mode
                    config = self.config_manager.get_config()
                    collection_name = config.get("collection_name", "registrations")
                    
                    # Load registrations
                    collection_ref = db.collection(collection_name)
                    docs = list(collection_ref.stream())
                    self.all_loaded_data = {
                        doc.id: doc.to_dict()
                        for doc in docs if doc.exists
                    }
                    
                    # Load attendance data
                    try:
                        attendance_ref = db.collection("attendance")
                        attendance_docs = list(attendance_ref.stream())
                        self.attendance_data = {
                            doc.id: doc.to_dict()
                            for doc in attendance_docs if doc.exists
                        }
                        logging.info(f"Loaded {len(self.attendance_data)} attendance records.")
                    except Exception as e:
                        logging.warning(f"Could not load attendance data: {e}")
                        self.attendance_data = {}
                    
                    logging.info(f"Loaded {len(self.all_loaded_data)} documents from Firestore.")

                if self.unsaved_changes:
                    reply = QMessageBox.question(
                        self, "Unsaved Changes",
                        f"You have {len(self.unsaved_changes)} unsaved change(s). "
                        "Reloading will discard them. Continue?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.No:
                        self.update_status("Reload cancelled.")
                        return
                    self.unsaved_changes.clear()

            except Exception as e:
                logging.error(f"Error loading data: {e}\n{traceback.format_exc()}")
                # Fallback to demo mode
                self.demo_mode = True
                self.all_loaded_data = DemoDataGenerator.generate_demo_delegates()
                self.attendance_data = DemoDataGenerator.generate_demo_attendance()
                QMessageBox.warning(self, "Connection Error", 
                                  f"Could not connect to database. Running in demo mode.\nError: {e}")
                self.update_status("Ready • Demo Mode • MatterID - Manager v2.5")

        self.populate_table()
        self.user_view.update_users(self.all_loaded_data)
        self.attendance_view.update_attendance_data(self.all_loaded_data, self.attendance_data)
        self.analytics_view.update_data(self.all_loaded_data, self.attendance_data)

    def populate_table(self):
        self.update_status("Filtering and displaying data…")
        QApplication.processEvents()

        self.table.setSortingEnabled(False)
        self.table.blockSignals(True)
        self.table.clearContents()

        search_field = self.search_field_combo.currentText()
        search_value = self.search_text_edit.text().strip().lower()
        filter_field = self.filter_field_combo.currentText()
        filter_value = self.filter_text_edit.text().strip().lower()

        filtered_ids = []
        for doc_id, data in self.all_loaded_data.items():
            if data is None:
                continue
            match = True

            if search_value:
                if search_field == "Document ID":
                    if search_value not in doc_id.lower():
                        match = False
                else:
                    field_data = str(data.get(search_field, "")).lower()
                    if search_value not in field_data:
                        match = False

            if match and filter_value:
                field_data = str(data.get(filter_field, "")).lower()
                if filter_value != field_data:
                    match = False

            if match:
                filtered_ids.append(doc_id)

        self.table.setRowCount(len(filtered_ids))
        visible_row_count = 0

        config = self.config_manager.get_config()
        table_columns = config.get("table_columns", [])

        for row_position, doc_id in enumerate(filtered_ids):
            data = self.all_loaded_data.get(doc_id)
            if data is None:
                continue

            is_unsaved = (doc_id in self.unsaved_changes)
            visible_row_count += 1

            for col_index, column_config in enumerate(table_columns):
                field_name = column_config.get("field")
                display_name = column_config.get("display", "")
                editable = column_config.get("editable", True)
                
                item = None
                if field_name is None:
                    item_widget = QTableWidgetItem(doc_id)
                    item_widget.setFlags(item_widget.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(row_position, col_index, item_widget)
                    item = item_widget
                else:
                    current_value = data.get(field_name, "")
                    is_readonly_col = not editable

                    if field_name == "finalCommittee":
                        combo = QComboBox()
                        options = ["allot", "Lok Sabha", "UNHRC", "UNGA-Disec", "UNCSW", "Continuous Crisis Committee", "International Press"]
                        combo.addItems(options)
                        if current_value not in options and current_value:
                            combo.addItem(current_value)
                        idx = combo.findText(str(current_value), Qt.MatchFlag.MatchFixedString)
                        combo.setCurrentIndex(idx if idx != -1 else 0)
                        combo.currentIndexChanged.connect(
                            lambda state, r=row_position, d=doc_id: self.mark_unsaved(d, r)
                        )
                        self.table.setCellWidget(row_position, col_index, combo)
                        item = combo
                    else:
                        display_text = str(current_value)
                        if field_name == "updatedAt" and current_value:
                            display_text = format_timestamp(current_value)

                        item_widget = QTableWidgetItem(display_text)
                        if is_readonly_col:
                            item_widget.setFlags(item_widget.flags() & ~Qt.ItemFlag.ItemIsEditable)
                            item_widget.setForeground(QBrush(QColor("lightgray")))
                        self.table.setItem(row_position, col_index, item_widget)
                        item = item_widget

                if is_unsaved and item:
                    self.set_row_color(row_position, UNSAVED_COLOR)

        self.table.resizeColumnsToContents()
        self.table.blockSignals(False)
        self.table.setSortingEnabled(True)

        total_docs = len(self.all_loaded_data)
        self.status_row_count_label.setText(f"Rows: {visible_row_count} / {total_docs}")
        status_msg = "Ready • Demo Mode • MatterID - Manager v2.5" if self.demo_mode else "Ready • MatterID - Manager v2.5"
        self.update_status(status_msg)
        logging.info(f"Table populated with {visible_row_count} rows (of {total_docs} total).")

    def refresh_data(self):
        logging.info("Refresh requested.")
        self.load_data(reload_all=True)

    def reset_view(self):
        logging.info("Reset view requested.")
        self.search_text_edit.blockSignals(True)
        self.filter_text_edit.blockSignals(True)
        self.search_text_edit.clear()
        self.filter_text_edit.clear()
        self.search_text_edit.blockSignals(False)
        self.filter_text_edit.blockSignals(False)
        self.load_data(reload_all=False)

    def on_search_text_changed(self):
        self.search_debounce_timer.start(DEBOUNCE_TIME_MS)

    def on_filter_text_changed(self):
        self.filter_debounce_timer.start(DEBOUNCE_TIME_MS)

    def trigger_search(self):
        logging.info("Triggering search…")
        self.search_debounce_timer.stop()
        self.load_data(reload_all=False)

    def trigger_filter(self):
        logging.info("Triggering filter…")
        self.filter_debounce_timer.stop()
        self.load_data(reload_all=False)

    def mark_unsaved(self, doc_id, row_hint=-1):
        if doc_id not in self.unsaved_changes:
            self.unsaved_changes.add(doc_id)
            logging.debug(f"Marked doc '{doc_id}' as unsaved.")
            row = self.find_row_by_doc_id(doc_id)
            if row != -1:
                self.set_row_color(row, UNSAVED_COLOR)

    def handle_cell_change(self, item):
        if item and not self.table.signalsBlocked():
            row = item.row()
            col = item.column()
            if col in self.field_mapping:
                doc_id_item = self.table.item(row, 0)
                if doc_id_item:
                    doc_id = doc_id_item.text()
                    self.mark_unsaved(doc_id, row)

    def set_row_color(self, row, color=None):
        if row < 0 or row >= self.table.rowCount():
            return
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            widget = self.table.cellWidget(row, col)
            if widget:
                if color:
                    widget.setStyleSheet(f"background-color: {color.name()};")
                else:
                    widget.setStyleSheet("")
            elif item:
                if color:
                    item.setBackground(color)
                else:
                    item.setBackground(QColor(Qt.GlobalColor.transparent))

    def flash_row_color(self, row, color):
        self.set_row_color(row, color)
        QTimer.singleShot(SAVE_FEEDBACK_DURATION_MS, lambda: self.reset_flashed_color(row))

    def reset_flashed_color(self, row):
        doc_id_item = self.table.item(row, 0)
        if doc_id_item:
            doc_id = doc_id_item.text()
            if doc_id in self.unsaved_changes:
                self.set_row_color(row, UNSAVED_COLOR)
            else:
                self.set_row_color(row)

    def find_row_by_doc_id(self, doc_id):
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item and item.text() == doc_id:
                return r
        return -1

    def save_row(self, row):
        if row < 0 or row >= self.table.rowCount():
            return
        doc_id_item = self.table.item(row, 0)
        if not doc_id_item:
            return
        doc_id = doc_id_item.text()
        if not doc_id:
            return

        updated_data = {}
        validation_ok = True
        invalid_field_info = None

        config = self.config_manager.get_config()
        table_columns = config.get("table_columns", [])

        self.table.blockSignals(True)
        try:
            for col_index, column_config in enumerate(table_columns):
                field_name = column_config.get("field")
                editable = column_config.get("editable", True)
                
                if field_name and editable:
                    widget = self.table.cellWidget(row, col_index)
                    value = None
                    if widget is not None and isinstance(widget, QComboBox):
                        value = widget.currentText()
                    else:
                        item = self.table.item(row, col_index)
                        if item is not None:
                            value = item.text().strip()

                    if value is not None:
                        updated_data[field_name] = value
                        if field_name == "email" and value and not is_valid_email(value):
                            validation_ok = False
                            invalid_field_info = (row, col_index, "Invalid email format")
                            break

            if validation_ok and updated_data:
                updated_data["updatedAt"] = firestore.SERVER_TIMESTAMP if db else datetime.now()

        finally:
            self.table.blockSignals(False)

        if not validation_ok:
            if invalid_field_info:
                r, c, msg = invalid_field_info
                logging.warning(f"Validation failed for doc '{doc_id}': {msg}")
                QMessageBox.warning(self, "Validation Error", f"Cannot save row {r+1}:\n{msg}")
                self.flash_row_color(r, SAVE_ERROR_COLOR)
            return

        if not updated_data:
            logging.debug(f"Save skipped for row {row}: No data changes detected.")
            return

        try:
            self.update_status(f"Saving {doc_id}…")
            QApplication.processEvents()

            if db:
                config = self.config_manager.get_config()
                collection_name = config.get("collection_name", "registrations")
                logging.info(f"Updating Firestore doc '{doc_id}' with data: "
                           f"{ {k: v for k, v in updated_data.items() if k != 'updatedAt'} }")
                db.collection(collection_name).document(doc_id).update(updated_data)

                updated_doc = db.collection(collection_name).document(doc_id).get()
                if updated_doc.exists:
                    self.all_loaded_data[doc_id] = updated_doc.to_dict()
            else:
                # Demo mode - just update local data
                self.all_loaded_data[doc_id].update(updated_data)

            if doc_id in self.unsaved_changes:
                self.unsaved_changes.remove(doc_id)

            self.flash_row_color(row, SAVE_SUCCESS_COLOR)
            self.update_status(f"Saved {doc_id}")
            logging.info(f"Successfully updated document {doc_id}")

        except Exception as e:
            logging.error(f"Error updating document {doc_id}: {e}\n{traceback.format_exc()}")
            QMessageBox.critical(self, "Save Error", f"Error saving document {doc_id}:\n{e}")
            self.flash_row_color(row, SAVE_ERROR_COLOR)
            self.update_status(f"Error saving {doc_id}", error=False)

    def autosave_all_rows(self):
        doc_ids_to_save = list(self.unsaved_changes)
        if not doc_ids_to_save:
            QMessageBox.information(self, "Save All", "No unsaved changes to save.")
            return

        rows_to_save = []
        for doc_id in doc_ids_to_save:
            row = self.find_row_by_doc_id(doc_id)
            if row != -1:
                rows_to_save.append(row)

        if not rows_to_save:
            QMessageBox.information(self, "Save All", "No unsaved changes found in the current view.")
            return

        progress = QProgressDialog("Saving all changes…", "Cancel", 0, len(rows_to_save), self)
        progress.setWindowTitle("Saving Data")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)
        QApplication.processEvents()

        saved_count = 0
        error_count = 0

        for i, row in enumerate(rows_to_save):
            if progress.wasCanceled():
                logging.info("Save All cancelled by user.")
                break
            self.save_row(row)
            doc_id_item = self.table.item(row, 0)
            if doc_id_item and doc_id_item.text() not in self.unsaved_changes:
                saved_count += 1
            else:
                error_count += 1
            progress.setValue(i + 1)
            QApplication.processEvents()

        progress.close()
        summary_msg = f"Save All finished.\nSuccessful: {saved_count}\nErrors: {error_count}"
        if error_count > 0:
            QMessageBox.warning(self, "Save All Complete", summary_msg)
        else:
            QMessageBox.information(self, "Save All Complete", summary_msg)
        self.update_status("Ready")

    def delete_selected_documents(self):
        selected_rows = sorted(
            list(set(item.row() for item in self.table.selectedItems())),
            reverse=True
        )
        if not selected_rows:
            QMessageBox.warning(self, "Delete Error", "No rows selected.")
            return

        doc_ids_to_delete = []
        for row in selected_rows:
            doc_id_item = self.table.item(row, 0)
            if doc_id_item and doc_id_item.text():
                doc_ids_to_delete.append(doc_id_item.text())

        if not doc_ids_to_delete:
            QMessageBox.warning(self, "Delete Error", "Could not find valid Document IDs.")
            return

        confirm = QMessageBox.question(
            self, "Confirm Deletion",
            f"Are you sure you want to delete {len(doc_ids_to_delete)} document(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        progress = QProgressDialog(f"Deleting {len(doc_ids_to_delete)} items…", "Cancel", 0, len(doc_ids_to_delete), self)
        progress.setWindowTitle("Deleting Data")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)

        deleted_count = 0
        error_count = 0

        for i, doc_id in enumerate(doc_ids_to_delete):
            if progress.wasCanceled():
                break
            try:
                if db and not self.demo_mode:
                    config = self.config_manager.get_config()
                    collection_name = config.get("collection_name", "registrations")
                    db.collection(collection_name).document(doc_id).delete()
                
                self.all_loaded_data.pop(doc_id, None)
                self.unsaved_changes.discard(doc_id)
                deleted_count += 1
            except Exception as e:
                error_count += 1
                logging.error(f"Error deleting document {doc_id}: {e}")
            progress.setValue(i + 1)
            QApplication.processEvents()

        progress.close()

        if deleted_count > 0:
            self.load_data(reload_all=False)

        summary_msg = f"Deletion finished.\nDeleted: {deleted_count}\nErrors: {error_count}"
        if error_count > 0:
            QMessageBox.warning(self, "Deletion Complete", summary_msg)
        else:
            QMessageBox.information(self, "Deletion Complete", summary_msg)
        self.update_status("Ready")
        self.update_button_states()

    def download_filtered_data(self):
        visible_rows = self.table.rowCount()
        if visible_rows == 0:
            QMessageBox.information(self, "Export Error", "No data currently visible in the table.")
            return

        default_filename = f"matterid_filtered_export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Filtered Data", default_filename, "CSV Files (*.csv)")
        if not file_path:
            return

        try:
            config = self.config_manager.get_config()
            table_columns = config.get("table_columns", [])
            
            with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file)
                header = [col.get("display", "") for col in table_columns]
                writer.writerow(header)
                
                for row in range(visible_rows):
                    row_data = []
                    doc_id_item = self.table.item(row, 0)
                    if doc_id_item:
                        doc_id = doc_id_item.text()
                        data = self.all_loaded_data.get(doc_id, {})
                        
                        for col_index, column_config in enumerate(table_columns):
                            field_name = column_config.get("field")
                            if field_name is None:
                                row_data.append(doc_id)
                            elif field_name == "updatedAt":
                                row_data.append(format_timestamp(data.get(field_name)))
                            else:
                                row_data.append(data.get(field_name, ""))
                        writer.writerow(row_data)
            
            QMessageBox.information(self, "Success", f"Visible data successfully exported to:\n{file_path}")
            self.update_status("Export complete.")
        except Exception as e:
            logging.error(f"Error exporting filtered data: {e}")
            QMessageBox.critical(self, "Export Error", f"Error exporting data:\n{e}")

    def export_all_data(self):
        total_docs = len(self.all_loaded_data)
        if total_docs == 0:
            QMessageBox.information(self, "Export Error", "No data loaded to export.")
            return

        default_filename = f"matterid_full_export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, "Export All Data", default_filename, "CSV Files (*.csv)")
        if not file_path:
            return

        try:
            config = self.config_manager.get_config()
            table_columns = config.get("table_columns", [])
            
            with open(file_path, mode="w", newline="", encoding="utf-8") as csv_file:
                writer = csv.writer(csv_file)
                header = [col.get("display", "") for col in table_columns]
                writer.writerow(header)
                
                for doc_id, data in self.all_loaded_data.items():
                    if data is None:
                        continue
                    row_data = []
                    for column_config in table_columns:
                        field_name = column_config.get("field")
                        if field_name is None:
                            row_data.append(doc_id)
                        elif field_name == "updatedAt":
                            row_data.append(format_timestamp(data.get(field_name)))
                        else:
                            row_data.append(data.get(field_name, ""))
                    writer.writerow(row_data)
            
            QMessageBox.information(self, "Success", f"All {total_docs} documents exported to:\n{file_path}")
            self.update_status("Full export complete.")
        except Exception as e:
            logging.error(f"Error exporting all data: {e}")
            QMessageBox.critical(self, "Export Error", f"Error exporting data:\n{e}")

    def update_status(self, message, timeout=0, error=False):
        if error:
            self.statusBar.setStyleSheet(f"color: {MATTERID_COLORS['error']};")
        else:
            self.statusBar.setStyleSheet("")
        self.statusBar.showMessage(message, timeout)
        logging.info(f"Status: {message}")

    def update_button_states(self):
        has_selection = bool(self.table.selectedItems())
        self.delete_button.setEnabled(has_selection)

    def keyPressEvent(self, event):
        key = event.key()
        modifiers = event.modifiers()

        if key == Qt.Key.Key_Slash and modifiers == Qt.KeyboardModifier.NoModifier:
            self.search_text_edit.setFocus()
            self.search_text_edit.selectAll()
            event.accept()
        elif event.matches(QKeySequence.StandardKey.Save):
            current_row = self.table.currentRow()
            if current_row >= 0:
                self.save_row(current_row)
            event.accept()
        else:
            super().keyPressEvent(event)

    def show_table_context_menu(self, pos):
        row = self.table.rowAt(pos.y())
        if row < 0:
            return

        menu = QMenu()
        save_action = QAction("💾 Save Row", self)
        delete_action = QAction("🗑️ Delete Row", self)
        save_action.triggered.connect(lambda: self.save_row(row))
        delete_action.triggered.connect(lambda: self.delete_row_from_context(row))
        menu.addAction(save_action)
        menu.addAction(delete_action)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def delete_row_from_context(self, row):
        self.table.clearSelection()
        self.table.selectRow(row)
        self.delete_selected_documents()

    def closeEvent(self, event):
        if self.unsaved_changes:
            reply = QMessageBox.question(
                self, 'Unsaved Changes',
                f"You have {len(self.unsaved_changes)} unsaved change(s). Quit anyway?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

# Main Execution
def main():
    global LOGIN_TOKEN

    app = QApplication(sys.argv)
    logging.info("MatterID - Manager v2.5 application starting…")

    config_manager = ConfigManager()

    # Show MatterID splash screen
    show_matterid_splash_screen(app)

    download_splash = DownloadSplashScreen(config_manager)
    if download_splash.exec() != QDialog.DialogCode.Accepted:
        logging.warning("Key download failed. Continuing in demo mode.")

    # Skip login for demo mode, otherwise show login
    if db is not None:
        login_dialog = WebLoginDialog()
        result = login_dialog.exec()
        if result != QDialog.DialogCode.Accepted or LOGIN_TOKEN is None:
            logging.warning("Login failed or canceled. Continuing in demo mode.")
            LOGIN_TOKEN = "demo_token"

        # Verify token if we have one and database connection
        if LOGIN_TOKEN != "demo_token":
            max_retries = 1
            decoded = None

            while max_retries >= 0:
                try:
                    # Use the auth app for token verification
                    decoded = auth.verify_id_token(
                        LOGIN_TOKEN,
                        app=firebase_admin.get_app('auth')
                    )
                    break

                except ExpiredIdTokenError:
                    if max_retries > 0:
                        QMessageBox.information(
                            None, "Session Expired",
                            "Your login token has expired. Click OK to log in again."
                        )
                        max_retries -= 1
                        LOGIN_TOKEN = None
                        login_dialog = WebLoginDialog()
                        ret = login_dialog.exec()
                        if ret != QDialog.DialogCode.Accepted or LOGIN_TOKEN is None:
                            logging.warning("Re‐login failed or was cancelled. Continuing in demo mode.")
                            LOGIN_TOKEN = "demo_token"
                            break
                        continue
                    else:
                        QMessageBox.warning(None, "Auth Error", "Your login has expired. Continuing in demo mode.")
                        LOGIN_TOKEN = "demo_token"
                        break

                except InvalidIdTokenError as invalid_err:
                    logging.error(f"Invalid ID token: {invalid_err}")
                    QMessageBox.warning(None, "Auth Error", "Invalid ID token detected. Continuing in demo mode.")
                    LOGIN_TOKEN = "demo_token"
                    break

                except Exception as e:
                    logging.error(f"Token verification failed: {e}")
                    QMessageBox.warning(None, "Auth Error", "An error occurred verifying your token. Continuing in demo mode.")
                    LOGIN_TOKEN = "demo_token"
                    break

            if LOGIN_TOKEN != "demo_token" and decoded:
                uid = decoded.get("uid")
                if uid:
                    logging.info(f"User {uid} successfully authenticated.")
                else:
                    logging.warning("Could not extract UID from token. Continuing in demo mode.")

    try:
        window = MainWindow(config_manager)

        # MatterID Color Scheme Stylesheet
        stylesheet = f"""
          QMainWindow, QWidget {{ 
              background-color: {MATTERID_COLORS['background']}; 
              color: {MATTERID_COLORS['text_primary']}; 
          }}
          
          QPushButton {{
            background-color: {MATTERID_COLORS['primary']}; 
            color: white; 
            border: none;
            padding: 8px 16px; 
            border-radius: 6px; 
            min-width: 80px;
            font-weight: bold;
            font-size: 12px;
          }}
          QPushButton:hover {{ 
              background-color: {MATTERID_COLORS['hover']}; 
          }}
          QPushButton:pressed {{ 
              background-color: #2c4a9e; 
          }}
          QPushButton:disabled {{ 
              background-color: {MATTERID_COLORS['border']}; 
              color: {MATTERID_COLORS['text_muted']}; 
          }}

          QPushButton#delete_button {{ 
              background-color: {MATTERID_COLORS['error']}; 
          }}
          QPushButton#delete_button:hover {{ 
              background-color: #d32f2f; 
          }}
          QPushButton#delete_button:disabled {{ 
              background-color: #8f4048; 
              color: {MATTERID_COLORS['text_muted']}; 
          }}

          QLineEdit, QComboBox {{
            background-color: {MATTERID_COLORS['card_bg']}; 
            border: 2px solid {MATTERID_COLORS['border']};
            border-radius: 6px; 
            padding: 6px; 
            color: {MATTERID_COLORS['text_primary']}; 
            min-height: 20px;
            font-size: 12px;
          }}
          QLineEdit:focus, QComboBox:focus {{
            border-color: {MATTERID_COLORS['primary']};
          }}
          QComboBox::drop-down {{ 
              border: none; 
              background-color: {MATTERID_COLORS['primary']};
              width: 20px;
              border-radius: 3px;
          }}
          QComboBox::down-arrow {{
              width: 12px;
              height: 12px;
          }}

          QTableWidget {{
            background-color: {MATTERID_COLORS['card_bg']}; 
            alternate-background-color: #464646;
            color: {MATTERID_COLORS['text_primary']}; 
            gridline-color: {MATTERID_COLORS['border']}; 
            border: 2px solid {MATTERID_COLORS['primary']};
            selection-background-color: {MATTERID_COLORS['primary']}; 
            selection-color: white;
            border-radius: 6px;
          }}
          QHeaderView::section {{
            background-color: {MATTERID_COLORS['primary']}; 
            color: white; 
            padding: 8px;
            border: 1px solid {MATTERID_COLORS['background']}; 
            font-weight: bold;
            font-size: 12px;
          }}

          QLabel {{ 
              color: {MATTERID_COLORS['text_secondary']}; 
          }}
          QStatusBar {{ 
              color: {MATTERID_COLORS['text_secondary']}; 
              background-color: {MATTERID_COLORS['card_bg']};
              border-top: 1px solid {MATTERID_COLORS['border']};
          }}

          QTabWidget::pane {{ 
              border: 2px solid {MATTERID_COLORS['primary']}; 
              border-radius: 6px;
              background-color: {MATTERID_COLORS['background']};
          }}
          QTabBar::tab {{
            background-color: {MATTERID_COLORS['card_bg']}; 
            color: {MATTERID_COLORS['text_primary']}; 
            padding: 10px 20px;
            border: 2px solid {MATTERID_COLORS['border']}; 
            margin-right: 2px;
            border-radius: 6px 6px 0px 0px;
            font-weight: bold;
            font-size: 12px;
          }}
          QTabBar::tab:selected {{ 
              background-color: {MATTERID_COLORS['primary']}; 
              color: white;
              border-color: {MATTERID_COLORS['primary']};
          }}
          QTabBar::tab:hover {{ 
              background-color: {MATTERID_COLORS['hover']}; 
              color: white;
          }}

          QGroupBox {{
            border: 2px solid {MATTERID_COLORS['primary']}; 
            border-radius: 8px;
            margin-top: 1ex; 
            padding-top: 15px;
            background-color: {MATTERID_COLORS['card_bg']};
            font-weight: bold;
          }}
          QGroupBox::title {{
            subcontrol-origin: margin; 
            left: 15px; 
            padding: 0 8px 0 8px;
            color: {MATTERID_COLORS['primary']};
            font-size: 14px;
            font-weight: bold;
          }}

          QListWidget {{
            background-color: {MATTERID_COLORS['card_bg']}; 
            border: 2px solid {MATTERID_COLORS['border']};
            color: {MATTERID_COLORS['text_primary']}; 
            selection-background-color: {MATTERID_COLORS['primary']};
            border-radius: 6px;
          }}
          QListWidget::item {{
            padding: 8px;
            border-bottom: 1px solid {MATTERID_COLORS['border']};
          }}
          QListWidget::item:selected {{
            background-color: {MATTERID_COLORS['primary']};
            color: white;
          }}
          QListWidget::item:hover {{
            background-color: {MATTERID_COLORS['hover']};
            color: white;
          }}

          QScrollArea {{ 
              border: 2px solid {MATTERID_COLORS['border']}; 
              border-radius: 6px;
              background-color: {MATTERID_COLORS['background']};
          }}
          QScrollBar:vertical {{
            border: 1px solid {MATTERID_COLORS['border']}; 
            background: {MATTERID_COLORS['card_bg']}; 
            width: 15px;
            border-radius: 6px;
          }}
          QScrollBar::handle:vertical {{
            background: {MATTERID_COLORS['primary']}; 
            min-height: 20px; 
            border-radius: 6px;
          }}
          QScrollBar::handle:vertical:hover {{
            background: {MATTERID_COLORS['hover']};
          }}
          QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
          }}
          
          QTextEdit {{
            background-color: {MATTERID_COLORS['card_bg']};
            color: {MATTERID_COLORS['text_primary']};
            border: 2px solid {MATTERID_COLORS['border']};
            border-radius: 6px;
            padding: 8px;
          }}
          
          QProgressBar {{
            border: 2px solid {MATTERID_COLORS['border']};
            border-radius: 6px;
            text-align: center;
            background-color: {MATTERID_COLORS['card_bg']};
            color: {MATTERID_COLORS['text_primary']};
            font-weight: bold;
          }}
          QProgressBar::chunk {{
            background-color: {MATTERID_COLORS['primary']};
            border-radius: 4px;
          }}
          
          QMessageBox {{
            background-color: {MATTERID_COLORS['background']};
            color: {MATTERID_COLORS['text_primary']};
          }}
          QMessageBox QPushButton {{
            min-width: 80px;
            padding: 8px 16px;
          }}
        """
        app.setStyleSheet(stylesheet)

        window.show()
        exit_code = app.exec()
        logging.info(f"MatterID - Manager v2.5 application exiting with code {exit_code}.")
        sys.exit(exit_code)

    except Exception as e:
        logging.critical(f"Fatal error: {e}\n{traceback.format_exc()}")
        QMessageBox.critical(None, "Fatal Error", f"A critical error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()