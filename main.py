# main.py

import sys
import os
from PyQt5.QtWidgets import (QApplication, QVBoxLayout, QPushButton, QMessageBox, QLabel)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from pathlib import Path
import logging
import requests
import auto_updater
import time
from retro_style import RetroWindow, RetroDialog

# Constants
APP_NAME = "Joshs_Overnight_Oats"

def get_user_data_dir() -> Path:
    """Get user data directory"""
    home = Path.home()
    if os.name == 'nt':
        return home / "AppData" / "Local" / APP_NAME
    return home / f".{APP_NAME.lower()}"

def ensure_user_data_dir() -> Path:
    """Ensure user data directory exists"""
    data_dir = get_user_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir

def setup_logging():
    """Setup logging"""
    log_dir = ensure_user_data_dir() / "logs"
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / "app.log"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(str(log_file)),
            logging.StreamHandler(sys.stdout)
        ]
    )

def resource_path(relative_path: str) -> str:
    """Get resource path"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class LoginWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        from retro_style import create_retro_central_widget
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Create label
        self.title_label = QLabel("", self)
        self.title_label.setObjectName("login_label")
        self.title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.title_label)

        # Create the typing animation
        self.create_typing_animation(self.title_label, "Overnight Codes")

        layout.addSpacing(50)

        login_button = QPushButton('START')
        login_button.clicked.connect(self.start_app)
        layout.addWidget(login_button)

        self.setWindowTitle('Overnight Codes')
        self.setFixedSize(650, 480)
        self.center()

    def start_app(self):
        from main_window import MainWindow
        functions = [
            "ap_process",
            "cnb_transfer",
            "toast_reconcile",
            "due_to_from",
            "ap_reconcile",
            "tips_reconcile",
            "grubhub_process",
            "doordash_process",
            "ubereats_process",
            "royalties_process",
            "payroll_automation"  # Added the new function
        ]
        self.main_window = MainWindow(
            username="user",
            available_functions=functions
        )
        self.main_window.show()
        self.close()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

def main():
    from single_instance import SingleInstanceManager
    from retro_style import RetroSplash
    try:
        # Single instance check
        instance_manager = SingleInstanceManager("Joshs_Overnight_Oats")
        if instance_manager.is_running():
            logging.info("Another instance is running")
            return

        instance_manager.create_lock()

        app = QApplication(sys.argv)

        # Set icon
        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        app.setWindowIcon(QIcon(icon_path))

        # Create and show splash screen immediately
        splash = RetroSplash()
        splash.show()
        splash.showMessage("Starting application...")
        app.processEvents()

        # Setup logging
        setup_logging()

        # Check for updates
        splash.showMessage("Checking for updates...")
        app.processEvents()
        time.sleep(0.5)

        should_exit = auto_updater.check_and_update()
        if should_exit:
            splash.showMessage("Update found! Starting update process...")
            app.processEvents()
            time.sleep(0.5)
            logging.info("Exiting for update")
            instance_manager.release_lock()
            return

        # Show login window
        login_window = LoginWindow()
        splash.showMessage("Loading application...")
        app.processEvents()
        time.sleep(0.3)

        splash.finish(login_window)
        login_window.show()

        result = app.exec_()
        instance_manager.release_lock()
        sys.exit(result)

    except Exception as e:
        logging.critical(f"Fatal error: {str(e)}")
        QMessageBox.critical(None, "Error", f"An unexpected error occurred: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
