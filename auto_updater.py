"""
Auto-updater module for Joshs Overnight Oats application.
Handles version checking and automatic updates with backup safety.
"""

import os
import sys
import shutil
import subprocess
from typing import Optional, Dict, Any
import requests
from packaging import version
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QThread, pyqtSignal, QEventLoop
import time
import logging
from pathlib import Path

# Constants
OWNER = "JoshOats"
REPO = "Joshs-Overnight-Oats"
CURRENT_VERSION = "v1.2.37"
UPDATE_TIMEOUT = 5

# Setup logging
log_path = os.path.join(os.path.expanduser('~'), 'josh_oats_update.log')
logging.basicConfig(
    filename=log_path,
    level=logging.DEBUG,
    format='%(asctime)s - %(message)s'
)

class UpdaterThread(QThread):
    update_found = pyqtSignal()
    no_update_needed = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.release_data: Optional[Dict[str, Any]] = None

    def run(self):
        try:
            if self.check_for_updates():
                self.update_found.emit()
                self.perform_update()
            else:
                self.no_update_needed.emit()
        except Exception as e:
            logging.exception("Error in update thread")
            self.error_occurred.emit(str(e))
            self.no_update_needed.emit()

    def check_for_updates(self) -> bool:
        try:
            logging.debug("Checking for updates...")
            response = requests.get(
                f"https://api.github.com/repos/{OWNER}/{REPO}/releases/latest",
                timeout=UPDATE_TIMEOUT
            )
            response.raise_for_status()
            release = response.json()

            latest_version = version.parse(release['tag_name'].lstrip('v'))
            current_version = version.parse(CURRENT_VERSION.lstrip('v'))

            logging.debug(f"Current version: {current_version}, Latest version: {latest_version}")

            if latest_version > current_version and release.get('assets'):
                self.release_data = release
                return True
            return False
        except Exception as e:
            logging.error(f"Update check failed: {e}")
            return False

    def get_app_data_dir(self) -> str:
        """Get the AppData directory where the app is installed"""
        if os.name == 'nt':
            app_dir = os.path.join(os.environ['LOCALAPPDATA'], "Joshs_Overnight_Oats")
        else:
            app_dir = os.path.join(str(Path.home()), '.local', 'share', "Joshs_Overnight_Oats")
        return app_dir

    def perform_update(self):
        try:
            logging.info("Starting update process...")
            if not self.release_data:
                raise ValueError("No release data available")

            app_dir = self.get_app_data_dir()
            temp_dir = os.path.join(app_dir, 'temp_update')
            backup_dir = os.path.join(app_dir, 'backup')

            logging.debug(f"App directory: {app_dir}")
            logging.debug(f"Temp directory: {temp_dir}")
            logging.debug(f"Backup directory: {backup_dir}")

            # Create backup before starting update
            logging.info("Creating backup before update...")
            if os.path.exists(backup_dir):
                shutil.rmtree(backup_dir)
            shutil.copytree(app_dir, backup_dir, ignore=shutil.ignore_patterns('temp_update', 'backup'))

            # Create backup completion marker
            with open(os.path.join(backup_dir, 'backup_complete'), 'w') as f:
                f.write(str(time.time()))

            # Create update marker
            with open(os.path.join(app_dir, 'update_in_progress'), 'w') as f:
                f.write(str(time.time()))

            # Create temp directory
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            # Download update
            download_url = self.release_data['assets'][0]['browser_download_url']
            update_zip = os.path.join(temp_dir, 'update.zip')
            logging.debug(f"Downloading from: {download_url}")

            response = requests.get(download_url, timeout=UPDATE_TIMEOUT)
            response.raise_for_status()

            with open(update_zip, 'wb') as f:
                f.write(response.content)
            logging.debug("Download complete")

            # Create batch script
            batch_path = os.path.join(temp_dir, 'updater.bat')
            batch_content = f'''@echo off
title Installing Update - Josh's Overnight Oats
cd /d "{app_dir}"

echo Installing update...
echo Please wait...

rem Kill running process
taskkill /F /IM "Joshs_Overnight_Oats.exe" >nul 2>&1
timeout /t 2 /nobreak >nul

rem Remove current files (except temp, backup and markers)
for %%i in (*) do (
    if not "%%i"=="temp_update" if not "%%i"=="backup" if not "%%i"=="update_in_progress" (
        if exist "%%i" (
            if /I not "%%~nxi"=="Joshs_Overnight_Oats.exe" (
                del /F /Q "%%i" >nul 2>&1
            )
        )
    )
)
for /D %%i in (*) do (
    if not "%%i"=="temp_update" if not "%%i"=="backup" (
        rmdir /S /Q "%%i" >nul 2>&1
    )
)

rem Extract update
echo Extracting update...
powershell -Command "Expand-Archive -Force '{update_zip}' '{app_dir}\\temp_extract'" >nul 2>&1

rem Copy files from temp_extract/app
xcopy /S /E /H /Y "temp_extract\\app\\*" "." >nul 2>&1

rem Check if update was successful
if exist "Joshs_Overnight_Oats.exe" (
    echo Update completed successfully!
    del /F /Q "update_in_progress" >nul 2>&1
    rmdir /S /Q backup >nul 2>&1
    rmdir /S /Q temp_update >nul 2>&1
    rmdir /S /Q temp_extract >nul 2>&1
    echo Starting application...
    start "" "Joshs_Overnight_Oats.exe"
) else (
    echo Update failed - restoring previous version...
    xcopy /S /E /H /Y "backup\\*" "." >nul 2>&1
    del /F /Q "update_in_progress" >nul 2>&1
    rmdir /S /Q backup >nul 2>&1
    rmdir /S /Q temp_update >nul 2>&1
    rmdir /S /Q temp_extract >nul 2>&1
    start "" "Joshs_Overnight_Oats.exe"
)
exit'''

            with open(batch_path, 'w', encoding='utf-8') as f:
                f.write(batch_content)

            # Execute update script
            logging.info("Launching update script")
            subprocess.Popen(
                ['cmd', '/c', batch_path],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            time.sleep(1)
            logging.info("Update script launched, exiting application")
            os._exit(0)

        except Exception as e:
            # If anything fails during update prep, clean up and restore
            logging.exception("Update failed")
            marker_path = os.path.join(app_dir, 'update_in_progress')
            if os.path.exists(marker_path):
                os.remove(marker_path)
            if os.path.exists(backup_dir) and os.path.exists(os.path.join(backup_dir, 'backup_complete')):
                try:
                    # Restore from backup
                    for item in os.listdir(app_dir):
                        if item != 'backup':
                            item_path = os.path.join(app_dir, item)
                            if os.path.isfile(item_path):
                                os.remove(item_path)
                            elif os.path.isdir(item_path):
                                shutil.rmtree(item_path)

                    # Copy backup files back
                    for item in os.listdir(backup_dir):
                        src = os.path.join(backup_dir, item)
                        dst = os.path.join(app_dir, item)
                        if os.path.isfile(src):
                            shutil.copy2(src, dst)
                        elif os.path.isdir(src):
                            shutil.copytree(src, dst)
                finally:
                    # Clean up backup
                    shutil.rmtree(backup_dir)
            raise

def check_and_update():
    """Check for updates and handle the update process."""
    try:
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        # Check for interrupted update
        app_dir = UpdaterThread().get_app_data_dir()
        marker_path = os.path.join(app_dir, 'update_in_progress')
        backup_dir = os.path.join(app_dir, 'backup')

        if os.path.exists(marker_path):
            logging.warning("Detected interrupted update, attempting recovery")
            if os.path.exists(backup_dir) and os.path.exists(os.path.join(backup_dir, 'backup_complete')):
                try:
                    # Restore from backup
                    for item in os.listdir(app_dir):
                        if item != 'backup':
                            item_path = os.path.join(app_dir, item)
                            if os.path.isfile(item_path):
                                os.remove(item_path)
                            elif os.path.isdir(item_path):
                                shutil.rmtree(item_path)

                    # Copy backup files back
                    for item in os.listdir(backup_dir):
                        src = os.path.join(backup_dir, item)
                        dst = os.path.join(app_dir, item)
                        if os.path.isfile(src):
                            shutil.copy2(src, dst)
                        elif os.path.isdir(src):
                            shutil.copytree(src, dst)

                    logging.info("Recovery completed successfully")
                finally:
                    # Clean up markers and backup
                    if os.path.exists(marker_path):
                        os.remove(marker_path)
                    if os.path.exists(backup_dir):
                        shutil.rmtree(backup_dir)
                return False

        updater = UpdaterThread()
        loop = QEventLoop()

        updater.update_found.connect(lambda: None)
        updater.no_update_needed.connect(loop.quit)
        updater.error_occurred.connect(lambda msg: logging.error(f"Update error: {msg}"))

        updater.start()
        loop.exec_()

        return False
    except Exception as e:
        logging.exception("Error in check_and_update")
        return False

if __name__ == "__main__":
    check_and_update()
