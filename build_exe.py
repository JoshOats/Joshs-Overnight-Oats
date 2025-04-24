"""
Build script for Joshs Overnight Oats application.
Handles PyInstaller compilation and distribution packaging.
"""

import os
import sys
import shutil
import platform
import zipfile
from typing import List
from pathlib import Path
import PyInstaller.__main__

# Constants
APP_NAME = "Joshs_Overnight_Oats"
HIDDEN_IMPORTS = [
    'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.sip',
    'requests', 'auto_updater', 'future_projects', 'retro_style'
]

class BuildManager:
    def __init__(self):
        self.script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
        self.dist_dir = self.script_dir / 'dist'
        self.build_dir = self.script_dir / 'build'
        self.separator = ";" if platform.system() == "Windows" else ":"
        self.icon_path = self.script_dir / 'assets' / 'icon.ico'

    def get_pyinstaller_args(self) -> List[str]:
        """Generate PyInstaller arguments for main application."""
        python_version = f"{sys.version_info.major}{sys.version_info.minor}"

        args = [
            'main.py',
            '--onedir',
            '--windowed',
            '--clean',
            '--log-level', 'DEBUG',
            '--name', APP_NAME,
            '--target-arch=x64',
            '--noconfirm',
            f'--icon={self.icon_path}',
            f'--distpath={self.dist_dir / "temp_build"}',
            f'--workpath={self.build_dir}',
        ]

        # Add hidden imports
        for import_name in HIDDEN_IMPORTS:
            args.extend(['--hidden-import', import_name])

        # Add data files
        data_files = [
            ('assets/icon.png', 'assets'),
            ('assets', 'assets'),
            ('auto_updater.py', '.'),
        ]

        for src, dest in data_files:
            args.extend(['--add-data', f'{src}{self.separator}{dest}'])

        return args

    def create_launcher_script(self):
        """Create the launcher script."""
        launcher_script = '''
import os
import sys
import shutil
import subprocess
import time
from pathlib import Path

APP_NAME = "Joshs_Overnight_Oats"

class SingleInstanceChecker:
    def __init__(self, app_name):
        if os.name == 'nt':  # Windows
            self.lock_file = Path(os.environ['LOCALAPPDATA']) / f".{app_name.lower()}.lock"
        else:  # Linux/Mac
            self.lock_file = Path.home() / f".{app_name.lower()}.lock"

    def is_running(self):
        try:
            if self.lock_file.exists():
                with open(self.lock_file, 'r') as f:
                    pid = int(f.read().strip())
                    try:
                        os.kill(pid, 0)
                        return True
                    except OSError:
                        self.lock_file.unlink(missing_ok=True)
            return False
        except:
            self.lock_file.unlink(missing_ok=True)
            return False

    def create_lock(self):
        try:
            with open(self.lock_file, 'w') as f:
                f.write(str(os.getpid()))
        except:
            pass

    def release_lock(self):
        try:
            self.lock_file.unlink(missing_ok=True)
        except:
            pass

def get_app_directory():
    """Get or create application directory in AppData"""
    if os.name == 'nt':  # Windows
        app_dir = Path(os.environ['LOCALAPPDATA']) / APP_NAME
    else:  # Linux/Mac
        app_dir = Path.home() / '.local' / 'share' / APP_NAME

    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir

def install_or_update_app():
    """Install or update application files in AppData"""
    try:
        launcher_dir = Path(os.path.dirname(os.path.abspath(sys.executable)))
        app_dir = get_app_directory()
        bundled_app = launcher_dir / 'app'
        backup_dir = app_dir / 'backup'

        # Check for partial installation
        if app_dir.exists():
            exe_path = app_dir / f"{APP_NAME}.exe"
            internal_path = app_dir / "_internal"

            # If partial installation and backup exists, restore from backup
            if (not exe_path.exists() or not internal_path.exists()) and backup_dir.exists():
                print("Restoring previous version...")

                # Clean current partial installation
                for item in os.listdir(app_dir):
                    if item != 'backup':  # Keep backup folder
                        item_path = app_dir / item
                        if item_path.is_file():
                            item_path.unlink()
                        elif item_path.is_dir():
                            shutil.rmtree(item_path)

                # Restore from backup
                for item in os.listdir(backup_dir):
                    src = backup_dir / item
                    dst = app_dir / item
                    if src.is_file():
                        shutil.copy2(src, dst)
                    elif src.is_dir():
                        shutil.copytree(src, dst)

                # Clean up backup after successful restore
                shutil.rmtree(backup_dir)
                return True

        # Proceed with fresh installation if bundled app exists
        if bundled_app.exists():
            if app_dir.exists():
                shutil.rmtree(app_dir)
            shutil.copytree(bundled_app, app_dir)
            shutil.rmtree(bundled_app)

        return True
    except Exception as e:
        print(f"Installation error: {str(e)}")
        return False

def main():
    try:
        print("Starting application...")
        time.sleep(1)  # Show startup message briefly

        # Check for single instance
        instance_checker = SingleInstanceChecker(APP_NAME)
        if instance_checker.is_running():
            print("Multiple instances detected... Try opening the app only once")
            time.sleep(2)
            return

        instance_checker.create_lock()

        app_dir = get_app_directory()
        if not install_or_update_app():
            return

        app_exe = app_dir / f"{APP_NAME}.exe"
        if app_exe.exists():
            subprocess.Popen([str(app_exe)])
        else:
            print("Error: Application not found.")
            time.sleep(2)

    except Exception as e:
        print("Multiple instances detected... Try opening the app only once")
        time.sleep(2)

if __name__ == '__main__':
    main()
'''
        return launcher_script

    def build(self):
        """Execute the complete build process."""
        try:
            print("Starting build process...")
            self.clean_directories()

            # Build main application
            print("Compiling main application...")
            PyInstaller.__main__.run(self.get_pyinstaller_args())

            # Create and build launcher
            print("Building launcher...")
            launcher_script_path = self.script_dir / 'launcher_temp.py'
            with open(launcher_script_path, 'w') as f:
                f.write(self.create_launcher_script())

            PyInstaller.__main__.run([
                str(launcher_script_path),
                '--onefile',
                '--console',
                '--clean',
                '--name', APP_NAME,
                f'--icon={self.icon_path}',
                f'--distpath={self.dist_dir / APP_NAME}',
                '--noconfirm'
            ])

            # Clean up launcher script
            launcher_script_path.unlink()

            # Move main application files
            app_dir = self.dist_dir / APP_NAME / 'app'
            app_dir.mkdir(parents=True, exist_ok=True)

            temp_build_dir = self.dist_dir / 'temp_build' / APP_NAME
            if temp_build_dir.exists():
                for item in temp_build_dir.iterdir():
                    shutil.move(str(item), str(app_dir / item.name))
                shutil.rmtree(self.dist_dir / 'temp_build')

            self.create_distribution_zip()
            print("\nBuild completed successfully!")

        except Exception as e:
            print(f"\nBuild failed: {e}")
            sys.exit(1)

    def clean_directories(self):
        """Clean build and dist directories."""
        print("Cleaning previous build files...")
        for directory in [self.dist_dir, self.build_dir]:
            if directory.exists():
                shutil.rmtree(directory)
                print(f"Cleaned {directory}")

    def create_distribution_zip(self):
        """Create ZIP file from the dist directory."""
        print("\nCreating distribution package...")
        zip_path = self.script_dir / f'{APP_NAME}.zip'
        app_dist_dir = self.dist_dir / APP_NAME

        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(app_dist_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(app_dist_dir)
                        zipf.write(file_path, arcname)
            print(f"Successfully created {zip_path}")
        except Exception as e:
            print(f"Error creating ZIP file: {e}")
            raise

if __name__ == "__main__":
    builder = BuildManager()
    builder.build()
