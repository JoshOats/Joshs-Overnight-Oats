import os
from pathlib import Path
import time

class SingleInstanceManager:
    def __init__(self, app_name):
        self.lock_file = Path(os.path.join(Path.home(), f".{app_name.lower()}.lock"))
        self.update_lock_file = Path(os.path.join(Path.home(), f".{app_name.lower()}.update.lock"))

    def is_running(self):
        if self.lock_file.exists():
            try:
                with open(self.lock_file, 'r') as f:
                    pid = int(f.read().strip())
                os.kill(pid, 0)  # Test if process is running
                return True
            except (OSError, ValueError):
                self.lock_file.unlink(missing_ok=True)

        # Check if update is in progress
        if self.update_lock_file.exists():
            try:
                with open(self.update_lock_file, 'r') as f:
                    timestamp = float(f.read().strip())
                    if time.time() - timestamp < 60:  # Lock valid for 60 seconds
                        return True
            except:
                pass
            self.update_lock_file.unlink(missing_ok=True)

        return False

    def create_lock(self):
        with open(self.lock_file, 'w') as f:
            f.write(str(os.getpid()))

    def create_update_lock(self):
        with open(self.update_lock_file, 'w') as f:
            f.write(str(time.time()))

    def release_lock(self):
        self.lock_file.unlink(missing_ok=True)
        self.update_lock_file.unlink(missing_ok=True)
