import os
import shutil
from cryptography.fernet import Fernet
import json
import hashlib
import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)

class SubscriptionManager:
    def __init__(self):
        self.data_dir = self._get_data_dir()
        self.key = self._load_or_create_key()
        self.cipher = Fernet(self.key)
        self.users_file = os.path.join(self.data_dir, "users_data.encrypted")
        self.load_users()
    
    def _get_data_dir(self):
        """Get or create the data directory"""
        # Create data directory in the current directory when setting up
        data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
        
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        print(f"Using data directory: {data_dir}")
        return data_dir
    
    def _load_or_create_key(self):
        key_file = os.path.join(self.data_dir, "encryption.key")
        try:
            if os.path.exists(key_file):
                with open(key_file, 'rb') as f:
                    return f.read()
            else:
                key = Fernet.generate_key()
                with open(key_file, 'wb') as f:
                    f.write(key)
                return key
        except Exception as e:
            print(f"Error with encryption key: {e}")
            key = Fernet.generate_key()
            with open(key_file, 'wb') as f:
                f.write(key)
            return key
    
    def load_users(self):
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'rb') as f:
                    encrypted_data = f.read()
                    decrypted_data = self.cipher.decrypt(encrypted_data)
                    self.users = json.loads(decrypted_data)
            except Exception as e:
                print(f"Error loading users: {e}")
                self.users = {}
        else:
            self.users = {}
    
    def save_users(self):
        encrypted_data = self.cipher.encrypt(json.dumps(self.users).encode())
        with open(self.users_file, 'wb') as f:
            f.write(encrypted_data)

    def add_local_user(self, username, password, available_functions):
        """Add a new user locally (for testing or initial setup)"""
        hashed_password = hashlib.sha256(password.encode()).hexdigest()
        self.users[username] = {
            'password': hashed_password,
            'available_functions': available_functions,
            'subscription_expiry': (datetime.datetime.now() + 
                                  datetime.timedelta(days=30)).isoformat(),
            'last_verified': datetime.datetime.now().isoformat()
        }
        self.save_users()
        print(f"Added user: {username}")

def setup_initial_users():
    # First, ensure the data directory is clean
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    if os.path.exists(data_dir):
        shutil.rmtree(data_dir)
    os.makedirs(data_dir)
    
    manager = SubscriptionManager()
    
    print("Setting up initial users...")
    
    # Add Josh's account with all functions
    manager.add_local_user(
        username="josh",
        password="carrot123",
        available_functions=[
            "cnb_transfer",
            "toast_reconcile",
            "doordash_payout",
            "olo_payout",
            "future_projects"
        ]
    )
    
    # Add Company A
    manager.add_local_user(
        username="companyA",
        password="companyA_password",
        available_functions=[
            "cnb_transfer", 
            "toast_reconcile"
        ]
    )
    
    # Add Company B
    manager.add_local_user(
        username="companyB",
        password="companyB_password",
        available_functions=[
            "doordash_payout", 
            "olo_payout"
        ]
    )
    
    print("\nVerifying users were created:")
    print(f"Users in file: {list(manager.users.keys())}")
    print(f"\nFiles created in: {manager.data_dir}")
    print("\nSetup complete!")

if __name__ == "__main__":
    setup_initial_users()