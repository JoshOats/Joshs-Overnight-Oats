import hashlib
import json
import os
from cryptography.fernet import Fernet
import datetime
import logging

class CustomerSetup:
    def __init__(self):
        self.data_dir = 'data'
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        # Initialize encryption
        self.key = self._load_or_create_key()
        self.cipher = Fernet(self.key)
        self.users_file = os.path.join(self.data_dir, "users_data.encrypted")
        self.load_users()

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
                print(f"Loaded existing users: {list(self.users.keys())}")
            except Exception as e:
                print(f"Error loading users: {e}")
                self.users = {}
        else:
            print("No existing users file found, starting fresh")
            self.users = {}

    def save_users(self):
        encrypted_data = self.cipher.encrypt(json.dumps(self.users).encode())
        with open(self.users_file, 'wb') as f:
            f.write(encrypted_data)
        print(f"Saved users to file: {list(self.users.keys())}")

    def create_user(self):
        print("\n=== New Customer Setup ===")
        username = input("Enter customer username: ")
        password = input("Enter customer password: ")
        company_name = input("Enter company name: ")
        
        print("\nAvailable functions:")
        print("1. cnb_transfer")
        print("2. toast_reconcile")
        print("3. doordash_payout")
        print("4. olo_payout")
        print("5. future_projects")
        
        function_input = input("\nEnter function numbers (separated by spaces, e.g., '1 2 3'): ")
        selected_functions = []
        
        function_map = {
            "1": "cnb_transfer",
            "2": "toast_reconcile",
            "3": "doordash_payout",
            "4": "olo_payout",
            "5": "future_projects"
        }
        
        for num in function_input.split():
            if num in function_map:
                selected_functions.append(function_map[num])
        
        # Create the subscription data for GitHub
        subscription_data = {
            username: {
                "expiry_date": (datetime.datetime.now() + 
                               datetime.timedelta(days=365)).isoformat(),
                "available_functions": selected_functions,
                "company_name": company_name
            }
        }
        
        # Create the user credentials
        hashed_password = hashlib.sha256(password.encode()).hexdigest()
        
        # Add user to local data
        self.users[username] = {
            'password': hashed_password,
            'available_functions': selected_functions,
            'subscription_expiry': (datetime.datetime.now() + 
                                  datetime.timedelta(days=365)).isoformat(),
            'last_verified': datetime.datetime.now().isoformat()
        }
        
        # Save the updated user data
        self.save_users()
        
        print("\n=== Customer Information ===")
        print(f"Username: {username}")
        print(f"Password: {password}")
        print(f"Company: {company_name}")
        print(f"Selected functions: {selected_functions}")
        
        print("\n=== Add this to your GitHub subscriptions.json ===")
        print(json.dumps(subscription_data, indent=4))
        
        print("\n=== User has been added locally ===")
        print(f"Current users in system: {list(self.users.keys())}")
        
        return username, password, hashed_password, subscription_data

if __name__ == "__main__":
    setup = CustomerSetup()
    setup.create_user()