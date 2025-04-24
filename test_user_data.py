import os
from cryptography.fernet import Fernet
import json
import hashlib

def test_user_data():
    print("\nTesting user data files...")
    
    # Check if files exist
    print("\nChecking files:")
    if os.path.exists('encryption.key'):
        print("✓ encryption.key exists")
    else:
        print("✗ encryption.key missing")
        
    if os.path.exists('users_data.encrypted'):
        print("✓ users_data.encrypted exists")
    else:
        print("✗ users_data.encrypted missing")
    
    # Try to read and decrypt user data
    try:
        # Load key
        with open('encryption.key', 'rb') as f:
            key = f.read()
        print("\nLoaded encryption key successfully")
        
        # Create cipher
        cipher = Fernet(key)
        
        # Load and decrypt user data
        with open('users_data.encrypted', 'rb') as f:
            encrypted_data = f.read()
        decrypted_data = cipher.decrypt(encrypted_data)
        users = json.loads(decrypted_data)
        
        print("\nUser data decrypted successfully")
        print(f"Found {len(users)} users:")
        for username, data in users.items():
            print(f"\nUsername: {username}")
            print(f"Available functions: {data['available_functions']}")
            print(f"Subscription expiry: {data['subscription_expiry']}")
            print(f"Password hash: {data['password']}")
            
            # Test josh's password
            if username == 'josh':
                test_password = 'carrot123'
                test_hash = hashlib.sha256(test_password.encode()).hexdigest()
                print(f"\nTesting 'josh' password:")
                print(f"Expected hash: {test_hash}")
                print(f"Stored hash:   {data['password']}")
                if test_hash == data['password']:
                    print("✓ Password hash matches")
                else:
                    print("✗ Password hash doesn't match")
        
    except Exception as e:
        print(f"\nError reading user data: {e}")

if __name__ == "__main__":
    test_user_data()