import json
from cryptography.fernet import Fernet

def verify_user_files():
    print("Verifying user data files...")
    
    # Read the encryption key
    with open('encryption.key', 'rb') as f:
        key = f.read()
        print(f"Encryption key loaded: {key}")
    
    # Create cipher
    cipher = Fernet(key)
    
    # Read and decrypt user data
    with open('users_data.encrypted', 'rb') as f:
        encrypted_data = f.read()
        decrypted_data = cipher.decrypt(encrypted_data)
        users = json.loads(decrypted_data)
        
    print("\nDecrypted user data:")
    for username, data in users.items():
        print(f"\nUser: {username}")
        for key, value in data.items():
            print(f"{key}: {value}")

if __name__ == "__main__":
    verify_user_files()