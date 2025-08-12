# Import necessary libraries
import os
import base64
import getpass
import pyodbc

from cryptography.fernet                       import Fernet
from cryptography.hazmat.primitives            import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

# -----------------------------------------------------------------------------
# This file contains various functions/procedures to handle credentials.
# These will be stored in locally, if the files are removed, the users 
# will be prompted to re-enter their credentials.
# -----------------------------------------------------------------------------
     
# -----------------------------------------------------------------------------
# Function: get_encryption_key, save secure string of encryption key to local file
# -----------------------------------------------------------------------------
def get_encryption_key():
    """    Retrieves the encryption key from a secure location. 
    This key is used to encrypt and decrypt sensitive data.
    The key is stored in a file named 'encryption_key.txt' in the current directory.
    If the file exists, it will read the key from the file.
    If the file does not exist, it will prompt the user to enter a new key.
    Returns:
    str: The encryption key.
    """   
    
    # set encryption key fill
    fp_secure_files = get_current_file_folder() + "\\secure_files"
    fp_key_file     = fp_secure_files + "\\encryption_key.txt"
    
    # ensure the secure_files folder exists otherwise create it.
    if not(os.path.exists(fp_secure_files)):
        os.makedirs(fp_secure_files)

    # Check if encryption key file exists
    if os.path.exists(fp_key_file):
        # Read the encrypted key from file
        try:
            with open(fp_key_file, 'rb') as f:
                encrypted_data = f.read()
            
            # For this implementation, we'll assume the key is base64 encoded
            # In a real scenario, you might want additional encryption
            tx_encryption_key = base64.b64decode(encrypted_data).decode('utf-8')
            return tx_encryption_key
            
        except Exception as e:
            print(f"Error reading encryption key file: {e}")
            print("Please re-enter your encryption key.")
    
    # If file doesn't exist or couldn't be read, prompt user for new key
    print("Encryption key file not found or corrupted.")
    tx_encryption_key = getpass.getpass("Please enter your encryption key: ")
    
    # Store the key in a secure format (base64 encoded for basic obfuscation)
    try:
        encoded_key = base64.b64encode(tx_encryption_key.encode('utf-8'))
        with open(fp_key_file, 'wb') as f:
            f.write(encoded_key)
        print(f"Encryption key saved to {fp_key_file}")

    except Exception as e:
        print(f"Error saving encryption key: {e}")
    
    tx_encryption_key = base64.b64decode(tx_encryption_key).decode('utf-8')
    return tx_encryption_key

# -----------------------------------------------------------------------------
# Function: get_fernet_key
# -----------------------------------------------------------------------------
def get_fernet_key():
    """
    Creates a Fernet encryption key from a password using PBKDF2.
    
    Parameters:
    password (str): The password to derive the key from.
    
    Returns:
    bytes: A Fernet-compatible encryption key.
    """
    # Use a fixed salt for consistency (in production, store this securely)
    
    # In production, use a random salt and store it securely
    cd_password = get_encryption_key()  # Get the encryption key as password
    bt_salt     = cd_password.encode()  # Convert to bytes for PBKDF2
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=bt_salt,
        iterations=100000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(cd_password.encode()))
    return key

# -----------------------------------------------------------------------------
# Function: encrypt_data
# -----------------------------------------------------------------------------
def encrypt_data(ip_tx_value):
    
    # Create Fernet key from the encryption key
    fernet_key = get_fernet_key()
    fernet = Fernet(fernet_key)
    
    # Return Encrypt the data
    return fernet.encrypt(ip_tx_value.encode()).decode('utf-8')

# -----------------------------------------------------------------------------
# Function: decrypt_data
# -----------------------------------------------------------------------------
def decrypt_data(ip_tx_encrypted):

    # Create Fernet key from the encryption key
    fernet_key = get_fernet_key()
    fernet = Fernet(fernet_key)
    
    # Decrypt the data
    return fernet.decrypt(ip_tx_encrypted.encode()).decode()

# -----------------------------------------------------------------------------
# Function: get_secure_information
# -----------------------------------------------------------------------------
def get_secure_information(ip_cd_information_type, ip_nm_database):
    
    # filepath to the encrypted information
    fp_encrypted = os.path.join(get_current_file_folder(), f"secure_files\\{ip_nm_database}_{ip_cd_information_type}.txt")

    # Check if encrypted file exists
    if os.path.exists(fp_encrypted):
        # Read the encrypted data from file
        try:
            with open(fp_encrypted, 'rb') as f:
                tx_encrypted = f.read()
            
            # Decode the base64 encoded data
            tx_encrypted = base64.b64decode(tx_encrypted).decode('utf-8')

            # Decrypt the data using the encryption key
            return tx_encrypted
            
        except Exception as e:
            tx_error = f"Error reading or decrypting file: {e}"
    
    # Prompt user for new information if file doesn't exist or couldn't be read
    if (ip_cd_information_type == "password"):
        # If the information type is password, prompt for secure input
        tx_value = getpass.getpass(f"Please enter your {ip_cd_information_type}: ")
    
    else:
        # For other types of information, prompt normally
        tx_value = input(f"Please enter your {ip_cd_information_type}: ")
      
    
    # Encrypt and store the information
    try:
        # Encrypt the value using the encryption function
        tx_encrypted = encrypt_data(tx_value)

        # Encode the encrypted information to base64 for storage
        tx_encrypted = base64.b64encode(tx_encrypted.encode('utf-8'))

        # Save the encrypted information to the file
        with open(fp_encrypted, 'wb') as f:
            f.write(tx_encrypted)

    except Exception as e:
        print(f"Error saving secure information: {e}")
    
    # Return the encrypted information
    # Note: This is a simplified example; in production, you might want to handle the encryption key more securely.
    tx_encrypted = base64.b64decode(tx_encrypted).decode('utf-8')
    return tx_encrypted

# -----------------------------------------------------------------------------
# Function: get_current_file_folder
# -----------------------------------------------------------------------------
def get_current_file_folder():
    return os.path.dirname(os.path.abspath(__file__))

# -----------------------------------------------------------------------------
# Function: credentials for the "Secret"-database
# -----------------------------------------------------------------------------
def credentials_secret(ip_is_override=None):
    
    # Return the credentials as a dictionary
    return credentials("secrets", ip_is_override)  # Use the function to get credentials for the secrets database

# -----------------------------------------------------------------------------
# Function: credentials for the "Secret"-database
# -----------------------------------------------------------------------------
def credentials(ip_nm_database, ip_is_override=None):
    
    # if override is provided, delete all secure-files related to secrets
    if ip_is_override != None:
        # Logic to delete secure files related to secrets
        secure_files = [f"server_for_{ip_nm_database}_db.txt", f"database_for_{ip_nm_database}_db.txt", f"username_for_{ip_nm_database}_db.txt", f"password_for_{ip_nm_database}_db.txt"]
        for file in secure_files:
            file_path = os.path.join(get_current_file_folder(), file)
            if os.path.exists(file_path):
                os.remove(file_path)    
    
    # Varables of the credentials
    nm_server   = get_secure_information("server",   ip_nm_database)
    nm_database = get_secure_information("database", ip_nm_database)
    nm_username = get_secure_information("username", ip_nm_database)
    nm_password = get_secure_information("password", ip_nm_database)

    # Return the credentials as a dictionary
    return {
        "server"   : nm_server,
        "database" : nm_database,
        "username" : nm_username,
        "password" : nm_password
    }

# -----------------------------------------------------------------------------
# Function: add_secret
# -----------------------------------------------------------------------------
def add_secret(ip_nm_secret, ip_ds_secret):
    """
    Adds a secret to the database. If the secret already exists, it will be updated.
    
    Parameters:
    ip_nm_secret (str): The name of the secret.
    ip_tx_secret (str): The value of the secret.
    """
    
    try:
        # Get decrypted credentials for the secrets database
        credentials = credentials_secret()
        
        # Credentials are now properly decrypted by get_secure_information
        server   = decrypt_data(credentials["server"])
        database = decrypt_data(credentials["database"])
        username = decrypt_data(credentials["username"])
        password = decrypt_data(credentials["password"])
        
        # Encrypt the secret value
        ip_ds_secret = encrypt_data(ip_ds_secret)

        # Create database connection string
        connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        
        # Connect to the database
        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            
            # Delete existing record
            delete_query = "DELETE FROM dbo.secrets WHERE nm_secret = ?"
            cursor.execute(delete_query, (ip_nm_secret))
            
            # Insert the new secret
            insert_query = "INSERT INTO dbo.secrets (nm_secret, ds_secret) VALUES (?, ?)"
            cursor.execute(insert_query, (ip_nm_secret, ip_ds_secret))
            
            # Commit the transaction
            conn.commit()
                        
    except pyodbc.Error as db_error:
        print(f"Database error: {db_error}")

    except Exception as e:
        print(f"Error adding secret to database: {e}")

# -----------------------------------------------------------------------------
# Function: read_secret
# -----------------------------------------------------------------------------
def read_secret(ip_nm_secret):
    """
    Reads a secret from the database and returns it decrypted.
    
    Parameters:
    ip_nm_secret (str): The name of the secret to retrieve.
    
    Returns:
    str: The decrypted secret value, or None if not found.
    """
    
    try:
        # Get decrypted credentials for the secrets database
        credentials = credentials_secret()
        
        # Credentials are now properly decrypted by get_secure_information
        server   = decrypt_data(credentials["server"])
        database = decrypt_data(credentials["database"])
        username = decrypt_data(credentials["username"])
        password = decrypt_data(credentials["password"])

        # Create database connection string
        connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        
        # Connect to the database
        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            
            # Query for the secret
            select_query = "SELECT ds_secret FROM dbo.secrets WHERE nm_secret = ?"
            cursor.execute(select_query, (ip_nm_secret,))
            
            result = cursor.fetchone()
            
            if result:
                # Decrypt and return the secret value
                encrypted_secret = result[0]
                decrypted_secret = decrypt_data(encrypted_secret)
                return decrypted_secret
            
            else:
                print(f"Secret '{ip_nm_secret}' not found in database.")
                return None
            
    except pyodbc.Error as db_error:
        print(f"Database error: {db_error}")
        return None

    except Exception as e:
        print(f"Error reading secret from database: {e}")
        return None

# -----------------------------------------------------------------------------
# Function: add_secret
# -----------------------------------------------------------------------------
def del_secret(ip_nm_secret):
   
    try:
        # Get decrypted credentials for the secrets database
        credentials = credentials_secret()
        
        # Credentials are now properly decrypted by get_secure_information
        server   = decrypt_data(credentials["server"])
        database = decrypt_data(credentials["database"])
        username = decrypt_data(credentials["username"])
        password = decrypt_data(credentials["password"])
        
        # Create database connection string
        connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        
        # Connect to the database
        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            
            # Delete existing record
            delete_query = "DELETE FROM dbo.secrets WHERE nm_secret = ?"
            cursor.execute(delete_query, (ip_nm_secret,))
                        
            # Commit the transaction
            conn.commit()
            
    except pyodbc.Error as db_error:
        print(f"Database error: {db_error}")

    except Exception as e:
        print(f"Error adding secret to database: {e}")
 