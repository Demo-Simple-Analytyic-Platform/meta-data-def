# Secrets Management Module Documentation

This module provides comprehensive functionality for managing encrypted credentials and secrets. It handles local file-based credential storage and database-based secret management with strong encryption capabilities.

## Overview

The `secrets.py` module contains functions for:
- Encryption key management
- Data encryption/decryption using Fernet symmetric encryption
- Local credential storage and retrieval
- Database connection credential management
- Secret storage and retrieval from database

## Dependencies

```python
import os
import base64
import getpass
import pyodbc
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
```

## Functions Documentation

### Core Encryption Functions

#### `get_encryption_key()`
**Purpose**: Retrieves or prompts for the master encryption key used for all cryptographic operations.

**Parameters**: None

**Returns**: `str` - The encryption key

**Functionality**:
- Checks for existing `encryption_key.txt` file in the current directory
- If file exists, reads and decodes the base64-encoded key
- If file doesn't exist or is corrupted, prompts user for new key using `getpass`
- Stores the new key in base64-encoded format for basic obfuscation
- Provides error handling for file operations

**Security Features**:
- Uses `getpass` for secure password input (hidden from terminal)
- Base64 encoding for basic key obfuscation
- Automatic key file creation and management

---

#### `get_fernet_key()`
**Purpose**: Creates a Fernet-compatible encryption key using PBKDF2 key derivation.

**Parameters**: None

**Returns**: `bytes` - A Fernet-compatible encryption key

**Functionality**:
- Uses the master encryption key as password source
- Applies PBKDF2 key derivation with SHA256 hashing
- Uses 100,000 iterations for enhanced security
- Uses the password itself as salt (note: this creates a circular dependency issue)
- Returns URL-safe base64 encoded key

**Security Features**:
- PBKDF2 with SHA256 algorithm
- 100,000 iterations for brute-force resistance
- URL-safe base64 encoding

---

#### `encrypt_data(ip_tx_value)`
**Purpose**: Encrypts plain text data using Fernet symmetric encryption.

**Parameters**:
- `ip_tx_value` (str): Plain text data to encrypt

**Returns**: `str` - Encrypted data as UTF-8 string

**Functionality**:
- Creates Fernet encryption object using derived key
- Encrypts input data
- Returns encrypted data as UTF-8 decoded string

**Security Features**:
- Fernet symmetric encryption (AES 128 in CBC mode with HMAC-SHA256)
- Automatic IV generation for each encryption

---

#### `decrypt_data(ip_tx_encrypted)`
**Purpose**: Decrypts Fernet-encrypted data back to plain text.

**Parameters**:
- `ip_tx_encrypted` (str): Encrypted data to decrypt

**Returns**: `str` - Decrypted plain text data

**Functionality**:
- Creates Fernet decryption object using derived key
- Decrypts the provided encrypted data
- Returns plain text as UTF-8 decoded string

**Security Features**:
- Automatic authentication verification (HMAC validation)
- Consistent key derivation ensures successful decryption

---

### Local File Management Functions

#### `get_secure_information(ip_cd_information_type, ip_nm_database)`
**Purpose**: Retrieves or prompts for secure information (credentials) and stores them encrypted.

**Parameters**:
- `ip_cd_information_type` (str): Type of information (server, database, username, password)
- `ip_nm_database` (str): Database identifier for file naming

**Returns**: `str` - Encrypted information (base64 encoded)

**Functionality**:
- Checks for existing encrypted file: `{database}_{type}.txt`
- If file exists, reads and decodes the encrypted data
- If file doesn't exist, prompts user for input
- Uses `getpass` for password fields, regular input for others
- Encrypts new data and stores it in base64-encoded format
- Returns encrypted data

**File Naming Convention**: `{ip_nm_database}_{ip_cd_information_type}.txt`

**Security Features**:
- Secure password input for sensitive fields
- Double encoding (Fernet encryption + base64 encoding)
- Automatic file-based credential management

---

#### `get_current_file_folder()`
**Purpose**: Returns the absolute path of the directory containing the current Python file.

**Parameters**: None

**Returns**: `str` - Absolute directory path

**Functionality**:
- Uses `__file__` to get current script location
- Returns directory path using `os.path.dirname()` and `os.path.abspath()`

**Usage**: Used by other functions to determine where to store credential files

---

### Database Credential Management Functions

#### `credentials_secret(ip_is_override=None)`
**Purpose**: Wrapper function to get credentials specifically for the secrets database.

**Parameters**:
- `ip_is_override` (optional): If provided, deletes existing credential files

**Returns**: `dict` - Dictionary containing database credentials

**Functionality**:
- Calls `credentials("secrets", ip_is_override)`
- Provides simplified interface for secrets database access

---

#### `credentials(ip_nm_database, ip_is_override=None)`
**Purpose**: Retrieves complete database connection credentials for any database.

**Parameters**:
- `ip_nm_database` (str): Database identifier
- `ip_is_override` (optional): If provided, deletes existing credential files

**Returns**: `dict` - Dictionary with keys: server, database, username, password

**Functionality**:
- If override is provided, deletes all related credential files
- Retrieves four credential components: server, database, username, password
- Uses `get_secure_information()` for each credential type
- Returns structured credential dictionary

**Override Functionality**:
- Deletes files: `server_for_{database}_db.txt`, `database_for_{database}_db.txt`, etc.
- Forces re-entry of all credentials

---

### Database Secret Management Functions

#### `add_secret(ip_nm_secret, ip_tx_secret)`
**Purpose**: Adds or updates a secret in the database with encryption.

**Parameters**:
- `ip_nm_secret` (str): Name/identifier of the secret
- `ip_tx_secret` (str): Secret value to store

**Returns**: None

**Functionality**:
- Retrieves database credentials using `credentials_secret()`
- Decrypts database connection credentials
- Encrypts the secret value before storage
- Connects to SQL Server using ODBC Driver 17
- Removes existing secret with same name (if exists)
- Inserts new encrypted secret into `dbo.secrets` table
- Commits transaction

**Database Operations**:
- `DELETE FROM dbo.secrets WHERE nm_secret = ?`
- `INSERT INTO dbo.secrets (nm_secret, ds_secret) VALUES (?, ?)`

**Security Features**:
- Parameterized queries prevent SQL injection
- Secret value is encrypted before database storage
- Automatic credential decryption

---

#### `read_secret(ip_nm_secret)`
**Purpose**: Retrieves and decrypts a secret from the database.

**Parameters**:
- `ip_nm_secret` (str): Name/identifier of the secret to retrieve

**Returns**: `str` or `None` - Decrypted secret value, or None if not found/error

**Functionality**:
- Retrieves database credentials using `credentials_secret()`
- Decrypts database connection credentials
- Connects to SQL Server using ODBC Driver 17
- Queries for the specified secret
- Decrypts the retrieved secret value
- Returns plain text secret or None if not found

**Database Operations**:
- `SELECT ds_secret FROM dbo.secrets WHERE nm_secret = ?`

**Error Handling**:
- Database connection errors
- Secret not found scenarios
- Decryption failures

---

#### `del_secret(ip_nm_secret)`
**Purpose**: Deletes a secret from the database.

**Parameters**:
- `ip_nm_secret` (str): Name/identifier of the secret to delete

**Returns**: None

**Functionality**:
- Retrieves database credentials using `credentials_secret()`
- Decrypts database connection credentials
- Connects to SQL Server using ODBC Driver 17
- Deletes the specified secret from database
- Commits transaction

**Database Operations**:
- `DELETE FROM dbo.secrets WHERE nm_secret = ?`

**Note**: Function header comment incorrectly says "add_secret" - should be corrected to "del_secret"

---

## Security Considerations

### Strengths
- **Fernet Encryption**: Industry-standard symmetric encryption with authentication
- **PBKDF2 Key Derivation**: 100,000 iterations with SHA256 provide strong key derivation
- **Parameterized Queries**: All database operations use parameterized queries
- **Secure Input**: Uses `getpass` for password input
- **Multi-layer Encoding**: Base64 encoding plus Fernet encryption

### Areas for Improvement
- **Salt Usage**: Currently uses password as salt in PBKDF2, creating circular dependency
- **Key Storage**: Master encryption key is only base64 encoded, not encrypted
- **Error Handling**: Some functions could benefit from more specific error handling
- **Key Rotation**: No mechanism for changing encryption keys
- **Database Table**: Assumes `dbo.secrets` table exists with specific schema

## Database Schema Requirements

The module expects a SQL Server database with the following table structure:

```sql
CREATE TABLE dbo.secrets (
    nm_secret NVARCHAR(255) PRIMARY KEY,
    ds_secret NVARCHAR(MAX)
);
```

## Usage Examples

```python
# Initialize and store database credentials
credentials = credentials_secret()

# Add a new secret
add_secret("api_key", "your-secret-api-key")

# Retrieve a secret
api_key = read_secret("api_key")

# Delete a secret
del_secret("old_api_key")

# Override credentials (force re-entry)
credentials = credentials_secret("override")
```

## File Structure

The module creates the following files in the current directory:
- `secure_files\encryption_key.txt` - Base64 encoded master encryption key
- `secure_files\{database}_server.txt` - Encrypted database server information
- `secure_files\{database}_database.txt` - Encrypted database name
- `secure_files\{database}_username.txt` - Encrypted database username  
- `secure_files\{database}_password.txt` - Encrypted database password

**Note**: Ensure these files are added to `.gitignore` to prevent accidental commit of sensitive data.