# This is to store username to the database which is used in the publish profile.
$nm_model = "<id_model>"
$nm_windows_user = [System.Environment]::UserName
$fp_secure_folder   = "c:\users\$nm_windows_user\secure\$nm_model"
$fp_secure_username = "$fp_secure_folder\secure-username.txt"
$secureUsername = Read-Host "Enter your username for development database access" -AsSecureString
$secureUsername | ConvertFrom-SecureString | Set-Content "$fp_secure_username"