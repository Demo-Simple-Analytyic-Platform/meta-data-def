# This is to store PASSWORD to the database which is used in the publish profile.
$nm_model = "meta-def-example"
$nm_windows_user = [System.Environment]::UserName
$fp_secure_folder   = "c:\users\$nm_windows_user\secure\$nm_model"
$fp_secure_password = "$fp_secure_folder\secure-Password.txt"
$securePassword = Read-Host "Enter your Password for development database access" -AsSecureString
$securePassword | ConvertFrom-SecureString | Set-Content "$fp_secure_password"