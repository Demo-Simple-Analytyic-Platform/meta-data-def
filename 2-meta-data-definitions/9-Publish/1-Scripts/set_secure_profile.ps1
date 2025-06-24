# This is to store PASSWORD to the database which is used in the publish profile.
$nm_model = "<id_model>"
$nm_windows_user = [System.Environment]::UserName
$fp_secure_folder   = "c:\users\$nm_windows_user\secure\$nm_model"
$fp_secure_profile = "$fp_secure_folder\secure-Profile.txt"
$secureProfile = Read-Host "Enter your Profile for development database access" -AsSecureString
$secureProfile | ConvertFrom-SecureString | Set-Content "$fp_secure_profile"