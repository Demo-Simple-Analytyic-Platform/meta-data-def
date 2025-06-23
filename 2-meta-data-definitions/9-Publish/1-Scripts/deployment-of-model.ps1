# This Script is used to deploy the model to SQL Server Database.
# If requires your to setup publish profiel for your database, make sure you save in the correct folder.
# !!! Make sure you have re-named to "re-name-to-name-of-model"-project !!!
$nm_model      = "<nm_model>"                    # Replace with your model name
$nm_profile    = "<name-of-profile.publish.xml>" # Replace with your publish profile name
$ds_git_folder = "<git-folder>"                  # Replace with your git folder path
$nm_server     = "<nm_server>"                   # Replace with your server name
$nm_database   = "<nm_database>"                 # Replace with your database name   

# -----------------------------------------------------------------------------
# After setting the above variables, you can run this script to deploy the model. 
# by : powershell -ExecutionPolicy Bypass -File "C:\git\Demo-Simple-Analytyic-Platform\meta-data-model\.attachments\scripts\deployment-of-model.ps1"
# -----------------------------------------------------------------------------
echo "powershell -ExecutionPolicy Bypass -File "ds_git_folder\nm_model\.attachments\scripts\deployment-of-model.ps1"

# Determine filepath to Publich profile
$fp_publish    = "$ds_git_folder\$nm_model\2-meta-data-definitions\9-Publish\2-Publish\$nm_profile" # Replace with your publish profile name
if (-not (Test-Path $fp_publish)) { throw "File not found: $fp_publish"}


# 0. Save password if secure-password file is missing
echo "# 0. Set password for deployment."

$nm_windows_user = [System.Environment]::UserName
$fp_secure_folder   = "c:\users\$nm_windows_user\secure"
$fp_secure_password = "$fp_secure_folder\secure-password.txt"
$fp_secure_username = "$fp_secure_folder\secure-username.txt"

if (-not (Test-Path $fp_secure_password)) { 
    if (-Not (Test-Path -Path $fp_secure_folder)) { New-Item -Path $fp_secure_folder -ItemType Directory -Force }
    echo "Secure password file not found. Creating a new one."
    $securePassword = Read-Host "Enter your password" -AsSecureString
    $securePassword | ConvertFrom-SecureString | Set-Content "$fp_secure_password"
    echo "Secure password saved to $fp_secure_password"
} 

if (-not (Test-Path $fp_secure_username)) { 
    if (-Not (Test-Path -Path $fp_secure_folder)) { New-Item -Path $fp_secure_folder -ItemType Directory -Force }
    echo "Secure username file not found. Creating a new one."
    $secureUsername = Read-Host "Enter your username" -AsSecureString
    $secureUsername | ConvertFrom-SecureString | Set-Content "$fp_secure_username"
    echo "Secure Username saved to $fp_secure_username"
}

# Determine the project file path
echo "# 1. Build `Meta-Data-Model`."

# Search for SqlPackage.exe and store the first match in a variable
$msbuild = Get-ChildItem -Path "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter MSBuild.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
echo "$msbuild"

# cahnge directory
#Set-Location -Path "$msbuild"
& "$msbuild" "$ds_git_folder\$nm_model\2-meta-data-definitions\$nm_model.sqlproj" `
    /p:Configuration=Debug `
    /p:DeployOnBuild=true `
    /p:PublishProfile="$fp_publish"

echo "# 2. Publish `Meta-Data-Model` to database."

# Search for SqlPackage.exe and store the first match in a variable
$sqlPackagePath = Get-ChildItem -Path "C:\" -Filter "SqlPackage.exe" -Recurse -ErrorAction SilentlyContinue -Force |
    Where-Object { $_.FullName -match "SqlPackage.exe" } |
    Select-Object -First 1 -ExpandProperty FullName
echo "$sqlPackagePath"

# Extract Credentials from secure password file
$secureUsername = Get-Content $fp_secure_username | ConvertTo-SecureString
$securePassword = Get-Content $fp_secure_password | ConvertTo-SecureString

# Run SqlPackage.exe to publish
& "$sqlPackagePath" /Action:Publish `
    /SourceFile:"$ds_git_folder\$nm_model\2-meta-data-definitions\bin\Debug\_2_meta_data_definitions.dacpac" `
    /Profile:"$fp_publish" `
    /TargetServerName:"$nm_server" `
    /TargetDatabaseName:"$nm_database" `
    /TargetUser:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureUsername)))" `
    /TargetPassword:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)))"
      
# Check if the deployment was successful
if ($LASTEXITCODE -ne 0) {
    throw "Deployment failed with exit code $LASTEXITCODE. Please check the logs for more details."
}
echo "Change Have been deployed to with profile '$nm_profile' to the database."
