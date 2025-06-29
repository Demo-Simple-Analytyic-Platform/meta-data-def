# This Script is used to deploy the model to SQL Server Database.
# If requires your to setup publish profiel for your database, make sure you save in the correct folder.
# !!! Make sure you have re-named to "re-name-to-name-of-model"-project !!!
$nm_model           = "<nm_model>"                # Should be replace by ms-access-frontend tool.
$fp_model           = "<tx_git_folder>\$nm_model" # Should be replace by ms-access-frontend tool.

# ---------------------------------------------------------------------------- #
# Set these parameters to your needs.
# ---------------------------------------------------------------------------- #
$Encrypt                = "False"  # Set to "True" if you want to encrypt the connection.
$TrustServerCertificate = "False"  # Set to "True" if you want to trust the server certificate.
# ---------------------------------------------------------------------------- #

# Ensure the secure folder exists
$sfp_secure = "c:\users\$([System.Environment]::UserName)\secure"
if (-Not (Test-Path -Path $sfp_secure)) { New-Item -Path $sfp_secure -ItemType Directory -Force }

# Ensure the model folder exists
$sfp_model = "$sfp_secure\$nm_model"
if (-Not (Test-Path -Path $sfp_model)) { New-Item -Path $sfp_model -ItemType Directory -Force }

# Ensure "Server" information is stored securely.
$sfp_server = "$sfp_model\server.txt"
if (-not (Test-Path $sfp_server)) { 
    echo "Secure server file not found. Creating a new one."
    $secure_nm_server = Read-Host "Enter your server name" -AsSecureString
    $secure_nm_server | ConvertFrom-SecureString | Set-Content "$sfp_server"
    echo "Secure server saved to $sfp_server"
}

# Ensure "Database" information is stored securely.
$sfp_database = "$sfp_model\database.txt"
if (-not (Test-Path $sfp_database)) { 
    echo "Secure Database file not found. Creating a new one."
    $secure_nm_database = Read-Host "Enter your Database name" -AsSecureString
    $secure_nm_database | ConvertFrom-SecureString | Set-Content "$sfp_database"
    echo "Secure server saved to $sfp_database"
}

# Ensure "Username" information is stored securely.
$sfp_username = "$sfp_model\username.txt"
if (-not (Test-Path $sfp_username)) { 
    echo "Secure Username file not found. Creating a new one."
    $secure_nm_username = Read-Host "Enter your Username" -AsSecureString
    $secure_nm_username | ConvertFrom-SecureString | Set-Content "$sfp_username"
    echo "Secure server saved to $sfp_username"
}

# Ensure "Password" information is stored securely.
$sfp_password = "$sfp_model\password.txt"
if (-not (Test-Path $sfp_password)) { 
    echo "Secure Password file not found. Creating a new one."
    $secure_cd_password = Read-Host "Enter your Password" -AsSecureString
    $secure_cd_password | ConvertFrom-SecureString | Set-Content "$sfp_password"
    echo "Secure server saved to $sfp_password"
}

# Extract Credentials from secure files (this is still in secure format)
$nm_server   = Get-Content $sfp_server   | ConvertTo-SecureString
$nm_database = Get-Content $sfp_database | ConvertTo-SecureString
$nm_username = Get-Content $sfp_username | ConvertTo-SecureString
$cd_password = Get-Content $sfp_password | ConvertTo-SecureString

# Ensure folder path to "9-Publish"-folder exists
$fp_publish = "$fp_model\2-meta-data-definitions\9-Publish"
if (-Not (Test-Path -Path $fp_publish)) { New-Item -Path $fp_publish -ItemType Directory -Force }

# Ensure folder path to "2_deployment"-folder exists
$fp_deployment = "$fp_publish\2-Deployment"
if (-Not (Test-Path -Path $fp_deployment)) { New-Item -Path $fp_deployment -ItemType Directory -Force }

# Ensure file path to "deployment-from-ms-access.publish.xml" exists
$fp_profile = "$fp_deployment\deployment-from-ms-access.publish.xml"
$xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="4.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
    <PropertyGroup>
    <TargetDatabaseName>$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_database)))</TargetDatabaseName>
    <TargetConnectionString>Data Source=($([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_server))));Initial Catalog=$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_database)))Integrated Security=True;Encrypt=$Encrypt;TrustServerCertificate=$TrustServerCertificate;</TargetConnectionString>
    <DeployScriptFileName>$nm_model.sql</DeployScriptFileName>
    <ProfileVersionNumber>1</ProfileVersionNumber>
    </PropertyGroup>
</Project>
"@

if (-not (Test-Path $fp_profile)) { # Save the file
    $xmlContent | Out-File -FilePath $fp_profile -Encoding utf8
} else {
    Remove-Item $fp_profile -Force
    $xmlContent | Out-File -FilePath $fp_profile -Encoding utf8
}

# Determine the project file path
echo "# 1. Build `Meta-Data-Model`."

# Search for SqlPackage.exe and store the first match in a variable
$msbuild = Get-ChildItem -Path "C:\Program Files\Microsoft Visual Studio" -Recurse -Filter MSBuild.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName

# cahnge directory
#Set-Location -Path "$msbuild"
& "$msbuild" "$fp_model\2-meta-data-definitions\$nm_model.sqlproj" `
    /p:Configuration=Debug `
    /p:DeployOnBuild=true `
    /p:PublishProfile="$fp_profile"

echo "# 2. Publish `Meta-Data-Model` to database."

# Search for SqlPackage.exe and store the first match in a variable
$sqlPackagePath = Get-ChildItem -Path "C:\" -Filter "SqlPackage.exe" -Recurse -ErrorAction SilentlyContinue -Force |
    Where-Object { $_.FullName -match "SqlPackage.exe" } |
    Select-Object -First 1 -ExpandProperty FullName

# Run SqlPackage.exe to publish
& "$sqlPackagePath" /Action:Publish `
    /SourceFile:"$fp_model\2-meta-data-definitions\bin\Debug\_2_meta_data_definitions.dacpac" `
    /Profile:"$fp_profile" `
    /TargetServerName:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_server)))" `
    /TargetDatabaseName:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_database)))" `
    /TargetUser:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($nm_username)))" `
    /TargetPassword:"$([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cd_password)))"
      
# All most done, ust have to delete "waiting" file so ms access frontend tool knows that the deployment is done.
$fp_waiting = "C:\Temp\deployment_of_$nm_model.wait"

# Delete the waiting file if it exists
if ($fp_waiting -and (Test-Path $fp_waiting)) { Remove-Item $fp_waiting -Force}

# Check if the deployment was successful
if ($LASTEXITCODE -ne 0) {
    throw "Deployment failed with exit code $LASTEXITCODE. Please check the logs for more details."
}
echo "Change Have been deployed to with profile '$fp_profile' to the database."