# powershell -ExecutionPolicy Bypass -File "C:\git\Demo-Simple-Analytyic-Platform\meta-data-def\2-meta-data-definitions\1-Frontend\Development-Version\Code\PowerShell\helpers.ps1"
#
# Global variables for the script
$global:tx_error     = $null;
$global:is_debugging = $true;
#
# Debugging mode
if ($global:is_debugging) {Write-Host "Debugging Mode: ON"} else {Write-Host "Debugging Mode: OFF"}
#
# Function to remove all files from a specified folder
function Remove-All-Files-From-Folder($folderPath) {
    if (Test-Path -Path $folderPath) {
        Get-ChildItem -Path $folderPath -File | Remove-Item -Force
        if ($global:is_debugging) { Write-Host "All files removed from folder: $folderPath" }
    } else {
        if ($global:is_debugging) { Write-Host "Folder does not exist: $folderPath" }
    }
}
#
# Get the repository path
# This function returns the path of the repository by extracting it from the script root.
# It uses the $PSScriptRoot variable to find the path up to the "2-meta-data-definitions" folder.
# The path is determined by finding the index of "2-meta-data-definitions" in the script root and taking the substring up to that point.
# The function returns the full path of the repository.
function Get-Repository-Path() {
    $fp_repository = $PSScriptRoot.Substring(0, ($PSScriptRoot.IndexOf("2-meta-data-definitions")-1))
    return $fp_repository
}
#
#
# Get the repository name
# This function returns the name of the repository by extracting it from the path.
# It uses the Split-Path cmdlet to get the last part of the repository path.
function Get-Repository-Name() {
    $nm_repository = Split-Path -Path "$(Get-Repository-Path)" -Leaf
    return $nm_repository
}
#
# Get the target path for the application
# This function returns the target path for the application by combining the repository path with the specific folder structure.
# It uses the Get-Repository-Path function to get the repository path and appends the specific folder structure for the frontend application.
# The function returns the full path of the target application.
function Get-Target-Path-Application() {
    $fp_repository = "$(Get-Repository-Path)"
    $fp_target_appl = "$fp_repository\2-meta-data-definitions\1-Frontend"
    return $fp_target_appl
}
#
# Get the target path for the code
# This function returns the target path for the code by combining the target application path with the specific folder structure for the code.
# It uses the Get-Target-Path-Application function to get the target application path and appends the "Development-Version\Code" folder structure.
# The function returns the full path of the target code.
function Get-Target-Path-Code() {
    $fp_target_appl = "$(Get-Target-Path-Application)"
    $fp_target_code = "$fp_target_appl\Development-Version\Code"
    return $fp_target_code
}
#
# Get the target path for the build
# This function returns the target path for the build by combining the target application path with the specific
function Get-Target-Path-Build() {
    $fp_target_appl = "$(Get-Target-Path-Application)"
    $fp_target_build = "$fp_target_appl\ms-access-frontend.accdb"
    return $fp_target_build
}
#
# This function builds a new Access application by creating a new database file.
# It uses the NewCurrentDatabase method of the Access.Application COM object to create a new database.
# The function returns the path of the newly created database file.
function Build-New-Access-Application() {
    #
    $fp_target_build = "$(Get-Target-Path-Build)"
    if ($global:is_debugging) { Write-Host "Target path for build: $fp_target_build" }
    #
    # Path to accdb-file with "windows script host object model" (WScript) is used to create a new Access application.
    $fp_template_build = "$(Get-Target-Path-Code)" + "\build.accdb"
    # 
    # before creating a new database, ensure the target path is clean
    if (Test-Path -Path $fp_target_build) { 
        try { Remove-Item -Path $fp_target_build -Force }
        catch { $global:tx_error = "Error 1: Failed to remove existing Access application at: $fp_target_build. Error: $_" }
    }
    #
    # Copy the template build file to the target path
    Copy-Item -Path $fp_template_build -Destination $fp_target_build -Force
    #
    # Create a new Access application object
    $ob_access_appl = New-Object -ComObject Access.Application
    #
    # Create a new Access database application, if not error occurred
    if ($null -eq $global:tx_error) { 
        try { 
            #
            # Open the newly copied Access file
            $ob_access_appl.OpenCurrentDatabase($fp_target_build, $true)
            $ob_access_appl.Visible = $true
            if ($global:is_debugging) { Write-Host "Access file opened successfully: $fp_target_build" }
            #
        }
        catch { $global:tx_error = "Error 2: Failed to create new Access application at: $fp_target_build. Error: $_" }
    }
    #
    # If no error occurred, return the Access application object
    if ($null -eq $global:tx_error) { $return = $ob_access_appl } else { $return = $null }
    return $return
    #
}
#
# This function imports a module into the Access application.
# It uses the LoadFromText method of the Access.Application COM object to load the module from a file.
# The function returns an error message if the module fails to load.
function Import-Module-into-Access($ip_ob_access_appl) {
    # 
    $nm_object_name = "mdl_Import_Export_Code"
    $cd_object_type = 5 # acModule
    $pf_object_path = "$(Get-Target-Path-Code)\Modules\$nm_object_name.bas"
    if ($global:is_debugging) { Write-Host "Path to code module: $pf_object_path" }
    #
    # Check if the object file exists
    if (-Not(Test-Path -Path $pf_object_path)) {
        $global:tx_error = "Error 3: Code module file does not exist at: $pf_object_path"
    } 
    else { # Load the Object into the Access application
        try { 
            $ip_ob_access_appl.Application.LoadFromText($cd_object_type, "$nm_object_name", "$pf_object_path") 
            if ($global:is_debugging) { Write-Host "Code module '$nm_object_name' loaded successfully." }
        } 
        catch { $global:tx_error = "Error 4: $($_.Exception.Message)" } 
    }
    #
    # If no error occurred, return the Access application object
    if ($null -eq $global:tx_error) { $return = $ip_ob_access_appl } else { $return = $null }
    return $return
    #
}
#
#
function Exec-Access-internal-Function-Import-All($ip_ob_access_appl) {
    # This function executes the internal function "Import_All" in the Access application.
    # It uses the DoCmd.Run method of the Access.Application COM object to run the function.
    # The function returns an error message if the function fails to execute.
    try { 
        $ip_ob_access_appl.Run("ImportAll")
        if ($global:is_debugging) { Write-Host "Internal function 'ImportAll' executed successfully." }
    } 
    catch { $global:tx_error = "Error 5: Failed to execute ImportAll function. Error: $_" }
    #
    # If no error occurred, return the Access application object
    if ($null -eq $global:tx_error) { $return = $ip_ob_access_appl } else { $return = $null }
    return $return
    #
}
# 
#
# This function retrieves the root path of where Git repository is located.
function Get-Git-Root() {
    $fp_git_root = $PSScriptRoot.Substring(0, ($PSScriptRoot.IndexOf("$(Get-Repository-Name)")-1))
    return $fp_git_root
}
#
# This function retrieves folder path to the template folder.
function Get-Template-Path() {
    #
    # Build the path to the template folder
    $fp_template = "$(Get-Git-Root)" + "\Template"
    #
    # If the template path does not exist, create it
    if (-not(Test-Path -Path $fp_template)) { New-Item -ItemType Directory -Path $fp_template -Force | Out-Null }
    #
    # Return the template path
    return $fp_template
    #
}
#
# This function retrieves the path to the "meta-data-def" folder.
function Get-Meta-Data-Def() {
    #
    # Build the path to the template folder
    $fp_meta_data_def = "$(Get-Template-Path)\meta-data-def"
    #
    # If the template path does not exist, create it
    if (-not(Test-Path -Path $fp_meta_data_def)) { 
        New-Item -ItemType Directory -Path $fp_meta_data_def -Force | Out-Null
        git clone https://github.com/Demo-Simple-Analytyic-Platform/meta-data-def.git "$fp_meta_data_def" 2>$null
    } 
    else { # If the folder already exists, check if there should be a git repository, if not clone the repository else update it
        if (-not(Test-Path -Path "$fp_meta_data_def\.git")) {
            # If the folder exists but is not a git repository, clone the repository
            git clone https://github.com/Demo-Simple-Analytyic-Platform/meta-data-def.git "$fp_meta_data_def" 2>$null
        }
        else { # If the folder is a git repository, update it
            #
            # Save the current location
            $fp_current_location = Get-Location
            #
            # Change to the meta data definition folder, pull the latest changes from the main branch
            Set-Location -Path $fp_meta_data_def
            # 
            # Pull the latest changes from the main branch 
            git pull origin main 2>$null
            #
            # Change back to the original location
            Set-Location -Path $fp_current_location
            #
        }
    }
    #
    # Return the "meta-data-def" folder path
    $return = "$(Get-Template-Path)\meta-data-def"
    return "$return"
    #
}
#
# Get Secure Folder Path on user folder
function Get-Secure-Folder-Path() {
    #
    # Build the path to the secure folder
    $fp_secure_folder = "C:\users\$([System.Environment]::UserName)\secure"
    #
    # If the secure folder does not exist, create it
    if (-not(Test-Path -Path $fp_secure_folder)) { New-Item -ItemType Directory -Path $fp_secure_folder -Force | Out-Null }
    #
    # Return the secure folder path
    return $fp_secure_folder
    #
}
#
# Get Secure Model Folder Path on user folder
function Get-Secure-Model-Folder-Path() {
    #
    # Build the path to the secure model folder
    $nm_model               = "$(Get-Repository-Name)"
    $fp_secure_folder       = "$(Get-Secure-Folder-Path)"
    $fp_secure_model_folder = "$fp_secure_folder\$nm_model"
    #
    # If the secure model folder does not exist, create it
    if (-not(Test-Path -Path $fp_secure_model_folder)) { New-Item -ItemType Directory -Path $fp_secure_model_folder -Force | Out-Null }
    #
    # Return the secure model folder path
    return $fp_secure_model_folder
}   
#
# Get Remove Secure Model Folder Path on user folder
function Get-Remote-Secure-Model-Folder-Path() {
    #
    # Build the path to the remote secure model folder
    $nm_model               = "$(Get-Repository-Name)"
    $fp_secure_folder       = "$(Get-Secure-Folder-Path)"
    $fp_secure_model_folder = "$fp_secure_folder\$nm_model"
    #
    # Remove all Files containting "secure" info.
    Remove-All-Files-From-Folder -folderPath $fp_secure_model_folder
    #
}
#
# Store secure information in the secure model folder
function Set-Secure-Information($ip_nm_information) {
    #
    # Get the secure model folder path
    $fp_secure_model_folder = "$(Get-Secure-Model-Folder-Path)"
    $fp_information = "$fp_secure_model_folder\$ip_nm_information.txt"
    #
    # Store the secure information in the secure model folder
    if (-not (Test-Path $fp_information)) { 
        $secure_tx_information = Read-Host "Provide $ip_nm_information : " -AsSecureString
        $secure_tx_information | ConvertFrom-SecureString | Set-Content "$fp_information"
        if ($global:is_debugging) { Write-Host "Secure information stored in: $fp_information" }
    }
    #
    # Get Information from the secure model file
    $secure_tx_information = Get-Content $fp_information | ConvertTo-SecureString
    if ($global:is_debugging) { Write-Host "Secure information retrieved from: $fp_information" }
    #
    # Return the secure information
    return $secure_tx_information
    #
}
#
# Convert secure information to plain text
function Convert-Secure-Information-To-PlainText($ip_tx_secure_information) {
    #
    # Convert the secure string to plain text
    $plain_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ip_tx_secure_information))
    if ($global:is_debugging) { Write-Host "Secure information converted to plain text." }
    #
    # Return the plain text
    return $plain_text
} 
#
# Remove all files for give folder path and copy all files form source path to target path
function Remove-All-Files-And-Copy($ip_source_path, $ip_target_path) {  
    #
    # Remove all files from the target path
    if (Test-Path -Path $ip_target_path) { 
        Remove-Item -Path $ip_target_path\* -Recurse -Force 
        if ($global:is_debugging) { Write-Host "All files removed from folder: $ip_target_path" }
    }
    # Copy all files from the source path to the target path
    Copy-Item -Path ip_source_path\* -Destination $ip_target_path -Recurse -Force
    if ($global:is_debugging) { Write-Host "Files copied from $ip_source_path to $ip_target_path" }
    #
}
#
# Remoce single file from the target path and copy file from source path to target path
function Remove-File-And-Copy($ip_source_path, $ip_target_path) {
    #
    # Remove the file from the target path
    if (Test-Path -Path $ip_target_path) { 
        Remove-Item -Path $ip_target_path -Force 
        if ($global:is_debugging) { Write-Host "File removed from: $ip_target_path" }
    }
    # Copy the file from the source path to the target path
    Copy-Item -Path $ip_source_path -Destination $ip_target_path -Force
    if ($global:is_debugging) { Write-Host "File copied from $ip_source_path to $ip_target_path" }
    #
}