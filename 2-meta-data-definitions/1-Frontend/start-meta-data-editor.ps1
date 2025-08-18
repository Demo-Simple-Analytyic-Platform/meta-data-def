# -----------------------------------------------------------------------------
# This script is used to build the frontend application to be used in when ever
# the application is started.
# 
# Example: powershell -ExecutionPolicy Bypass -File "testing.ps1"
#
# This script is part of the Meta Data Definition project and is used to
# 1. Create a clean Microsoft Access Database Application, remove "previous" versions of the application if needed.
# 2. Load code module for importing all required components.
# 3. load all components into the Access application.
#
# ----------------------------------------------------------------------------- 
#
# Import Helper functions for the script
."$PSScriptRoot\Development-Version\Code\PowerShell\helpers.ps1"
#
# Set global variables for the script
$global:tx_error     = $null
$global:is_debugging = $true
#
# -----------------------------------------------------------------------------
# Main script execution starts here
# -----------------------------------------------------------------------------
#
$tx_message_1 = "1. Create a clean Microsoft Access Database Application, remove `previous` versions of the application if needed."
$tx_message_2 = "2. Load code module for importing all required components."
$tx_message_3 = "3. load all components into the Access application."
#
# Create a new Access application object
if ($null -eq $global:tx_error) { if ($global:is_debugging) { Write-Host $tx_message_1 }; $ob_access_appl = Build-New-Access-Application }
if ($null -eq $ob_access_appl) { $global:tx_error = "Error A: Access application object is null."; }
#
# Load the code module for importing all required components
if ($null -eq $global:tx_error) { if ($global:is_debugging) { Write-Host $tx_message_2 };    $ob_access_appl = Import-Module-into-Access($ob_access_appl); }
if ($null -eq $ob_access_appl) { $global:tx_error = "Error B: Access application object is null."; }
#
# If no error occurred, load all components into the Access application
if ($null -eq $global:tx_error) { if ($global:is_debugging) { Write-Host $tx_message_3 }; $ob_access_appl = Exec-Access-internal-Function-Import-All($ob_access_appl); }
#
# Raise an error if the code module fails to load
if ($null -ne $global:tx_error) { throw "$($global:tx_error)" } else { Write-Host "Build of `Access` application completed!" }
#
# Optional: Prevent script from exiting immediately
Read-Host "Press Enter to close the 'meta-data-editor'"
#