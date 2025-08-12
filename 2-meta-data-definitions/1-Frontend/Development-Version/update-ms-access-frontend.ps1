# This powershell script updates the MS Access frontend application by copying the latest code files from the development version to the build version.
# It ensures that the build version is up-to-date with the latest changes made in the development version.
#
# Example: powershell -ExecutionPolicy Bypass -File "update-ms-access-frontend.ps1"
#
# Import Helper functions for the script
. "$PSScriptRoot\Code\PowerShell\helpers.ps1"
#
# Ensure the "meta-data-def" repository is cloned or updated
$fp_meta_data_def = ("$(Get-Meta-Data-Def)").Replace("Already up to date. ", "")

# 
# Get the target path for the code
$fp_forms             = "$(Get-Target-Path-Code)\Forms"
$fp_modules           = "$(Get-Target-Path-Code)\Modules" 
$fp_queries           = "$(Get-Target-Path-Code)\Queries"
$fp_tables            = "$(Get-Target-Path-Code)\Tables"
$fp_target_appl       = "$(Get-Target-Path-Application)"
$fp_frontend          = "$fp_meta_data_def\2-meta-data-definitions\1-Frontend" # this will also clone of update the repository there
$fp_template          = "$fp_frontend\Development-Version\Code"
$fp_post_deployment_s = "$fp_meta_data_def" + "\9-Publish\1-Scripts\Script.PostDeployment.sql"
$fp_post_deployment_t = "$(Get-Repository-Path)\9-Publish\1-Scripts\Script.PostDeployment.sql"
$fp_deploy_of_model_s = "$fp_meta_data_def" + "\9-Publish\1-Scripts\deployment-of-model.ps1"
$fp_deploy_of_model_t = "$(Get-Repository-Path)\9-Publish\1-Scripts\deployment-of-model.ps1"
#
if ($global:is_debugging) {
    Write-Host "Target path for Forms             : '$fp_forms'"
    Write-Host "Target path for Modules           : '$fp_modules'"
    Write-Host "Target path for Queries           : '$fp_queries'"
    Write-Host "Target path for Tables            : '$fp_tables'"
    Write-Host "Target path for Application       : '$fp_target_appl'"
    Write-Host "Template path for Code            : '$fp_template'"
    Write-Host "Template path for Frontend        : '$fp_frontend'"
    Write-Host "Post Deployment Script Source     : '$fp_post_deployment_s'"
    Write-Host "Post Deployment Script Target     : '$fp_post_deployment_t'"
    Write-Host "Deployment of Model Script Source : '$fp_deploy_of_model_s'"
    Write-Host "Deployment of Model Script Target : '$fp_deploy_of_model_t'"
}
#
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# !!! START: Self-updating mechanism                                !!!
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
$currentScriptPath = $MyInvocation.MyCommand.Path
$newScriptPath     = "$fp_frontend\Development-Version\update-ms-access-frontend.ps1"
$hashCurrent       = (Get-FileHash -Path $currentScriptPath).Hash
$hashNew           = (Get-FileHash -Path $newScriptPath).Hash
#
# Check if the new script exists and if the hashes are different
if ($hashCurrent -ne $hashNew) {
    if ($global:is_debugging) { Write-Host "New     script path : $newScriptPath" }
    if ($global:is_debugging) { Write-Host "Current script path : $currentScriptPath" }
    if ($global:is_debugging) { Write-Host "New     script hash : $hashNew" }
    if ($global:is_debugging) { Write-Host "Current script hash : $hashCurrent" }
    Write-Host "Updating script with the latest version..."
    Copy-Item -Path $newScriptPath -Destination $currentScriptPath -Force
    Write-Host "Restarting script..."
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$currentScriptPath`"" -NoNewWindow
    exit
} else { Write-Host "No update needed. Current script is up-to-date." }   
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# !!! ENDED: Self-updating mechanism                                !!!
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#
# 1. Remove old "code"-files from all relevant folders.
Remove-All-Files-And-Copy("$fp_template\Forms", $fp_forms)
Remove-All-Files-And-Copy("$fp_template\Modules", $fp_modules)
Remove-All-Files-And-Copy("$fp_template\Queries", $fp_queries)
Remove-All-Files-And-Copy("$fp_template\Tables", $fp_tables)
Remove-File-And-Copy("$fp_frontend\start-meta-data-editor.bat", "$fp_target_appl\start-meta-data-editor.bat") 
Remove-File-And-Copy("$fp_frontend\start-meta-data-editor.ps1", "$fp_target_appl\start-meta-data-editor.ps1")
Remove-File-And-Copy($fp_post_deployment_s, $fp_post_deployment_t) 
Remove-File-And-Copy($fp_deploy_of_model_s, $fp_deploy_of_model_t)
#
# Done
Write-Host "Meta Data Definition Frontend application updated successfully!"