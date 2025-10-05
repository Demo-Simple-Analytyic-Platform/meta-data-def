# PowerShell script for processing database operations
# This script contains functions to interact with the SQL database for processing operations

# Import Helper functions for the script
$helpersPath = "$($PSScriptRoot.Replace('\powershell', ''))\2-meta-data-definitions\1-Frontend\Development-Version\Code\PowerShell\helpers.ps1"
if (Test-Path $helpersPath) {
    . $helpersPath
}

# Function to get database connection credentials (similar to the Python credentials module)
function Get-Database-Credentials {
    param(
        [string]$CredentialType = "target"  # "target" or "secrets"
    )
    
    try {
        # Extract Credentials from secure files (this follows the same pattern as deployment script)
        $tx_secure_server   = Set-Secure-Information -ip_nm_information "Server"
        $tx_secure_database = Set-Secure-Information -ip_nm_information "Database"
        $tx_secure_username = Set-Secure-Information -ip_nm_information "Username"
        $tx_secure_password = Set-Secure-Information -ip_nm_information "Password"
        
        # Return credentials object
        return @{
            Server   = Convert-Secure-Information-To-PlainText($tx_secure_server)
            Database = Convert-Secure-Information-To-PlainText($tx_secure_database)
            Username = Convert-Secure-Information-To-PlainText($tx_secure_username)
            Password = Convert-Secure-Information-To-PlainText($tx_secure_password)
        }
    }
    catch {
        Write-Error "Failed to retrieve database credentials: $_"
        return $null
    }
}

# Function to create SQL connection string
function Get-Connection-String {
    param(
        [hashtable]$Credentials,
        [string]$Encrypt = "False",
        [string]$TrustServerCertificate = "False"
    )
    
    if (-not $Credentials) {
        throw "Database credentials are required"
    }
    
    # Build connection string (similar to Python sql.py module)
    $connectionString = "Server=$($Credentials.Server);" +
                       "Database=$($Credentials.Database);" +
                       "User Id=$($Credentials.Username);" +
                       "Password=$($Credentials.Password);" +
                       "Encrypt=$Encrypt;" +
                       "TrustServerCertificate=$TrustServerCertificate;" +
                       "Connection Timeout=30;" +
                       "Command Timeout=300;"
    
    return $connectionString
}

# Function to execute SQL query and return results
function Invoke-SQL-Query {
    param(
        [string]$Query,
        [hashtable]$Credentials = $null,
        [switch]$IsDebugging = $false
    )
    
    try {
        # Get credentials if not provided
        if (-not $Credentials) {
            $Credentials = Get-Database-Credentials
            if (-not $Credentials) {
                throw "Failed to obtain database credentials"
            }
        }
        
        # Get connection string
        $connectionString = Get-Connection-String -Credentials $Credentials
        
        if ($IsDebugging) {
            Write-Host "Executing SQL Query: $($Query.Substring(0, [Math]::Min(100, $Query.Length)))..." -ForegroundColor Yellow
        }
        
        # Create SQL connection
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()
        
        # Create and execute command
        $command = New-Object System.Data.SqlClient.SqlCommand($Query, $connection)
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataset = New-Object System.Data.DataSet
        
        # Fill dataset with results
        $adapter.Fill($dataset) | Out-Null
        
        # Close connection
        $connection.Close()
        
        if ($IsDebugging) {
            Write-Host "Query executed successfully. Rows returned: $($dataset.Tables[0].Rows.Count)" -ForegroundColor Green
        }
        
        # Return results as array of objects
        return $dataset.Tables[0]
    }
    catch {
        Write-Error "SQL Query execution failed: $_"
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
        throw $_
    }
}

# Function to extract array of process group numbers from the SQL database
function Get-Process-Group-Numbers {
    param(
        [switch]$IsDebugging = $false,
        [string]$IdModel = $null,
        [hashtable]$Credentials = $null
    )
    
    try {
        if ($IsDebugging) {
            Write-Host "Extracting process group numbers from database..." -ForegroundColor Cyan
        }
        
        # Build SQL query (based on the pattern from Python run.py module)
        $query  = "SELECT DISTINCT ni_process_group "
        $query += "`nFROM dta.process_group"
        $query += "`nWHERE 1=1"

        # Add filters if provided (following the Python pattern)
        if ($IdModel) {
            $query += "`nAND id_model = '$IdModel'"
        }
        
        # Add Ordering Ascending
        $query += "`nORDER BY ni_process_group ASC"
        
        if ($IsDebugging) {
            Write-Host "SQL Query:" -ForegroundColor Yellow
            Write-Host $query -ForegroundColor White
        }
        
        # Execute query
        $results = Invoke-SQL-Query -Query $query -Credentials $Credentials -IsDebugging:$IsDebugging
        
        # Extract process group numbers into array
        $processGroupNumbers = @()
        foreach ($row in $results.Rows) {
            $processGroupNumbers += $row["ni_process_group"]
        }
        
        if ($IsDebugging) {
            Write-Host "Found $($processGroupNumbers.Count) process groups: $($processGroupNumbers -join ', ')" -ForegroundColor Green
        }
        
        return $processGroupNumbers
    }
    catch {
        Write-Error "Failed to extract process group numbers: $_"
        throw $_
    }
}

# Function to get all process groups with detailed information
function Get-Process-Groups-Detailed {
    param(
        [string]$IdModel = $null,
        [switch]$IsDebugging = $false,
        [hashtable]$Credentials = $null
    )
    
    try {
        if ($IsDebugging) {
            Write-Host "Extracting detailed process group information from database..." -ForegroundColor Cyan
        }
        
        # Build comprehensive SQL query (based on the pattern from Python run.py module)
        $query   = "SELECT ni_process_group, "
        $query  += "`n       id_dataset, "
        $query  += "`n       is_ingestion, "
        $query  += "`n       nm_procedure, "
        $query  += "`n       nm_tsl_schema, "
        $query  += "`n       nm_tsl_table, "
        $query  += "`n       nm_tgt_schema, "
        $query  += "`n       nm_tgt_table"
        $query  += "`nFROM dta.process_group"
        $query  += "`nWHERE 1=1"
        
        # Add filters if provided
        if ($IdModel) {
            $query += "`nAND id_model = '$IdModel'"
        }
                
        $query += "`nORDER BY ni_process_group ASC"
        
        if ($IsDebugging) {
            Write-Host "SQL Query:" -ForegroundColor Yellow
            Write-Host $query -ForegroundColor White
        }
        
        # Execute query and return results
        $results = Invoke-SQL-Query -Query $query -Credentials $Credentials -IsDebugging:$IsDebugging
        
        if ($IsDebugging) {
            Write-Host "Found $($results.Rows.Count) process group records" -ForegroundColor Green
        }
        
        return $results
    }
    catch {
        Write-Error "Failed to extract detailed process group information: $_"
        throw $_
    }
}

# Example usage function
function Test-Process-Group-Functions {
    param(
        [switch]$IsDebugging = $true
    )
    
    try {
        Write-Host "Testing Process Group Functions..." -ForegroundColor Magenta
        
        # Get Database Credentials
        $credentials = Get-Database-Credentials
       

        # Test getting process group numbers
        Write-Host "`n1. Getting Process Group Numbers:" -ForegroundColor Yellow
        $processGroups = Get-Process-Group-Numbers -IsDebugging:$IsDebugging -IdModel "5f4a1942465c575a1f5a5a575d1e191c" -Credentials $credentials
        Write-Host "Process Groups Array: [$($processGroups -join ', ')]" -ForegroundColor White
        
        # Test getting detailed information
        Write-Host "`n2. Getting Detailed Process Group Information:" -ForegroundColor Yellow
        $detailedInfo = Get-Process-Groups-Detailed -IsDebugging:$IsDebugging
        
        if ($detailedInfo.Rows.Count -gt 0) {
            Write-Host "Sample records:" -ForegroundColor White
            $detailedInfo.Rows | Select-Object -First 3 | ForEach-Object {
                Write-Host "  Group: $($_.ni_process_group), Dataset: $($_.id_dataset), Ingestion: $($_.is_ingestion)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`nTesting completed successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Testing failed: $_"
        return $false
    }
}



# Main execution example (uncomment to test)
Test-Process-Group-Functions -IsDebugging:$true

Write-Host "Processing.ps1 loaded successfully. Available functions:" -ForegroundColor Green
Write-Host "  - Get-Process-Group-Numbers" -ForegroundColor White
Write-Host "  - Get-Process-Groups-Detailed" -ForegroundColor White
Write-Host "  - Get-Database-Credentials" -ForegroundColor White
Write-Host "  - Invoke-SQL-Query" -ForegroundColor White
Write-Host "  - Test-Process-Group-Functions" -ForegroundColor White
