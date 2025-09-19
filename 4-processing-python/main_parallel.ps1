# -----------------------------------------------------------------------------
# PowerShell script to run multiple instances of a Python script in parallel
# -----------------------------------------------------------------------------

# Initialize the processing environment, waiting for it to complete
function run-python-script-initialization() {

    #$fp_python = "D:\git\Misset-Data-Analytics\My-Financial-Stock-Information\my-stock-info\4-processing-python\initialize.py"
    $fp_python = "initialize.py"
    Start-Process "python" -ArgumentList "$fp_python" -Wait

    try {
        # Extract max number of process groups from control file
        $fp_control_max_process_group = "C:/Temp/control_max_process_group.txt"
        if (Test-Path $fp_control_max_process_group) {
            $ni_process_group = Get-Content $fp_control_max_process_group
            Write-Host "Max process group number extracted: $ni_process_group" -ForegroundColor Green
        } else {
            Write-Host "Control file for max process group not found: $fp_control_max_process_group" -ForegroundColor Red
            exit 1
        }

        # Extact max numeber of process groups from control file
        $fp_control_id_model = "C:/Temp/control_id_model.txt"
        if (Test-Path $fp_control_id_model) {
            $id_model = Get-Content $fp_control_id_model
            Write-Host "ID model extracted: $id_model" -ForegroundColor Green
        } else {
            Write-Host "Control file for ID model not found: $fp_control_id_model" -ForegroundColor Red
            exit 1
        }

        # Return both id_model and ni_process_group
        return @{ 
            id_model = $id_model; 
            ni_process_group = $ni_process_group }
    }
    catch {
        Write-Error "Failed to retrieve processing environment variables: $_"
        throw $_
    }
}

# Processing environment initialization
$result = run-python-script-initialization
$id_model = $result.id_model
$ni_process_group_max = $result.ni_process_group
Write-Host "Processing environment initialized with ID Model: '$id_model' and Max Process Groups: $ni_process_group" -ForegroundColor Green

# Set python script path
$fp_main_parallel = "main_parallel.py"
$ni_sessions      = 5 # --> Number of parallel sessions
$ni_process_group = 0 # --> Current process group to be assigned

# Set progress bar style
#$PSStyle.Progress.View = 'Minimal'

try {
    # Add empty line for better readability so the progress bars do not overwrite the last line of output
    1..$ni_sessions | ForEach-Object { Write-Host "`n" }

    # While ni_process_group is less or equal than ni_process_group_max
    while ($ni_process_group -le $ni_process_group_max) {

        # Start # ni_sessions to process the dataset of process group
        
        Write-Host "Starting $ni_sessions parallel sessions for process group $ni_process_group..." -ForegroundColor Cyan
        1..$ni_sessions | ForEach-Object {
            #
            # Set unique session ID
            $id_session = $_
            #
            # Create a control file for this session
            $fp_control = "C:/Temp/control_parallel_$id_session.txt"
            New-Item -Path $fp_control -ItemType File -Force | Out-Null
            #
            # Start the Python script in a new process
            Start-Process "python" -ArgumentList "$fp_main_parallel --ni-process-group $ni_process_group --ni-sessions $ni_sessions --id-session $id_session"
            #
            # Optional: Add a small delay to avoid overwhelming the system
            Start-Sleep -Milliseconds 100
            #
        }
        #
        # Initalize progress monitoring
        Write-Host "Monitoring progress of $ni_sessions parallel sessions..."
        #
        ## Wait for all sessions to complete by monitoring the control files
        $ni_control  = $ni_sessions
        $ar_sessions = @{}
        1..$ni_sessions | ForEach-Object { $ar_sessions[$_] = 0 }
        #
        while ($ni_control -gt 0) {
            #
            # Reset control files counter
            $ni_control = 0
            #
            # Update progress bar (this is a simple simulation, adjust as needed) for each session a progress bar is shown
            # Simulate progress for demonstration purposes, if progress is hard 10 dots it will reset, until control file for session is gone
            # if no control file for session is found then progress should shown "Completed"
            1..$ni_sessions | ForEach-Object {
                $id_session = $_
                $fp_control = "C:/Temp/control_parallel_$id_session.txt"
                #    
                if (Test-Path $fp_control) {
                    $ni_control += 1
                    $ar_sessions[$id_session] = ($ar_sessions[$id_session] + 1) % 10
                    $dots = "." * ($ar_sessions[$id_session] + 1)
                    Write-Progress -Id $id_session -Activity "Session $id_session Processing for Process Group $ni_process_group" -Status "Working$dots" 
                } else {
                    # Session completed - control file removed
                    Write-Progress -Id $id_session -Activity "Session $id_session Processing for Process Group $ni_process_group" -Status "Completed"
                }
            }
            Start-Sleep -Milliseconds 500
        }
        ##
        ## Clear all progress bars when completed
        1..$ni_sessions | ForEach-Object {
            Write-Progress -Id $_ -Activity "Session $_ Processing" -Completed
        }
        ##
        Write-Host "All $ni_sessions sessions completed successfully! for process group $ni_process_group" -ForegroundColor Green
        #
        # Increment process group number
        $ni_process_group += 1 
    }
}
catch {
    Write-Host "An error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "Stack Trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
## Optional: Prevent script from exiting immediately
Read-Host "Wait for all sessions to complete and then press Enter to continue."
