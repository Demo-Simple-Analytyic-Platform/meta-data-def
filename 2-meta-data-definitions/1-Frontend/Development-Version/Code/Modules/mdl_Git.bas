Option Compare Database
Option Explicit
'
'### Basic Assumtions:
'- Ensure Git is installed and added to the system's PATH.
'- If authentication is required for pushing, make sure your Git credentials are configured (e.g., using a credential manager or SSH key).
'
' Function to check if Git is installed
Public Function IsGitInstalled() As Boolean
    '
    '### Explanation:
    '1. **`git --version`**:
    '   - This command outputs the installed Git version if Git is available on the system.
    '   - If Git is not installed, the command will fail.
    '
    '2. **Error Handling**:
    '   - If the command fails (e.g., Git is not found), the function will return `False`.
    '
    '3. **Output Check**:
    '   - The function checks if the output contains the text `git version` to confirm that Git is installed.
    '
    '### Notes:
    '- Ensure the system's PATH environment variable includes the Git executable directory.
    '- This function uses error handling to gracefully handle cases where Git is not installed.
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim wshShell    As Object
    Dim wshExec     As Object
    Dim shellOutput As String
    '
On Error GoTo ErrorHandler
    '
    ' Construct the Git command to check the version
    gitCommand = "cmd.exe /c git --version"
    '
    ' Create a WScript Shell object to capture the output
    Set wshShell = CreateObject("WScript.Shell")
    Set wshExec = wshShell.Exec(gitCommand)
    '
    ' Read the output of the command
    shellOutput = wshExec.StdOut.ReadAll
    '
    ' If the output contains "git version", Git is installed
    If InStr(LCase(shellOutput), "git version") > 0 Then
        IsGitInstalled = True
    Else
        IsGitInstalled = False
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), Git is not installed
    IsGitInstalled = False
    '
End Function
'
'
' Function to clone a Git repository
Public Function CloneGitRepo(repoUrl As String, targetDir As String) As String
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim wshShell    As Object
    Dim wshExec     As Object
    Dim shellOutput As String
    '
On Error GoTo ErrorHandler
    '
    ' Description
    '### How to Use:
    'You can call this function to clone a Git repository by providing the repository URL and the target directory. For example:
    '
    '```vba
    'Function TestCloneGitRepo()
    '    Dim repoUrl As String
    '    Dim targetDir As String
    '
    '    repoUrl = "https://github.com/folders/repo_name.git"
    '    targetDir = "C:\git\repo_name"
    '
    '    Call CloneGitRepo(repoUrl, targetDir)
    'End Function
    '```
    '
    '### Explanation:
    '1. **`git clone <url> <directory>`**:
    '   - Clones the repository from the specified URL into the target directory.
    '
    '2. **Error Handling**:
    '   - If the command fails (e.g., due to an invalid URL or directory), the function displays an error message.
    '
    '3. **Output Check**:
    '   - The function checks the output for keywords like `fatal` or `error` to determine if the cloning operation failed.
    '
    '### Notes:
    '- Ensure Git is installed and added to the system's PATH.
    '- The target directory should not already contain a Git repository; otherwise, the command will fail.
    '- If authentication is required (e.g., for private repositories), ensure your Git credentials are configured (e.g., using a credential manager or SSH key).
    '
    ' Construct the Git command to clone the repository
    gitCommand = "cmd.exe /c git clone " & Chr(34) & repoUrl & Chr(34) & " " & Chr(34) & targetDir & Chr(34)
    '
    ' Create a WScript Shell object to execute the command
    Set wshShell = CreateObject("WScript.Shell")
    Set wshExec = wshShell.Exec(gitCommand)
    '
    ' Read the output of the command
    shellOutput = wshExec.StdOut.ReadAll & wshExec.StdErr.ReadAll
    '
    ' Check if the cloning was successful
    If InStr(LCase(shellOutput), "fatal") > 0 Or InStr(LCase(shellOutput), "error") > 0 Then
        CloneGitRepo = "Failed: to clone the repository. Please check the URL and target directory." & vbCrLf & vbCrLf & shellOutput
    Else
        CloneGitRepo = "Success: Repository cloned to: " & targetDir
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), display an error message
    CloneGitRepo = "Error: An error occurred while trying to clone the repository."
    '
End Function
'
' Function to commit changes to the local repository
Public Function CommitChanges(repoPath As String, commitMessage As String) As String
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim wshShell    As Object
    Dim wshExec     As Object
    Dim shellOutput As String
    '
On Error GoTo ErrorHandler
    '
    ' Construct the Git command to stage and commit changes
    gitCommand = "cmd.exe /c cd /d " & Chr(34) & repoPath & Chr(34) & _
                 " && git add . && git commit -m " & Chr(34) & commitMessage & Chr(34)
    '
    ' Create a WScript Shell object to capture the output
    Set wshShell = CreateObject("WScript.Shell")
    Set wshExec = wshShell.Exec(gitCommand)
    '
    ' Read the output of the command
    shellOutput = wshExec.StdOut.ReadAll & wshExec.StdErr.ReadAll
    '
    ' Check if the commit was successful
    If InStr(shellOutput, "nothing to commit") > 0 Then
        CommitChanges = "Warning: No changes to commit."
        '
    ElseIf InStr(shellOutput, "error") > 0 Or InStr(shellOutput, "fatal") > 0 Then
        CommitChanges = "Failed: Commit failed. Please check the repository path or commit message." & vbCrLf & vbCrLf & shellOutput
        '
    Else
        CommitChanges = "Success: Changes committed successfully with message: " & commitMessage
        '
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), display an error message
    CommitChanges = "Error: An error occurred while trying to Commit change to the repository."
    '
End Function
'
' Function to push changes to the remote repository
Public Function PushChanges(repoPath As String, branchName As String)
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim wshShell    As Object
    Dim wshExec     As Object
    Dim shellOutput As String
    '
On Error GoTo ErrorHandler
    '
    ' Construct the Git command to push changes
    gitCommand = "cmd.exe /c cd /d " & Chr(34) & repoPath & Chr(34) & _
                 " && git push origin " & branchName
    '
    ' Create a WScript Shell object to capture the output
    Set wshShell = CreateObject("WScript.Shell")
    Set wshExec = wshShell.Exec(gitCommand)
    '
    ' Read the output of the command
    shellOutput = wshExec.StdOut.ReadAll & wshExec.StdErr.ReadAll
    '
    ' Check if the push was successful
    If InStr(shellOutput, "error") > 0 Or InStr(shellOutput, "fatal") > 0 Then
        PushChanges = "Failed: Pushing change to Remote. Please check the repository path, branch name, or authentication settings." & vbCrLf & vbCrLf & shellOutput
    Else
        PushChanges = "Sucess: Changes pushed to branch: " & branchName
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), display an error message
    PushChanges = "Error: An error occurred while trying to Commit change to the repository."
    '
End Function
'
' Function to set the branch of a Git repository
Public Function SetGitBranch(repoPath As String, branchName As String) As String
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim shellResult As Long
    '
On Error GoTo ErrorHandler
    '
    ' Construct the Git command to switch branches
    gitCommand = "cmd.exe /c cd /d " & Chr(34) & repoPath & Chr(34) & " && git checkout " & branchName
    '
    ' Execute the command using Shell
    shellResult = shell(gitCommand, vbHide)
    '
    ' Check if the command executed successfully
    If shellResult = 0 Then
        SetGitBranch = "Success: Branch switched to: " & branchName
        '
    Else
        SetGitBranch = "Failed: Switching branch failed. Please check the repository path and branch name."
        '
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), display an error message
    SetGitBranch = "Error: An error occurred while trying to switch branch."
    '
End Function
'
' Function to sync the main branch of a Git repository and discard local changes
Public Function SyncWithRemote(repoPath As String, branchName As String) As String
    '
    ' Local Variables
    Dim gitCommand  As String
    Dim shellResult As Long
    '
On Error GoTo ErrorHandler
    '
    ' Construct the Git command to discard local changes and sync the branch
    gitCommand = "cmd.exe /c cd /d " & Chr(34) & repoPath & Chr(34) & " " & _
                 "&& git checkout " & branchName & " " & _
                 "&& git reset --hard " & _
                 "&& git pull origin " & branchName

    ' Execute the command using Shell
    shellResult = shell(gitCommand, vbHide)

    ' Check if the command executed successfully
    If shellResult = 0 Then
        SyncWithRemote = "Success: Branch synced successfully, and local changes were discarded."
    Else
        SyncWithRemote = "Failed: Synchonization with Remote failed. Please check the repository path."
    End If
    '
    ' All is well, function executed successfully
    Exit Function
    '
ErrorHandler:
    '
    ' If an error occurs (e.g., command not found), display an error message
    SyncWithRemote = "Error: An error occurred while trying to synchonize branch."
    '
End Function
'