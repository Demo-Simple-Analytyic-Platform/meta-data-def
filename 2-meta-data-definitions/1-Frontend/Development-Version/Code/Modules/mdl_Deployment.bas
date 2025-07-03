Option Compare Database
Option Explicit

Public Sub run_powershell_to_deploy_metadata_defintitions()
    '
    ' Local Variables
    Dim shell As Object
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fp_script As String:   fp_script = tx_git_folder & "\" & nm_repository & "\2-meta-data-definitions\9-Publish\1-Scripts\deployment-of-model.ps1"
    Dim fp_logging As String:  fp_logging = "C:\temp\deployment_of_" & nm_repository & "_" & Format(Now, "yyyyMMddhhmmss") & ".log"
    Dim fp_waiting As String:  fp_waiting = "C:\temp\deployment_of_" & nm_repository & ".wait"
    Dim old As TextStream
    Dim txt As TextStream
    Dim line As String
    '
    ' Rename "Exiting" to old
    Call fso.CopyFile(fp_script, Replace(fp_script, ".ps1", ".old"), True)
    Call fso.DeleteFile(fp_script, True)
    '
    ' Open "Old"-version and replace "placeholder"
    Set old = fso.OpenTextFile(Replace(fp_script, ".ps1", ".old"), ForReading, False, TristateTrue)
    Set txt = fso.OpenTextFile(fp_script, ForWriting, True, TristateTrue)
    '
    ' Loop though all lines and replace placeholder.
    Do While Not old.AtEndOfStream
        line = old.ReadLine
        line = Replace(line, "<nm_model>", nm_repository)
        line = Replace(line, "<tx_git_folder>", Mid(tx_git_folder, 1, Len(tx_git_folder) - 1))
        txt.WriteLine line
    Loop
    txt.Close
    old.Close
    '
    ' set windows Command: start script and write output to log.
    Dim cmd As String: cmd = "cmd.exe /k powershell -ExecutionPolicy Bypass -File """ & fp_script & """ >> """ & fp_logging & """"
    '
    ' Start Wait file
    Set txt = fso.CreateTextFile(fp_waiting, True, True): txt.WriteLine "wait": txt.Close
    '
    'Create Shell to start.execute command from
    DoCmd.SetWarnings False
    Set shell = CreateObject("WScript.Shell")
    shell.Run cmd, False, False  ' 1 = show window, False = do not wait for completion
    DoCmd.SetWarnings True
    '
    'Start wait screen
    DoCmd.OpenForm "waiting"
    '
End Sub