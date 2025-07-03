Option Compare Database
Option Explicit
'
Public Function repos(Optional ByRef relative As Boolean = False) As String
    'Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM [repo]")
    repos = tx_repo_folderpath(nm_repository)
    If (relative = True) Then: repos = relative_path(repos)
End Function
Public Function srd(Optional ByRef relative As Boolean = False) As String:   srd = repos(relative) & "1-Static-Reference-Data\": End Function
Public Function ohg(Optional ByRef relative As Boolean = False) As String:   ohg = repos(relative) & "2-Organization-Hierarchies-and-Groups\": End Function
Public Function dta(Optional ByRef relative As Boolean = False) As String:   dta = repos(relative) & "3-Data-Transformation-Area\": End Function
Public Function dqm(Optional ByRef relative As Boolean = False) As String:   dqm = repos(relative) & "4-Data-Quality-Model\": End Function

Public Function fld(cd_domain As String) As String

    If cd_domain = "srd" Then fld = srd()
    If cd_domain = "ohg" Then fld = ohg()
    If cd_domain = "dta" Then fld = dta()
    If cd_domain = "dqm" Then fld = dqm()
            
End Function
Public Function relative_path(tx_path) As String
    
    Dim ni_start As Integer: ni_start = InStr(1, tx_path, "\2-Definitions", vbTextCompare)
    relative_path = "." & Mid(tx_path, ni_start + Len("\2-Definitions"))

End Function

Public Sub save_tx_path_repos(ByVal tx_path_repos As String)
    '
    ' Local Variables
    Dim rst As Recordset
    Dim sql As String
    '
    ' clean up existing info
    sql = "DELETE * FROM repo"
    DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Insert "current" repo info
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM [repo]")
    rst.AddNew
    rst!nm_repo = nm_repository()
    rst!tx_folderpath = Replace(Replace(tx_path_repos, "1-Frontend", "2-Definitions"), "\Development-Version", "")
    rst.Update
    rst.Close
    '
End Sub


Public Function tx_user_folder() As String
    
    tx_user_folder = "C:\Users\" & nm_user_name() & "\" & nm_repository() & "\"

End Function

Public Function nm_user_name() As String

    nm_user_name = Environ("Username")
    
End Function
Public Function tx_git_folder() As String
    tx_git_folder = Mid(CurrentProject.Path, 1, InStr(1, CurrentProject.Path, "\" & nm_repository(), vbTextCompare))
End Function
Public Function tx_repo_folderpath(ip_nm_repository As String) As String
    '
    ' this will return the expacted folder path to the metadata defintions of the "model".
    '
    tx_repo_folderpath = tx_git_folder() & ip_nm_repository & "\2-meta-data-definitions\2-Definitions\"
    '
End Function
Public Function tx_repo_folderpath_exists(ip_nm_repository As String) As Boolean
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    tx_repo_folderpath_exists = fso.FolderExists(tx_repo_folderpath(ip_nm_repository))
    
End Function

Public Function is_development_version() As Boolean
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fld As folder:           Set fld = fso.GetFolder(CurrentProject.Path)
    is_development_version = (fld.Name = "Development-Version")
End Function
Public Function nm_repository() As String
    '
    ' This will extract the name of the model from the currentProject path.
    '
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fld As folder:           Set fld = fso.GetFolder(CurrentProject.Path)
    nm_repository = IIf(is_development_version, fld.ParentFolder.ParentFolder.ParentFolder.Name, fld.ParentFolder.ParentFolder.Name)
    '
    ' Check lengte of model name, max 16 charactere allowed!
    If (Len(nm_repository) > 16) Then
        MsgBox "The model name of `" & nm_repository & "` is to long, mx 16 characters allowed!" & vbNewLine & "Action: Change the name of the repository into something with a length of max 16 characters.", vbCritical
        Application.Quit
        '
    End If
    '
End Function
Public Function nm_model() As String
    nm_model = nm_repository
End Function
Public Function id_model(ip_nm_repository As String) As String
    id_model = generate_id_model(ip_nm_repository)
End Function
Public Function generate_id_model(ip_nm_repository As String) As String
    '
    ' THIS WILL "Encrypt" the nm_repository so id_model is alway the same
    '
    If (Len(ip_nm_repository) > 16) Then
        MsgBox "Modelname is to long, max length allowed is 16 characters!", vbCritical
        Exit Function
    End If
    '
    ' Hash/encripy Model Name
    '
    ' Local Vaiables
    Dim ark_data          As Long
    Dim encryption_key    As Long
    Dim encrypted_value   As String
    Dim index_xor_value_1 As Integer
    Dim index_xor_value_2 As Integer
    Dim temp              As Integer
    Dim tempstring        As String
    '
    ' if shorte then 16 append with space!
    ip_nm_repository = ip_nm_repository & String(16 - Len(ip_nm_repository), "-")
    '
    ' Get Encryption Key
    encryption_key = 1234567890
    '
    ' Loop throught all characters of the to be encrypted value
    For ark_data = 1 To Len(ip_nm_repository)
        '
        'The first value to be XOr-ed comes from the data to be encrypted
        index_xor_value_1 = Asc(Mid$(ip_nm_repository, ark_data, 1))
        '
        'The second value comes from the code key
        index_xor_value_2 = Asc(Mid$(encryption_key, ((ark_data Mod Len(encryption_key)) + 1), 1))
        '
        ' Rememeber tempory result
        temp = (index_xor_value_1 Xor index_xor_value_2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        '
        ' Add to Encrypted result
        encrypted_value = encrypted_value + tempstring
        '
    Next ark_data
    '
    ' Return De-crypted Result
    generate_id_model = LCase(encrypted_value)
    '
End Function