Attribute VB_Name = "vsp_visual_studio_project"
Option Compare Database
Option Explicit
'
' Functions: Public Folderpaths
Public Function fp_datasets(Optional relative As Boolean = False) As String:        fp_datasets = Replace(mdl_Folders.repos(relative) & "3-Datasets\", "\2-Definitions", ""): End Function
Public Function fp_ingestions(Optional relative As Boolean = False) As String:      fp_ingestions = fp_datasets(relative) & "1-Ingestions\":           End Function
Public Function fp_transformations(Optional relative As Boolean = False) As String: fp_transformations = fp_datasets(relative) & "2-Transformations\": End Function
Public Function fp_dqcontrols(Optional relative As Boolean = False) As String:      fp_dqcontrols = fp_datasets(relative) & "3-DQ-Controls\":          End Function
Public Function fp_sql_project(Optional relative As Boolean = False) As String:     fp_sql_project = Replace(mdl_Folders.repos(relative) & "2-meta-data-definitions.sqlproj", "\2-Definitions", ""): End Function
'
' Sub: Build Folder structure for Schema/Datasets -> Ingestions/Transformations/DQ-Controls
Public Sub build_folder_structure()
  '
  ' Build Folder Structure
  Call create_folder_if_not_exists(fp_datasets())
  Call create_folder_if_not_exists(fp_ingestions())
  Call create_folder_if_not_exists(fp_dqcontrols())
  Call create_folder_if_not_exists(fp_transformations())
  '
  ' Add Folderpaths to Project
  Call AddFolderToSqlProj(Replace(fp_datasets(True), ".\", ""))
  Call AddFolderToSqlProj(Replace(fp_ingestions(True), ".\", ""))
  Call AddFolderToSqlProj(Replace(fp_dqcontrols(True), ".\", ""))
  Call AddFolderToSqlProj(Replace(fp_transformations(True), ".\", ""))
  '
End Sub
'
' Add a folder to SQL Server Database Project (.sqlproj file)
Public Function AddFolderToSqlProj(folderPath As String, Optional buildAction As String = "None", Optional is_debugging As Boolean = False) As Boolean: On Error GoTo ErrorHandler

    Dim sqlProjPath As String: sqlProjPath = fp_sql_project
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim xmlDoc As Object
    Dim txtFile As Object
    Dim xmlContent As String
    Dim itemGroupNode As Object
    Dim folderNode As Object
    Dim newLineNode   As Object
    Dim nsMgr As Object
    
    ' Check if sqlproj file exists
    If Not fso.FileExists(sqlProjPath) Then
        Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
        AddFolderToSqlProj = False
        Exit Function
    End If
    
    ' Load XML document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"
    Set newLineNode = xmlDoc.createTextNode(vbCrLf & vbTab)
    
    If Not xmlDoc.Load(sqlProjPath) Then
        Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
        AddFolderToSqlProj = False
        Exit Function
    End If
    
    ' Find or create ItemGroup for folders
    Set itemGroupNode = FindOrCreateFolderItemGroup(xmlDoc)
    
    ' Check if folder already exists
    If Not FolderExistsInProject(xmlDoc, folderPath) Then
        ' Create Folder element
        Set folderNode = xmlDoc.createNode(1, "Folder", "http://schemas.microsoft.com/developer/msbuild/2003")
        folderNode.setAttribute "Include", folderPath
        
        ' Add BuildAction child element if not "None"
        If buildAction <> "" And buildAction <> "None" Then
            Dim buildActionNode As Object
            Set buildActionNode = xmlDoc.createNode(1, "BuildAction", "http://schemas.microsoft.com/developer/msbuild/2003")
            buildActionNode.Text = buildAction
            folderNode.appendChild buildActionNode
            itemGroupNode.appendChild newLineNode
        End If
        
        ' Append folder node to ItemGroup
        itemGroupNode.appendChild folderNode
        itemGroupNode.appendChild newLineNode
        
        ' Save the modified XML back to file with proper formatting
        xmlDoc.save sqlProjPath
        
        If (is_debugging) Then Debug.Print "Successfully added folder '" & folderPath & "' to project"
        AddFolderToSqlProj = True
    Else
        If (is_debugging) Then Debug.Print "Folder '" & folderPath & "' already exists in project"
        AddFolderToSqlProj = True
    End If
    
    Set folderNode = Nothing
    Set itemGroupNode = Nothing
    Set xmlDoc = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "Error in AddFolderToSqlProj: " & Err.Number & " - " & Err.Description
    AddFolderToSqlProj = False

End Function
'
' Add a SQL file to SQL Server Database Project with Build action
Public Function AddSqlFileToSqlProj(sqlFilePath As String, Optional buildAction As String = "Build", Optional is_debugging As Boolean = False) As Boolean: On Error GoTo ErrorHandler
    
    Dim fso           As FileSystemObject: Set fso = New FileSystemObject
    Dim sqlProjPath   As String:       sqlProjPath = fp_sql_project
    Dim xmlDoc        As Object
    Dim itemGroupNode As Object
    Dim newLineNode   As Object
    Dim buildNode     As Object
    
    ' Check if sqlproj file exists
    If Not fso.FileExists(sqlProjPath) Then
        Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
        AddSqlFileToSqlProj = False
        Exit Function
    End If
    
    ' Load XML document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"
    Set newLineNode = xmlDoc.createTextNode(vbCrLf & vbTab)
    
    If Not xmlDoc.Load(sqlProjPath) Then
        Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
        AddSqlFileToSqlProj = False
        Exit Function
    End If
    
    ' Check if file already exists
    If FileExistsInProject(xmlDoc, sqlFilePath) Then
        If (is_debugging) Then Debug.Print "File '" & sqlFilePath & "' already exists in project"
        AddSqlFileToSqlProj = True
        Exit Function
    End If
    
    ' Find or create appropriate ItemGroup for Build items
    Set itemGroupNode = FindOrCreateBuildItemGroup(xmlDoc, buildAction)
    
    ' Create Build element based on buildAction
    Select Case UCase(buildAction)
        Case "BUILD"
            Set buildNode = xmlDoc.createNode(1, "Build", "http://schemas.microsoft.com/developer/msbuild/2003")
        Case "NONE"
            Set buildNode = xmlDoc.createNode(1, "None", "http://schemas.microsoft.com/developer/msbuild/2003")
        Case "POSTDEPLOY"
            Set buildNode = xmlDoc.createNode(1, "PostDeploy", "http://schemas.microsoft.com/developer/msbuild/2003")
        Case "PREDEPLOY"
            Set buildNode = xmlDoc.createNode(1, "PreDeploy", "http://schemas.microsoft.com/developer/msbuild/2003")
        Case Else
            Set buildNode = xmlDoc.createNode(1, "Build", "http://schemas.microsoft.com/developer/msbuild/2003")
    End Select
    
    buildNode.setAttribute "Include", sqlFilePath
    
    ' Append build node to ItemGroup
    itemGroupNode.appendChild buildNode
    itemGroupNode.appendChild newLineNode
    
    ' Save the modified XML back to file
    xmlDoc.save sqlProjPath
    
    If (is_debugging) Then Debug.Print "Successfully added file '" & sqlFilePath & "' to project with BuildAction='" & buildAction & "'"
    AddSqlFileToSqlProj = True
    
    Set buildNode = Nothing
    Set itemGroupNode = Nothing
    Set xmlDoc = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "Error in AddSqlFileToSqlProj: " & Err.Number & " - " & Err.Description
    AddSqlFileToSqlProj = False

End Function
'

' Helper function to find or create ItemGroup for folders
Private Function FindOrCreateFolderItemGroup(xmlDoc As Object) As Object
    Dim itemGroups As Object
    Dim itemGroup As Object
    Dim projectNode As Object
    
    ' Find existing ItemGroup that contains Folder elements
    Set itemGroups = xmlDoc.selectNodes("//msb:ItemGroup[msb:Folder]")
    
    If itemGroups.Length > 0 Then
        ' Use existing ItemGroup with Folder elements
        Set FindOrCreateFolderItemGroup = itemGroups(0)
    Else
        ' Create new ItemGroup for folders
        Set projectNode = xmlDoc.selectSingleNode("//msb:Project")
        Set itemGroup = xmlDoc.createNode(1, "ItemGroup", "http://schemas.microsoft.com/developer/msbuild/2003")
        
        ' Add some whitespace for formatting
        Dim textNode As Object
        Set textNode = xmlDoc.createTextNode(vbCrLf & "  ")
        projectNode.insertBefore textNode, Nothing
        
        ' Append the ItemGroup
        projectNode.appendChild itemGroup
        
        ' Add trailing whitespace
        Set textNode = xmlDoc.createTextNode(vbCrLf & vbCrLf & "  ")
        projectNode.appendChild textNode
        
        Set FindOrCreateFolderItemGroup = itemGroup
    End If
End Function

' Helper function to find or create ItemGroup for Build items
Private Function FindOrCreateBuildItemGroup(xmlDoc As Object, buildAction As String) As Object
    Dim itemGroups As Object
    Dim itemGroup As Object
    Dim projectNode As Object
    Dim elementName As String
    
    ' Determine element name based on build action
    Select Case UCase(buildAction)
        Case "BUILD"
            elementName = "Build"
        Case "NONE"
            elementName = "None"
        Case "POSTDEPLOY"
            elementName = "PostDeploy"
        Case "PREDEPLOY"
            elementName = "PreDeploy"
        Case Else
            elementName = "Build"
    End Select
    
    ' Find existing ItemGroup that contains the specific build element
    Set itemGroups = xmlDoc.selectNodes("//msb:ItemGroup[msb:" & elementName & "]")
    
    If itemGroups.Length > 0 Then
        ' Use existing ItemGroup
        Set FindOrCreateBuildItemGroup = itemGroups(0)
    Else
        ' Create new ItemGroup
        Set projectNode = xmlDoc.selectSingleNode("//msb:Project")
        Set itemGroup = xmlDoc.createNode(1, "ItemGroup", "http://schemas.microsoft.com/developer/msbuild/2003")
        
        ' Add some whitespace for formatting
        Dim textNode As Object
        Set textNode = xmlDoc.createTextNode(vbCrLf & "  ")
        projectNode.insertBefore textNode, Nothing
        
        projectNode.appendChild itemGroup
        
        Set textNode = xmlDoc.createTextNode(vbCrLf & vbCrLf & "  ")
        projectNode.appendChild textNode
        
        Set FindOrCreateBuildItemGroup = itemGroup
    End If
End Function
'
' Helper function to check if folder already exists in project
Private Function FolderExistsInProject(xmlDoc As Object, folderPath As String) As Boolean
    Dim folderNodes As Object
    Dim query As String
    
    query = "//msb:Folder[@Include='" & folderPath & "']"
    Set folderNodes = xmlDoc.selectNodes(query)
    
    FolderExistsInProject = (folderNodes.Length > 0)
End Function
'
' Helper function to check if file already exists in project
Private Function FileExistsInProject(xmlDoc As Object, filePath As String) As Boolean
    Dim fileNodes As Object
    Dim query As String
    
    ' Normalize path separators for comparison
    filePath = Replace(filePath, "/", "\")
    
    ' Check in all possible build action types
    query = "//msb:Build[@Include='" & filePath & "'] | " & _
            "//msb:None[@Include='" & filePath & "'] | " & _
            "//msb:PostDeploy[@Include='" & filePath & "'] | " & _
            "//msb:PreDeploy[@Include='" & filePath & "']"
    
    Set fileNodes = xmlDoc.selectNodes(query)
    
    FileExistsInProject = (fileNodes.Length > 0)
End Function

' Remove a folder and all its contents (subfolders and files) from SQL Server Database Project
Public Function RemoveFolderFromSqlProj(Optional is_debugging As Boolean = False) As Boolean: On Error GoTo ErrorHandler

  Dim fso                    As FileSystemObject: Set fso = New FileSystemObject
  Dim tx_search_file         As String:     tx_search_file = "<Build Include=""3-Datasets\"
  Dim tx_search_folder       As String:     tx_search_folder = "<Folder Include=""3-Datasets\"
  Dim tx_search_end_tag      As String:     tx_search_end_tag = "/>"
  
  Dim fp_xml_project_current As String:     fp_xml_project_current = fp_sql_project()
  Dim fp_xml_project_replace As String:     fp_xml_project_replace = fp_sql_project() & ".txt"
  
  '
  ' Delete Existing Files
  If (fso.FileExists(fp_xml_project_current & ".old")) Then fso.DeleteFile fp_xml_project_current & ".old", True
  If (fso.FileExists(fp_xml_project_current & ".txt")) Then fso.DeleteFile fp_xml_project_current & ".txt", True
  '
  ' Open ProjectFile and Start Replacement File
  Dim tx_xml_project_current As TextStream: Set tx_xml_project_current = fso.OpenTextFile(fp_xml_project_current, ForReading, False, TristateMixed)
  Dim tx_xml_project_replace As TextStream: Set tx_xml_project_replace = fso.OpenTextFile(fp_xml_project_replace, ForWriting, True, TristateMixed)
  Dim tx_line                As String
  '
  ' Position Info
  Dim ni_tag_begin As Integer
  Dim ni_tag_end   As Integer
  Dim is_tag_found As Boolean
  '
  Do While Not tx_xml_project_current.AtEndOfStream
    '
    tx_line = tx_xml_project_current.ReadLine
    '
    ' Find Tag Begin
    is_tag_found = False
    If (is_tag_found = False And InStr(1, tx_line, tx_search_folder, vbTextCompare) <> 0) Then is_tag_found = True
    If (is_tag_found = False And InStr(1, tx_line, tx_search_file, vbTextCompare) <> 0) Then is_tag_found = True
    '
    ' If tag is NOT found copy the line to the replacement file
    If (Not is_tag_found) Then tx_xml_project_replace.WriteLine tx_line
    '
  Loop
  tx_xml_project_replace.Close
  tx_xml_project_current.Close
  '
  ' Rename Current Project-file to Old
  Call fso.MoveFile(fp_xml_project_current, fp_xml_project_current & ".old")
  Call fso.MoveFile(fp_xml_project_replace, fp_xml_project_current)
  '
  ' Check if file the moved proper, delete the "old"
  If (fso.FileExists(fp_xml_project_current)) Then fso.DeleteFile fp_xml_project_current & ".old", True
  '
  RemoveFolderFromSqlProj = True
  '
Exit Function
ErrorHandler:
  Debug.Print "Error in RemoveFolderFromSqlProj: " & Err.Number & " - " & Err.Description
  Stop
  Resume
    RemoveFolderFromSqlProj = False
End Function

' Helper function to remove all items (files and folders) starting with a specific path
Private Function RemoveItemsStartingWith(xmlDoc As Object, folderPath As String, is_debugging As Boolean) As Integer
  Dim itemsRemoved As Integer: itemsRemoved = 0
  Dim allNodes As Object
  Dim node As Object
  Dim includePath As String

  ' Query for all items with Include attribute (Build, None, Folder, PostDeploy, PreDeploy)
  Dim query As String
  
  query = "//msb:Build[@Include] | " & _
          "//msb:None[@Include] | " & _
          "//msb:Folder[@Include] | " & _
          "//msb:PostDeploy[@Include] | " & _
          "//msb:PreDeploy[@Include]"
    
    Set allNodes = xmlDoc.selectNodes(query)
    
    ' Iterate backwards to safely remove nodes
    Dim i As Integer
  For i = allNodes.Length - 1 To 0 Step -1
        Set node = allNodes(i)
        includePath = node.getAttribute("Include")

    ' Normalize path
    includePath = Replace(includePath, "/", "\")

    ' Check if this item is under the folder being removed
    If Left(includePath, Len(folderPath)) = folderPath Or includePath = Left(folderPath, Len(folderPath) - 1) Then
      If is_debugging Then Debug.Print "Removing: " & node.nodeName & " - " & includePath
            node.parentNode.removeChild node
            itemsRemoved = itemsRemoved + 1
    End If
  Next i

  RemoveItemsStartingWith = itemsRemoved
End Function

' Remove a specific file from SQL Server Database Project
Public Function RemoveSqlFileFromSqlProj(sqlFilePath As String, Optional is_debugging As Boolean = False) As Boolean
  On Error GoTo ErrorHandler

  Dim sqlProjPath As String: sqlProjPath = fp_sql_project()
  Dim fso As FileSystemObject:   Set fso = New FileSystemObject
    Dim xmlDoc As Object

  ' Check if sqlproj file exists
  If Not fso.FileExists(sqlProjPath) Then
    Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
        RemoveSqlFileFromSqlProj = False
    Exit Function
  End If
    
    ' Load XML document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"

    If Not xmlDoc.Load(sqlProjPath) Then
    Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
        RemoveSqlFileFromSqlProj = False
    Exit Function
  End If

  ' Normalize path
  sqlFilePath = Replace(sqlFilePath, "/", "\")

  ' Query for the file in all possible build action types
  Dim query As String
  query = "//msb:Build[@Include='" & sqlFilePath & "'] | " & _
          "//msb:None[@Include='" & sqlFilePath & "'] | " & _
          "//msb:PostDeploy[@Include='" & sqlFilePath & "'] | " & _
          "//msb:PreDeploy[@Include='" & sqlFilePath & "']"

  Dim fileNodes As Object
    Set fileNodes = xmlDoc.selectNodes(query)
    
    If fileNodes.Length > 0 Then
    Dim i As Integer
    For i = fileNodes.Length - 1 To 0 Step -1
      fileNodes(i).parentNode.removeChild fileNodes(i)
        Next i

    ' Save the modified XML
    xmlDoc.save sqlProjPath

        If is_debugging Then Debug.Print "Successfully removed file '" & sqlFilePath & "' from project"
        RemoveSqlFileFromSqlProj = True
  Else
    If is_debugging Then Debug.Print "File '" & sqlFilePath & "' not found in project"
        RemoveSqlFileFromSqlProj = True ' Consider it success if file doesn't exist
  End If
    
    Set xmlDoc = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
  Debug.Print "Error in RemoveSqlFileFromSqlProj: " & Err.Number & " - " & Err.Description
    RemoveSqlFileFromSqlProj = False
End Function

' Remove folder from both filesystem and project
Public Function DeleteFolderAndRemoveFromProject(folderPath As String, Optional is_debugging As Boolean = False) As Boolean: On Error GoTo ErrorHandler

  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim fullFolderPath As String: fullFolderPath = tx_repo_folderpath(nm_repository()) & IIf(Left(folderPath, 2) = ".\", Mid(folderPath, 3, Len(folderPath) - 2), folderPath)

  ' Remove from project first
  If Not RemoveFolderFromSqlProj(is_debugging) Then
    Debug.Print "Warning: Could not remove folder from project"
    End If

  ' Delete from filesystem
  If fso.FolderExists(fullFolderPath) Then
    fso.DeleteFolder fullFolderPath, Force:=True
        If is_debugging Then Debug.Print "Deleted folder from filesystem: " & fullFolderPath
    Else
    If is_debugging Then Debug.Print "Folder does not exist on filesystem: " & fullFolderPath
    End If

  DeleteFolderAndRemoveFromProject = True
    
    Set fso = Nothing
    Exit Function

ErrorHandler:
  Debug.Print "Error in DeleteFolderAndRemoveFromProject: " & Err.Number & " - " & Err.Description
    DeleteFolderAndRemoveFromProject = False
    Set fso = Nothing
End Function

' Recursively scan folder structure and store all folders/files in hlp_folder_and_files table
Public Sub scan_folder_structure(ip_fp_folder As String, Optional ip_is_debugging As Boolean = False): On Error GoTo ErrorHandler
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim flr As folder
    Dim fp_project As String: fp_project = Replace(tx_repo_folderpath(nm_repository()), "\2-Definitions\", "")

    ' Clear the helper table before starting
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM hlp_folders_and_files"
    DoCmd.SetWarnings True
        
    ' Validate input folder path
    If Not fso.FolderExists(ip_fp_folder) Then
        If ip_is_debugging Then Debug.Print "Error: Folder does not exist: " & ip_fp_folder
        Exit Sub
    End If
    
    ' Get folder object
    Set flr = fso.GetFolder(ip_fp_folder)
    
    If ip_is_debugging Then Debug.Print "Starting recursive scan of: " & ip_fp_folder
    
    ' Start recursive processing
    Call process_folder_recursive(flr, fp_project, ip_is_debugging)
    
    If ip_is_debugging Then Debug.Print "Folder scan complete"
    
    Set flr = Nothing
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "Error in scan_folder_structure: " & Err.Number & " - " & Err.Description
    Set flr = Nothing
    Set fso = Nothing
End Sub
Public Sub test_scan_folder_structure()
  scan_folder_structure "D:\git\Misset-Data-Analytics\My-Financial-Stock-Information\my-stock-info\2-meta-data-definitions\3-Datasets", True
End Sub

' Helper function to recursively process folders and files
Private Sub process_folder_recursive(flr As folder, fp_project As String, is_debugging As Boolean): On Error GoTo ErrorHandler
    
    Dim subFolder As folder
    Dim file As file
    Dim rst As Recordset
    Dim fp_relative As String
    Dim fp_full As String
    
    ' Process current folder
    fp_full = flr.Path
    fp_relative = Mid(fp_full, Len(fp_project) + 2)
    
    ' Insert folder record
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM hlp_folders_and_files WHERE 1=2")
    rst.AddNew
    rst!cd_type = "folder"
    rst!fp_relative = fp_relative
    rst!fp_full = fp_full
    rst!ni_depth = (Len(fp_relative) - Len(Replace(fp_relative, "\", "")))
    rst.Update
    rst.Close
    
    If is_debugging Then Debug.Print "Folder: " & fp_relative
    
    ' Process all files in current folder
    For Each file In flr.Files
        fp_full = file.Path
        fp_relative = Mid(fp_full, Len(fp_project) + 2)
        
        Set rst = CurrentDb.OpenRecordset("SELECT * FROM hlp_folders_and_files WHERE 1=2")
        rst.AddNew
        rst!cd_type = "file"
        rst!fp_relative = fp_relative
        rst!ni_depth = (Len(fp_relative) - Len(Replace(fp_relative, "\", "")))
        rst!fp_full = fp_full
        rst.Update
        rst.Close
        
        If is_debugging Then Debug.Print "  File: " & fp_relative
    Next file
    
    ' Recursively process all subfolders
    For Each subFolder In flr.SubFolders
        Call process_folder_recursive(subFolder, fp_project, is_debugging)
    Next subFolder
    
    Exit Sub

ErrorHandler:
    Debug.Print "Error in process_folder_recursive: " & Err.Number & " - " & Err.Description
End Sub

' Exclude a specific file from SQL Server Database Project (remove from project but keep file on disk)
Public Function ExcludeSqlFileFromSqlProj(sqlFilePath As String, Optional is_debugging As Boolean = False) As Boolean: On Error GoTo ErrorHandler

  Dim sqlProjPath As String: sqlProjPath = fp_sql_project()
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim xmlDoc As Object

  ' Check if sqlproj file exists
  If Not fso.FileExists(sqlProjPath) Then
    Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
    ExcludeSqlFileFromSqlProj = False
    Exit Function
  End If
    
  ' Load XML document
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.Async = False
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"

  If Not xmlDoc.Load(sqlProjPath) Then
    Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
    ExcludeSqlFileFromSqlProj = False
    Exit Function
  End If

  ' Normalize path
  sqlFilePath = Replace(sqlFilePath, "/", "\")

  ' Query for the file in all possible build action types
  Dim query As String
  query = "//msb:Build[@Include='" & sqlFilePath & "'] | " & _
          "//msb:None[@Include='" & sqlFilePath & "'] | " & _
          "//msb:PostDeploy[@Include='" & sqlFilePath & "'] | " & _
          "//msb:PreDeploy[@Include='" & sqlFilePath & "']"

  Dim fileNodes As Object
  Set fileNodes = xmlDoc.selectNodes(query)
    
  If fileNodes.Length > 0 Then
    Dim i As Integer
    For i = fileNodes.Length - 1 To 0 Step -1
      fileNodes(i).parentNode.removeChild fileNodes(i)
    Next i

    ' Save the modified XML
    xmlDoc.save sqlProjPath

    If is_debugging Then Debug.Print "Successfully excluded file '" & sqlFilePath & "' from project (file kept on disk)"
    ExcludeSqlFileFromSqlProj = True
  Else
    If is_debugging Then Debug.Print "File '" & sqlFilePath & "' not found in project"
    ExcludeSqlFileFromSqlProj = True ' Consider it success if file doesn't exist in project
  End If
    
  Set xmlDoc = Nothing
  Set fso = Nothing
  Exit Function

ErrorHandler:
  Debug.Print "Error in ExcludeSqlFileFromSqlProj: " & Err.Number & " - " & Err.Description
  ExcludeSqlFileFromSqlProj = False
End Function
Public Sub test_ExcludeSqlFileFromSqlProj()
  
  Dim fp_file As String: fp_file = "3-Datasets\1-Ingestions\psa_yahoo_stock_info\Tables\nio.sql"

  If ExcludeSqlFileFromSqlProj(fp_file, True) Then
      Debug.Print "File successfully excluded from project"
  Else
      Debug.Print "File failed excluded from project"
  End If

End Sub

' Exclude a folder and all its contents from SQL Server Database Project (keep folder on disk)
Public Function ExcludeFolderFromSqlProj(folderPath As String, Optional is_debugging As Boolean = False) As Boolean
  On Error GoTo ErrorHandler
  
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  Dim sqlProjPath As String: sqlProjPath = fp_sql_project()
  Dim itemsRemoved As Integer: itemsRemoved = 0

  ' Check if sqlproj file exists
  If Not fso.FileExists(sqlProjPath) Then
    Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
    ExcludeFolderFromSqlProj = False
    Exit Function
  End If
    
  ' Load XML document
  xmlDoc.Async = False
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"

  If Not xmlDoc.Load(sqlProjPath) Then
    Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
    ExcludeFolderFromSqlProj = False
    Exit Function
  End If

  ' Normalize folder path
  folderPath = Replace(folderPath, "/", "\")
  If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

  ' Remove all files and subfolders that start with this folder path
  itemsRemoved = itemsRemoved + RemoveItemsStartingWith(xmlDoc, folderPath, is_debugging)

  ' Remove the folder itself
  Dim folderPathWithoutSlash As String: folderPathWithoutSlash = Left(folderPath, Len(folderPath) - 1)
  Dim query As String: query = "//msb:Folder[@Include='" & folderPathWithoutSlash & "']"
  Dim folderNodes As Object: Set folderNodes = xmlDoc.selectNodes(query)
    
  If folderNodes.Length > 0 Then
    Dim i As Integer
    For i = folderNodes.Length - 1 To 0 Step -1
      folderNodes(i).parentNode.removeChild folderNodes(i)
      itemsRemoved = itemsRemoved + 1
      If is_debugging Then Debug.Print "Excluded folder: " & folderPathWithoutSlash
    Next i
  End If

  ' Save the modified XML back to file
  If itemsRemoved > 0 Then
    xmlDoc.save sqlProjPath
    If is_debugging Then Debug.Print "Successfully excluded " & itemsRemoved & " items from project for folder: " & folderPath & " (folder kept on disk)"
    ExcludeFolderFromSqlProj = True
  Else
    If is_debugging Then Debug.Print "No items found to exclude for folder: " & folderPath
    ExcludeFolderFromSqlProj = True ' Consider it success if nothing to exclude
  End If
    
  Set xmlDoc = Nothing
  Set fso = Nothing
  Exit Function

ErrorHandler:
  Debug.Print "Error in ExcludeFolderFromSqlProj: " & Err.Number & " - " & Err.Description
  ExcludeFolderFromSqlProj = False
End Function

Public Sub ExcludeAllFilesAndFolersFrom_3_Datasets(Optional ip_is_debugging As Boolean = False)

  scan_folder_structure "D:\git\Misset-Data-Analytics\My-Financial-Stock-Information\my-stock-info\2-meta-data-definitions\3-Datasets\", ip_is_debugging
  
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim sql As String
  Dim rst As Recordset
  
  sql = "SELECT fp_relative, fp_full FROM hlp_folders_and_files WHERE cd_type='file' ORDER BY ni_depth DESC"
  Set rst = CurrentDb.OpenRecordset(sql): Do Until rst.EOF
    If (ExcludeSqlFileFromSqlProj(rst!fp_relative, ip_is_debugging) = True) Then
      Call fso.DeleteFile(rst!fp_full, True)
    End If
    rst.MoveNext
  Loop
  
  sql = "SELECT fp_relative, fp_full FROM hlp_folders_and_files WHERE cd_type='folder' ORDER BY ni_depth DESC"
  Set rst = CurrentDb.OpenRecordset(sql): Do Until rst.EOF
    If (ExcludeFolderFromSqlProj(rst!fp_relative, ip_is_debugging) = True) Then
      Call fso.DeleteFolder(rst!fp_full, True)
    End If
    rst.MoveNext
    
  Loop
  
End Sub

' Remove all folder and file references from SQL project that no longer exist on filesystem
Public Function RemoveNonExistingReferencesFromSqlProj(Optional is_debugging As Boolean = False) As Boolean
  On Error GoTo ErrorHandler
  
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  Dim sqlProjPath As String: sqlProjPath = fp_sql_project()
  Dim projectBasePath As String: projectBasePath = Replace(tx_repo_folderpath(nm_repository()), "\2-Definitions\", "") & "\"
  Dim itemsRemoved As Integer: itemsRemoved = 0
  
  ' Check if sqlproj file exists
  If Not fso.FileExists(sqlProjPath) Then
    Debug.Print "Error: .sqlproj file not found at: " & sqlProjPath
    RemoveNonExistingReferencesFromSqlProj = False
    Exit Function
  End If
    
  ' Load XML document
  xmlDoc.Async = False
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:msb='http://schemas.microsoft.com/developer/msbuild/2003'"

  If Not xmlDoc.Load(sqlProjPath) Then
    Debug.Print "Error parsing XML: " & xmlDoc.parseError.reason
    RemoveNonExistingReferencesFromSqlProj = False
    Exit Function
  End If

  If is_debugging Then Debug.Print "Starting cleanup of non-existing references..."
  
  ' First pass: Remove non-existing files
  Dim fileQuery As String
  fileQuery = "//msb:Build[@Include] | " & _
              "//msb:None[@Include] | " & _
              "//msb:PostDeploy[@Include] | " & _
              "//msb:PreDeploy[@Include]"
    
  Dim fileNodes As Object: Set fileNodes = xmlDoc.selectNodes(fileQuery)
  Dim node As Object
  Dim includePath As String
  Dim fullPath As String
  Dim i As Integer
  
  If is_debugging Then Debug.Print "Checking " & fileNodes.Length & " file references..."
  
  ' Iterate backwards to safely remove nodes
  For i = fileNodes.Length - 1 To 0 Step -1
    Set node = fileNodes(i)
    includePath = node.getAttribute("Include")
    
    ' Normalize path
    includePath = Replace(includePath, "/", "\")
    
    ' Build full path (handle relative paths)
    If Left(includePath, 2) = ".\" Then
      fullPath = projectBasePath & Mid(includePath, 3)
    Else
      fullPath = projectBasePath & includePath
    End If
    
    ' Normalize path separators
    fullPath = Replace(fullPath, "\\", "\")
    
    ' Check if file exists
    If Not fso.FileExists(fullPath) Then
      If is_debugging Then
        Debug.Print "Removing non-existing file: " & includePath
        Debug.Print "Fullpath non Existing     : " & fullPath
      End If
      node.parentNode.removeChild node
      itemsRemoved = itemsRemoved + 1
    End If
  Next i
  
  ' Second pass: Remove non-existing folders
  Dim folderQuery As String
  folderQuery = "//msb:Folder[@Include]"
    
  Dim folderNodes As Object: Set folderNodes = xmlDoc.selectNodes(folderQuery)
  
  If is_debugging Then Debug.Print "Checking " & folderNodes.Length & " folder references..."
  
  ' Iterate backwards to safely remove nodes
  For i = folderNodes.Length - 1 To 0 Step -1
    Set node = folderNodes(i)
    includePath = node.getAttribute("Include")
    
    ' Normalize path
    includePath = Replace(includePath, "/", "\")
    
    ' Build full path (handle relative paths)
    If Left(includePath, 2) = ".\" Then
      fullPath = projectBasePath & Mid(includePath, 3)
    Else
      fullPath = projectBasePath & includePath
    End If
    
    ' Normalize path separators
    fullPath = Replace(fullPath, "\\", "\")
    
    ' Check if folder exists
    If Not fso.FolderExists(fullPath) Then
      If is_debugging Then
        Debug.Print "Removing non-existing folder: " & includePath
        Debug.Print "Fullpath non existing folder: " & fullPath
      End If
      node.parentNode.removeChild node
      itemsRemoved = itemsRemoved + 1
    End If
  Next i

  ' Save the modified XML back to file if changes were made
  If itemsRemoved > 0 Then
    xmlDoc.save sqlProjPath
    If is_debugging Then Debug.Print "Successfully removed " & itemsRemoved & " non-existing references from project"
    RemoveNonExistingReferencesFromSqlProj = True
  Else
    If is_debugging Then Debug.Print "No non-existing references found in project"
    RemoveNonExistingReferencesFromSqlProj = True
  End If
    
  Set xmlDoc = Nothing
  Set fso = Nothing
  Exit Function

ErrorHandler:
  Debug.Print "Error in RemoveNonExistingReferencesFromSqlProj: " & Err.Number & " - " & Err.Description
  RemoveNonExistingReferencesFromSqlProj = False
End Function