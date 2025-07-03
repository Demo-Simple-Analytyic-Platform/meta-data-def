Attribute VB_Name = "mdl_Import_Export_Code"
Option Compare Database
Option Explicit
'
' Private Module Variables
Private fp_exported_code    As String
Private fp_exported_modules As String
Private fp_exported_forms   As String
Private fp_exported_reports As String
Private fp_exported_macros  As String
Private fp_exported_tables  As String
Private fp_exported_queries As String

Sub set_folder_paths()
    '
    ' Local Variables
    fp_exported_code = CurrentProject.Path & "\Code"
    fp_exported_modules = fp_exported_code & "\Modules\"
    fp_exported_forms = fp_exported_code & "\Forms\"
    fp_exported_reports = fp_exported_code & "\Reports\"
    fp_exported_macros = fp_exported_code & "\Macros\"
    fp_exported_tables = fp_exported_code & "\Tables\"
    fp_exported_queries = fp_exported_code & "\Queries\"
    '
End Sub

Public Sub ExportAllCodeModules()
    '
    ' Set Folder paths
    Call set_folder_paths
    '
    ' Create FileSystemObject to handle folder creation
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim obj As Object
    
    If fso.FolderExists(fp_exported_code) = False Then fso.CreateFolder fp_exported_code
    If fso.FolderExists(fp_exported_modules) = False Then fso.CreateFolder fp_exported_modules
    If fso.FolderExists(fp_exported_forms) = False Then fso.CreateFolder fp_exported_forms
    If fso.FolderExists(fp_exported_reports) = False Then fso.CreateFolder fp_exported_reports
    If fso.FolderExists(fp_exported_macros) = False Then fso.CreateFolder fp_exported_macros
    If fso.FolderExists(fp_exported_queries) = False Then fso.CreateFolder fp_exported_queries
    If fso.FolderExists(fp_exported_tables) = False Then fso.CreateFolder fp_exported_tables
    
    For Each obj In Application.CurrentProject.AllModules
        Application.SaveAsText acModule, obj.Name, fp_exported_modules & obj.Name & ".bas"
        Debug.Print "Module/Class definition for '" & obj.Name & "' exported to: " & fp_exported_modules & obj.Name & ".bas"
    Next

    For Each obj In Application.CurrentProject.AllForms
        Application.SaveAsText acForm, obj.Name, fp_exported_forms & obj.Name & ".frm"
        Debug.Print "Form definition for '" & obj.Name & "' exported to: " & fp_exported_forms & obj.Name & ".frm"
    Next

    For Each obj In Application.CurrentProject.AllReports
        Application.SaveAsText acReport, obj.Name, fp_exported_reports & obj.Name & ".frm"
        Debug.Print "Report definition for '" & obj.Name & "' exported to: " & fp_exported_reports & obj.Name & ".bas"
    Next

    For Each obj In Application.CurrentProject.AllMacros
        Application.SaveAsText acMacro, obj.Name, fp_exported_macros & obj.Name & ".bas"
        Debug.Print "Marco definition for '" & obj.Name & "' exported to: " & fp_exported_macros & obj.Name & ".bas"
    Next

    For Each obj In Application.CurrentDb.TableDefs
        If Left(obj.Name, 4) <> "MSys" And Left(obj.Name, 1) <> "~" Then
            Call ExportSingleTableDef(obj.Name)
        End If
    Next

    For Each obj In Application.CurrentDb.QueryDefs
        If Left(obj.Name, 4) <> "MSys" And Left(obj.Name, 1) <> "~" Then
            Call ExportSingleQueryDef(obj.Name)
        End If
    Next
    
    Debug.Print "Exported all code modules, forms, reports, and macros to the following directories:"

End Sub

Public Function ImportAllCodeModules() As Boolean
    '
    ' Set Folder paths
    Call set_folder_paths
    '
    ' Create FileSystemObject to handle folder creation
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fil As file
    '
    ' Import Modules, Forms, Reports and Marcos
    For Each fil In fso.GetFolder(fp_exported_modules).Files
        '
        ' Exclude the Import/Export Code module itself
        If fil.name <> "mdl_Import_Export_Code" then Application.LoadFromText acModule, Mid(fil.Name, 1, Len(fil.Name) - 4), fil.Path
        '
    Next fil
    For Each fil In fso.GetFolder(fp_exported_forms).Files: Application.LoadFromText acForm, Mid(fil.Name, 1, Len(fil.Name) - 4), fil.Path: Next fil
    For Each fil In fso.GetFolder(fp_exported_reports).Files: Application.LoadFromText acReport, Mid(fil.Name, 1, Len(fil.Name) - 4), fil.Path: Next fil
    For Each fil In fso.GetFolder(fp_exported_macros).Files: Application.LoadFromText acMacro, Mid(fil.Name, 1, Len(fil.Name) - 4), fil.Path: Next fil
    For Each fil In fso.GetFolder(fp_exported_tables).Files: ImportSingleTableDef Mid(fil.Name, 1, Len(fil.Name) - 4): Next fil
    For Each fil In fso.GetFolder(fp_exported_queries).Files: ImportSingleQueryDef Mid(fil.Name, 1, Len(fil.Name) - 4): Next fil
    '
    ' All Done
    Debug.Print "All code has been imported!"
    '
    ' Open the "Start Up"-form
    DoCmd.OpenForm "StartUp"
    '
End Function


' Function to generate CREATE TABLE SQL statement
Private Function GenerateCreateTableSQL(tdf As DAO.TableDef) As String
    On Error GoTo ErrorHandler
    
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim strSQL As String
    Dim strFields As String
    Dim strIndexes As String
    Dim strDataType As String
    
    ' Start building the CREATE TABLE statement
    strSQL = "-- Table: " & tdf.Name & vbCrLf
    strSQL = strSQL & "-- Created: " & Format(Now(), "yyyy-mm-dd hh:nn:ss") & vbCrLf
    strSQL = strSQL & "-- Records: " & tdf.RecordCount & vbCrLf & vbCrLf
    
    strSQL = strSQL & "CREATE TABLE [" & tdf.Name & "] (" & vbCrLf
    
    ' Loop through fields
    For Each fld In tdf.fields
        If strFields <> "" Then strFields = strFields & "," & vbCrLf
        
        ' Get field data type
        strDataType = GetAccessDataType(fld)
        
        strFields = strFields & "    [" & fld.Name & "] " & strDataType
        
        ' Add field properties
        If fld.Required And Not fld.AllowZeroLength Then
            strFields = strFields & " NOT NULL"
        End If
        
        If fld.DefaultValue <> "" Then
            strFields = strFields & " DEFAULT " & fld.DefaultValue
        End If
        
    Next fld
    
    strSQL = strSQL & strFields & vbCrLf & ");" & vbCrLf & vbCrLf
    
    ' Add indexes
    For Each idx In tdf.Indexes
        If idx.Primary Then
            strSQL = strSQL & "-- Primary Key: " & idx.Name & vbCrLf
            strSQL = strSQL & "ALTER TABLE [" & tdf.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY ("
        ElseIf idx.Unique Then
            strSQL = strSQL & "-- Unique Index: " & idx.Name & vbCrLf
            strSQL = strSQL & "CREATE UNIQUE INDEX [" & idx.Name & "] ON [" & tdf.Name & "] ("
        Else
            strSQL = strSQL & "-- Index: " & idx.Name & vbCrLf
            strSQL = strSQL & "CREATE INDEX [" & idx.Name & "] ON [" & tdf.Name & "] ("
        End If
        
        Dim fldIndex As DAO.Field
        Dim strIndexFields As String
        For Each fldIndex In idx.fields
            If strIndexFields <> "" Then strIndexFields = strIndexFields & ", "
            strIndexFields = strIndexFields & "[" & fldIndex.Name & "]"
        Next fldIndex
        
        strSQL = strSQL & strIndexFields & ");" & vbCrLf & vbCrLf
    Next idx
    
    GenerateCreateTableSQL = strSQL
    Exit Function

ErrorHandler:
    Debug.Print strSQL
    GenerateCreateTableSQL = "-- Error generating SQL for table: " & tdf.Name & vbCrLf & _
                              "-- Error: " & Err.Number & " - " & Err.Description
End Function

' Function to convert Access data types to SQL data types
Private Function GetAccessDataType(fld As DAO.Field) As String
    Dim strDataType As String
    
    Select Case fld.Type
        Case dbBoolean
            strDataType = "BIT"
        Case dbByte
            strDataType = "TINYINT"
        Case dbInteger
            strDataType = "SMALLINT"
        Case dbLong
            If (fld.Attributes And dbAutoIncrField) Then
                strDataType = "IDENTITY(1,1) INT"
            Else
                strDataType = "INT"
            End If
        Case dbCurrency
            strDataType = "MONEY"
        Case dbSingle
            strDataType = "REAL"
        Case dbDouble
            strDataType = "FLOAT"
        Case dbDate
            strDataType = "DATETIME"
        Case dbText
            If fld.Size > 0 Then
                strDataType = "VARCHAR(" & fld.Size & ")"
            Else
                strDataType = "VARCHAR(255)"
            End If
        Case dbMemo
            strDataType = "TEXT"
        Case dbLongBinary
            strDataType = "IMAGE"
        Case dbGUID
            strDataType = "UNIQUEIDENTIFIER"
        Case Else
            strDataType = "VARCHAR(255)" ' Default fallback
    End Select
    
    GetAccessDataType = strDataType
End Function

' Alternative function to export table structure to CSV format
Public Sub ExportTableDefsToCSV()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strExportPath As String
    Dim strFileName As String
    Dim intFileNum As Integer
    
    Set db = CurrentDb()
    strExportPath = CurrentProject.Path & "\TableStructures\"
    
    ' Create directory if it doesn't exist
    If Dir(strExportPath, vbDirectory) = "" Then
        MkDir strExportPath
    End If
    
    ' Create a master file with all table structures
    strFileName = strExportPath & "AllTableStructures.csv"
    intFileNum = FreeFile
    Open strFileName For Output As #intFileNum
    
    ' Write header
    Print #intFileNum, "TableName,FieldName,DataType,Size,Required,DefaultValue,Description"
    
    ' Loop through all tables
    For Each tdf In db.TableDefs
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" And Left(tdf.Name, 4) <> "tmp" Then
            For Each fld In tdf.fields
                Print #intFileNum, """" & tdf.Name & """,""" & fld.Name & """,""" & _
                    GetAccessDataType(fld) & """," & fld.Size & "," & fld.Required & _
                    ",""" & Nz(fld.DefaultValue, "") & """,""" & Nz(fld.Properties("Description"), "") & """"
            Next fld
        End If
    Next tdf
    
    Close #intFileNum
    
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    
    MsgBox "Table structures exported to CSV: " & strFileName, vbInformation
    Exit Sub

ErrorHandler:
    If intFileNum > 0 Then Close #intFileNum
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Function to export a single table definition
Public Sub ExportSingleTableDef(ip_nm_table As String): On Error GoTo ErrorHandler
    '
    ' Local Variables
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim dbs As DAO.Database:     Set dbs = CurrentDb
    Dim tdf As DAO.TableDef:     Set tdf = dbs.TableDefs(ip_nm_table)
    Dim txt As TextStream:       Set txt = fso.CreateTextFile(fp_exported_tables & ip_nm_table & ".sql", True, True)
    Dim sql As String:               sql = GenerateCreateTableSQL(tdf)
    '
    ' Write SQL to TextStream
    txt.Write sql
    '
    ' Save/Close File
    txt.Close
    '
    ' All Done
    Debug.Print "Table definition for '" & ip_nm_table & "' exported to: " & fp_exported_tables
    '
Exit Sub
'
ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    '
End Sub
'
'
' Function to import a single table definition from SQL file
Public Function ImportSingleTableDef(strTableName As String) As Boolean
On Error GoTo ErrorHandler
    '
    ' Validate Fp (Folderpath) is NOT empty
    If (fp_exported_tables = "") Then set_folder_paths
    
    Dim strSQL As String: strSQL = ReadSQLFile(fp_exported_tables & strTableName & ".sql")
    
    
    Dim db As DAO.Database
    Dim strCreateSQL As String
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb()
    
    ' Read the SQL file
    
    
    If strSQL = "" Then
        Debug.Print "Error: Could not read SQL file for table " & strTableName
        ImportSingleTableDef = False
        Exit Function
    End If
    
    ' Check if table already exists
    If TableExists(strTableName) Then
        If MsgBox("Table '" & strTableName & "' already exists. Do you want to replace it?", _
                  vbYesNo + vbQuestion) = vbYes Then
            db.TableDefs.Delete strTableName
        Else
            ImportSingleTableDef = False
            Exit Function
        End If
    End If
    
    ' Parse and execute the CREATE TABLE statement
    strCreateSQL = ExtractCreateTableSQL(strSQL)
    
    If strCreateSQL <> "" Then
        ' Convert SQL Server syntax to Access syntax
        strCreateSQL = ConvertSQLServerToAccess(strCreateSQL)
        
        ' Execute the CREATE TABLE statement
        db.Execute strCreateSQL, dbFailOnError
        
        ' Add indexes separately (after table creation)
        AddIndexesFromSQL strTableName, strSQL
        
        ImportSingleTableDef = True
        Debug.Print "Successfully created table: " & strTableName
    Else
        Debug.Print "Error: Could not extract CREATE TABLE statement for " & strTableName
        ImportSingleTableDef = False
    End If
    
    Set db = Nothing
    Exit Function

ErrorHandler:
    Set db = Nothing
    Debug.Print "Error creating table " & strTableName & ": " & Err.Number & " - " & Err.Description & vbNewLine & "SQL:" & vbNewLine & strCreateSQL
    ImportSingleTableDef = False
End Function

' Function to read SQL file content using FileSystemObject
Private Function ReadSQLFile(strFilePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim txtFile As Object
    Dim strContent As String
    
    ' Check if file exists
    If Not fso.FileExists(strFilePath) Then
        Debug.Print "File not found: " & strFilePath
        ReadSQLFile = ""
        Exit Function
    End If
    
    ' Open and read the file
    Set txtFile = fso.OpenTextFile(strFilePath, IOMode.ForReading, False, TristateTrue)
    
    strContent = txtFile.ReadAll
    txtFile.Close
    
    ReadSQLFile = strContent
    
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    Debug.Print "Error reading file " & strFilePath & ": " & Err.Description
    ReadSQLFile = ""
End Function

' Function to extract CREATE TABLE statement from SQL content
Private Function ExtractCreateTableSQL(strSQL As String) As String
    On Error GoTo ErrorHandler
    
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim strCreateSQL As String
    
    ' Find the CREATE TABLE statement
    intStart = InStr(UCase(strSQL), "CREATE TABLE")
    
    If intStart > 0 Then
        ' Find the end of the CREATE TABLE statement (look for ");")
        intEnd = InStr(intStart, strSQL, ");")
        
        If intEnd > 0 Then
            strCreateSQL = Mid(strSQL, intStart, intEnd - intStart + 2)
            ExtractCreateTableSQL = strCreateSQL
        End If
    End If
    
    Exit Function

ErrorHandler:
    ExtractCreateTableSQL = ""
End Function

' Function to convert SQL Server syntax to Access syntax
Private Function ConvertSQLServerToAccess(strSQL As String) As String
    On Error GoTo ErrorHandler
    
    Dim strResult As String
    strResult = strSQL
    
    ' Convert SQL Server data types to Access equivalents
    strResult = Replace(strResult, "IDENTITY(1,1) INT", "AUTOINCREMENT")
    strResult = Replace(strResult, "IDENTITY(1,1)", "AUTOINCREMENT")
    strResult = Replace(strResult, "BIT", "YESNO")
    strResult = Replace(strResult, "TINYINT", "BYTE")
    strResult = Replace(strResult, "SMALLINT", "INTEGER")
    strResult = Replace(strResult, "INT", "LONG")
    strResult = Replace(strResult, "MONEY", "CURRENCY")
    strResult = Replace(strResult, "REAL", "SINGLE")
    strResult = Replace(strResult, "FLOAT", "DOUBLE")
    strResult = Replace(strResult, "DATETIME", "DATETIME")
    strResult = Replace(strResult, "TEXT", "MEMO")
    strResult = Replace(strResult, "VARCHAR", "TEXT")
    strResult = Replace(strResult, "IMAGE", "OLEOBJECT")
    strResult = Replace(strResult, "UNIQUEIDENTIFIER", "TEXT(36)")
    
    ' Remove NOT NULL constraints (Access handles this differently)
    strResult = Replace(strResult, " NOT NULL", "")
    
    ' Remove DEFAULT constraints for now (can be added manually later)
    Dim intPos As Integer
    Do
        intPos = InStr(strResult, " DEFAULT ")
        If intPos > 0 Then
            Dim intNextComma As Integer
            Dim intNextParen As Integer
            intNextComma = InStr(intPos, strResult, ",")
            intNextParen = InStr(intPos, strResult, ")")
            
            If intNextComma > 0 And (intNextComma < intNextParen Or intNextParen = 0) Then
                strResult = Left(strResult, intPos - 1) & Mid(strResult, intNextComma)
            ElseIf intNextParen > 0 Then
                strResult = Left(strResult, intPos - 1) & Mid(strResult, intNextParen)
            Else
                Exit Do
            End If
        End If
    Loop While intPos > 0
    
    ConvertSQLServerToAccess = strResult
    Exit Function

ErrorHandler:
    ConvertSQLServerToAccess = strSQL
End Function

' Function to add indexes from SQL content
Private Sub AddIndexesFromSQL(strTableName As String, strSQL As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strLines() As String
    Dim strLine As String
    Dim i As Integer
    Dim strIndexName As String
    Dim strIndexFields As String
    Dim bPrimaryKey As Boolean
    Dim bUnique As Boolean
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTableName)
    
    ' Split SQL into lines
    strLines = Split(strSQL, vbCrLf)
    
    ' Process each line looking for index definitions
    For i = 0 To UBound(strLines)
        strLine = Trim(strLines(i))
        
        ' Check for Primary Key
        If InStr(UCase(strLine), "PRIMARY KEY") > 0 Then
            strIndexName = "PrimaryKey"
            strIndexFields = ExtractIndexFields(strLine)
            If strIndexFields <> "" Then
                CreateAccessIndex tdf, strIndexName, strIndexFields, True, True
            End If
        
        ' Check for Unique Index
        ElseIf InStr(UCase(strLine), "CREATE UNIQUE INDEX") > 0 Then
            strIndexName = ExtractIndexName(strLine)
            strIndexFields = ExtractIndexFields(strLine)
            If strIndexName <> "" And strIndexFields <> "" Then
                CreateAccessIndex tdf, strIndexName, strIndexFields, False, True
            End If
        
        ' Check for Regular Index
        ElseIf InStr(UCase(strLine), "CREATE INDEX") > 0 And InStr(UCase(strLine), "UNIQUE") = 0 Then
            strIndexName = ExtractIndexName(strLine)
            strIndexFields = ExtractIndexFields(strLine)
            If strIndexName <> "" And strIndexFields <> "" Then
                CreateAccessIndex tdf, strIndexName, strIndexFields, False, False
            End If
        End If
    Next i
    
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Set tdf = Nothing
    Set db = Nothing
    Debug.Print "Error adding indexes for table " & strTableName & ": " & Err.Description
End Sub

' Helper function to create an index in Access
Private Sub CreateAccessIndex(tdf As DAO.TableDef, strName As String, strFields As String, bPrimary As Boolean, bUnique As Boolean)
    On Error GoTo ErrorHandler
    
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim arrFields() As String
    Dim i As Integer
    
    ' Create new index
    Set idx = tdf.CreateIndex(strName)
    idx.Primary = bPrimary
    idx.Unique = bUnique
    
    ' Parse field names
    strFields = Replace(strFields, "[", "")
    strFields = Replace(strFields, "]", "")
    arrFields = Split(strFields, ",")
    
    ' Add fields to index
    For i = 0 To UBound(arrFields)
        Set fld = idx.CreateField(Trim(arrFields(i)))
        idx.fields.Append fld
    Next i
    
    ' Append index to table
    tdf.Indexes.Append idx
    
    Set fld = Nothing
    Set idx = Nothing
    Exit Sub

ErrorHandler:
    Set fld = Nothing
    Set idx = Nothing
    Debug.Print "Error creating index " & strName & ": " & Err.Description
End Sub

' Helper function to extract index name from SQL line
Private Function ExtractIndexName(strLine As String) As String
    On Error GoTo ErrorHandler
    
    Dim intStart As Integer
    Dim intEnd As Integer
    
    ' Look for pattern like "CREATE INDEX [IndexName]" or "CREATE UNIQUE INDEX [IndexName]"
    intStart = InStr(UCase(strLine), "INDEX [")
    If intStart > 0 Then
        intStart = intStart + 6 ' Move past "INDEX "
        intEnd = InStr(intStart, strLine, "]")
        If intEnd > 0 Then
            ExtractIndexName = Mid(strLine, intStart + 1, intEnd - intStart - 1)
        End If
    End If
    
    Exit Function

ErrorHandler:
    ExtractIndexName = ""
End Function

' Helper function to extract index fields from SQL line
Private Function ExtractIndexFields(strLine As String) As String
    On Error GoTo ErrorHandler
    
    Dim intStart As Integer
    Dim intEnd As Integer
    
    ' Look for fields within parentheses
    intStart = InStr(strLine, "(")
    intEnd = InStr(strLine, ")")
    
    If intStart > 0 And intEnd > intStart Then
        ExtractIndexFields = Mid(strLine, intStart + 1, intEnd - intStart - 1)
    End If
    
    Exit Function

ErrorHandler:
    ExtractIndexFields = ""
End Function

' Helper function to check if table exists
Private Function TableExists(strTableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTableName)
    
    TableExists = True
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TableExists = False
    Set tdf = Nothing
    Set db = Nothing
End Function

' =============================================================================
' QUERY DEFINITION EXPORT/IMPORT FUNCTIONS
' =============================================================================

' Main subroutine to export all query definitions to SQL files
Public Sub ExportAllQueryDefs()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim fso As Object
    Dim txtFile As Object
    Dim strExportPath As String
    Dim strFileName As String
    Dim intCount As Integer
    
    ' Set the database reference
    Set db = CurrentDb()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set export path
    strExportPath = CurrentProject.Path & "\QueryDefinitions\"
    
    ' Create directory if it doesn't exist
    If Not fso.FolderExists(strExportPath) Then
        fso.CreateFolder strExportPath
    End If
    
    ' Loop through all query definitions
    For Each qdf In db.QueryDefs
        ' Skip system queries and temporary queries
        If Left(qdf.Name, 1) <> "~" And Left(qdf.Name, 4) <> "MSys" Then
            strFileName = strExportPath & qdf.Name & ".sql"
            
            ' Create and write to file using FileSystemObject
            Set txtFile = fso.CreateTextFile(strFileName, True) ' True = overwrite if exists
            txtFile.Write GenerateQuerySQL(qdf)
            txtFile.Close
            Set txtFile = Nothing
            
            intCount = intCount + 1
            Debug.Print "Exported query: " & qdf.Name & " to " & strFileName
        End If
    Next qdf
    
    ' Clean up
    Set txtFile = Nothing
    Set fso = Nothing
    Set qdf = Nothing
    Set db = Nothing
    
    MsgBox "Successfully exported " & intCount & " query definitions to: " & strExportPath, vbInformation
    Exit Sub

ErrorHandler:
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    Set qdf = Nothing
    Set db = Nothing
    MsgBox "Error exporting queries: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Function to generate query SQL with metadata
Private Function GenerateQuerySQL(qdf As DAO.QueryDef) As String
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    Dim prm As DAO.Parameter
    Dim strParameters As String
    
    ' Start building the query definition
    strSQL = "-- Query: " & qdf.Name & vbCrLf
    strSQL = strSQL & "-- Created: " & Format(Now(), "yyyy-mm-dd hh:nn:ss") & vbCrLf
    strSQL = strSQL & "-- Type: " & GetQueryType(qdf.Type) & vbCrLf
    
    ' Add parameters if any
    If qdf.Parameters.Count > 0 Then
        strSQL = strSQL & "-- Parameters:" & vbCrLf
        For Each prm In qdf.Parameters
            strParameters = strParameters & "--   " & prm.Name & " (" & GetParameterType(prm.Type) & ")" & vbCrLf
        Next prm
        strSQL = strSQL & strParameters
    End If
        
    ' Add the actual SQL statement
    strSQL = strSQL & "-- SQL Statement:" & vbCrLf
    strSQL = strSQL & qdf.sql & vbCrLf
    
    GenerateQuerySQL = strSQL
    Exit Function

ErrorHandler:
    GenerateQuerySQL = "-- Error generating SQL for query: " & qdf.Name & vbCrLf & _
                       "-- Error: " & Err.Number & " - " & Err.Description
End Function

' Function to get query type description
Private Function GetQueryType(intType As Integer) As String
    Select Case intType
        Case dbQSelect
            GetQueryType = "Select Query"
        Case dbQAction
            GetQueryType = "Action Query"
        Case dbQCrosstab
            GetQueryType = "Crosstab Query"
        Case dbQDelete
            GetQueryType = "Delete Query"
        Case dbQUpdate
            GetQueryType = "Update Query"
        Case dbQAppend
            GetQueryType = "Append Query"
        Case dbQMakeTable
            GetQueryType = "Make Table Query"
        Case dbQDDL
            GetQueryType = "DDL Query"
        Case dbQSQLPassThrough
            GetQueryType = "SQL Pass-Through Query"
        Case dbQSetOperation
            GetQueryType = "Union Query"
        Case Else
            GetQueryType = "Unknown Query Type (" & intType & ")"
    End Select
End Function

' Function to get parameter type description
Private Function GetParameterType(intType As Integer) As String
    Select Case intType
        Case dbBoolean
            GetParameterType = "Boolean"
        Case dbByte
            GetParameterType = "Byte"
        Case dbInteger
            GetParameterType = "Integer"
        Case dbLong
            GetParameterType = "Long"
        Case dbCurrency
            GetParameterType = "Currency"
        Case dbSingle
            GetParameterType = "Single"
        Case dbDouble
            GetParameterType = "Double"
        Case dbDate
            GetParameterType = "Date/Time"
        Case dbText
            GetParameterType = "Text"
        Case dbMemo
            GetParameterType = "Memo"
        Case dbGUID
            GetParameterType = "GUID"
        Case Else
            GetParameterType = "Unknown Type (" & intType & ")"
    End Select
End Function

' Function to export a single query definition
Public Sub ExportSingleQueryDef(strQueryName As String)
    On Error GoTo ErrorHandler
    If fp_exported_queries = "" Then set_folder_paths
    
    Dim strExportPath As String: strExportPath = fp_exported_queries
    Dim strFileName   As String: strFileName = strExportPath & strQueryName & ".sql"
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim db  As DAO.Database:     Set db = CurrentDb()
    Dim qdf As DAO.QueryDef:     Set qdf = db.QueryDefs(strQueryName)
    Dim txtFile As Object:       Set txtFile = fso.CreateTextFile(strFileName, True, True)
    '
    ' Write Generated SQL for Query to file
    txtFile.Write GenerateQuerySQL(qdf)
    '
    ' Save/close
    txtFile.Close
    '
    ' All Done
    Set txtFile = Nothing
    Set fso = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Debug.Print "Query definition for '" & strQueryName & "' exported to: " & strFileName
    '
Exit Sub

ErrorHandler:
    If Not txtFile Is Nothing Then txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    Set qdf = Nothing
    Set db = Nothing
    MsgBox "Error exporting query: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' =============================================================================
' QUERY DEFINITION IMPORT FUNCTIONS
' =============================================================================

' Main subroutine to import all query definitions from SQL files
Public Sub ImportAllQueryDefs()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim strImportPath As String
    Dim strQueryName As String
    Dim intCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set import path
    strImportPath = CurrentProject.Path & "\QueryDefinitions\"
    
    ' Check if directory exists
    If Not fso.FolderExists(strImportPath) Then
        MsgBox "QueryDefinitions folder not found at: " & strImportPath, vbCritical
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(strImportPath)
    
    ' Loop through all SQL files in the folder
    For Each file In folder.Files
        If UCase(fso.GetExtensionName(file.Name)) = "SQL" Then
            ' Extract query name from filename
            strQueryName = fso.GetBaseName(file.Name)
            
            ' Import the query
            If ImportSingleQueryDef(strQueryName) Then
                intCount = intCount + 1
                Debug.Print "Imported query: " & strQueryName
            End If
        End If
    Next file
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    MsgBox "Successfully imported " & intCount & " query definitions from: " & strImportPath, vbInformation
    Exit Sub

ErrorHandler:
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    MsgBox "Error importing queries: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' Function to import a single query definition from SQL file
Public Function ImportSingleQueryDef(strQueryName As String) As Boolean
On Error GoTo ErrorHandler
    If fp_exported_queries = "" Then set_folder_paths

    Dim strFilePath    As String:    strFilePath = fp_exported_queries & strQueryName & ".sql"
    Dim strFileContent As String: strFileContent = ReadSQLFile(strFilePath)
    Dim db             As DAO.Database:   Set db = CurrentDb()
    Dim qdf            As DAO.QueryDef
    Dim strSQL         As String
    
    If strFileContent = "" Then
        Debug.Print "Error: Could not read SQL file for query " & strQueryName
        ImportSingleQueryDef = False
        Exit Function
    End If
    
    ' Extract the SQL statement from the file content
    strSQL = ExtractQuerySQL(strFileContent)
    
    If strSQL = "" Then
        Debug.Print "Error: Could not extract SQL statement for query " & strQueryName
        ImportSingleQueryDef = False
        Exit Function
    End If
    
    ' Check if query already exists
    If QueryExists(strQueryName) Then
        If MsgBox("Query '" & strQueryName & "' already exists. Do you want to replace it?", _
                  vbYesNo + vbQuestion) = vbYes Then
            db.QueryDefs.Delete strQueryName
        Else
            ImportSingleQueryDef = False
            Exit Function
        End If
    End If
    
    ' Create the new query
    Set qdf = db.CreateQueryDef(strQueryName, strSQL)
    
    ' Set description if available in the file
    SetQueryDescription qdf, strFileContent
    
    ImportSingleQueryDef = True
    Debug.Print "Successfully created query: " & strQueryName
    
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    Set qdf = Nothing
    Set db = Nothing
    Debug.Print "Error creating query " & strQueryName & ": " & Err.Number & " - " & Err.Description
    ImportSingleQueryDef = False
End Function

' Function to extract SQL statement from file content
Private Function ExtractQuerySQL(strFileContent As String) As String
    On Error GoTo ErrorHandler
    
    Dim arrLines() As String
    Dim strLine As String
    Dim strSQL As String
    Dim bSQLSection As Boolean
    Dim i As Integer
    
    ' Split content into lines
    arrLines = Split(strFileContent, vbCrLf)
    
    ' Look for the SQL statement section
    For i = 0 To UBound(arrLines)
        strLine = Trim(arrLines(i))
        
        ' Check if we've reached the SQL statement section
        If InStr(UCase(strLine), "-- SQL STATEMENT:") > 0 Then
            bSQLSection = True
        ElseIf bSQLSection And Left(strLine, 2) <> "--" And strLine <> "" Then
            ' Start collecting SQL lines (non-comment, non-empty)
            strSQL = strSQL & strLine & vbCrLf
        ElseIf bSQLSection And strSQL <> "" And Left(strLine, 2) = "--" Then
            ' Stop if we hit another comment section after collecting SQL
            Exit For
        End If
    Next i
    
    ' Clean up the SQL statement
    strSQL = Trim(strSQL)
    If Right(strSQL, 2) = vbCrLf Then
        strSQL = Left(strSQL, Len(strSQL) - 2)
    End If
    
    ExtractQuerySQL = strSQL
    Exit Function

ErrorHandler:
    ExtractQuerySQL = ""
End Function

' Function to set query description from file content
Private Sub SetQueryDescription(qdf As DAO.QueryDef, strFileContent As String)
    On Error GoTo ErrorHandler
    
    Dim arrLines() As String
    Dim strLine As String
    Dim strDescription As String
    Dim i As Integer
    
    ' Split content into lines
    arrLines = Split(strFileContent, vbCrLf)
    
    ' Look for description line
    For i = 0 To UBound(arrLines)
        strLine = Trim(arrLines(i))
        
        If InStr(UCase(strLine), "-- DESCRIPTION:") > 0 Then
            strDescription = Trim(Mid(strLine, InStr(UCase(strLine), "-- DESCRIPTION:") + 15))
            If strDescription <> "No description" And strDescription <> "" Then
                qdf.Properties("Description") = strDescription
            End If
            Exit For
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    ' Description property might not exist, ignore errors
End Sub

' Helper function to check if query exists
Private Function QueryExists(strQueryName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb()
    Set qdf = db.QueryDefs(strQueryName)
    
    QueryExists = True
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    QueryExists = False
    Set qdf = Nothing
    Set db = Nothing
End Function

' Function to import query definitions from a specific folder
Public Sub ImportQueryDefsFromFolder(strFolderPath As String)
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim strQueryName As String
    Dim intCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure folder path is valid and exists
    If Not fso.FolderExists(strFolderPath) Then
        MsgBox "Folder not found: " & strFolderPath, vbCritical
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(strFolderPath)
    
    ' Loop through all SQL files in the folder
    For Each file In folder.Files
        If UCase(fso.GetExtensionName(file.Name)) = "SQL" Then
            ' Extract query name from filename
            strQueryName = fso.GetBaseName(file.Name)
            
            ' Import the query
            If ImportSingleQueryDef(strQueryName) Then
                intCount = intCount + 1
                Debug.Print "Imported query: " & strQueryName
            End If
        End If
    Next file
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    MsgBox "Successfully imported " & intCount & " query definitions from: " & strFolderPath, vbInformation
    Exit Sub

ErrorHandler:
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    MsgBox "Error importing queries from folder: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

