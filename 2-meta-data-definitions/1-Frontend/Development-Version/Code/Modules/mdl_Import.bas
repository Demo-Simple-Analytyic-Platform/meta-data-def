Option Compare Database
Option Explicit
'
' Check if setup mode
Public Function is_setup_mode() As Boolean
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim pth As String:               pth = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "3-Data-Transformation-Area\"
    Dim tsm As TextStream:       Set tsm = fso.OpenTextFile(pth & "model.sql", ForReading, True)
    Dim txt As String
    '
    is_setup_mode = False
    If tsm.AtEndOfStream Then
        is_setup_mode = True
    Else
        txt = tsm.ReadAll
        If Len(txt) < 3 Then
            is_setup_mode = True
        End If
    End If
End Function
'
Public Sub process_all_sql_files(fm As Form_StartUp)
    '
    ' Local Variables
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim ms As String: ms = "Processing for Model `<nm_repository>`" & vbNewLine & "SQL-file : `<tx_sql_filepath>`"
    Dim rp As String
    Dim ml As Recordset
    Dim fs As FileSystemObject: Set fs = New FileSystemObject
    Dim fd As folder
    Dim fl As file
    Dim sp As Integer
    Dim tx As String
    Dim is_setup_mode_detected As Boolean: is_setup_mode_detected = is_setup_mode
    '
    ' Reset feedback
    fm.tx_loading_feedback.Caption = ""
    '
    ' Truncate all tables in Access DB
    Call truncate_all_meta_datasets
    '
    ' Insert "Model/Repository-record if "Setup-Mode".
    If is_setup_mode_detected = True Then
        '
        tx = "INSERT INTO dta_model (id_model, nm_repository) VALUES ('" & id_model_default & "', '" & nm_repository & "')"
        DoCmd.SetWarnings False: DoCmd.RunSQL tx: DoCmd.SetWarnings True
        '
        'Export the "Model"-definitions to the repository file structure!
        Call mdl_Export.export_table("dta", "model")
        '
        ' Inform user!
        If (fso.FileExists(CurrentProject.Path & "\PowerShellStart.txt") = False) Then
            Call MsgBox("This iniital setup from Template in `meta-data-def`-repository, the `Model` named `" & nm_repository & "` has been initialized.'" _
                      & "Please remeber to assign the correct database credentials under the `Model`-settings.", vbInformation)
        End If
        '
    Else
        '
        ' Process Model and Database
        rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "3-Data-Transformation-Area\"
        Call process_sql_file(rp & "model.sql", fm)
        Call load_dta_database(fm)
    End If
    '
    ' Determine how much is to be done
    fm.ni_progress_todo = count_number_of_lines()
    '
    ' Load "Static Reference Data"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "1-Static-Reference-Data\"
    Call process_sql_file(rp & "datatype.sql", fm)
    Call process_sql_file(rp & "development_status.sql", fm)
    Call process_sql_file(rp & "dq_dimension.sql", fm)
    Call process_sql_file(rp & "dq_result_status.sql", fm)
    Call process_sql_file(rp & "dq_review_status.sql", fm)
    Call process_sql_file(rp & "dq_risk_level.sql", fm)
    Call process_sql_file(rp & "parameter_group.sql", fm)
    Call process_sql_file(rp & "parameter.sql", fm)
    Call process_sql_file(rp & "processing_status.sql", fm)
    Call process_sql_file(rp & "processing_step.sql", fm)
    '
    ' Load "Organisation, Hierarchy and Group"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "2-Organization-Hierarchies-and-Groups\"
    Call process_sql_file(rp & "group.sql", fm)
    Call process_sql_file(rp & "hierarchy.sql", fm)
    '
    ' Load "Data Quality Requirements"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "4-Data-Quality-Model\"
    Call process_sql_file(rp & "dq_requirement.sql", fm)
    '
    ' Load "Data Quality Requirements"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "3-Data-Transformation-Area\"
    For Each fl In fs.GetFolder(rp).Files
        If (fl.Name <> "datasets.sql") And (fl.Name <> "database.sql") And (fl.Name <> "model.sql") Then
            Call process_sql_file(fl.Path, fm)
        End If
    Next fl
    '
    ' loading "referenced"-models!
    Set ml = CurrentDb.OpenRecordset("SELECT * FROM [models] WHERE tx_repo_folderpath_exists <> 0 AND id_model <> '" & mdl_Folders.id_model(mdl_Folders.nm_repository()) & "'")
    Do While Not ml.EOF
        '
        ' Load "Static Reference Data"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "1-Static-Reference-Data\"
        Call process_sql_file(rp & "datatype.sql", fm)
        Call process_sql_file(rp & "development_status.sql", fm)
        Call process_sql_file(rp & "dq_dimension.sql", fm)
        Call process_sql_file(rp & "dq_result_status.sql", fm)
        Call process_sql_file(rp & "dq_review_status.sql", fm)
        Call process_sql_file(rp & "dq_risk_level.sql", fm)
        Call process_sql_file(rp & "parameter_group.sql", fm)
        Call process_sql_file(rp & "parameter.sql", fm)
        Call process_sql_file(rp & "processing_status.sql", fm)
        Call process_sql_file(rp & "processing_step.sql", fm)
        '
        ' Load "Organisation, Hierarchy and Group"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "2-Organization-Hierarchies-and-Groups\"
        Call process_sql_file(rp & "group.sql", fm)
        Call process_sql_file(rp & "hierarchy.sql", fm)
        '
        ' Load "Data Quality Requirements"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "4-Data-Quality-Model\"
        Call process_sql_file(rp & "dq_requirement.sql", fm)
        '
        ' Load "Data Quality Requirements"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "3-Data-Transformation-Area\"
        For Each fl In fs.GetFolder(rp).Files
            If (fl.Name <> "datasets.sql") And (fl.Name <> "database.sql") And (fl.Name <> "model.sql") Then
                Call process_sql_file(fl.Path, fm)
            End If
        Next fl
        '
    ml.MoveNext: Loop
    '
    ' Do "Export" because "Setup"-mode was detected!
    If is_setup_mode_detected Then
        mdl_Export.export_all
    End If
    '
End Sub
Public Function process_all_sql_files_no_feedback() As Boolean
    '
    ' Local Variables
    Dim fm As Form: Set fm = Nothing
    Dim ms As String: ms = "Processing for Model `<nm_repository>`" & vbNewLine & "SQL-file : `<tx_sql_filepath>`"
    Dim rp As String
    Dim ml As Recordset
    Dim fs As FileSystemObject: Set fs = New FileSystemObject
    Dim fd As folder
    Dim fl As file
    Dim sp As Integer
    Dim tx As String
    Dim is_setup_mode_detected As Boolean: is_setup_mode_detected = is_setup_mode
    '
    ' Truncate all tables in Access DB
    Call truncate_all_meta_datasets
    '
    ' Insert "Model/Repository-record if "Setup-Mode".
    If is_setup_mode_detected = True Then
        '
        tx = "INSERT INTO dta_model (id_model, nm_repository) VALUES ('" & id_model_default & "', '" & nm_repository & "')"
        DoCmd.SetWarnings False: DoCmd.RunSQL tx: DoCmd.SetWarnings True
        '
        'Export the "Model"-definitions to the repository file structure!
        Call mdl_Export.export_table("dta", "model")
        '
    Else
        '
        ' Process Model and Database
        rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "3-Data-Transformation-Area\"
        Call process_sql_file(rp & "model.sql", fm)
        Call load_dta_database(fm)
    End If
    '
    ' Load "Static Reference Data"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "1-Static-Reference-Data\"
    Call process_sql_file(rp & "datatype.sql", fm)
    Call process_sql_file(rp & "development_status.sql", fm)
    Call process_sql_file(rp & "dq_dimension.sql", fm)
    Call process_sql_file(rp & "dq_result_status.sql", fm)
    Call process_sql_file(rp & "dq_review_status.sql", fm)
    Call process_sql_file(rp & "dq_risk_level.sql", fm)
    Call process_sql_file(rp & "parameter_group.sql", fm)
    Call process_sql_file(rp & "parameter.sql", fm)
    Call process_sql_file(rp & "processing_status.sql", fm)
    Call process_sql_file(rp & "processing_step.sql", fm)
    '
    ' Load "Organisation, Hierarchy and Group"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "2-Organization-Hierarchies-and-Groups\"
    Call process_sql_file(rp & "group.sql", fm)
    Call process_sql_file(rp & "hierarchy.sql", fm)
    '
    ' Load "Data Quality Requirements"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "4-Data-Quality-Model\"
    Call process_sql_file(rp & "dq_requirement.sql", fm)
    '
    ' Load "Data Quality Requirements"
    rp = mdl_Folders.tx_repo_folderpath(mdl_Folders.nm_repository()) & "3-Data-Transformation-Area\"
    For Each fl In fs.GetFolder(rp).Files
        If (fl.Name <> "datasets.sql") And (fl.Name <> "database.sql") And (fl.Name <> "model.sql") Then
            Call process_sql_file(fl.Path, fm)
        End If
    Next fl
    '
    ' loading "referenced"-models!
    Set ml = CurrentDb.OpenRecordset("SELECT * FROM [models] WHERE tx_repo_folderpath_exists <> 0 AND id_model <> '" & mdl_Folders.id_model(mdl_Folders.nm_repository()) & "'")
    Do While Not ml.EOF
        '
        ' Load "Static Reference Data"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "1-Static-Reference-Data\"
        Call process_sql_file(rp & "datatype.sql", fm)
        Call process_sql_file(rp & "development_status.sql", fm)
        Call process_sql_file(rp & "dq_dimension.sql", fm)
        Call process_sql_file(rp & "dq_result_status.sql", fm)
        Call process_sql_file(rp & "dq_review_status.sql", fm)
        Call process_sql_file(rp & "dq_risk_level.sql", fm)
        Call process_sql_file(rp & "parameter_group.sql", fm)
        Call process_sql_file(rp & "parameter.sql", fm)
        Call process_sql_file(rp & "processing_status.sql", fm)
        Call process_sql_file(rp & "processing_step.sql", fm)
        '
        ' Load "Organisation, Hierarchy and Group"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "2-Organization-Hierarchies-and-Groups\"
        Call process_sql_file(rp & "group.sql", fm)
        Call process_sql_file(rp & "hierarchy.sql", fm)
        '
        ' Load "Data Quality Requirements"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "4-Data-Quality-Model\"
        Call process_sql_file(rp & "dq_requirement.sql", fm)
        '
        ' Load "Data Quality Requirements"
        rp = mdl_Folders.tx_repo_folderpath(ml!nm_repository) & "3-Data-Transformation-Area\"
        For Each fl In fs.GetFolder(rp).Files
            If (fl.Name <> "datasets.sql") And (fl.Name <> "database.sql") And (fl.Name <> "model.sql") Then
                Call process_sql_file(fl.Path, fm)
            End If
        Next fl
        '
    ml.MoveNext: Loop
    '
    ' Do "Export" because "Setup"-mode was detected!
    If is_setup_mode_detected Then
        Call mdl_Export.export_all
    End If
    '
    ' Return
    process_all_sql_files_no_feedback = True
    '
End Function

Public Sub truncate_all_meta_datasets()
    '
    ' Empty out the "DQ" related tables
    Call empty_table("dqm", "dq_threshold")
    Call empty_table("dqm", "dq_control")
    Call empty_table("dqm", "dq_requirement")
    '
    ' Empty out the "Groups" related tables
    Call empty_table("ohg", "hierarchy")
    Call empty_table("ohg", "related")
    '
    ' Empty out the "Dataset" related tables
    Call empty_table("dta", "parameter_value")
    Call empty_table("srd", "parameter_group")
    Call empty_table("srd", "parameter")
    '
    ' Empty out the "Dataset" related tables
    Call empty_table("dta", "ingestion_etl")
    Call empty_table("dta", "attribute")
    Call empty_table("dta", "schedule")
    Call empty_table("dta", "database")
    Call empty_table("dta", "dataset")
    '
    ' Empty out the "Other" related tables
    Call empty_table("srd", "dq_risk_level")
    Call empty_table("srd", "dq_dimension")
    Call empty_table("srd", "dq_result_status")
    Call empty_table("srd", "dq_review_status")
    Call empty_table("ohg", "group")
    Call empty_table("srd", "datatype")
    Call empty_table("srd", "development_status")
    Call empty_table("srd", "processing_status")
    Call empty_table("srd", "processing_step")
    '
    ' Empty out the "Models"
    Call empty_table("dta", "model")
    '
End Sub
Public Sub empty_table(nm_schema As String, nm_table As String)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM " & nm_schema & "_" & nm_table
    DoCmd.SetWarnings True
End Sub
Public Sub process_sql_file(ByVal tx_path_sql_file As String, Optional fm As Form_StartUp = Nothing)
    '
    ' Local Variables
    Dim ext As String: ext = Right(CurrentProject.Name, 5)
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim txt As TextStream:       Set txt = fso.OpenTextFile(tx_path_sql_file, ForReading, False, TristateMixed)
    Dim fgo As Boolean: fgo = False
    Dim sql As String:  sql = ""
    Dim txl As String:  txl = ""
    Dim ins As String:  ins = "  INSERT INTO "
    Dim gos As String:  gos = "GO"
    Dim ms  As String:
    '
    ' In case this is "setup"-mode, the "<id_model>"-placeholder must be replaced with the id_model of this model.
    Dim id_model As String: id_model = id_model_default
    '
    ' set feedback message
    If (Not (fm Is Nothing)) Then
        fm.tx_loading_feedback.Caption = Replace(Replace(ms, "<nm_repository>", mdl_Folders.nm_repository()), "<tx_sql_filepath>", relative_path(tx_path_sql_file))
    End If
    '
    ' Check is file has "text"-lines.
    If txt.AtEndOfStream = True Then
        Exit Sub
    End If
    '
    '  Read one line from SQL-file.
    txl = txt.ReadLine: If (Not (fm Is Nothing)) Then Call fm.tx_loading_progress_add
    txl = Replace(txl, "<id_model>", id_model)
    '
    ' Turn off Warnings
    DoCmd.SetWarnings False
    '
    ' Read from the file till the end.
    Do While Not txt.AtEndOfStream
        '
        ' Reset "SQL"-statement.
        sql = ""
        '
        ' Determine if "GO;"-statement needs to be found.
        ' check if line start with "INSERT"-statement
        If (Left(txl, Len(ins)) = ins) Then
            '
            ' Copy "Line" to SQL-statemen.
            sql = txl
            '
            ' Read " next"-line
            txl = txt.ReadLine: If (Not (fm Is Nothing)) Then Call fm.tx_loading_progress_add
            txl = Replace(txl, "<id_model>", id_model)
            Do Until ((Left(txl, Len(ins)) = ins) Or (Left(txl, Len(gos)) = gos) Or (Left(txl, 3) = "end") Or (Len(Trim(txl)) = 0) Or txt.AtEndOfStream)
                sql = sql & Chr(10) & txl
                txl = txt.ReadLine: Call fm.tx_loading_progress_add
                txl = Replace(txl, "<id_model>", id_model)
            Loop
            '
            ' Execute "SQL"-statemen
            sql = Trim(Replace(sql, "INSERT INTO tsa_", "INSERT INTO "))
            sql = Trim(Replace(sql, ".tsa_", "_"))
            '
            ' Custom operatie per table
            If (InStr(1, sql, "dqm_dq_threshold") > 0) Then
                sql = Replace(sql, ".", CheckDecimalSeparator)
            End If
            '
            ' Replace "newline"
            sql = Replace(sql, "<newline>", vbNewLine)
            '
            If (ext = "accdb") Then
                Debug.Print (sql)
            End If
            If (execute_sql(sql) = False) Then
                DoCmd.SetWarnings True
                Exit Sub
            End If
            '
            ' to see the feedback
            If (Not (fm Is Nothing)) Then
                Call mdl_Kernel32.Sleep(fm.ni_wait_time)
                DoEvents
            End If
            '
        Else
            txl = txt.ReadLine: If (Not (fm Is Nothing)) Then Call fm.tx_loading_progress_add
            txl = Replace(txl, "<id_model>", id_model)
        End If
        '
    Loop
    '
    ' Turn off Warnings
    DoCmd.SetWarnings True
    '
End Sub
Public Function count_number_of_lines() As Integer
    '
    ' Local Variables
    Dim rp As Recordset:        Set rp = CurrentDb.OpenRecordset("SELECT * FROM [models] WHERE tx_repo_folderpath_exists <> 0")
    Dim fs As FileSystemObject: Set fs = New FileSystemObject
    Dim fd As folder
    Dim fl As file
    Dim tx As TextStream
    Dim ln As String
    Dim cn As Integer
    '
    Do While Not rp.EOF
        '
        ' Process all the files in folders of all non directly "dataset"-related inserts.
        For Each fd In fs.GetFolder(rp!tx_repo_folderpath).SubFolders: For Each fl In fd.Files
            
            If (fl.Name <> "insert_definition_into_temp_tables.sql" _
            And fl.Name <> "datasets.sql" _
            And fl.Name <> "readme.md") Then
                Debug.Print fl.Name
                Set tx = fs.OpenTextFile(fl.Path, ForReading, False, TristateMixed)
                Do While Not tx.AtEndOfStream: ln = tx.ReadLine: cn = cn + 1: Loop
            End If
        Next fl: Next fd
        '
    rp.MoveNext: Loop
    '
    ' return number of lines to be processed
    count_number_of_lines = cn
    '
End Function
Public Sub skil_line(tx_line As String)
    '
    ' Handle all "known" non-sql-statements
    If (Left(tx_line, 2) = ":r") Then: Exit Sub
    If (Left(tx_line, 2) = ":/") Then: Exit Sub
    If (Left(tx_line, 1) = Chr(10)) Then: Exit Sub
    If (Left(tx_line, 1) = Chr(13)) Then: Exit Sub
    '
End Sub

Public Function execute_sql(tx_sql As String) As Boolean
    '
    ' Local Variables
    Dim shm As String:        shm = Mid(tx_sql, 13, 3)
    Dim tbl As String:        tbl = Mid(tx_sql, 17, InStr(1, tx_sql, " (", vbTextCompare) - 16)
    Dim sql As String:        sql = "SELECT COUNT(*) AS ni_records FROM " & shm & "_" & tbl
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    Dim exp As Integer:       exp = rst.fields("ni_records") + 1: rst.Close
    Dim log As Recordset
    Dim pky As Integer
    '
    ' Execute SQL
    execute_sql = True
    DoCmd.SetWarnings False
    DoCmd.RunSQL tx_sql
    DoCmd.SetWarnings True
    '
    ' Validate if record was inserted and did NOT fail!
    Set rst = CurrentDb.OpenRecordset(sql)
    If (exp <> rst.fields("ni_records")) Then
        '
        ' Log "failed" SQL statement.
        Set log = CurrentDb.OpenRecordset("SELECT * FROM sql_insert_log WHERE 1=2"): With log
            '
            '  Add record and save schema/table en SQL
            .AddNew
            .fields("nm_schema") = shm
            .fields("nm_table") = tbl
            .fields("tx_sql") = tx_sql
            pky = .fields("id_log")
            .Update
            '
            ' remeber id_log to filter log-from
            DoCmd.OpenForm "sql_insert_log", acNormal, , "id_log=" & str(pky), acFormReadOnly, acWindowNormal
            '
            ' Set to false to stop processing
            execute_sql = False
            DoCmd.CancelEvent
            '
            ' Close "StartUp"-from
            DoCmd.Close acForm, "StartUp", acSaveNo
            '
            ' Raise Error
            Err.Raise Number:=9999, Description:="SQL Insert Stattement failt!"
            '
        End With: log.Close
        '
    End If
    '
End Function
'
' Load database metadata information into tooling
Public Sub load_dta_database(fm As Form)
    '
    ' Local Variables
    Dim n   As String: n = vbNewLine
    Dim pth As String
    Dim sql As String: sql = "SELECT * FROM [models] ORDER BY [is_current_model] ASC, [nm_repository] ASC"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    '
    'Step 0: empty hlp_dta_database
    sql = "DELETE FROM [hlp_dta_database]"
    DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Process all the models
    Do While Not rst.EOF
        '
        ' Step 1: first load metadata "dta_database" of "current"-model.
        pth = rst!tx_repo_folderpath
        Call empty_table("dta", "database")
        Call process_sql_file(pth & "3-Data-Transformation-Area\database.sql", fm)
        '
        ' Step 2: remove all but "dta_database"-info, exept the model of the currrent record in recordset.
        sql = "DELETE FROM [dta_database] WHERE [id_model] <> '" & rst!id_model & "'"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        '
        'Step 3: copy left "dta_database"-metadata to "helper"
        sql = "INSERT INTO hlp_dta_database (" _
        & n & "  id_model, id_database, id_environment, nm_server, nm_database, nm_username, nm_secret" _
        & n & ")" _
        & n & "SELECT " _
        & n & "  id_model, id_database, id_environment, nm_server, nm_database, nm_username, nm_secret" _
        & n & "FROM dta_database"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        '
    rst.MoveNext: Loop
    '
    ' Step 4: copy "helper" data of the "dta_database".
    Call empty_table("dta", "database")
    sql = "INSERT INTO dta_database (" _
    & n & "  id_model, id_database, id_environment, nm_server, nm_database, nm_username, nm_secret" _
    & n & ")" _
    & n & "SELECT " _
    & n & "  id_model, id_database, id_environment, nm_server, nm_database, nm_username, nm_secret" _
    & n & "FROM hlp_dta_database"
    DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
End Sub