Option Compare Database
Option Explicit
'
' Module Variables
Private fso     As New FileSystemObject

Public Sub export_all()
    '
    ' Initialize fso
    Set fso = New FileSystemObject
    '
    ' Create Folders if not already there
    Call create_folder_structure
    '
    ' Export all data of "non" direct related to "dataset".
    Call export_non_direct_related_to_dataset
    '
    ' Export all "Dataset"-defintions.
    Call export_all_dataset_and_related_definitions
    Call build_sql_file_dataset
    '
    ' Local Variables
    Dim txt As TextStream: Set txt = fso.OpenTextFile(mdl_Folders.repos() & "insert_definition_into_temp_tables.sql", ForWriting, True, TristateTrue)
    '
    '
    txt.WriteLine ""
    txt.WriteLine "/* Static Reference Data */"
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "datatype.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "development_status.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "dq_dimension.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "dq_result_status.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "dq_review_status.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "dq_risk_level.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "processing_status.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "processing_step.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "parameter_group.sql"""
    txt.WriteLine ":r """ & mdl_Folders.srd(True) & "parameter.sql"""
    '
    txt.WriteLine ""
    txt.WriteLine "/* Organization, Hierarchies and Groups */"
    txt.WriteLine ":r """ & mdl_Folders.ohg(True) & "group.sql"""
    txt.WriteLine ":r """ & mdl_Folders.ohg(True) & "hierarchy.sql"""
    '
    txt.WriteLine ""
    txt.WriteLine "/* Data Quality Model */"
    txt.WriteLine ":r """ & mdl_Folders.dqm(True) & "dq_requirement.sql"""
    '
    txt.WriteLine ""
    txt.WriteLine "/* All Model(s), Database(s), Dataset(s) */"
    txt.WriteLine ":r """ & mdl_Folders.dta(True) & "model.sql"""
    txt.WriteLine ":r """ & mdl_Folders.dta(True) & "database.sql"""
    txt.WriteLine ":r """ & mdl_Folders.dta(True) & "datasets.sql"""
    '
    txt.WriteLine ""
    txt.WriteLine "BEGIN /* Name of Git Repository / Current Model */"
    txt.WriteLine "  "
    txt.WriteLine "  DELETE FROM mdm.current_model; INSERT INTO mdm.current_model (id_model, nm_repository) SELECT  "
    txt.WriteLine "    id_model      = convert(char(32),      '" & mdl_Folders.id_model(mdl_Folders.nm_repository()) & "'),"
    txt.WriteLine "    nm_repository = CONVERT(NVARCHAR(128), '" & mdl_Folders.nm_repository() & "');"
    txt.WriteLine "  "
    txt.WriteLine "END"
    txt.WriteLine "GO"
    txt.WriteLine ""
    '
    ' Close SQL-File
    txt.Close
    '
End Sub

Public Sub create_folder_structure()
    '
    ' Check if " 2-Definitions"-folder exists.
    If Not fso.FolderExists(mdl_Folders.repos) Then Call fso.CreateFolder(mdl_Folders.repos)
    '
    ' Check if "metadata"-domain-folder exists, if NOT create them.
    If Not fso.FolderExists(mdl_Folders.srd) Then Call fso.CreateFolder(mdl_Folders.srd)
    If Not fso.FolderExists(mdl_Folders.ohg) Then Call fso.CreateFolder(mdl_Folders.ohg)
    If Not fso.FolderExists(mdl_Folders.dta) Then Call fso.CreateFolder(mdl_Folders.dta)
    If Not fso.FolderExists(mdl_Folders.dqm) Then Call fso.CreateFolder(mdl_Folders.dqm)

End Sub
Public Sub export_non_direct_related_to_dataset()
    '
    ' Static Reference Data
    export_table "srd", "datatype"
    export_table "srd", "development_status"
    export_table "srd", "dq_dimension"
    export_table "srd", "dq_result_status"
    export_table "srd", "dq_review_status"
    export_table "srd", "dq_risk_level"
    export_table "srd", "processing_status"
    export_table "srd", "processing_step"
    export_table "srd", "parameter_group"
    export_table "srd", "parameter"
    '
    ' Organization, Hierarchies and Groups
    export_table "ohg", "group"
    export_table "ohg", "hierarchy"
    '
    ' Data Quality Model
    export_table "dqm", "dq_requirement"
    '
End Sub
Public Sub export_table(nm_schema As String, nm_table As String)
    '
    ' Local Variables
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM " & nm_schema & "_" & nm_table & IIf(nm_table = "model", "", " WHERE id_model = '" & id_model_default & "'"))
    Dim txt As TextStream: Set txt = fso.OpenTextFile(mdl_Folders.fld(nm_schema) & nm_table & ".sql", ForWriting, True, TristateTrue)
    '
    ' Export all data to SQL-file.
    txt.WriteLine "BEGIN"
    If (Not rst.EOF) Then
        Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert(nm_schema, nm_table, .fields): rst.MoveNext: End With: Loop
    Else
        txt.WriteLine "PRINT('No Metadata');"
    End If
    txt.WriteLine "END"
    txt.WriteLine "GO"
    txt.WriteLine ""
    '
    ' Close SQL-File
    txt.Close
    '
End Sub
Public Sub build_sql_file_dataset()
    '
    ' Local Variables
    Dim txt As TextStream: Set txt = fso.OpenTextFile(mdl_Folders.dta & "datasets.sql", ForWriting, True, TristateTrue)
    '
    ' Local Variables
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'")
    '
    ' Export all data to SQL-file.
    Do Until rst.EOF: With rst: txt.WriteLine ":r "".\" & .fields("id_dataset") & ".sql""": .MoveNext: End With: Loop
    '
    ' Close SQL-File
    txt.Close
    '
End Sub
Public Sub export_all_dataset_and_related_definitions()
    '
    ' Local Variables
    Dim txt As TextStream: Set txt = fso.OpenTextFile(mdl_Folders.fld("dta") & "datasets.sql", ForWriting, True, TristateTrue)
    '
    ' Local Variables
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'")
    '
    ' Export all data to SQL-file.
    Do Until rst.EOF: With rst: export_dataset_and_related_definitions .fields("id_dataset"): .MoveNext: End With: Loop
    '
    ' Close SQL-File
    txt.Close
    '
End Sub

Public Sub export_dataset_and_related_definitions(id_dataset As String)
    '
    ' build sql for filtering on model.
    Dim tx_where As String: tx_where = "WHERE id_dataset = '" & id_dataset & "' AND id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'"
    '
    If (id_dataset = "") Then
        Exit Sub
    End If
    '
    ' Initialize fso
    Set fso = New FileSystemObject
    '
    ' Local Variables
    Dim txt As TextStream: Set txt = fso.OpenTextFile(mdl_Folders.dta() & id_dataset & ".sql", ForWriting, True, TristateTrue)
    Dim rst As Recordset
    '
    ' Write to SQL-file
    txt.WriteLine "/* -------------------------------------------------------------------------- */"
    txt.WriteLine "/* Definitions for `Dataset` and `related`-objects like `attributes`,         */"
    txt.WriteLine "/* `DQ Controls`, `DQ Thresholds` and `related Group(s)`.                     */"
    txt.WriteLine "/* -------------------------------------------------------------------------- */"
    txt.WriteLine "/*                                                                            */"
    txt.WriteLine "/* ID Dataset : `" & id_dataset & "`                            */"
    txt.WriteLine "/*                                                                            */"
    txt.WriteLine "/* -------------------------------------------------------------------------- */"
    txt.WriteLine "BEGIN"
    txt.WriteLine ""
    '
    ' Export record to "Dataset"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset " & tx_where)
    txt.WriteLine "  /* --------------------- */"
    txt.WriteLine "  /* `Dataset`-definitions */"
    txt.WriteLine "  /* --------------------- */"
    txt.WriteLine build_sql_insert("dta", "dataset", rst.fields)
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `Dataset`"
    txt.WriteLine "  "
    '
    ' Export record to "Attribute"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_attribute " & tx_where)
    txt.WriteLine "  /* ----------------------- */"
    txt.WriteLine "  /* `Attribute`-definitions */"
    txt.WriteLine "  /* ----------------------- */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dta", "attribute", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `Attribute`"
    txt.WriteLine ""
    '
    ' Export record to "Parameter Values"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_parameter_value " & tx_where)
    txt.WriteLine "  /* ------------------------------ */"
    txt.WriteLine "  /* `Parameter Values`-definitions */"
    txt.WriteLine "  /* ------------------------------ */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dta", "parameter_value", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `Parameter Values`"
    txt.WriteLine ""
    '
    ' Export record to "SQL for ETL"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_ingestion_etl " & tx_where)
    txt.WriteLine "  /* ------------------------------ */"
    txt.WriteLine "  /* `SQL for ETL`-definitions      */"
    txt.WriteLine "  /* ------------------------------ */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dta", "ingestion_etl", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `SQL for ETL`"
    txt.WriteLine ""
    '
    ' Export record to "Schedule"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_schedule " & tx_where)
    txt.WriteLine "  /* ------------------------------ */"
    txt.WriteLine "  /* `Schedule`-definitions         */"
    txt.WriteLine "  /* ------------------------------ */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dta", "schedule", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `SQL for ETL`"
    txt.WriteLine ""
    '
    ' Export record to "Related (Groups)"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM ohg_related " & tx_where)
    txt.WriteLine "  /* -------------------------------- */"
    txt.WriteLine "  /* `Related (Group(s))`-definitions */"
    txt.WriteLine "  /* -------------------------------- */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("ohg", "related", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `Related (Group(s))`"
    txt.WriteLine ""
    '
    ' Export record to "DQ Controls"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dqm_dq_control " & tx_where)
    txt.WriteLine "  /* ------------------------ */"
    txt.WriteLine "  /* `DQ Control`-definitions */"
    txt.WriteLine "  /* ------------------------ */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dqm", "dq_control", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `DQ Control`"
    txt.WriteLine ""
    '
    ' Export record to "DQ Threshold"-definitions
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dqm_dq_threshold WHERE id_dq_control IN (SELECT id_dq_control FROM dqm_dq_control " & tx_where & ") AND id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'")
    txt.WriteLine "  /* -------------------------- */"
    txt.WriteLine "  /* `DQ Threshold`-definitions */"
    txt.WriteLine "  /* -------------------------- */"
    Do Until rst.EOF: With rst: txt.WriteLine build_sql_insert("dqm", "dq_threshold", .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: txt.WriteLine "  -- No Defintions for `DQ Threshold`"
    txt.WriteLine "  "
    txt.WriteLine "END"
    txt.WriteLine "GO"
    txt.WriteLine ""
    '
    ' Close "SQL"-file.
    txt.Close
    '
End Sub

Public Function build_sql_insert(ByVal nm_schema As String, ByVal nm_table As String, ByRef fields As fields) As String
    '
    ' Local Variables
    Dim tx_fields As String: tx_fields = ""
    Dim tx_values As String: tx_values = ""
    '
    ' Export record to SQL-file
    Dim fld As Field: For Each fld In fields
        '
        tx_fields = tx_fields & IIf(tx_fields = "", "", ", ")
        tx_fields = tx_fields & fld.Name
        '
        tx_values = tx_values & IIf(tx_values = "", "", ", ")
        Select Case fld.Type
            Case DAO.DataTypeEnum.dbDate
                tx_values = tx_values & "'" & Format(fld.Value, "yyyy-mm-dd") & "'"
            
            Case DAO.DataTypeEnum.dbDecimal
                tx_values = tx_values & "'" & Replace(Format(fld.Value, "0.000000"), CheckDecimalSeparator, ".") & "'"
                                
            
            Case DAO.DataTypeEnum.dbBoolean
                tx_values = tx_values & "'" & IIf(fld.Value = True, "1", "0") & "'"
                                
            Case Else
                tx_values = tx_values & "'" & Replace(Nz(fld.Value, ""), "'", "<quot>") & "'"
                
        End Select
        '
    Next fld
    '
    ' Replace 'double-single quot`s with NULL
    tx_values = Replace(tx_values, Chr(13) & Chr(10), "<newline>")
    tx_values = Replace(tx_values, Chr(10) & Chr(13), "<newline>")
    tx_values = Replace(tx_values, Chr(10), "<newline>")
    tx_values = Replace(tx_values, Chr(13), "<newline>")
    tx_values = Replace(tx_values, "''", "NULL")
    tx_values = Replace(tx_values, "<quot>", "''")
    '
    ' Build SQL: Insert part
    build_sql_insert = "  INSERT INTO tsa_" & nm_schema & ".tsa_" & nm_table & " (" & tx_fields & ") VALUES (" & tx_values & ");"
    '
End Function