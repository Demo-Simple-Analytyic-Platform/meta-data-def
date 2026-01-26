Attribute VB_Name = "mdl_Export"
Option Compare Database
Option Explicit
'
' Module Variables
Private fso As New FileSystemObject

' Export all metadata definitions to repository .sql files.
' Creates folder structure, exports non-dataset + dataset definitions.
' Builds datasets.sql and insert_definition_into_temp_tables.sql includes.
' Params: none.
' Example: Call export_all
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
    Dim l_id_model      As String:      l_id_model = mdl_Folders.id_model(mdl_Folders.nm_repository())
    Dim l_nm_repository As String: l_nm_repository = mdl_Folders.nm_repository()
    Dim txt            As TextStream:      Set txt = fso.OpenTextFile(mdl_Folders.repos() & "insert_definition_into_temp_tables.sql", ForWriting, True, TristateTrue)
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
    txt.WriteLine "  DELETE FROM deployment.current_model;"
    txt.WriteLine "  INSERT INTO deployment.current_model (id_model, nm_repository) SELECT"
    txt.WriteLine "    id_model      = CONVERT(CHAR(32),      '" & l_id_model & "'),"
    txt.WriteLine "    nm_repository = CONVERT(NVARCHAR(128), '" & l_nm_repository & "');"
    txt.WriteLine "  "
    txt.WriteLine "  DELETE FROM deployment.last_deployment;"
    txt.WriteLine "  INSERT INTO deployment.last_deployment (id_model, dt_deployment) SELECT"
    txt.WriteLine "    id_model       = CONVERT(CHAR(32),      '" & l_id_model & "'),"
    txt.WriteLine "    dt_deployment  = (SELECT ISNULL(MAX(meta_dt_valid_from), CONVERT(DATETIME, '1979-01-01')) FROM dta.dataset WHERE id_model = '" & l_id_model & "')"
    txt.WriteLine "  "
    txt.WriteLine "END"
    txt.WriteLine "GO"
    txt.WriteLine ""
    '
    ' Close SQL-File
    txt.Close
    '
    ' Build All SQL Schemes, Tables and Procedures files
    Call all_create_dataset_specified_procedure
    '
End Sub

' Ensure repository export folders exist (create if missing).
' Creates base repos folder plus domain folders (srd/ohg/dta/dqm).
' Uses mdl_Folders.* paths and FileSystemObject.FolderExists/CreateFolder.
' Params: none.
' Example: Call create_folder_structure
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


' Export metadata not directly tied to datasets.
' Writes static reference data (srd), groups/hierarchies (ohg), and DQ model.
' Calls export_table for each fixed table name.
' Params: none.
' Example: Call export_non_direct_related_to_dataset
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

' Export one schema/table to a .sql file under the repo folder.
' Params: nm_schema = schema prefix (srd/ohg/dta/dqm); nm_table = table name.
' Filters by id_model_default() for non-model tables.
' Writes BEGIN/END + INSERT lines via build_sql_insert.
' Example: Call export_table("srd", "datatype")
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

' Build the datasets.sql include file for all exported dataset definition files.
' Writes ":r .\<id_dataset>.sql" lines based on dta_dataset rows for model.
' Output: mdl_Folders.dta() & "datasets.sql".
' Params: none.
' Example: Call build_sql_file_dataset
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
    ' Now remove those file that nolonger have a related dataset in the dta_dataset-table.
    Dim is_found As Boolean
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'")
    Dim ob_files As Files: Set ob_files = fso.GetFolder(mdl_Folders.dta).Files
    Dim ob_file  As file:
    Dim id_dataset  As String
    For Each ob_file In ob_files
      '
      ' extract "id_datasaet"
      id_dataset = Replace(ob_file.Name, ".sql", "")
      '
      ' Exclude database.sql, model.sql and datasets.sql
      If (id_dataset <> "database" And id_dataset <> "model" And id_dataset <> "datasets") Then
        '
        ' Find id_dataset in dta_dataset-recordset
        is_found = False: rst.MoveFirst: Do Until (rst.EOF Or is_found):  is_found = (rst!id_dataset = id_dataset):      rst.MoveNext: Loop
        '
        ' If NOT found then delete the file
        If (is_found = False) Then
          ob_file.Delete True
        End If
      End If
      '
    Next ob_file
    '
End Sub

' Export all datasets and their related definitions to individual .sql files.
' Iterates dta_dataset for current model and calls export_dataset_and_related_definitions.
' Rebuilds datasets.sql include file at the end.
' Params: none.
' Example: Call export_all_dataset_and_related_definitions
Public Sub export_all_dataset_and_related_definitions()
    '
    ' Local Variables
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE id_model = '" & mdl_Folders.id_model(mdl_Folders.nm_repository) & "'")
    '
    ' Export all data to SQL-file.
    Do Until rst.EOF: With rst: export_dataset_and_related_definitions .fields("id_dataset"): .MoveNext: End With: Loop
    '
    ' Build SQL file to exec als dta-dataset-sql-files
    Call build_sql_file_dataset
    '
End Sub

' Export one dataset and its related objects to a single .sql file.
' Param: id_dataset = dataset id to export (skips when empty).
' Writes dta/ohg/dqm rows and transformation-related objects when applicable.
' Output: mdl_Folders.dta() & "<id_dataset>.sql".
' Example: Call export_dataset_and_related_definitions(id_dataset)
Public Sub export_dataset_and_related_definitions(id_dataset As String)
    '
    ' build sql for filtering on model.
    Dim id_model As String: id_model = mdl_Folders.id_model(mdl_Folders.nm_repository)
    Dim tx_where As String: tx_where = "WHERE id_dataset = '" & id_dataset & "' AND id_model = '" & id_model & "'"
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
    Dim rst As Recordset:  Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset " & tx_where)
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
    Call add_to_export_file(txt, tx_where, "dta", "Dataset")
    Call add_to_export_file(txt, tx_where, "dta", "Attribute")
    Call add_to_export_file(txt, tx_where, "dta", "Parameter Value")
    Call add_to_export_file(txt, tx_where, "dta", "Ingestion Etl")
    Call add_to_export_file(txt, tx_where, "dta", "Schedule")
    Call add_to_export_file(txt, tx_where, "ohg", "Related")
    Call add_to_export_file(txt, tx_where, "dqm", "DQ Control")
    '
    ' Re-build WHERE Clause
    tx_where = "WHERE id_dq_control IN (SELECT id_dq_control FROM dqm_dq_control " & tx_where & ") AND id_model = '" & id_model & "'"
    Call add_to_export_file(txt, tx_where, "dqm", "DQ Threshold")
    '
    ' If Transformation (NOT Ingestion) then Export Transformation Parts, Datasets, Column Mappings and Attribute References.
    If (Not rst!is_ingestion) Then
        '
        ' Build WHERE Clause for Transformation Part
        tx_where = "WHERE id_dataset = '" & id_dataset & "' AND id_model = '" & id_model & "'"
        Call add_to_export_file(txt, tx_where, "dta", "Transformation Part")
        '
        ' Build WHERE Clause for Transformation Dataset and Column Mapping
        Dim tx_where_tpt As String:  tx_where_tpt = "WHERE id_transformation_part IN (SELECT id_transformation_part FROM dta_transformation_part " & tx_where & ") AND id_model = '" & id_model & "'"
        Call add_to_export_file(txt, tx_where_tpt, "dta", "Transformation Dataset")
        Call add_to_export_file(txt, tx_where_tpt, "dta", "Transformation Column Mapping")
        Call add_to_export_file(txt, tx_where_tpt, "dta", "Transformation Part Attribute")
        '
        ' Build WHERE Clause for Transformation Dataset and Column Mapping
        Dim tx_where_tds As String:  tx_where_tds = "WHERE id_transformation_dataset IN (SELECT id_transformation_dataset FROM dta_transformation_dataset " & tx_where_tpt & ") AND id_model = '" & id_model & "'"
        Call add_to_export_file(txt, tx_where_tds, "dta", "Transformation Dataset Attribute")
        '
        ' Build WHERE Clause for Transformation Dataset and Column Mapping
        Dim tx_where_tcm As String:  tx_where_tcm = "WHERE id_transformation_column_mapping IN (SELECT id_transformation_column_mapping FROM dta_transformation_column_mapping " & tx_where_tpt & ") AND id_model = '" & id_model & "'"
        Call add_to_export_file(txt, tx_where_tcm, "dta", "Transformation Column Mapping Attribute")
        '
    End If
    '
    ' End the SQL Block
    txt.WriteLine "  "
    txt.WriteLine "END"
    txt.WriteLine "GO"
    txt.WriteLine ""
    '
    ' Close "SQL"-file.
    txt.Close
    '
End Sub

' Append INSERT statements for one table to an open export TextStream.
' Params: ip_ob_txt = output stream; ip_tx_where = WHERE clause incl WHERE.
' ip_nm_schema = schema prefix; ip_nm_table = display name (spaces allowed).
' Writes a small section header and rows via build_sql_insert.
' Example: Call add_to_export_file(txt, "WHERE 1=1", "dta", "Dataset")
Public Sub add_to_export_file(ByRef ip_ob_txt As TextStream, ByVal ip_tx_where As String, ByVal ip_nm_schema As String, ByVal ip_nm_table As String)
    Dim rst As DAO.Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM " & Trim(ip_nm_schema) & "_" & Trim(LCase(Replace(ip_nm_table, " ", "_"))) & " " & ip_tx_where)
    Dim fdt As String:            fdt = "`" & Trim(ip_nm_table) & "`-definitions"
    ip_ob_txt.WriteLine "  /* " & String(Len(fdt), "-") & " */"
    ip_ob_txt.WriteLine "  /* " & fdt & " */"
    ip_ob_txt.WriteLine "  /* " & String(Len(fdt), "-") & " */"
    Do Until rst.EOF: With rst: ip_ob_txt.WriteLine build_sql_insert(Trim(ip_nm_schema), Trim(LCase(Replace(ip_nm_table, " ", "_"))), .fields): .MoveNext: End With: Loop
    If rst.RecordCount = 0 Then: ip_ob_txt.WriteLine "  -- No Defintions for `" & Trim(ip_nm_table) & "`": rst.Close
    ip_ob_txt.WriteLine ""
End Sub

' Build one SQL INSERT statement from a DAO.Fields collection.
' Params: nm_schema/nm_table = target table; fields = current record fields.
' Formats dates/decimals/bools and escapes quotes/newlines for export.
' Returns: INSERT statement text for tsa_<schema>.tsa_<table>.
' Example: sql = build_sql_insert("srd", "datatype", rst.Fields)
Public Function build_sql_insert(ByVal nm_schema As String, ByVal nm_table As String, ByRef fields As fields) As String
    '
    ' Local Variables
    Dim tx_fields As String: tx_fields = ""
    Dim tx_values As String: tx_values = ""
    '
    ' Export record to SQL-file
    Dim fld As Field: For Each fld In fields
        '
        ' Filter Meta-Attributes
        If fld.Name <> "meta_created_at" And fld.Name <> "meta_updated_at" Then
            '
            tx_fields = tx_fields & IIf(tx_fields = "", "", ", ")
            tx_fields = tx_fields & fld.Name
            tx_values = tx_values & IIf(tx_values = "", "", ", ")
            '
            ' Handling Dataypes
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
        End If
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