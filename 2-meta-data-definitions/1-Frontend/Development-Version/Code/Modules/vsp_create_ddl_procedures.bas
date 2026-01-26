Attribute VB_Name = "vsp_create_ddl_procedures"
Option Compare Database
Option Explicit
Public Sub test_create_dataset_specified_procedure()
  Dim id_dataset   As String:  id_dataset = "05040c05070d0d030707080706190d09"
  Dim is_debugging As Boolean: is_debugging = True
  Call build_folder_structure
  Call add_or_update_create_table_sql_file(id_dataset)
  Call create_dataset_specified_procedure(id_dataset, is_debugging)
End Sub

Public Sub all_create_dataset_specified_procedure(): On Error GoTo errHandle
  '
  If (1 = 1) Then ' Remove all Folders and Files
    '
    ' Remove all the files and folder
    Dim fso As FileSystemObject:   Set fso = New FileSystemObject
    Dim fp_datasets As String: fp_datasets = Replace(tx_repo_folderpath(nm_repository), "2-Definitions\", "3-Datasets")
    If (fso.FolderExists(fp_datasets)) Then fso.DeleteFolder fp_datasets, True
    '
    ' Remove Files and Folders
    Call RemoveFolderFromSqlProj(True)
    '
  End If
  '
  If (1 = 1) Then ' Re-Build all Folders and Files
    '
    ' Build Folder structure
    Call build_folder_structure
    '
    ' Build Schemas Folders with ddl-files (this includes tables)
    Call add_all_schemas
    '
    ' Loop through all Dataset
    Dim sql As String: sql = "SELECT id_dataset, nm_target_schema, nm_target_table FROM dta_dataset GROUP BY id_dataset, nm_target_schema, nm_target_table"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql): Do Until rst.EOF: Call create_dataset_specified_procedure(rst!id_dataset, False): rst.MoveNext: Loop
    '
  End If
  '
Exit Sub
'
errHandle:
  Debug.Print "--- Error ------------------------------------------------"
  Debug.Print "Number      : " & CStr(Err.Number)
  Debug.Print "Description : " & Err.Description
  Stop
  Resume
  '
End Sub


Public Sub create_dataset_specified_procedure(ip_id_dataset As String, Optional ip_is_debugging As Boolean = False): On Error GoTo errHandle
  '
  ' Set other Variables
  Dim id_model   As String: id_model = id_model_default()
  Dim id_dataset As String: id_dataset = ip_id_dataset
  '
  ' Local Variables
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim idx As Integer
  Dim max As Integer: max = 100
  '
  ' Local Variable Helpers for building SQL Statemens
  Dim nwl As String: nwl = vbNewLine
  Dim emp As String: emp = ""
  Dim sql As String: sql = ""
  Dim qry As String: qry = ""
  '
  ' Local Variable for helpers with strings
  Dim tb1 As String: tb1 = nwl & "  "
  Dim tb2 As String: tb2 = nwl & "    "
  Dim tb3 As String: tb3 = nwl & "      "
  '
  ' Recorsets for Metadata
  Dim rs_etl As Recordset: Set rs_etl = CurrentDb.OpenRecordset("SELECT * FROM dta_ingestion_etl WHERE id_dataset = '" & ip_id_dataset & "' AND id_model = '" & id_model_default() & "'")
  Dim rs_att As Recordset: Set rs_att = CurrentDb.OpenRecordset("SELECT * FROM dta_attribute     WHERE id_dataset = '" & ip_id_dataset & "' AND id_model = '" & id_model_default() & "' ORDER BY ni_ordering ASC")
  Dim rs_dst As Recordset: Set rs_dst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset       WHERE id_dataset = '" & ip_id_dataset & "' AND id_model = '" & id_model_default() & "'")
  '
  ' Get "Schema"-folder full and relative
  Dim cd_dataset_type As String: cd_dataset_type = get_cd_dataset_type(rs_dst!nm_target_schema, rs_dst!is_ingestion)
  '
  ' Add Tables
  Call add_or_update_create_table_sql_file(rs_dst!id_dataset)
  '
  ' Build Folderpath to Procedures
  Dim rp_folder As String: rp_folder = get_fp_schema(rs_dst!nm_target_schema, cd_dataset_type, True) & "Procedures\": Call AddFolderToSqlProj(rp_folder)
  Dim fp_folder As String: fp_folder = get_fp_schema(rs_dst!nm_target_schema, cd_dataset_type, False) & "Procedures\": Call create_folder_if_not_exists(fp_folder)
  '
  ' Set "Schema/Tables"-names
  Dim tgt As String: tgt = Trim("[    ") & rs_dst!nm_target_schema & Trim("].[    ") & rs_dst!nm_target_table & "]"
  Dim src As String: src = Trim("[tsa_") & rs_dst!nm_target_schema & Trim("].[tsa_") & rs_dst!nm_target_table & "]"
  Dim tsa As String: tsa = Trim("[tsa_") & rs_dst!nm_target_schema & Trim("].[tsa_") & rs_dst!nm_target_table & "]"
  Dim usp As String: usp = Trim("[    ") & rs_dst!nm_target_schema & Trim("].[usp_") & rs_dst!nm_target_table & "]"
  '
  ' Set other Variables
  Dim is_ingestion                  As Boolean: is_ingestion = rs_dst!is_ingestion
  Dim nm_data_flow_type             As String:  nm_data_flow_type = IIf(rs_dst!is_ingestion, "Ingestion", "Transformation")
  Dim tx_source_query               As String:  tx_source_query = Replace(rs_dst!tx_source_query, "<newline>", nwl)
  Dim nm_processing_type            As String:  nm_processing_type = IIf(rs_dst!nm_target_schema = "dq_totals", "Fullload", rs_etl!nm_processing_type)
  Dim tx_sql_for_meta_dt_valid_from As String:  tx_sql_for_meta_dt_valid_from = Nz(rs_etl!tx_sql_for_meta_dt_valid_from, "n/a")
  Dim tx_sql_for_meta_dt_valid_till As String:  tx_sql_for_meta_dt_valid_till = Nz(rs_etl!tx_sql_for_meta_dt_valid_till, "n/a")
  '
  ' Show Extracted Variables
  If (ip_is_debugging) Then
    Debug.Print "/* Extract schema and Table. */"
    Debug.Print "id_dataset                    : " & ip_id_dataset; ""
    Debug.Print "is_ingestion                  : " & CStr(is_ingestion)
    Debug.Print "nm_data_flow_type             : " & Nz(nm_data_flow_type, "n/a")
    Debug.Print "tx_query_source               : " & Nz(tx_source_query, "n/a")
    Debug.Print "nm_processing_type            : " & Nz(nm_processing_type, "n/a")
    Debug.Print "tx_sql_for_meta_dt_valid_from : " & Nz(tx_sql_for_meta_dt_valid_from, "n/a")
    Debug.Print "tx_sql_for_meta_dt_valid_till : " & Nz(tx_sql_for_meta_dt_valid_till, "n/a")
  End If
  '
  ' /* Extract "temp"-table with Columns of "Target"-table, exclude the "meta-attributes. */
  Dim tx_attributes As String: rs_att.MoveFirst: Do Until rs_att.EOF: tx_attributes = tx_attributes & "s.[" & rs_att!nm_target_column & "], ": rs_att.MoveNext: Loop
  Dim tx_pk_fields  As String: rs_att.MoveFirst: Do Until rs_att.EOF: tx_pk_fields = tx_pk_fields & IIf(rs_att!is_businesskey, ", s.[" & rs_att!nm_target_column & "], '|'", ""): rs_att.MoveNext: Loop
  '
  If (1 = 1) Then ' /* Generate the Mapping for meta_ch_rh, meta_ch_bk and meta_ch_pk.*/
    '
    If (1 = 1) Then ' /* Build SQL Statment for Column "meta_ch_rh. */
      Dim att As String: att = ""
      Dim rwh As String: rwh = "CONCAT(CONVERT(NVARCHAR(MAX), '')," & nwl & "  CONCAT("
      idx = 0: rs_att.MoveFirst: Do Until rs_att.EOF
        att = att & "[main].[" & rs_att!nm_target_column & "] AS [" & rs_att!nm_target_column & "]," & nwl
        If (idx = max) Then idx = idx + 1: rwh = rwh & "'|')"
        If (idx > max) Then idx = 0:       rwh = rwh & "," & nwl & " CONCAT("
        If (idx < max) Then idx = idx + 1: rwh = rwh & " '|', [main].[" & rs_att!nm_target_column & "],"
      rs_att.MoveNext: Loop
      rwh = rwh & "'|')" & nwl & "))"
    End If
    '
    If (1 = 1) Then ' /* Build SQL Statment for Column "meta_ch_bk"  and "meta_ch_pk". */
      Dim pks As String
      Dim bks As String: bks = "CONCAT(CONVERT(NVARCHAR(MAX), '')," & nwl & "  CONCAT("
      idx = 0: rs_att.MoveFirst: Do Until rs_att.EOF
        If (rs_att!is_businesskey) Then
          If (idx = max) Then idx = idx + 1: bks = bks & "'|')"
          If (idx > max) Then idx = 0:       bks = bks & "," & nwl & " CONCAT("
          If (idx < max) Then idx = idx + 1: bks = bks & " '|', [main].[" & rs_att!nm_target_column & "],"
        End If
      rs_att.MoveNext: Loop
      pks = bks & " '|', [main].[meta_dt_valid_from], '|')" & nwl & "))"
      bks = bks & "'|')" & nwl & "))"
    End If
    '
  End If
  '
  ' Adding some spaces
  rwh = Replace(rwh, nwl, nwl & "                       ")
  pks = Replace(pks, nwl, nwl & "                       ")
  bks = Replace(bks, nwl, nwl & "                       ")
  '
  ' Show the Generated rhw, bks and pks
  If (ip_is_debugging) Then
    Debug.Print "rwh : " & rwh
    Debug.Print "bks : " & bks
    Debug.Print "pks : " & pks
  End If
  '
  ' /* For Ingestions: Extent the `Source`-query */
  If (is_ingestion) Then
    '
    ' /* Build SQL Statment for "Ingestion"  to handle the "TSL"-table. in correct and desired way. */
    sql = emp & emp & "SELECT"
    sql = sql & nwl & "  " & Replace(att, nwl, tb1)
    sql = sql & emp & "[main].[meta_dt_valid_from] AS [meta_dt_valid_from],"
    sql = sql & nwl & "  [main].[meta_dt_valid_till] AS [meta_dt_valid_till],"
    sql = sql & nwl & "  CONVERT(BIT, CASE WHEN [main].[meta_dt_valid_till] > '9999-12-31' THEN 1 ELSE 0 END) AS [meta_is_active],"
    sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & rwh & ", 2) AS [meta_ch_rh],"
    sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & bks & ", 2) AS [meta_ch_bk],"
    sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & pks & ", 2) AS [meta_ch_pk]"
    sql = sql & nwl & "FROM ("
    sql = sql & nwl & "  " & Replace(tx_source_query, "SELECT ", _
                nwl & "  SELECT meta_dt_valid_from = CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_from & ")" & _
                nwl & "       , meta_dt_valid_till = CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_till & ")" & _
                nwl & "       , ")
    sql = sql & nwl & ") AS [main]"
    '
    ' /* Show SQL Statement for "Ingestion" if in Debugging mode. */
    If (ip_is_debugging) Then
      Debug.Print "/* Extent the `Source`-query */"
      Debug.Print "Ingestion Source Query: " & nwl & sql
    End If
    '
  Else ' /* For Transformations: Extent the "Source"-query */
    '
    ' /* Show Extracted Transformation Parts if in Debugging mode. */
    If (ip_is_debugging) Then Debug.Print "/* Extract `Parts` for the `Transformations`. */"
    '
    ' Local Variables for looping through Transformation Parts
    Dim mx_prt As Integer: mx_prt = DCount("id_transformation_part", "dta_transformation_part", "id_model = '" & id_model & "' AND id_dataset = '" & ip_id_dataset & "'")
    '
    ' Local Variables for Transformation Parts
    Dim is_aggregate_function_used            As Boolean: is_aggregate_function_used = False ' Default set to False
    Dim is_aggregate_function_used_valid_from As Boolean: is_aggregate_function_used_valid_from = False ' Default set to False
    Dim is_aggregate_function_used_valid_till As Boolean: is_aggregate_function_used_valid_till = False  ' Default set to False
    '
    ' /* Declare "Transformation"-part variables. */
    Dim is_utilized_column_used_in_valid_from As Boolean
    Dim is_utilized_column_used_in_valid_till As Boolean
    '
    ' Local Variable for Extractio of "Query-Parts" of the "Transformation-Part"
    Dim ni_pos_from     As Integer: ni_pos_from = 0
    Dim ni_pos_where    As Integer: ni_pos_where = 0
    Dim ni_pos_group_by As Integer: ni_pos_group_by = 0
    Dim ni_pos_having   As Integer: ni_pos_having = 0
    Dim ni_length       As Integer: ni_length = 0
    Dim tx_sql_select   As String:  tx_sql_select = ""
    Dim tx_sql_from     As String:  tx_sql_from = ""
    Dim tx_sql_where    As String:  tx_sql_where = ""
    Dim tx_sql_group_by As String:  tx_sql_group_by = ""
    Dim tx_sql_having   As String:  tx_sql_having = ""
    '
    ' /* Extract "Parts" for the "Transformations". */
    sql = "SELECT * FROM dta_transformation_part WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "' ORDER BY ni_transformation_part ASC"
    Dim tpt As Recordset: Set tpt = CurrentDb.OpenRecordset(sql): Do Until tpt.EOF
      '
      If (1 = 1) Then ' Determine if an Aggregate Function is used in the Transformation Part
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "COUNT(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "SUM(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "AVG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "MAX(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "MIN(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "STDEV(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "STDEVP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "VAR(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "VARP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "GROUPING(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "CHECKSUM_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "STRING_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tpt!tx_transformation_part, " ", "")), "JSON_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
      End If
      '
      If (1 = 1) Then ' Determine if an Aggregate Function is used in the ETL SQL for Meta Dt Valid From
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "COUNT(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "SUM(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "AVG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "MAX(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "MIN(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "STDEV(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "STDEVP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "VAR(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "VARP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "GROUPING(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "CHECKSUM_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "STRING_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_from, " ", "")), "JSON_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
      End If
      '
      If (1 = 1) Then ' Determine if an Aggregate Function is used in the ETL SQL for Meta Dt Valid Till
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "COUNT(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "SUM(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "AVG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "MAX(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "MIN(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "STDEV(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "STDEVP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "VAR(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "VARP(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "GROUPING(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "CHECKSUM_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "STRING_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
        If (InStr(1, UCase(Replace(tx_sql_for_meta_dt_valid_till, " ", "")), "JSON_AGG(", vbTextCompare) > 0) Then is_aggregate_function_used = True
      End If
      '
      If (1 = 1) Then ' /* Calculating postion on FROM, WHERE, GROUP BY and/or HAVING. */
        '
        ni_pos_from = InStr(1, tpt!tx_transformation_part, " FROM ", vbTextCompare)
        ni_pos_where = InStr(1, tpt!tx_transformation_part, "WHERE ", vbTextCompare)
        ni_pos_group_by = InStr(1, tpt!tx_transformation_part, "GROUP BY ", vbTextCompare)
        ni_pos_having = InStr(1, tpt!tx_transformation_part, "HAVING ", vbTextCompare)
        ni_length = Len(tpt!tx_transformation_part) + 1       ' !!! Otherwise the substring at the end will miss 1 character !!!
        '
        If (ni_pos_from = 0) Then
          tx_sql_select = tpt!tx_transformation_part
        Else
          tx_sql_select = Trim(Mid(tpt!tx_transformation_part, 1, ni_pos_from - 1))
        End If
        '
        tx_sql_from = ""
        If (tx_sql_from = "" And ni_pos_from <> 0 And ni_pos_where <> 0) Then tx_sql_from = Mid(tpt!tx_transformation_part, ni_pos_from, ni_pos_where - ni_pos_from)
        If (tx_sql_from = "" And ni_pos_from <> 0 And ni_pos_group_by <> 0) Then tx_sql_from = Mid(tpt!tx_transformation_part, ni_pos_from, ni_pos_group_by - ni_pos_from)
        If (tx_sql_from = "" And ni_pos_from <> 0 And ni_pos_where = 0 And ni_pos_group_by = 0) Then tx_sql_from = Mid(tpt!tx_transformation_part, ni_pos_from, ni_length - ni_pos_from)
        '
        tx_sql_where = ""
        If (tx_sql_where = "" And ni_pos_where <> 0 And ni_pos_group_by <> 0) Then tx_sql_where = Mid(tpt!tx_transformation_part, ni_pos_where, ni_pos_group_by - ni_pos_where)
        If (tx_sql_where = "" And ni_pos_where <> 0 And ni_pos_having <> 0) Then tx_sql_where = Mid(tpt!tx_transformation_part, ni_pos_where, ni_pos_having - ni_pos_where)
        If (tx_sql_where = "" And ni_pos_where <> 0 And ni_pos_group_by = 0 And ni_pos_having = 0) Then tx_sql_where = Mid(tpt!tx_transformation_part, ni_pos_where, ni_length - ni_pos_where)
        '
        tx_sql_group_by = ""
        If (tx_sql_group_by = "" And ni_pos_group_by <> 0 And ni_pos_having <> 0) Then tx_sql_group_by = Mid(tpt!tx_transformation_part, ni_pos_group_by, ni_pos_having - ni_pos_group_by)
        If (tx_sql_group_by = "" And ni_pos_group_by <> 0 And ni_pos_having = 0) Then tx_sql_group_by = Mid(tpt!tx_transformation_part, ni_pos_group_by, ni_length - ni_pos_group_by)
        '
        tx_sql_having = ""
        If (tx_sql_having = "" And ni_pos_having <> 0) Then tx_sql_having = Mid(tpt!tx_transformation_part, ni_pos_having, ni_length - ni_pos_having)
        '
      End If

      If (1 = 1) Then ' /* Determing if the "Transformation"-part is using a "Source"-attributes in ETL valid from/till definitons. */
        '
        ' Initilize the Utilized Column Used in Valid From/Till Flags
        is_utilized_column_used_in_valid_from = False: is_utilized_column_used_in_valid_till = False
        '
        ' Loop through all utilized Source Columns in the Transformation Partyh b /
        sql = emp & emp & "SELECT"
        sql = sql & nwl & "    dta_attribute.nm_target_column AS nm_source_column"
        sql = sql & nwl & "FROM"
        sql = sql & nwl & "    ("
        sql = sql & nwl & "        ("
        sql = sql & nwl & "            dta_transformation_part As prt"
        sql = sql & nwl & "            LEFT JOIN dta_transformation_column_mapping AS map ON (prt.id_model = map.id_model)"
        sql = sql & nwl & "            AND ("
        sql = sql & nwl & "                prt.id_transformation_part = map.id_transformation_part"
        sql = sql & nwl & "            )"
        sql = sql & nwl & "        )"
        sql = sql & nwl & "        LEFT JOIN dta_transformation_column_mapping_attribute AS att ON ("
        sql = sql & nwl & "            map.id_transformation_column_mapping = att.id_transformation_column_mapping"
        sql = sql & nwl & "        )"
        sql = sql & nwl & "        AND (map.id_model = att.id_model)"
        sql = sql & nwl & "    )"
        sql = sql & nwl & "    LEFT JOIN dta_attribute ON ("
        sql = sql & nwl & "        att.id_source_attribute = dta_attribute.id_attribute"
        sql = sql & nwl & "    )"
        sql = sql & nwl & "    AND (att.id_source_model = dta_attribute.id_model)"
        sql = sql & nwl & "WHERE"
        sql = sql & nwl & "    ("
        sql = sql & nwl & "        ("
        sql = sql & nwl & "            (prt.id_model) = '" & id_model & "'"
        sql = sql & nwl & "        )"
        sql = sql & nwl & "        AND ("
        sql = sql & nwl & "            (prt.id_transformation_part) = '" & tpt!id_transformation_part & "'"
        sql = sql & nwl & "        )"
        sql = sql & nwl & "    )"
        Dim rs_map As Recordset: Set rs_map = CurrentDb.OpenRecordset(sql): Do Until rs_map.EOF
          If (InStr(1, tx_sql_for_meta_dt_valid_from, rs_map!nm_source_column, vbTextCompare) > 0) Then is_utilized_column_used_in_valid_from = True
          If (InStr(1, tx_sql_for_meta_dt_valid_till, rs_map!nm_source_column, vbTextCompare) > 0) Then is_utilized_column_used_in_valid_till = True
        rs_map.MoveNext: Loop
        '
      End If
      '
      ' Build SQL Query for the Transformation Part -> Main Source Query
      qry = qry & nwl & "  " & tx_sql_select
      qry = qry & nwl & "       , meta_dt_valid_from = CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_from & ")"
      qry = qry & nwl & "       , meta_dt_valid_till = CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_till & ")"
      If (tx_sql_from <> "") Then qry = qry & nwl & "  " & tx_sql_from
      If (tx_sql_where <> "") Then qry = qry & nwl & "  " & tx_sql_where
      If (tx_sql_group_by <> "") Then
        qry = qry & nwl & "  " & tx_sql_group_by
        If (is_utilized_column_used_in_valid_from And Not is_aggregate_function_used_valid_from) Then qry = qry & nwl & "         , CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_from & ")"
        If (is_utilized_column_used_in_valid_till And Not is_aggregate_function_used_valid_till) Then qry = qry & nwl & "         , CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_till & ")"
      End If
      If (tx_sql_group_by = "" And is_aggregate_function_used) Then
        If (is_utilized_column_used_in_valid_from And Not is_aggregate_function_used_valid_from) Then qry = qry & nwl & IIf(is_utilized_column_used_in_valid_from And Not is_aggregate_function_used_valid_from, "  GROUP BY CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_from & ")", "")
        If (is_utilized_column_used_in_valid_till And Not is_aggregate_function_used_valid_till) Then
          qry = qry & nwl & IIf(Not is_utilized_column_used_in_valid_from Or (is_utilized_column_used_in_valid_from And Not is_aggregate_function_used_valid_from), "  GROUP BY ", "         , ") & "CONVERT(DATETIME, " & tx_sql_for_meta_dt_valid_till & ")"
        End If
      End If
      If (tx_sql_having <> "") Then qry = qry & nwl & "  " & tx_sql_having
      '
      ' Add "UNION ALL" if there is a "Next" "Transformation"-part.
      If ((tpt!ni_transformation_part + 1) <= mx_prt) Then qry = qry & nwl & "  UNION ALL"
      '
      ' Show the Transformation Part and if an Aggregate Function is used
      If (ip_is_debugging) Then
        Debug.Print "/* Show Transformation Part # " & tpt!ni_transformation_part & " - Details */"
        Debug.Print "id_transformation_part                : " & tpt!id_transformation_part
        Debug.Print "ni_transformation_part                : " & tpt!ni_transformation_part
        Debug.Print "tx_transformation_part                : " & tpt!tx_transformation_part
        Debug.Print "is_aggregate_function_used            : " & CStr(is_aggregate_function_used)
        Debug.Print "is_utilized_column_used_in_valid_from : " & CStr(is_utilized_column_used_in_valid_from)
        Debug.Print "is_utilized_column_used_in_valid_till : " & CStr(is_utilized_column_used_in_valid_till)
        Debug.Print "tx_sql_select                         : " & tx_sql_select
        Debug.Print "tx_sql_from                           : " & tx_sql_from
        Debug.Print "tx_sql_where                          : " & tx_sql_where
        Debug.Print "tx_sql_group_by                       : " & tx_sql_group_by
        Debug.Print "tx_sql_having                         : " & tx_sql_having
      End If
      '
    ' Move to next Transformation Part
    tpt.MoveNext: Loop
    '
    If (1 = 1) Then ' Build the "Source"-query for the Dataset-Specific Stored Procedure
      sql = emp & emp & "SELECT"
      sql = sql & nwl & "  " & Replace(att, nwl, tb1)
      sql = sql & emp & "[main].[meta_dt_valid_from] AS [meta_dt_valid_from],"
      sql = sql & nwl & "  [main].[meta_dt_valid_till] AS [meta_dt_valid_till],"
      sql = sql & nwl & "  CONVERT(BIT, 1) AS [meta_is_active],"
      sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & rwh & ", 2) AS [meta_ch_rh],"
      sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & bks & ", 2) AS [meta_ch_bk],"
      sql = sql & nwl & "  CONVERT(CHAR(32), HASHBYTES('MD5', " & pks & ", 2) AS [meta_ch_pk]"
      sql = sql & nwl & "FROM ( -- '" & nm_processing_type & "' processing mode."
      sql = sql & nwl & qry
      sql = sql & nwl & ") AS [main]"
    End If
    '
    If (ip_is_debugging) Then
      Debug.Print "/* Final Source Query for the Dataset-Specific Stored Procedure */"
      Debug.Print "Transformation Query: " & nwl & sql
    End If
    '
  End If
  '
  If (1 = 1) Then ' /* Build SQL Statement for Insert into "Temporal Staging Area" or "Target" table. */
    sql = "INSERT INTO " & tsa & " (" & Replace(tx_attributes, "s.[", "[") & "meta_dt_valid_from, meta_dt_valid_till, meta_is_active, meta_ch_rh, meta_ch_bk, meta_ch_pk)" & nwl & sql
    '
    If (ip_is_debugging) Then
      Debug.Print "/* Insert Statement for Temporal Staging Area */"
      Debug.Print "tx_query_update : " & nwl & sql
    End If
    '
    ' Set the Insert Query for Temporal Staging Area
    Dim tx_query_source As String: tx_query_source = sql
    '
  End If
  '
  If (1 = 1) Then ' /* Build SQL Statements for Update and Insert into "Target" table. */
    sql = emp & emp & "UPDATE t SET"
    sql = sql & nwl & "  t.meta_is_active = 0, t.meta_dt_valid_till = ISNULL(s.meta_dt_valid_from, @dt_current_stand),"
    '
    ' /* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */
    ' /* Extent building SQL Statement for updating "meta_ch_pk", to handle case where data is updated */
    ' /* retrospectively, meaning the validity of the data has NOT changed, this can be due to         */
    ' /* corrections in the data. To handle this, without creating duplicate Primarykeys, the
    ' /* meta_ch_pk must be updated!!                                                                  */
    ' /* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */
    sql = sql & nwl & "  t.meta_ch_pk = CASE WHEN s.meta_ch_pk = t.meta_ch_pk"
    sql = sql & nwl & "                      THEN CONVERT(CHAR(32), HASHBYTES('MD5', CONCAT(t.meta_ch_pk, t.meta_dt_created)), 2)"
    sql = sql & nwl & "                      ELSE S.meta_ch_pk"
    sql = sql & nwl & "                 END"
    ' /* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */
    sql = sql & nwl & "FROM " & tgt & " AS t LEFT JOIN " & src & " AS s ON t.meta_ch_bk = s.meta_ch_bk"
    sql = sql & nwl & "WHERE t.meta_is_active = 1 AND t.meta_ch_rh != ISNULL(s.meta_ch_rh,'n/a')"
    sql = sql & nwl & IIf(nm_processing_type = "Incremental", "AND t.meta_ch_bk IN (SELECT meta_ch_bk FROM " & src & ")", "")
    '
    If (ip_is_debugging) Then
      Debug.Print "/* Update Statement for Target Table */"
      Debug.Print "tx_query_update : " & nwl & sql
    End If
    '
    ' Set the Update Query for Target Table
    Dim tx_query_update As String: tx_query_update = sql
    '
  End If
  '
  If (1 = 1) Then ' /* Build SQL Statement for Insert into "Target" table for "Fullload/Incremental" processing type. */
    sql = emp & "INSERT INTO " & tgt & " (" & Replace(tx_attributes, "s.[", "[") & "meta_dt_valid_from, meta_dt_valid_till, meta_is_active, meta_ch_rh, meta_ch_bk, meta_ch_pk)" & nwl
    sql = sql & "SELECT " & tx_attributes & " s.meta_dt_valid_from, s.meta_dt_valid_till, s.meta_is_active, s.meta_ch_rh, s.meta_ch_bk, s.meta_ch_pk" & nwl
    sql = sql & "FROM " & src & " AS s LEFT JOIN " & tgt & " AS t ON t.meta_is_active = 1 AND t.meta_ch_rh = s.meta_ch_rh" & nwl
    sql = sql & "WHERE t.meta_ch_pk IS NULL"
    '
    If (ip_is_debugging) Then
      Debug.Print "/* Insert Statement for Target Table */"
      Debug.Print "tx_query_insert : " & nwl & sql
    End If
    '
    ' Set the Insert Query for Target Table
    Dim tx_query_insert As String: tx_query_insert = sql
    '
  End If
  '
  If (1 = 1) Then ' /* Build SQL Statement for "Calculation"-dates */
    '
    ' Build SQL Statement for Calculation-dates
    sql = emp & emp & "  /* Initialization of the `Run` in the `rdp.run_start`, the  `Previous Stand` is Determined based on meta_dt_valid_from and meta_dt_valid_till, hereby `9999-12-31` and greater are excluded. */"
    sql = sql & nwl & "  SELECT @dt_previous_stand = CONVERT(DATETIME2(7), MAX(run.dt_previous_stand))"
    sql = sql & nwl & "       , @dt_current_stand  = CONVERT(DATETIME2(7), MAX(run.dt_current_stand))"
    sql = sql & nwl & "  FROM rdp.run AS run"
    sql = sql & nwl & "  WHERE run.id_model   = '" & id_model & "'"
    sql = sql & nwl & "  AND   run.id_dataset = '" & id_dataset & "'"
    sql = sql & nwl & "  AND   run.dt_run_started = ("
    sql = sql & nwl & "    /* Find the `Previous` run that NOT ended in `Failed`-status. */"
    sql = sql & nwl & "    SELECT MAX(dt_run_started)"
    sql = sql & nwl & "    FROM rdp.run"
    sql = sql & nwl & "    WHERE id_model             = '" & id_model & "'"
    sql = sql & nwl & "    AND   id_dataset           = '" & id_dataset & "'"
    sql = sql & nwl & "    AND   id_processing_status = gnc.id_processing_status('" & id_model & "', 'Finished')"
    sql = sql & nwl & "  )"
    If (nm_processing_type = "Incremental") Then
    sql = sql & nwl & "      IF (@is_override_fullload <> 0) BEGIN SET @dt_previous_stand = CONVERT(DATETIME2(7), '1970-01-01') END"
    sql = sql & nwl & "      "
    End If
    '
    ' /* Set SQL Statement for "Calculation"-dates */
    Dim tx_query_calculation As String: tx_query_calculation = sql
    '
  End If
  '
  If (1 = 1) Then ' /* Build SQL Statemen for creation of "Stored Procedure" */
    '
    ' /* Build SQL Statement for creation of "Stored Procedure" */
    sql = emp & emp & "CREATE PROCEDURE " & usp & " "
    If (is_ingestion) Then sql = sql & emp & " AS "
    If (Not is_ingestion) Then ' /* In case of "Transformation" */
      sql = sql & emp & "  /* Input Parameter(s) */"
      sql = sql & nwl & "  @ip_ds_external_reference_id NVARCHAR(999) = 'n/a'"
      sql = sql & nwl & " AS "
    End If
    sql = sql & nwl & ""
    sql = sql & nwl & "/* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */"
    sql = sql & nwl & "/* !!!                                                                            !!! */"
    sql = sql & nwl & "/* !!! This Stored Procdures has been generated by excuting the procedure of      !!! */"
    sql = sql & nwl & "/* !!! VBA module vsp_create_ddl_procedures.create_dataset_specified_procedure    !!! */"
    sql = sql & nwl & "/* !!!                                                                            !!! */"
    sql = sql & nwl & "/* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */"
    sql = sql & nwl & "/* "
    sql = sql & nwl & "-- VBA Example for 'Generation of " & nm_data_flow_type & " Procedure':"
    sql = sql & nwl & "CALL vsp_create_ddl_procedures.create_dataset_specified_procedure("
    sql = sql & nwl & "  """ & id_dataset & """"
    sql = sql & nwl & ")"
    sql = sql & nwl & ""
    sql = sql & nwl & "-- SQL Example for 'Executing the " & nm_data_flow_type & " Procedure':"
    sql = sql & nwl & "EXEC " & usp & ""
    sql = sql & nwl & "GO"
    sql = sql & nwl & ""
    sql = sql & nwl & "*/"
    sql = sql & nwl & "/* !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! */"
    sql = sql & nwl & ""
    sql = sql & nwl & "DECLARE /* Local Variables */"
    sql = sql & nwl & "  @id_dataset           CHAR(32)      = '" & id_dataset & "', "
    sql = sql & nwl & "  @nm_target_schema     NVARCHAR(128) = '" & rs_dst!nm_target_schema & "', "
    sql = sql & nwl & "  @nm_target_table      NVARCHAR(128) = '" & rs_dst!nm_target_table & "', "
    sql = sql & nwl & "  @tx_error_message     NVARCHAR(MAX),"
    sql = sql & nwl & "  @dt_previous_stand    DATETIME2(7),"
    sql = sql & nwl & "  @dt_current_stand     DATETIME2(7),"
    sql = sql & nwl & "  @id_run               CHAR(32)       = NULL,"
    sql = sql & nwl & "  @is_transaction       BIT            = 0, -- Helper to keep track if a transaction has been started."
    sql = sql & nwl & "  @ni_before            INT            = 0, -- # Record 'Before' processing."
    sql = sql & nwl & "  @ni_ingested          INT            = 0, -- # Record that were 'Ingested'."
    sql = sql & nwl & "  @ni_inserted          INT            = 0, -- # Record that were 'Inserted'."
    sql = sql & nwl & "  @ni_updated           INT            = 0, -- # Record that were 'Updated'."
    sql = sql & nwl & "  @ni_after             INT            = 0, -- # Record 'After' processing."
    If (nm_processing_type = "Incremental") Then
    sql = sql & nwl & "  @is_override_fullload INT            = 0,"
    End If
    sql = sql & nwl & "  @cd_procedure_step    NVARCHAR(32),"
    sql = sql & nwl & "  @ds_procedure_step    NVARCHAR(999)"
    sql = sql & nwl & "  "
    sql = sql & nwl & "BEGIN"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  /* Turn off Effected Row */"
    sql = sql & nwl & "  SET NOCOUNT ON"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  /* Turn off Warnings */"
    sql = sql & nwl & "  SET ANSI_WARNINGS OFF"
    sql = sql & nwl & "  "
    If (Not is_ingestion) Then ' /* In case of 'Transformation' */
    sql = sql & nwl & "  SET @cd_procedure_step = 'STR'"
    sql = sql & nwl & "  IF (1=1) BEGIN SET @ds_procedure_step = 'Start Run (only for `Transformations` needed, with `Ingestions` the `Run` is started via the `orchastration`-tool like `Azure Data Factory` for instance.)'"
    sql = sql & nwl & "    EXEC rdp.run_start '" & id_model & "', '" & id_dataset & "', @ip_ds_external_reference_id"
    sql = sql & nwl & "  END"
    sql = sql & nwl & "  "
    End If
    If (nm_processing_type = "Incremental") Then
    sql = sql & nwl & "  IF (1=1 ) BEGIN SET @ds_procedure_step = 'Set Override after Dataset has been modified on ider the Parameters or Source Query.'"
    sql = sql & nwl & "    SET @is_override_fullload = ("
    sql = sql & nwl & "      SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END"
    sql = sql & nwl & "      FROM dta.parameter_value AS pv"
    sql = sql & nwl & "      WHERE pv.id_dataset     = '" & id_dataset & "'"
    sql = sql & nwl & "      AND   pv.meta_is_active = 1"
    sql = sql & nwl & "      AND   pv.meta_dt_valid_from > (SELECT MAX(dt_previous_stand) FROM rdp.run WHERE id_dataset = pv.id_dataset)"
    sql = sql & nwl & "    )"
    sql = sql & nwl & "    IF (@is_override_fullload = 0) BEGIN"
    sql = sql & nwl & "      SET @is_override_fullload = ("
    sql = sql & nwl & "        SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END "
    sql = sql & nwl & "        FROM ("
    sql = sql & nwl & "            SELECT id_dataset, meta_dt_valid_from, meta_is_active, tx_source_query AS tx_curr_source_query,"
    sql = sql & nwl & "                   LEAD(tx_source_query) OVER (PARTITION BY id_dataset ORDER BY meta_dt_valid_from DESC) AS tx_prev_source_query"
    sql = sql & nwl & "            FROM dta.dataset AS dst"
    sql = sql & nwl & "        ) AS sq"
    sql = sql & nwl & "        WHERE sq.id_dataset            = '" & id_dataset & "'"
    sql = sql & nwl & "        AND   sq.meta_is_active        = 1"
    sql = sql & nwl & "        AND   sq.meta_dt_valid_from   >= (SELECT MAX(dt_previous_stand) FROM rdp.run WHERE run.id_dataset = sq.id_dataset)"
    sql = sql & nwl & "        AND   sq.tx_prev_source_query IS NOT NULL"
    sql = sql & nwl & "        AND   sq.tx_prev_source_query != sq.tx_curr_source_query"
    sql = sql & nwl & "      )"
    sql = sql & nwl & "    END"
    sql = sql & nwl & "  END"
    sql = sql & nwl & "  "
    End If
    sql = sql & nwl & "  IF (1=1 /* Extract 'Last' calculation datetime. */) BEGIN"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    " & Replace(tx_query_calculation, nwl, tb2)
    sql = sql & nwl & "    "
    sql = sql & nwl & "  END"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  /* Calculate # Records 'before' Processing. */"
    sql = sql & nwl & "  SELECT @ni_before = COUNT(1) FROM " & tgt & " WHERE [meta_is_active] = 1"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  SET @cd_procedure_step = 'SRC'"
    sql = sql & nwl & "  IF (1=1) BEGIN SET @ds_procedure_step = 'Execute `Source`-query and insert result into `Temporal Staging Area`-table.'"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    BEGIN TRY"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* 'Truncate of the 'Temporal (Landing and/or) Staging Area'-table(s). */"
    sql = sql & nwl & "      TRUNCATE TABLE " & tsa & ""
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Start the 'Transaction'. */"
    sql = sql & nwl & "      BEGIN TRANSACTION SET @is_transaction = 1"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      " & Replace(tx_query_source, nwl, tb3)
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Set # Ended records. */"
    sql = sql & nwl & "      SET @ni_ingested = @@ROWCOUNT"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Commit the 'Transaction'. */"
    sql = sql & nwl & "      COMMIT TRANSACTION SET @is_transaction = 0"
    sql = sql & nwl & "      "
    sql = sql & nwl & "    END TRY"
    sql = sql & nwl & "    BEGIN CATCH"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* An 'Error' occured', rollback the transaction and register the 'Error' in the Logging. */"
    sql = sql & nwl & "      IF (TRANCOUNT > 0) BEGIN ROLLBACK TRANSACTION EXEC rdp.run_failed '" & id_model & "', '" & id_dataset & "' END"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      IF (id_run IS NULL) BEGIN"
    sql = sql & nwl & "        SET @tx_error_message = 'ERROR: Loading of data to `Temporal Staging Area`-table `" & tgt & "` failed!'"
    sql = sql & nwl & "        RAISERROR(@tx_error_message, 18, 1)"
    sql = sql & nwl & "      END"
    sql = sql & nwl & "    END CATCH  "
    sql = sql & nwl & "  END"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  BEGIN TRY"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    SET @cd_procedure_step = 'RUN'"
    sql = sql & nwl & "    IF (1=1) BEGIN SET @ds_procedure_step = 'Check that there is an `Run-Dataset`-process running.'"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Fetch the Latest 'Run ID'. */"
    sql = sql & nwl & "      SET @id_run = rdp.get_id_run('" & id_model & "', '" & id_dataset & "')"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Raise Error to indicate that the process of 'Adding' and 'Ending' of records was not logged as started! */"
    sql = sql & nwl & "      IF (@id_run IS NULL) BEGIN"
    sql = sql & nwl & "        SET @tx_error_message = 'ERROR: NO running `process` for dataset `" & tgt & "`!'"
    sql = sql & nwl & "        RAISERROR(@tx_error_message, 18, 1)"
    sql = sql & nwl & "      END"
    sql = sql & nwl & "      "
    sql = sql & nwl & "    END"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Start the 'Transaction'. */"
    sql = sql & nwl & "    BEGIN TRANSACTION SET @is_transaction = 1"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    SET @cd_procedure_step = 'END'"
    sql = sql & nwl & "    IF (1=1) BEGIN SET @ds_procedure_step = '`End` records that are nolonger in `Source` and still in `Target`.'"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      " & Replace(tx_query_update, nwl, tb3)
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Set # Ended records. */"
    sql = sql & nwl & "      SET @ni_updated = @@ROWCOUNT"
    sql = sql & nwl & "      "
    sql = sql & nwl & "    END"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    SET @cd_procedure_step = 'ADD'"
    sql = sql & nwl & "    IF (1=1) BEGIN SET @ds_procedure_step = '`Add` records that are in the `Source` and NOT in `Target`.'"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      " & Replace(tx_query_insert, nwl, tb3)
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Set # Added records. */"
    sql = sql & nwl & "      SET @ni_inserted = @@ROWCOUNT"
    sql = sql & nwl & "      "
    sql = sql & nwl & "    END"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Calculate # Records 'After' Processing. */"
    sql = sql & nwl & "    SELECT @ni_after = COUNT(1) FROM " & tgt & " WHERE [meta_is_active] = 1"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    IF (1=1 /* Validate uniqueness of meta_ch_pk, if NOT Unique then Raise ERROR and rollback !!! */) BEGIN"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      /* Local Variable for Executing Check(s) */"
    sql = sql & nwl & "      DECLARE @ni_expected INT, @ni_measured INT"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      SET @cd_procedure_step = 'CPK'"
    sql = sql & nwl & "      IF (1=1) BEGIN SET @ds_procedure_step = 'Execute Check if `meta_ch_pk`-attribute values are unique.'"
    sql = sql & nwl & "        SELECT @ni_expected = COUNT(1), @ni_measured = COUNT(DISTINCT meta_ch_pk) FROM " & tgt & ""
    sql = sql & nwl & "        IF (@ni_expected != @ni_measured) BEGIN"
    sql = sql & nwl & "            SET @tx_error_message  = 'ERROR: meta_ch_pk NOT unique for `" & tgt & "`!' + CHAR(10) + '-- SQL Statement:'"
    sql = sql & nwl & "            SET @tx_error_message += CHAR(10) + 'SELECT * FROM " & tsa & "'"
    sql = sql & nwl & "            SET @tx_error_message += CHAR(10) + 'WHERE meta_ch_pk IN (SELECT meta_ch_pk FROM " & tsa & " GROUP BY meta_ch_pk HAVING COUNT(*) > 1)'"
    sql = sql & nwl & "            RAISERROR(@tx_error_message, 18, 1)"
    sql = sql & nwl & "        END"
    sql = sql & nwl & "      END"
    sql = sql & nwl & "      "
    sql = sql & nwl & "      SET @cd_procedure_step = 'APK'"
    sql = sql & nwl & "      IF (1=1) BEGIN SET @ds_procedure_step = 'Accuracy only 1 `Active` record per `Primarykey`.'"
    sql = sql & nwl & "        SELECT @ni_expected = COUNT(         CONCAT('|'" & tx_pk_fields & "))"
    sql = sql & nwl & "             , @ni_measured = COUNT(DISTINCT CONCAT('|'" & tx_pk_fields & "))"
    sql = sql & nwl & "        FROM " & tgt & " AS s"
    sql = sql & nwl & "        WHERE s.meta_is_active = 1"
    sql = sql & nwl & "        IF (@ni_expected != ni_measured) BEGIN"
    sql = sql & nwl & "            SET @tx_error_message  = 'ERROR: There should only be 1 record per `Primarykey(s)` for " & tgt & "!'"
    sql = sql & nwl & "            RAISERROR(@tx_error_message, 18, 1)"
    sql = sql & nwl & "        END"
    sql = sql & nwl & "      END"
    sql = sql & nwl & "      "
    sql = sql & nwl & "    END"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Commit the 'Transaction'. */"
    sql = sql & nwl & "    COMMIT TRANSACTION SET @is_transaction = 0"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Cleanup of the 'Temporal (Landing and/or) Staging Area'-table(s). */"
    sql = sql & nwl & "    TRUNCATE TABLE " & tsa & ""
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Set Run Dataset to Success */"
    sql = sql & nwl & "    EXEC rdp.run_finish '" & id_model & "', '" & id_dataset & "', @ni_before, @ni_ingested, @ni_inserted, @ni_updated, @ni_after"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* All done */"
    sql = sql & nwl & "    PRINT('Data Ingestion for Dataset `" & tgt & "` has been successfull.')"
    sql = sql & nwl & "    "
    sql = sql & nwl & "  END TRY"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  BEGIN CATCH"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* An 'Error' occured', rollback the transaction and register the 'Error' in the Logging. */"
    sql = sql & nwl & "    IF (TRANCOUNT > 0) BEGIN ROLLBACK TRANSACTION EXEC rdp.run_failed '" & id_model & "', '" & id_dataset & "'; END"
    sql = sql & nwl & "    "
    sql = sql & nwl & "    /* Ended in 'Error' ! */"
    sql = sql & nwl & "    PRINT('Data `" & IIf(is_ingestion = 1, "Ingestion", "Transformation") & "` for Dataset `" & tgt & "` has ended in `Error`.')"
    sql = sql & nwl & "    PRINT(ISNULL(tx_error_message, 'ERROR (' + cd_procedure_step + ') : ' & ds_procedure_step))"
    sql = sql & nwl & "    "
    sql = sql & nwl & "  END CATCH"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  /* Update 'Documentation' */"
    sql = sql & nwl & "  EXEC documentation.usp_build_html_file_dataset '" & id_model & "', '" & id_dataset & "';"
    sql = sql & nwl & "  "
    sql = sql & nwl & "END"
    '
    ' Show SQL Statement
    If (ip_is_debugging) Then
      Debug.Print "SQL Stored Procedure : " & nwl & sql
    End If
    '
  End If
  '
  Dim fil As TextStream
  Dim rp_file As String: rp_file = rp_folder & "usp_" & rs_dst!nm_target_table & ".sql": Call AddSqlFileToSqlProj(rp_file)
  Dim fp_file As String: fp_file = fp_folder & "usp_" & rs_dst!nm_target_table & ".sql": Set fil = fso.OpenTextFile(fp_file, ForWriting, True, TristateTrue)
  fil.Write sql: fil.Close
  '
Exit Sub
'
errHandle:
  Debug.Print "--- Error ------------------------------------------------"
  Debug.Print "Number      : " & CStr(Err.Number)
  Debug.Print "Description : " & Err.Description
  Stop
  Resume
  '
End Sub