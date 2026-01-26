Attribute VB_Name = "sql_dq_generation"
Option Compare Database
Option Explicit

' Dataset Structure for DQ Control Transformations
Public Type typ_dataset
    id_model              As String
    id_dataset            As String
    id_development_status As String
    id_group              As String
    is_ingestion          As Boolean
    fn_dataset            As String
    fd_dataset            As String
    nm_target_schema      As String
    nm_target_table       As String
    tx_source_query       As String
End Type

' Build DQ transformation datasets for all DQ controls.
' Clears old DQ datasets in dta_dataset (nm_target_schema LIKE 'dq*').
' Calls sql_insert_or_update_dq_control_transformation for each control.
' Params: none.
' Example: Call sql_insert_or_update_dq_control_transformation_all
Public Sub sql_insert_or_update_dq_control_transformation_all()
  On Error GoTo ErrorHandling
  '
  ' Drop all the existing "Transformation" related to "Data Quality"
  Dim del As String: del = "DELETE FROM dta_dataset WHERE nm_target_schema LIKE 'dq*'"
  DoCmd.SetWarnings False: DoCmd.RunSQL del: DoCmd.SetWarnings True
  '
  ' /* Declare Local Variables */
  Dim dqc As DAO.Recordset: Set dqc = CurrentDb.OpenRecordset("SELECT * FROM dqm_dq_control WHERE LEN(nz(tx_dq_control_query,'')) > 10 ORDER BY fd_dq_control ASC") ' -> DQ Control Recordset
  '
  ' /* Process all DQ Controls */
  Do While Not dqc.EOF:
    Debug.Print ("DQ Control:" & dqc!fn_dq_control)
    Call sql_insert_or_update_dq_control_transformation(dqc!id_dq_control, False)
  dqc.MoveNext: Loop
  '
  ' All is Well
  Exit Sub
  '
ErrorHandling:
  '
  ' Print Error
  Debug.Print ("--- Error --------------------------------------------------------------")
  Debug.Print ("Err Number      : " & CStr(Err.Number))
  Debug.Print ("Err Description : " & Err.Description)
  Stop
  Resume
  '
End Sub

' Return id_dq_result_status from srd_dq_result_status.
' Params: ip_id_model = model id; ip_fn_dq_result_status = code (OKE/NOK/OOS).
' Uses CurrentDb.OpenRecordset and returns the first matching id.
' Example: tsa_id_dq_result_status(id_model_default(), "OKE")
' Note: errors if no record matches the given inputs.
Public Function tsa_id_dq_result_status(ip_id_model As String, ip_fn_dq_result_status As String) As String
  '
  ' /* Declare Local Variables */
  Dim sql As String:            sql = "SELECT id_dq_result_status FROM srd_dq_result_status WHERE id_model = '" & ip_id_model & "' AND fn_dq_result_status = '" & ip_fn_dq_result_status & "'"
  Dim rst As DAO.Recordset: Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

  ' /* Return Result */
  tsa_id_dq_result_status = rst!id_dq_result_status

End Function

' Create default DQ groups used by generated DQ datasets.
' Ensures group codes dqg-001..dqg-022 exist for results/totals.
' Exports the ohg.group definitions to the repository.
' Params: none.
' Example: Call create_dq_group
Public Sub create_dq_group()
  '
  ' Add DQ Groups
  add_group_if_not_exists id_model_default, "dqg-001 result", "Data Quality Group for Results of DQ Control Transformations"
  add_group_if_not_exists id_model_default, "dqg-002-totals", "Data Quality Group for Totals of DQ Results"
  add_group_if_not_exists id_model_default, "dqg-011-aggregates-results", "Data Quality Group for Aggregetes of all DQ Results"
  add_group_if_not_exists id_model_default, "dqg-012-aggregates-totals", "Data Quality Group for Aggregates of all DQ Totals"
  add_group_if_not_exists id_model_default, "dqg-021-all-results", "Data Quality Group for All Results"
  add_group_if_not_exists id_model_default, "dqg-022-all-totals", "Data Quality Group for All Totals"
  '
  ' Export ohg_group definitions
  mdl_Export.export_table "ohg", "group"
  '
End Sub
'
'
' Drop table if it exists (temporal/temp helper).
' Param: ip_nm_table = table name in CurrentDb.TableDefs.
' Deletes TableDef when found; otherwise does nothing.
' Returns: none.
' Example: Call drop_temporal_table("tmp_dq_aggregate_level")
Public Sub drop_temporal_table(ip_nm_table As String)
    '
    ' Remove "Temporal"-table if exists
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = ip_nm_table Then CurrentDb.TableDefs.Delete ip_nm_table
        Exit For
    Next
    '
End Sub
'
'
' Replace status codes (OK/NOK/OOS) in a DQ query with status IDs.
' Param: ip_tx_dq_control_query = SQL text containing THEN/ELSE status codes.
' Returns adjusted SQL with ids from tsa_id_dq_result_status.
' Keeps other SQL unchanged; handles multiple case variants.
' Example: sql = replace_dq_result_status_with_id(sql)
Public Function replace_dq_result_status_with_id(ip_tx_dq_control_query As String) As String
    '
    ' Replace DQ Result Status values with their corresponding IDs
    Dim tx_dq_control_query As String: tx_dq_control_query = ip_tx_dq_control_query
    Dim id_model            As String: id_model = id_model_default()
    Dim id                  As String
        
    If (1 = 1) Then ' /* Convert "OKE" to "ID". */)
      id = "THEN '" & tsa_id_dq_result_status(id_model, "OKE") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'ok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'Ok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OKE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'Oke'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oKe'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'okE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OKe'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oKE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OkE'", id)
    End If
    If (1 = 1) Then ' /* Convert "NOK" to "ID". */
      id = "THEN '" & tsa_id_dq_result_status(id_model, "NOK") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'NOK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'nok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'Nok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'nOk'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'noK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'NOk'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'nOK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'NoK'", id)
    End If
    If (1 = 1) Then ' /* Convert "OOS" to "ID". */
      id = "THEN '" & tsa_id_dq_result_status(id_model, "OOS") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OOS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oos'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'Oos'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oOs'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'ooS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OOs'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'oOS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "THEN 'OoS'", id)
    End If
    '
    If (1 = 1) Then ' /* Convert "OKE" to "ID". */)
      id = "ELSE '" & tsa_id_dq_result_status(id_model, "OKE") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'ok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'Ok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OKE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'Oke'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oKe'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'okE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OKe'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oKE'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OkE'", id)
    End If
    If (1 = 1) Then ' /* Convert "NOK" to "ID". */
      id = "ELSE '" & tsa_id_dq_result_status(id_model, "NOK") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'NOK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'nok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'Nok'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'nOk'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'noK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'NOk'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'nOK'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'NoK'", id)
    End If
    If (1 = 1) Then ' /* Convert "OOS" to "ID". */
      id = "ELSE '" & tsa_id_dq_result_status(id_model, "OOS") & "'"
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OOS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oos'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'Oos'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oOs'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'ooS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OOs'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'oOS'", id)
      tx_dq_control_query = Replace(tx_dq_control_query, "ELSE 'OoS'", id)
    End If
    '
    ' Return Result adjusted DQ Control Query
    replace_dq_result_status_with_id = tx_dq_control_query
    '
End Function

' Lookup id_datatype from srd_datatype using a target datatype code.
' Param: ip_cd_target_datatype = target datatype (e.g. INT, DATE).
' Returns the matching id_datatype for id_model_default().
' Errors if no record matches the given datatype code.
' Example: id = get_id_datatype("INT")
Public Function get_id_datatype(ip_cd_target_datatype As String) As String
    '
    ' Get ID Datatype from CD Datatype
    Dim sql As String:            sql = "SELECT id_datatype FROM srd_datatype WHERE id_model = '" & id_model_default() & "' AND cd_target_datatype = '" & ip_cd_target_datatype & "'"
    Dim rst As DAO.Recordset: Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    '
    ' Return Result
    get_id_datatype = rst!id_datatype
    '
End Function

' Debug helper: prints a recordset's fields to Immediate Window.
' Params: ip_rst = recordset; ip_tx_message = header label.
' Optional ip_is_debugging = True to print, otherwise no output.
' Returns: none.
' Example: Call show_recordset_record(rst, "Dataset", True)
Public Sub show_recordset_record(ByRef ip_rst As Recordset, ip_tx_message As String, Optional ip_is_debugging As Boolean = False)
    If ip_is_debugging Then
        Debug.Print ""
        Debug.Print "--- " & ip_tx_message & " " & String(30 - (Len(ip_tx_message) + 5), "-")
        Dim fld As Field: For Each fld In ip_rst.fields: Debug.Print fld.Name & " : " & CStr(Nz(fld.Value, "n/a")): Next fld
        Debug.Print String(30, "-")
    End If
End Sub
'
'
' Validate DQ control SQL matches the expected generic SELECT...CASE...FROM.
' Param: ip_tx_dq_control_query = original SQL text (any whitespace).
' Returns True if pattern matches, else False and shows a MsgBox.
' Uses MinifySQL + RegExMatch with a fixed pattern.
' Example: ok = check_general_structue_dq_control_query(sql)
Public Function check_general_structue_dq_control_query(ip_tx_dq_control_query As String) As Boolean
  '
  ' Define Local Variables
  Dim msg As String
  Dim nwl As String: nwl = vbNewLine
  Dim sql As String: sql = MinifySQL(ip_tx_dq_control_query)
  '
  ' Build RegEx Pattern
  Dim rex As String: rex = "^SELECT" _
      & "\s+.+\s+AS\s+dt_dq_result," _
      & "\s+.+\s+AS\s+id_dataset_1_bk," _
      & "\s+.+\s+AS\s+id_dataset_2_bk," _
      & "\s+.+\s+AS\s+id_dataset_3_bk," _
      & "\s+.+\s+AS\s+id_dataset_4_bk," _
      & "\s+.+\s+AS\s+id_dataset_5_bk," _
      & "\s+CASE\s+WHEN\s+.+\s+THEN\s+.+\s+ELSE\s+.+\s+END\s+AS\s+id_dq_result_status" _
      & "\s+FROM\s+.+$"
  '
  ' Match RegEx pattern with provided Query
  If RegExMatch(sql, rex) = True Then
    check_general_structue_dq_control_query = True
  Else
    check_general_structue_dq_control_query = False
      msg = "DQ Control Query has not the correct format!" & _
      nwl & "Expected format for query: " & _
      nwl & "SELECT" & _
      nwl & "% AS dt_dq_result," & _
      nwl & "% AS id_dataset_1_bk," & _
      nwl & "% AS id_dataset_2_bk," & _
      nwl & "% AS id_dataset_3_bk," & _
      nwl & "% AS id_dataset_4_bk," & _
      nwl & "% AS id_dataset_5_bk," & _
      nwl & "%CASE%WHEN%ELSE%END AS id_dq_result_status" & _
      nwl & "%FROM%;" & _
      nwl & "DQ Control Query : """ & ip_tx_dq_control_query & """"
      MsgBox msg, vbCritical, "Invalid DQ Control Query"
  End If
  '
End Function

' Add one attribute row to the provided dta_attribute Recordset.
' Params: att=target Recordset; id_model/id_dataset/id_datatype; fields+flags.
' Sets id_attribute via CreateMD5, then writes ordering/name/column metadata.
' Calls show_recordset_record when ip_is_debugging=True, then att.Update.
' Example: add_dqc_dataset_attribute att, m, d, dt, 1, fn, fd, col, True, False
Public Sub add_dqc_dataset_attribute(ByRef att As Recordset, _
                                     ByRef id_model As String, _
                                     ByRef id_dataset As String, _
                                     ByRef id_datatype As String, _
                                     ByRef ni_ordering As Integer, _
                                     ByRef fn_attribute As String, _
                                     ByRef fd_attribute As String, _
                                     ByRef nm_target_column As String, _
                                     ByRef is_businesskey As Boolean, _
                                     ByRef is_nullable As Boolean, _
                            Optional ByRef ip_is_debugging As Boolean = False)
  att.AddNew
  att!id_model = id_model
  att!id_dataset = id_dataset
  att!id_datatype = id_datatype
  att!ni_ordering = ni_ordering
  att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(ni_ordering) & "|")
  att!fn_attribute = fn_attribute
  att!fd_attribute = fd_attribute
  att!nm_target_column = nm_target_column
  att!is_businesskey = is_businesskey
  att!is_nullable = is_nullable
  Call show_recordset_record(att, "Attribute", ip_is_debugging)
  att.Update
End Sub

' Insert/update one dataset definition plus its attributes/ingestion metadata.
' Params: id_model/id_dataset/id_development_status; names; target; source SQL.
' Creates rows in dta_dataset/dta_attribute/dta_ingestion_etl, then exports.
' Optional ip_is_debugging prints parameter + recordset details.
' Example: Call add_dqc_dataset(m, d, s, fn, fd, sch, tbl, sql, "result")
Public Sub add_dqc_dataset(ByRef id_model As String, _
                           ByRef id_dataset As String, _
                           ByRef id_development_status, _
                           ByRef fn_dataset As String, _
                           ByRef fd_dataset As String, _
                           ByRef nm_target_schema As String, _
                           ByRef nm_target_table As String, _
                           ByRef tx_source_query As String, _
                           ByRef dq_dataset_type As String, _
                           Optional ip_is_debugging As Boolean = False): On Error GoTo ErrorHandling
  '
  Dim sql As String
  '
  ' Show Input Parameters
  Debug.Print ("--- Add DQ Control Dataset ----------------------------------------------")
  Debug.Print ("dq_dataset_type       : " & dq_dataset_type)
  Debug.Print ("id_model              : " & id_model)
  Debug.Print ("id_dataset            : " & id_dataset)
  Debug.Print ("id_development_status : " & id_development_status)
  Debug.Print ("fn_dataset            : " & fn_dataset)
  Debug.Print ("fd_dataset            : " & fd_dataset)
  Debug.Print ("nm_target_schema      : " & nm_target_schema)
  Debug.Print ("nm_target_table       : " & nm_target_table)
  Debug.Print ("tx_source_query       : " & tx_source_query)
  Debug.Print ("--------------------------------------------------------------------------")
  '
  ' Drop "Existing" Dataset/Attribute/Ingestion_ETL definitions for the DQ Control Dataset
  DoCmd.SetWarnings False
  sql = "DELETE * FROM dta_ingestion_etl WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'": DoCmd.RunSQL sql
  sql = "DELETE * FROM dta_attribute     WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'": DoCmd.RunSQL sql
  sql = "DELETE * FROM dta_dataset       WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'": DoCmd.RunSQL sql
  sql = "DELETE * FROM dta_dataset       WHERE id_model = '" & id_model & "' AND fn_dataset = '" & fn_dataset & "'": DoCmd.RunSQL sql
  DoCmd.SetWarnings True
  '
  If (1 = 1) Then ' Add Dataset

    ' Open Dataset Recordset
    Dim dst As Recordset: Set dst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE 1=2")
    '
    dst.AddNew
    dst!id_model = id_model
    dst!id_dataset = id_dataset
    dst!id_group = IIf(dq_dataset_type = "result", get_id_group_by_fn_group(id_model, "dqg-001 result"), _
                   IIf(dq_dataset_type = "totals", get_id_group_by_fn_group(id_model, "dqg-002-totals"), _
                   IIf(dq_dataset_type = "result_agg", get_id_group_by_fn_group(id_model, "dqg-011-aggregates-results"), _
                   IIf(dq_dataset_type = "totals_agg", get_id_group_by_fn_group(id_model, "dqg-012-aggregates-totals"), ""))))
    dst!id_development_status = id_development_status
    dst!is_ingestion = False
    dst!fn_dataset = fn_dataset
    dst!fd_dataset = fd_dataset
    dst!nm_target_schema = nm_target_schema
    dst!nm_target_table = nm_target_table
    dst!tx_source_query = tx_source_query
    Call show_recordset_record(dst, "Dataset", ip_is_debugging)
    dst.Update
    dst.Close
    '
  End If
  '
  If (1 = 1) Then ' Add Attribute for Dataset
    '
    ' Open Attribute Recordset
    Dim att As Recordset: Set att = CurrentDb.OpenRecordset("SELECT * FROM dta_attribute WHERE 1=2")
    Dim id_datatype_int As String: id_datatype_int = get_id_datatype("INT")
    Dim id_datatype_dat As String: id_datatype_dat = get_id_datatype("DATE")
    Dim id_datatype_ide As String: id_datatype_ide = get_id_datatype("CHAR(32)")
    '
    ' -----------------------------------
    ' Result
    ' -----------------------------------
    If (dq_dataset_type = "result") Then
      '
      ' Add References for ID Dataset 1st till up on 5th Attribtue
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_dat, 1, "DQ Result Date", "DQ Result Date.", "dt_dq_result", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 2, "ID Dataset 1 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 1"" checked by the ""DQ Control"".", "id_dataset_1_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 3, "ID Dataset 2 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 2"" checked by the ""DQ Control"".", "id_dataset_2_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 4, "ID Dataset 3 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 3"" checked by the ""DQ Control"".", "id_dataset_3_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 5, "ID Dataset 4 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 4"" checked by the ""DQ Control"".", "id_dataset_4_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 6, "ID Dataset 5 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 5"" checked by the ""DQ Control"".", "id_dataset_5_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 7, "ID DQ Result Status", "Reference to ""DQ Result Status"".", "id_dq_result_status", False, False, ip_is_debugging)
      '
    End If
    '
    If (dq_dataset_type = "totals") Then
      '
      ' /* Add "DQ Result Date" Attribute */
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_dat, 1, "DQ Result Date", "DQ Result Date.", "dt_dq_result", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 2, "# Oke", "The # Records that have ""Status"" Oke.", "ni_oke", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 3, "# Not Oke", "The # Records that have ""Status"" Not Oke.", "ni_nok", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 4, "# Out of Scope", "The # Records that have ""Status"" Out of Scope.", "ni_oos", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 5, "# Total (excl. OOS)", "The # Total excluding ""Out of Scope""-status.", "ni_total_excl_oos", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 6, "# Total (incl. OOS)", "The # Total including ""Out of Scope""-status.", "ni_total_incl_oos", False, False, ip_is_debugging)
      '
      '
    End If
    '
    If (dq_dataset_type = "result_agg") Then
      '
      ' Add Attribute for "Dataset"
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 1, "DQ Control ID", "The ""Businesskey Hash""-value of the ""Record"" from ""DQ Control""", "id_dq_control", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_dat, 2, "DQ Result Date", "Date of the DQL Result", "dt_dq_result", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 3, "ID Dataset 1 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 2"" checked by the ""DQ Control"".", "id_dataset_1_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 4, "ID Dataset 2 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 2"" checked by the ""DQ Control"".", "id_dataset_2_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 5, "ID Dataset 3 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 3"" checked by the ""DQ Control"".", "id_dataset_3_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 6, "ID Dataset 4 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 4"" checked by the ""DQ Control"".", "id_dataset_4_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 7, "ID Dataset 5 BK", "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset 5"" checked by the ""DQ Control"".", "id_dataset_5_bk", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 8, "ID DQ Result Status", "Reference to ""DQ Result Status"".", "id_dq_result_status", False, False, ip_is_debugging)
      '
    End If
    '
    If (dq_dataset_type = "totals_agg") Then
      '
      ' Add Attribute for "Dataset"
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_ide, 1, "DQ Control ID", "The ""Businesskey Hash""-value of the ""Record"" from ""DQ Control""", "id_dq_control", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_dat, 2, "DQ Result Date", "DQ Result Date.", "dt_dq_result", True, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 3, "# Oke", "The # Records that have ""Status"" Oke.", "ni_oke", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 4, "# Not Oke", "The # Records that have ""Status"" Not Oke.", "ni_nok", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 5, "# Out of Scope", "The # Records that have ""Status"" Out of Scope.", "ni_oos", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 6, "# Total (excl. OOS)", "The # Total excluding ""Out of Scope""-status.", "ni_total_excl_oos", False, False, ip_is_debugging)
      Call add_dqc_dataset_attribute(att, id_model, id_dataset, id_datatype_int, 7, "# Total (incl. OOS)", "The # Total including ""Out of Scope""-status.", "ni_total_incl_oos", False, False, ip_is_debugging)
      '
    End If
    '
    ' Clean up
    att.Close
    '
  End If
  '
  If (1 = 1) Then ' Add Ingestion etl
    Dim etl As Recordset: Set etl = CurrentDb.OpenRecordset("SELECT * FROM dta_ingestion_etl WHERE 1=2")
    etl.AddNew
    etl!id_model = id_model
    etl!id_dataset = id_dataset
    etl!id_ingestion_etl = CreateMD5("|" & id_model & "|" & id_dataset & "|ingestion_etl|")
    etl!nm_processing_type = "Incremental"
    etl!tx_sql_for_meta_dt_valid_from = "meta_dt_valid_from"
    etl!tx_sql_for_meta_dt_valid_till = "meta_dt_valid_till"
    etl.Update
    etl.Close
  End If
  '
  ' Extrection of Transformation Metadata
  Call parse_transformation_parts(id_model, id_dataset, ip_is_debugging)
  '
  ' Update Repository file(s)
  Call export_dataset_and_related_definitions(id_dataset)
  '
  ' All is Well
  Exit Sub
  '
ErrorHandling:
  '
  ' Print Error
  Debug.Print ("--- Error --------------------------------------------------------------")
  Debug.Print ("Err Number      : " & CStr(Err.Number))
  Debug.Print ("Err Description : " & Err.Description)
  Call show_recordset_record(dst, "Dataset", True)
  '
End Sub
'
'------------------------------------------------------------------------------
' Function: get_record_count_of_aggregate_level
'
' Purpose:
'   Returns the number of records in the temporary table "tmp_dq_aggregate_level"
'   that match the given model ID, aggregation level, and aggregation type.
'
' Description:
'   - Rebuilds the temporary table by deleting it (if present) and recreating it
'     using the query "create_tmp_dq_aggregate_level".
'   - Constructs an SQL SELECT statement filtered by:
'         * id_model
'         * ni_level
'         * cd_aggregation_type
'   - Opens a recordset on the filtered data and counts the rows manually.
'   - Returns the total number of matching records.
'
' Parameters:
'   ip_id_model            String   Model identifier to filter on.
'   ip_ni_level            Integer  Aggregation level.
'   ip_cd_aggregation_type String   Aggregation type code.
'
' Returns:
'   Integer - Number of matching records.
'
' Error Handling:
'   - Error 2059 is ignored and execution continues.
'   - All other errors are printed to the Immediate Window and execution stops.
'------------------------------------------------------------------------------
Public Function get_record_count_of_aggregate_level(ip_id_model As String, ip_ni_level As Integer, ip_cd_aggregation_type As String) As Integer
On Error GoTo ErrorHandling
  '
  Dim cnt As Integer: cnt = 0
  Dim nwl As String:  nwl = vbNewLine
  Dim sql As String:  sql = "SELECT * FROM [tmp_dq_aggregate_level] AS [dql]" & _
                      nwl & "WHERE [dql].[id_model]            = '" & ip_id_model & "'" & _
                      nwl & "AND   [dql].[ni_level]            = " & CStr(ip_ni_level) & "" & _
                      nwl & "AND   [dql].[cd_aggregation_type] = '" & ip_cd_aggregation_type & "'"
  '
  ' Remove existing "tmp_dq_aggregate_level"
  DoCmd.SetWarnings False: DoCmd.DeleteObject acTable, "tmp_dq_aggregate_level": DoCmd.OpenQuery "create_tmp_dq_aggregate_level": DoCmd.SetWarnings True
  '
  Dim dst As Recordset: Set dst = CurrentDb.OpenRecordset(sql)
  Do While (Not dst.EOF): cnt = cnt + 1: dst.MoveNext: Loop
  get_record_count_of_aggregate_level = cnt
  '
  ' All Done here
  Exit Function
  '
ErrorHandling:
  '
  '
  If Err.Number = 2059 Then Err.Clear: Resume Next
  '
  ' Metadata Editor can't find the object 'tmp_dq_aggregate_level'.
  If Err.Number = 7874 Then Err.Clear: Resume Next
  '
  ' Print Error
  Debug.Print ("--- Error --------------------------------------------------------------")
  Debug.Print ("Err Number      : " & CStr(Err.Number))
  Debug.Print ("Err Description : " & Err.Description)
  Stop
  Resume
  '
End Function
'
'------------------------------------------------------------------------------
' Procedure: sql_insert_or_update_dq_control_transformation
'
' Purpose:
'   Builds or updates all dataset, attribute, and transformation definitions
'   associated with a given Data Quality Control (DQ Control). This includes:
'     � Generating the "Result" dataset for the DQ Control
'     � Generating the "Totals" dataset for aggregated DQ results
'     � Rebuilding the DQ Control SQL query by replacing internal status IDs
'       with their functional codes
'     � Updating the DQ Control record with the rebuilt query
'     � Triggering the creation/update of aggregation transformations
'
' Description:
'   - Retrieves the DQ Control definition from table `dqm_dq_control`.
'   - Validates that the stored DQ Control query follows the required structure.
'   - Creates dataset definitions for:
'         * The DQ Control result table (dq_result.dqr_<control>)
'         * The DQ Control totals table (dq_totals.dqt_<control>)
'   - Generates MD5-based dataset identifiers for both dataset types.
'   - Builds the source SQL for each dataset, including:
'         * Minified DQ Control query for the result dataset
'         * Aggregation query (OKE/NOK/OOS counts) for the totals dataset
'   - Rebuilds the DQ Control query by replacing internal result-status IDs
'     with their functional codes (e.g., OKE, NOK, OOS).
'   - Updates the DQ Control record with the rebuilt SQL.
'   - Calls transformation builders to generate aggregation/union logic.
'   - Optionally prints debugging output when `ip_is_debugging = True`.
'
' Parameters:
'   ip_id_dq_control   String   Identifier of the DQ Control to process.
'   ip_is_debugging    Boolean  Optional. When True, prints debug output.
'
' Behavior:
'   - If the DQ Control does not exist, a message box is shown and the
'     procedure exits.
'   - If the DQ Control query is not structurally valid, the procedure exits.
'   - All SQL updates are executed with warnings temporarily disabled.
'
' Side Effects:
'   - Inserts or updates dataset definitions in metadata tables.
'   - Updates the DQ Control SQL in `dqm_dq_control`.
'   - Generates transformation SQL files via `build_sql_file_dataset`.
'
' Error Handling:
'   - Any runtime error prints diagnostic information to the Immediate Window.
'   - No automatic recovery is attempted.
'
'------------------------------------------------------------------------------
Public Sub sql_insert_or_update_dq_control_transformation(ip_id_dq_control As String, Optional ip_is_debugging As Boolean = False): On Error GoTo ErrorHandling
  '
  ' Declare Local Variables for "Dataset"- and "Attribute"- definitions
  Dim rec As typ_dataset
  Dim id_model              As String: id_model = id_model_default()
  Dim id_dataset            As String
  Dim id_dq_control         As String
  Dim id_attribute          As String
  Dim id_datatype           As String
  Dim id_development_status As String: id_development_status = get_id_development_status("PRD")
  Dim fn_dq_control         As String
  Dim fd_dq_control         As String
  Dim nm_target_schema      As String
  Dim nm_target_table       As String
  Dim nm_target_column      As String
  Dim ni_ordering           As Integer
  Dim tx_dq_control_query   As String
  Dim id_dataset_result     As String
  '
  ' /* Declare Local Variables for Recordsets */
  Dim dqc As DAO.Recordset: ' -> DQ Control Recordset
  Dim dst As DAO.Recordset: ' -> Dataset Recordset
  Dim att As DAO.Recordset: ' -> Attribute Recordset
  Dim dqs As DAO.Recordset: ' -> DQ Result Status
  '
  ' /* Declare Local Variables for building "SQL"-statements and "Messages" */
  Dim msg As String: msg = ""
  Dim sql As String: sql = ""
  Dim emp As String: emp = ""
  Dim nwl As String: nwl = vbNewLine
  '
  ' Ensure OHG Groups exist
  Call create_dq_group
  '
  ' Set id_model
  rec.id_model = id_model_default()
  '
  ' /* Process all DQ Controls */
  sql = "SELECT * FROM dqm_dq_control" & _
  nwl & "WHERE id_model      = '" & rec.id_model & "'" & _
  nwl & "AND   id_dq_control = '" & ip_id_dq_control & "'"
  '
  Set dqc = CurrentDb.OpenRecordset(sql)
  If (dqc.EOF) Then
    msg = "DQ Control `" & ip_id_dq_control & "` not found in model `" & rec.id_model & "`!"
    MsgBox msg, vbCritical, "DQ Control not found"
    Exit Sub
  End If
  '
  ' Check if Provided DQ Control Query is compliant to Generic structure of "DQ Control Query"
  If (check_general_structue_dq_control_query(dqc!tx_dq_control_query) = False) Then
    Exit Sub
  End If
  '
  ' Check if Provided DQ Control Query is compliant to Generic structure of "DQ Control Query"
  If (1 = 1) Then ' /* Generate "Dataset"- and "Attribute"- definitions for the "DQ Control"-result. */
    rec.id_dataset = CreateMD5("|" & rec.id_model & "|" & ip_id_dq_control & "|result|")
    rec.id_development_status = get_id_development_status("PRD")
    rec.is_ingestion = False
    rec.fn_dataset = dqc!fn_dq_control & " (Results)"
    rec.fd_dataset = "`Dataset`-definition for `DQ Control`-result of `" & dqc!id_dq_control & "` with functional description of `" & dqc!fn_dq_control & "`."
    rec.nm_target_schema = "dq_result"
    rec.nm_target_table = "dqr_" & dqc!id_dq_control
    rec.tx_source_query = MinifySQL(replace_dq_result_status_with_id(dqc!tx_dq_control_query))
    Call add_dqc_dataset(rec.id_model, rec.id_dataset, rec.id_development_status, rec.fn_dataset, rec.fd_dataset, rec.nm_target_schema, rec.nm_target_table, rec.tx_source_query, "result", ip_is_debugging)
    '
    ' Remember Resulting Dataset ID
    id_dataset_result = rec.id_dataset
    '
  End If
  '
  If (1 = 1) Then ' /* Generate "Dataset"- and "Attribute"- definitions for the "DQ Control"-totals. */
    rec.id_dataset = CreateMD5("|" & rec.id_model & "|" & ip_id_dq_control & "|totals|")
    rec.id_development_status = get_id_development_status("PRD")
    rec.is_ingestion = False
    rec.fn_dataset = dqc!fn_dq_control & " (Totals)"
    rec.fd_dataset = "`Dataset`-definition for `DQ Control`-totals of `" & ip_id_dq_control & "` with functional description of `" & ip_id_dq_control & "`."
    rec.nm_target_schema = "dq_totals"
    rec.nm_target_table = "dqt_" & dqc!id_dq_control
    rec.tx_source_query = MinifySQL("SELECT dqr.dt_dq_result AS dt_dq_result," _
                            & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "OKE") + "' THEN 1 ELSE 0 END) AS ni_oke," _
                            & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "NOK") + "' THEN 1 ELSE 0 END) AS ni_nok," _
                            & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "OOS") + "' THEN 1 ELSE 0 END) AS ni_oos," _
                            & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status != '" + tsa_id_dq_result_status(id_model, "OOS") + "' THEN 1 ELSE 0 END) AS ni_total_excl_oos," _
                            & nwl & "       SUM(1) AS ni_total_incl_oos" _
                            & nwl & "FROM dq_result.dqr_" & ip_id_dq_control & " AS dqr" _
                            & nwl & "WHERE dqr.meta_dt_valid_from <= CONVERT(DATETIME, @dt_current_stand)" _
                            & nwl & "AND   dqr.meta_dt_valid_till >  CONVERT(DATETIME, @dt_current_stand)" _
                            & nwl & "GROUP BY dqr.dt_dq_result, CONVERT(DATETIME, meta_dt_valid_from), CONVERT(DATETIME, meta_dt_valid_till)")
                            ' "CONVERT(DATETIME, meta_dt_valid_from), CONVERT(DATETIME, meta_dt_valid_till)" are added to avoid "Msg 8120, Level 16, State 1, Line 1 Column 'meta_dt_valid_from' is invalid in the select list because it is not contained in either an aggregate function or the GROUP BY clause. For these are added in the "main"-query for the stored procedure generation
    Call add_dqc_dataset(rec.id_model, rec.id_dataset, rec.id_development_status, rec.fn_dataset, rec.fd_dataset, rec.nm_target_schema, rec.nm_target_table, rec.tx_source_query, "totals", ip_is_debugging)
    '
  End If
  '
  ' Rebuild Query
  sql = "SELECT tx_source_query FROM dta_dataset WHERE id_dataset = '" & id_dataset_result & "'"
  Set dst = CurrentDb.OpenRecordset(sql): Do While Not dst.EOF
    '
    ' Set "tx_dq_control_query"
    tx_dq_control_query = dst!tx_source_query
    '
    ' Replace the "DQ Result Status ID"
    sql = "SELECT id_dq_result_status, fn_dq_result_status FROM srd_dq_result_status WHERE id_model = '" & id_model & "'"
    Set dqs = CurrentDb.OpenRecordset(sql): Do While Not dqs.EOF
      '
      tx_dq_control_query = Replace(tx_dq_control_query, dqs!id_dq_result_status, dqs!fn_dq_result_status)
      If (ip_is_debugging) Then Debug.Print "DQ Result Status ID   : '" & dqs!id_dq_result_status & "'"
      If (ip_is_debugging) Then Debug.Print "DQ Result Status Code : '" & dqs!fn_dq_result_status & "'"
      If (ip_is_debugging) Then Debug.Print "SQL : " & tx_dq_control_query
      '
    dqs.MoveNext: Loop
    '
  dst.MoveNext: Loop
  '
  ' Build Update
  sql = "UPDATE dqm_dq_control " _
      & "SET tx_dq_control_query = '" & Replace(tx_dq_control_query, "'", "''") & "' " _
      & "WHERE id_dq_control = '" & ip_id_dq_control & "'"
  DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
  '
  ' Generate the "Transformation" to Aggregation/Union all the "DQ Control"-dataset.
  Call sql_insert_or_update_dq_control_aggregation("Result", ip_is_debugging)
  Call sql_insert_or_update_dq_control_aggregation("Totals", ip_is_debugging)
  '
  ' All is Well
  Call export_all_dataset_and_related_definitions
  Exit Sub
  '
ErrorHandling:
  '
  ' Print Error
  Debug.Print ("--- Error --------------------------------------------------------------")
  Debug.Print ("Err Number      : " & CStr(Err.Number))
  Debug.Print ("Err Description : " & Err.Description)
  '
End Sub
'
'
' Example: call sql_insert_or_update_dq_control_aggregation("Result", True)
'
Public Sub sql_insert_or_update_dq_control_aggregation(ip_cd_aggregation_type As String, Optional ip_is_debugging As Boolean = False): On Error GoTo ErrorHandling
  '
  ' /* Declare Local Variables */
  Dim rec As typ_dataset
  Dim dqr As DAO.Recordset
  Dim att As DAO.Recordset
  Dim agg As DAO.Recordset
  Dim sql As String
  Dim nwl As String: nwl = vbNewLine
  Dim qry As String
  '
  Dim id_model              As String: id_model = id_model_default()
  Dim id_development_status As String: id_development_status = get_id_development_status("PRD")
  ' Dim tx_source_query       As String
  Dim id_dataset_result     As String
  Dim ni_level              As Integer: ni_level = 0
  Dim ni_sub_level          As Integer: ni_sub_level = 0
  '
  If (1 = 1) Then ' Build Template for Aggregation Query
    If (ip_cd_aggregation_type = "Result") Then
      qry = "SELECT '<dqr!id_dq_control>' AS id_dq_control," & _
      nwl & "       dqr.dt_dq_result AS dt_dq_result," & _
      nwl & "       dqr.id_dataset_1_bk AS id_dataset_1_bk," & _
      nwl & "       dqr.id_dataset_2_bk AS id_dataset_2_bk," & _
      nwl & "       dqr.id_dataset_3_bk AS id_dataset_3_bk," & _
      nwl & "       dqr.id_dataset_4_bk AS id_dataset_4_bk," & _
      nwl & "       dqr.id_dataset_5_bk AS id_dataset_5_bk," & _
      nwl & "       dqr.id_dq_result_status AS id_dq_result_status" & _
      nwl & "FROM [<nm_source_schema>].[<nm_source_table>] AS dqr"
    End If
    If (ip_cd_aggregation_type = "Totals") Then
      qry = "SELECT '<dqr!id_dq_control>' AS id_dq_control," & _
      nwl & "       dqt.dt_dq_result             AS dt_dq_result," & _
      nwl & "       dqt.ni_oke                   AS ni_oke," & _
      nwl & "       dqt.ni_nok                   AS ni_nok," & _
      nwl & "       dqt.ni_oos                   AS ni_oos," & _
      nwl & "       dqt.ni_total_excl_oos        AS ni_total_excl_oos," & _
      nwl & "       dqt.ni_total_incl_oos        AS ni_total_incl_oos" & _
      nwl & "FROM [<nm_source_schema>].[<nm_source_table>] AS dqt"
    End If
  End If
  '
  If (1 = 1) Then ' Set Schema Name for Aggregated of the DQ Controls Results
    rec.nm_target_schema = "dq_" & LCase(ip_cd_aggregation_type) & "_agg"
    rec.id_development_status = get_id_development_status("PRD")
  End If
  '
  ' Remove existing "datasets" for "Aggregation" of "Result/Totals".
  sql = "DELETE * FROM dta_dataset" & _
  nwl & "WHERE nm_target_schema = '" & rec.nm_target_schema & "'" & _
  nwl & "OR (" & _
  nwl & "  nm_target_schema = 'dqm' AND " & _
  nwl & "  nm_target_table  = 'dq_" & LCase(ip_cd_aggregation_type) & "'" & _
  nwl & ")"
  DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
  '
  ' While Level 0 or there is still dataset to Aggregate.
  Do While (ni_level = 0 Or (get_record_count_of_aggregate_level(id_model, ni_level, ip_cd_aggregation_type) > 1))
    '
    ' Build SQL Statement for fetching DQ result/totals Transformations
    If (ni_level = 0) Then
      ' Step 1: Aggregate ever two "DQ Control"(Result)"-datasets into one "DQ Control (Result) - Level 1 - Aggregated N..9999"
      sql = "SELECT dst.id_dataset, Mid([nm_target_table], 5) AS id_dq_control " & _
      nwl & "FROM dta_dataset AS dst " & _
      nwl & "WHERE dst.nm_target_schema IN ('" & Replace(rec.nm_target_schema, "_agg", "") & "')" & _
      nwl & "ORDER BY dst.fn_dataset"
    Else
      ' Step 2: Aggregate ever two "DQ Control (Result/Totals) - Level N - Aggregated 9999"-datasets
      '         into one "DQ Control (Result/Totals) - Level N+1 - Aggregated 9999", until only one
      '         "DQ Control (Result/Totals) - Level N - Aggregated 9999"-dataset is left.
      sql = "SELECT dst.id_dataset AS id_dataset, dst.id_dataset AS id_dq_control" & _
      nwl & "FROM tmp_dq_aggregate_level AS dst" & _
      nwl & "WHERE dst.id_model            = '" & id_model & "'" & _
      nwl & "AND   dst.ni_level            =  " & CStr(ni_level) & "" & _
      nwl & "AND   dst.cd_aggregation_type = '" & ip_cd_aggregation_type & "'"
    End If
    '
    ' Setting Level and Sub Level, per iteration add 1 to the Level and restart Sub Level.
    ni_level = ni_level + 1
    ni_sub_level = 0
    '
    ' Show SQL for Fetch Dataset
    If (ip_is_debugging = True) Then Debug.Print "SQL for fetch Datasets:" & nwl & sql
    '
    Set dqr = CurrentDb.OpenRecordset(sql): Do While (Not dqr.EOF)
      '
      ' Add 1 to Sub Level, Started at 0.
      ni_sub_level = ni_sub_level + 1
      '
      ' Show Level and Sub Level
      If (ip_is_debugging = True) Then
        Debug.Print "# Level     :" & CStr(ni_level)
        Debug.Print "# Sub Level :" & CStr(ni_sub_level)
      End If
      '
      ' Generate id_dataset
      rec.id_dataset = CreateMD5("|" & id_model & "|" & CStr(ni_level) & "|" & CStr(ni_sub_level) & "|" & ip_cd_aggregation_type & "|")
      rec.fn_dataset = "DQ Control (" & ip_cd_aggregation_type & ") - Level " & CStr(ni_level) & " - Aggregated " & CStr(ni_sub_level)
      rec.fd_dataset = "DQ Control Dataset of '" & dqr!id_dataset & "'"
      rec.nm_target_table = "dqa_" & rec.id_dataset
      rec.tx_source_query = qry
      '
      ' Update "Source"-query with id_dq_control from source of hardcode based on source-table-name, Set Schema and Table names
      GoSub SettingSourceScehmaAndTable
      '
      ' Fetch Next "DQ Control (result)"-dataset
      dqr.MoveNext
      '
      ' If Not at EOF then Union to Next dataset.
      If (Not dqr.EOF) Then
        '
        ' Extent the "Functional Description" the the 2nd dataset info.
        rec.fd_dataset = rec.fd_dataset & " Unioned with DQ Controle Dataset of '" & dqr!id_dataset & "'"
        '
        ' Extent the "Source"-query with the "template"
        rec.tx_source_query = rec.tx_source_query & nwl & "UNION ALL" & nwl & qry
        '
        ' Update "Source"-query with id_dq_control from source of hardcode based on source-table-name, Set Schema and Table names
        GoSub SettingSourceScehmaAndTable
        '
      End If
      '
      ' Add New "DQ Control (Result) - Level N - Aggregate"
      Call add_dqc_dataset(id_model, rec.id_dataset, rec.id_development_status, rec.fn_dataset, rec.fd_dataset, rec.nm_target_schema, rec.nm_target_table, rec.tx_source_query, LCase(ip_cd_aggregation_type & "_agg"), ip_is_debugging)
      '
      ' Fetch Next DQ Control/Aggregation Dataset
      If (Not dqr.EOF) Then dqr.MoveNext
      '
    Loop: dqr.Close
    
  Loop
  '
  If (1 = 1) Then ' Updata the Last "Dataset" so it point to dqm.dq_result/totals
    '
    sql = "UPDATE dta_dataset" & _
    nwl & "SET fn_dataset       = 'DQ " & ip_cd_aggregation_type & " (Aggregated/Unioned)'" & _
    nwl & "  , id_group         = '" & get_id_group_by_fn_group(id_model, IIf(ip_cd_aggregation_type = "Result", "dqg-021-all-results", "dqg-022-all-totals")) & "'" & _
    nwl & "  , fd_dataset       = 'All the " & ip_cd_aggregation_type & " of the DQ Controls are unioned into this dataset.'" & _
    nwl & "  , nm_target_schema = 'dqm'" & _
    nwl & "  , nm_target_table  = 'dq_" & LCase(ip_cd_aggregation_type) & "'" & _
    nwl & "WHERE id_dataset = '" & rec.id_dataset & "'"
    DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Update Repository file(s)
    Call export_dataset_and_related_definitions(rec.id_dataset)
    '
  End If
  '
  ' All is Well
  Exit Sub
  '
SettingSourceScehmaAndTable:
  '
  ' Add Level = 1 the source is still teh DQ Control Transformation, the direct results of the DQ Controls. If Level > 1 the "source" becomes the "Aggregation"-dataset them selfs
  rec.tx_source_query = Replace(rec.tx_source_query, "<nm_source_schema>", "dq_" & LCase(ip_cd_aggregation_type) & IIf(ni_level = 1, "", "_agg"))
  '
  ' The name of the "Dataset" depends on the cd_aggregation_type,
  rec.tx_source_query = Replace(rec.tx_source_query, "<nm_source_table>", "dq" & IIf(ni_level = 1, LCase(Left(ip_cd_aggregation_type, 1)), "a") & "_" & dqr!id_dq_control)
  '
  ' If Level = 1, the id_dq_control is "hardcode" for this is not part of the DQ Control Transformation, this is because
  ' it is a generated code would only take up save in the result set. If the sources are the "Aggregation" themselfs the
  ' it is taken from the "source"-dataset.
  rec.tx_source_query = Replace(rec.tx_source_query, "'<dqr!id_dq_control>'", IIf(ni_level = 1, "'" & dqr!id_dq_control & "'", "dq" & LCase(Left(ip_cd_aggregation_type, 1)) & ".id_dq_control"))
  '
  ' All done Here
  Return ' to main Routine
  '
ErrorHandling:
  '
  ' Print Error
  Debug.Print ("--- Error --------------------------------------------------------------")
  Debug.Print ("Err Number      : " & CStr(Err.Number))
  Debug.Print ("Err Description : " & Err.Description)
  Stop
  Resume
  '
End Sub