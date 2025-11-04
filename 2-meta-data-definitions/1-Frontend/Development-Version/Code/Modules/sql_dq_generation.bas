Attribute VB_Name = "sql_dq_generation"
Option Compare Database
Option Explicit
'
' /* get tsa_id_dq_result_status */
Public Function tsa_id_dq_result_status(ip_id_model As String, ip_fn_dq_result_status As String) As String
  '
  ' /* Declare Local Variables */
  Dim sql As String:            sql = "SELECT id_dq_result_status FROM srd_dq_result_status WHERE id_model = '" & ip_id_model & "' AND fn_dq_result_status = '" & ip_fn_dq_result_status & "'"
  Dim rst As DAO.Recordset: Set rst = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

  ' /* Return Result */
  tsa_id_dq_result_status = rst!id_dq_result_status

End Function

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
    ' Return Result adjusted DQ Control Query
    replace_dq_result_status_with_id = tx_dq_control_query
    '
End Function

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

Public Sub show_recordset_record(ByRef ip_rst As Recordset, ip_tx_message As String, Optional ip_is_debugging As Boolean = False)
    If ip_is_debugging Then
        Debug.Print ""
        Debug.Print "--- " & ip_tx_message & " " & String(30 - (Len(ip_tx_message) + 5), "-")
        Dim fld As Field: For Each fld In ip_rst.fields: Debug.Print fld.Name & " : " & CStr(Nz(fld.Value, "n/a")): Next fld
        Debug.Print String(30, "-")
    End If
End Sub

Public Function check_general_structue_dq_control_query(ip_tx_dq_control_query As String) As Boolean
  '
  ' Define Local Variables
  Dim rex As String
  '
  ' Build RegEx Pattern
  rex = "^SELECT" _
      & "\s+.+\s+AS\s+id_dataset_1_bk," _
      & "\s+.+\s+AS\s+id_dataset_2_bk," _
      & "\s+.+\s+AS\s+id_dataset_3_bk," _
      & "\s+.+\s+AS\s+id_dataset_4_bk," _
      & "\s+.+\s+AS\s+id_dataset_5_bk," _
      & "\s+CASE\s+WHEN\s+.+\s+THEN\s+.+\s+ELSE\s+.+\s+END\s+AS\s+id_dq_result_status" _
      & "\s+FROM\s+.+$"
  '
  ' Match RegEx pattern with provided Query
  check_general_structue_dq_control_query = RegExMatch(ip_tx_dq_control_query, rex)
  '
End Function

' Generate Dataset (Transformations) for the DQ Controls
Public Sub sql_insert_or_update_dq_control_transformation(ip_id_dq_control As String, Optional ip_is_debugging As Boolean = False)
On Error GoTo ErrorHandling
  '
  ' Declare Local Variables for "Dataset"- and "Attribute"- definitions
  Dim id_model              As String: id_model = id_model_default()
  Dim id_dataset            As String
  Dim id_dq_control         As String
  Dim id_attribute          As String
  Dim id_datatype           As String
  Dim id_development_status As String
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
  ' Set id_model
  id_model = id_model_default()
  '
  ' /* Process all DQ Controls */
  Set dqc = CurrentDb.OpenRecordset("SELECT * FROM dqm_dq_control WHERE id_model = '" & id_model & "' AND id_dq_control = '" & ip_id_dq_control & "'")
  If (1 = 1) Then ' /* Generate "Dataset"- and "Attribute"- definitions for the "DQ Control"-result. */
    '
    ' Generate ID Dataset for "Transformation" that needs to be created or updated for the DQ Control
    id_dataset = CreateMD5("|" & id_model & "|" & ip_id_dq_control & "|result|")
    id_dataset_result = id_dataset
    '
    ' /* DQ Control Query , Convert "OKE", "NOK and "OOS"" to "ID". */)*/
    tx_dq_control_query = replace_dq_result_status_with_id(dqc!tx_dq_control_query)
    '
    ' Check if Provided DQ Control Query is compliant to Generic structure of "DQ Control Query"
    If (check_general_structue_dq_control_query(MinifySQL(tx_dq_control_query)) = False) Then
      msg = "DQ Control `" & dqc!id_dq_control & "` has not the correct format!" & nwl & "Expected format for query: " & nwl & "SELECT" & nwl & "% AS id_dataset_1_bk," & nwl & "% AS id_dataset_2_bk," & nwl & "% AS id_dataset_3_bk," & nwl & "% AS id_dataset_4_bk," & nwl & "% AS id_dataset_5_bk," & nwl & "%CASE%WHEN%ELSE%END AS id_dq_result_status" & nwl & "%FROM%" & nwl & "DQ Control Query : """ & tx_dq_control_query & """"
      Err.Raise vbObjectError + 513, "sql_insert_or_update_dq_control_transformation", msg
    End If
    '
    If (1 = 1) Then ' /* Add Dataset Definitions for DQ Control Result */
      '
      ' Drop "Existing" Dataset definitions for the DQ Control Dataset
      sql = "DELETE * FROM dta_dataset WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'"
      DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
      '
      ' Open Dataset Recordset
      Set dst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE 1=2")
      '
      ' /* Generate "Dataset"-definitions for the "DQ Control"-result. */
      dst.AddNew
      dst!id_model = id_model
      dst!id_dataset = id_dataset
      dst!id_development_status = dqc!id_development_status
      dst!is_ingestion = False
      dst!fn_dataset = dqc!fn_dq_control & " (Results)"
      dst!fd_dataset = "`Dataset`-definition for `DQ Control`-result of `" & dqc!id_dq_control & "` with functional description of `" & dqc!fn_dq_control & "`."
      dst!nm_target_schema = "dq_result"
      dst!nm_target_table = "dqr_" & dqc!id_dq_control
      dst!tx_source_query = tx_dq_control_query
      Call show_recordset_record(dst, "Dataset", ip_is_debugging)
      dst.Update
      dst.Close
      '
    End If
    '
    If (1 = 1) Then ' /* Add Attribute Definitions for DQ Control Result */
      '
      ' Drop "Existing" Attribute definitions for the DQ Control Dataset
      sql = "DELETE * FROM dta_attribute WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'"
      DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
      '
      ' Open Attribute Recordset
      Set att = CurrentDb.OpenRecordset("SELECT * FROM dta_attribute WHERE 1=2")
      '
      ' Get Datatype ID
      id_datatype = get_id_datatype("CHAR(32)")
      '
      ' Add References for ID Dataset 1st till up on 5th Attribtue
      For ni_ordering = 1 To 5
        '
        ' Add Attribtue
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = ni_ordering
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(ni_ordering) & "|")
        att!fn_attribute = "ID Dataset " & CStr(ni_ordering) & " BK"
        att!fd_attribute = "The ""Businesskey Hash""-value of the ""Record"" from ""Dataset " & CStr(ni_ordering) & """ checked by the ""DQ Control""."
        att!nm_target_column = "id_dataset_" & CStr(ni_ordering) & "_bk"
        att!is_businesskey = True
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
        '
      Next ni_ordering
      '
      ' Add "id_dq_result_status"
      att.AddNew
      att!id_model = id_model
      att!id_dataset = id_dataset
      att!id_datatype = id_datatype
      att!ni_ordering = 6
      att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
      att!fn_attribute = "ID DQ Result Status"
      att!fd_attribute = "Reference to ""DQ Result Status""."
      att!nm_target_column = "id_dq_result_status"
      att!is_businesskey = False
      att!is_nullable = False
      Call show_recordset_record(att, "Attribute", ip_is_debugging)
      att.Update
      '
    End If
    '
    ' Extraction of Transformation Parts/Dataset/Column-Mappings
    Call sql_transformation_part.parse_transformation_parts(id_model, id_dataset, ip_is_debugging, False)
    '
    ' Update Repository file(s)
    Call export_dataset_and_related_definitions(id_dataset)
    '
  End If
  '
  If (1 = 1) Then ' /* Generate "Dataset"- and "Attribute"- definitions for the "DQ Control"-totals. */
    '
    ' Generate ID Dataset for "Transformation" that needs to be created or updated for the DQ Control
    id_dataset = CreateMD5("|" & id_model & "|" & ip_id_dq_control & "|totals|")
    '
    ' /* DQ Control Query , Convert "OKE", "NOK and "OOS"" to "ID". */)*/
    tx_dq_control_query = "SELECT CONVERT(DATE, @dt_current_stand) AS dt_dq_result," _
                  & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "OKE") + "' THEN 1 ELSE 0 END) AS ni_oke," _
                  & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "NOK") + "' THEN 1 ELSE 0 END) AS ni_nok," _
                  & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status  = '" + tsa_id_dq_result_status(id_model, "OOS") + "' THEN 1 ELSE 0 END) AS ni_oos," _
                  & nwl & "       SUM(CASE WHEN dqr.id_dq_result_status != '" + tsa_id_dq_result_status(id_model, "OOS") + "' THEN 1 ELSE 0 END) AS ni_total_excl_oos," _
                  & nwl & "       SUM(1) AS ni_total_incl_oos" _
                  & nwl & "FROM dq_result.dqr_" & ip_id_dq_control & " AS dqr" _
                  & nwl & "WHERE dqr.meta_dt_valid_from <= CONVERT(DATETIME, @dt_current_stand)" _
                  & nwl & "AND   dqr.meta_dt_valid_till >  CONVERT(DATETIME, @dt_current_stand)"
    '
    If (1 = 1) Then ' /* Add Dataset Definitions for DQ Control Totals */
      '
      ' /* Drop "Existing" Dataset definitions for the DQ Control Dataset */
      sql = "DELETE * FROM dta_dataset WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'"
      DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
      '
      ' Open Dataset Recordset
      Set dst = CurrentDb.OpenRecordset("SELECT * FROM dta_dataset WHERE 1=2")
      '
      ' /* Generate "Dataset"-definitions for the "DQ Control"-result. */
      dst.AddNew
      dst!id_model = id_model
      dst!id_dataset = id_dataset
      dst!id_development_status = dqc!id_development_status
      dst!is_ingestion = False
      dst!fn_dataset = dqc!fn_dq_control & " (Totals)"
      dst!fd_dataset = "`Dataset`-definition for `DQ Control`-totals of `" & ip_id_dq_control & "` with functional description of `" & ip_id_dq_control & "`."
      dst!nm_target_schema = "dq_totals"
      dst!nm_target_table = "dqt_" & ip_id_dq_control
      dst!tx_source_query = tx_dq_control_query
      Call show_recordset_record(dst, "Dataset", ip_is_debugging)
      dst.Update
      dst.Close
      '
    End If
    '
    If (1 = 1) Then ' /* Add Attribute Definitions for DQ Control Totals */
      '
      ' Drop "Existing" Attribute definitions for the DQ Control Dataset
      sql = "DELETE * FROM dta_attribute WHERE id_model = '" & id_model & "' AND id_dataset = '" & id_dataset & "'"
      DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
      '
      ' Open Attribute Recordset
      Set att = CurrentDb.OpenRecordset("SELECT * FROM dta_attribute WHERE 1=2")
      '
      ' Get Datatype ID
      id_datatype = get_id_datatype("INT")
      '
      If (1 = 1) Then ' /* Add "DQ Result Date" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 1
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "DQ Result Date"
        att!fd_attribute = "DQ Result Date."
        att!nm_target_column = "dt_dq_result"
        att!is_businesskey = True
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
      If (1 = 1) Then ' /* Add "# Oke" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 2
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "# Oke"
        att!fd_attribute = "The # Records that have ""Status"" Oke."
        att!nm_target_column = "ni_oke"
        att!is_businesskey = False
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
      If (1 = 1) Then ' /* Add "# Not Oke" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 3
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "# Not Oke"
        att!fd_attribute = "The # Records that have ""Status"" Not Oke."
        att!nm_target_column = "ni_nok"
        att!is_businesskey = False
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
      If (1 = 1) Then ' /* Add "# Out of Scope" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 4
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "# Out of Scope"
        att!fd_attribute = "The # Records that have ""Status"" Out of Scope."
        att!nm_target_column = "ni_oos"
        att!is_businesskey = False
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
      If (1 = 1) Then ' /* Add "# Total (excl. OOS)" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 5
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "# Total (excl. OOS)"
        att!fd_attribute = "The # Total excluding ""Out of Scope""-status."
        att!nm_target_column = "ni_total_excl_oos"
        att!is_businesskey = False
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
      If (1 = 1) Then ' /* Add "# Total (incl. OOS)" Attribute */
        att.AddNew
        att!id_model = id_model
        att!id_dataset = id_dataset
        att!id_datatype = id_datatype
        att!ni_ordering = 6
        att!id_attribute = CreateMD5("|" & id_model & "|" & id_dataset & "|" & CStr(att!ni_ordering) & "|")
        att!fn_attribute = "# Total (incl. OOS)"
        att!fd_attribute = "The # Total including ""Out of Scope""-status."
        att!nm_target_column = "ni_total_incl_oos"
        att!is_businesskey = False
        att!is_nullable = False
        Call show_recordset_record(att, "Attribute", ip_is_debugging)
        att.Update
        '
      End If
      '
    End If
    '
    ' Extraction of Transformation Parts/Dataset/Column-Mappings
    Call sql_transformation_part.parse_transformation_parts(id_model, id_dataset, ip_is_debugging, False)
    '
    ' Update Repository file(s)
    Call export_dataset_and_related_definitions(id_dataset)
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
  ' All is Well
  Call build_sql_file_dataset
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
'
' Generate Dataset (Transformations) for the DQ Controls
Public Sub sql_insert_or_update_dq_control_transformation_all()
  On Error GoTo ErrorHandling
  '
  ' /* Declare Local Variables */
  Dim dqc As DAO.Recordset: Set dqc = CurrentDb.OpenRecordset("SELECT * FROM dqm_dq_control WHERE LEN(ISNULL(tx_dq_control_query,'')) > 10") ' -> DQ Control Recordset
  '
  ' /* Process all DQ Controls */
  Do While Not dqc.EOF:  Call sql_insert_or_update_dq_control_transformation(dqc!id_dq_control): dqc.MoveNext: Loop
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
  '
End Sub