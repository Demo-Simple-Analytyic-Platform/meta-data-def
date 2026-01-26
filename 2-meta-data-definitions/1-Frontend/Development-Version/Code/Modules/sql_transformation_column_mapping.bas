Attribute VB_Name = "sql_transformation_column_mapping"
Option Compare Database
Option Explicit

Public Type typ_transformation_column_mapping
    id_attribute              As String
    id_transformation_column_mapping As String
    tx_transformation_column_mapping As String
    is_in_group_by            As Boolean
End Type

Public Function get_group_by_text(ip_tx_transformation_part As String, Optional ip_is_debugging As Boolean = False) As String
    '
    ' Local Variables
    Dim tx_sql_group_by  As String: tx_sql_group_by = ""
    Dim ni_pos_begin     As Long
    Dim ni_pos_ended     As Long
    Dim ni_pos_length    As Long
    '
    ' Extract GROUP BY clause
    If InStr(1, ip_tx_transformation_part, "GROUP BY", vbTextCompare) > 0 Then
        '
        ' Group BY Text
        tx_sql_group_by = ip_tx_transformation_part
        '
        ni_pos_begin = InStr(1, UCase(tx_sql_group_by), "GROUP BY")
        ni_pos_ended = InStr(1, UCase(tx_sql_group_by), "HAVING")
        '
        If ni_pos_ended = 0 Then ni_pos_ended = Len(tx_sql_group_by)
        ni_pos_length = ni_pos_ended - ni_pos_begin - Len("GROUP BY")
        tx_sql_group_by = Trim(Mid(tx_sql_group_by, (ni_pos_begin + 9), ni_pos_length))
        '
        ' Show Group By text if Debugging Mode
        If ip_is_debugging Then Debug.Print "GROUP BY clause: " & tx_sql_group_by
        '
    End If
    '
    ' Return found text
    get_group_by_text = tx_sql_group_by
End Function

Public Function get_is_in_group_by(ip_tx_transformation_column_mapping As String, ip_tx_sql_group_by As String) As Boolean
  If (InStr(1, ip_tx_sql_group_by, ip_tx_transformation_column_mapping, vbTextCompare) > 0) Then
    get_is_in_group_by = True
  Else
    get_is_in_group_by = False
  End If
End Function

Public Sub parse_transformation_column_mapping(ip_id_model As String, ip_id_dataset As String, ip_id_transformation_part As String, ip_tx_transformation_part As String, Optional ip_is_debugging As Boolean = False, Optional ip_is_testing As Boolean = False)
On Error GoTo ErrorHandling
  '
  ' Declare Local Varaible
  Dim sql As String
  Dim nwl As String: nwl = vbNewLine
  Dim dbs As DAO.Database: Set dbs = CurrentDb
  Dim rst As DAO.Recordset
  Dim tcm As DAO.Recordset
  '
  If ip_is_debugging Then Debug.Print vbNewLine & String(40, "=") & vbNewLine & " Mappings:" & vbNewLine
  '
  If (1 = 1) Then ' /* "Update" transformation column mappings */
    '
    ' Drop existing records form transformation column mapping
    sql = "DELETE * FROM dta_transformation_column_mapping " _
        & "WHERE id_model               = '" & ip_id_model & "' " _
        & "AND   id_transformation_part = '" & ip_id_transformation_part & "'"
    If (ip_is_testing = False) Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Insert "new" transformation column mapping.
    sql = "SELECT" & _
    nwl & "  id_model, id_attribute, id_transformation_part," & _
    nwl & "  tx_transformation_column_mapping," & _
    nwl & "  tx_transformation_part" & _
    nwl & "FROM tmp_transformation_column_mapping" & _
    nwl & "WHERE id_model               = '" & ip_id_model & "' " & _
    nwl & "AND   id_transformation_part = '" & ip_id_transformation_part & "'"
    Set tcm = dbs.OpenRecordset("SELECT * FROM dta_transformation_column_mapping WHERE 1=2")
    Set rst = dbs.OpenRecordset(sql): Do While Not rst.EOF
      tcm.AddNew
      tcm!id_model = rst!id_model
      tcm!id_attribute = rst!id_attribute
      tcm!id_transformation_part = rst!id_transformation_part
      tcm!id_transformation_column_mapping = CreateMD5("|" & rst!id_model & "|" & rst!id_attribute & "|" & rst!id_transformation_part & "|")
      tcm!tx_transformation_column_mapping = rst!tx_transformation_column_mapping
      tcm!is_in_group_by = get_is_in_group_by(rst!tx_transformation_column_mapping, get_group_by_text(rst!tx_transformation_part))
      tcm.Update
    rst.MoveNext: Loop: tcm.Close: rst.Close
    '
    ' Get Rescordset for "Transformation Column Mappings"
    sql = "SELECT * FROM dta_transformation_column_mapping " _
        & "WHERE id_model               = '" & ip_id_model & "' " _
        & "AND   id_transformation_part = '" & ip_id_transformation_part & "'"
    Set rst = dbs.OpenRecordset(sql): With rst: Do While Not .EOF
      '
      ' Extract Attributes utilized in Column Mapping
      Call parse_transformation_column_mapping_attribute(!id_model, !id_transformation_column_mapping, ip_is_debugging, ip_is_testing)
      '
    .MoveNext: Loop: End With
    '
  End If
  '
  Set dbs = Nothing
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
  MsgBox "Parsing SQL Query failed!", vbCritical, "SQL Error Parsing"
  ' stop
  ' Resume
  '
End Sub