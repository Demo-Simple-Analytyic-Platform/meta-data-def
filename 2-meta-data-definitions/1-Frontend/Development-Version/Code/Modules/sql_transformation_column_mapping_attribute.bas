Attribute VB_Name = "sql_transformation_column_mapping_attribute"
Option Compare Database
Option Explicit

Public Sub test_parse_transformation_column_mapping_attribute()

    Call parse_transformation_column_mapping_attribute( _
      "5f4a1942465c575a1f5a5a575d1e191c", _
      "2272de7c9edadb179dbc6932c0f725a3", _
      True, _
      True _
    )

End Sub

Public Sub parse_transformation_column_mapping_attribute( _
  ip_id_model As String, _
  ip_id_transformation_column_mapping As String, _
  Optional ip_is_debugging As Boolean = False, _
  Optional ip_is_testing As Boolean = False _
)
    '
    ' Local Variables for Building SQL
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    Dim sql As String: sql = ""
    '
    ' Delete existing attributes for this transformation column mapping attribute
    sql = "DELETE FROM dta_transformation_column_mapping_attribute " & _
          "WHERE id_transformation_column_mapping = '" & ip_id_transformation_column_mapping & "';"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Build SQL Statment to Insert Utilized Attributes is given column mapping
    sql = emp & emp & "INSERT INTO dta_transformation_column_mapping_attribute ("
    sql = sql & nwl & "  id_model, id_transformation_column_mapping, id_transformation_column_mapping_attribute,"
    sql = sql & nwl & "  cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "  tx_source_attribute"
    sql = sql & nwl & ")"
    sql = sql & nwl & "SELECT id_model, id_transformation_column_mapping, id_transformation_column_mapping_attribute,"
    sql = sql & nwl & "       cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "       tx_source_attribute"
    sql = sql & nwl & "FROM tmp_transformation_column_mapping_attribute"
    sql = sql & nwl & "WHERE id_transformation_column_mapping = '" & ip_id_transformation_column_mapping & "';"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Fetch Transformation Part SQL Clauses
    sql = emp & emp & "SELECT tx_transformation_column_mapping, id_transformation_part"
    sql = sql & nwl & "FROM dta_transformation_column_mapping"
    sql = sql & nwl & "WHERE id_transformation_column_mapping = '" & ip_id_transformation_column_mapping & "'"
    Dim dbs As DAO.Database:  Set dbs = CurrentDb
    Dim rst As DAO.Recordset: Set rst = dbs.OpenRecordset(sql)
    '
    ' Update Transformation Part SQL Clauses
    rst.Edit
    rst!tx_transformation_column_mapping = source_attributes_to_placeholder(rst!tx_transformation_column_mapping, rst!id_transformation_part)
    rst.Update
    rst.Close
    '
End Sub