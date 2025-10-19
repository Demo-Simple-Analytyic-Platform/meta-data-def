Attribute VB_Name = "sql_transformation_part_attribute"
Option Compare Database
Option Explicit

Public Sub test_parse_transformation_part_attribute()

    Call parse_transformation_part_attribute( _
      "5f4a1942465c575a1f5a5a575d1e191c", _
      "2272de7c9edadb179dbc6932c0f725a3", _
      True, _
      True _
    )

End Sub

Public Sub parse_transformation_part_attribute( _
  ip_id_model As String, _
  ip_id_transformation_part As String, _
  Optional ip_is_debugging As Boolean = False, _
  Optional ip_is_testing As Boolean = False _
)
    '
    ' Local Variables for Building SQL
    Dim nwl As String: nwl = vbNewLine
    Dim emp As String: emp = ""
    Dim sql As String: sql = ""
    '
    ' Delete existing attributes for this transformation column mapping attribute
    sql = "DELETE FROM dta_transformation_part_attribute " & _
          "WHERE id_transformation_part = '" & ip_id_transformation_part & "';"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Build SQL Statment to Insert Utilized Attributes is given column mapping
    sql = emp & emp & "INSERT INTO dta_transformation_part_attribute ("
    sql = sql & nwl & "  id_model, id_transformation_part, id_transformation_part_attribute, cd_transformation_part_clause_type,"
    sql = sql & nwl & "  cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "  tx_source_attribute"
    sql = sql & nwl & ")"
    sql = sql & nwl & "SELECT"
    sql = sql & nwl & "  id_model, id_transformation_part, id_transformation_part_attribute, cd_transformation_part_clause_type,"
    sql = sql & nwl & "  cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "  tx_source_attribute"
    sql = sql & nwl & "FROM ("
    sql = sql & nwl & ""
    sql = sql & nwl & "  SELECT"
    sql = sql & nwl & "    id_model, id_transformation_part, id_transformation_part_attribute, cd_transformation_part_clause_type,"
    sql = sql & nwl & "    cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "    tx_source_attribute"
    sql = sql & nwl & "  FROM tmp_transformation_part_attribute_where"
    sql = sql & nwl & "  WHERE id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  UNION"
    sql = sql & nwl & "  SELECT"
    sql = sql & nwl & "    id_model, id_transformation_part, id_transformation_part_attribute, cd_transformation_part_clause_type,"
    sql = sql & nwl & "    cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "    tx_source_attribute"
    sql = sql & nwl & "  FROM tmp_transformation_part_attribute_group_by"
    sql = sql & nwl & "  WHERE id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & "  "
    sql = sql & nwl & "  UNION"
    sql = sql & nwl & "  SELECT"
    sql = sql & nwl & "    id_model, id_transformation_part, id_transformation_part_attribute, cd_transformation_part_clause_type,"
    sql = sql & nwl & "    cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "    tx_source_attribute"
    sql = sql & nwl & "  FROM tmp_transformation_part_attribute_having"
    sql = sql & nwl & "  WHERE id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & ") AS u;"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Fetch Transformation Part SQL Clauses
    sql = emp & emp & "SELECT tx_transformation_part_where_clause"
    sql = sql & nwl & "     , tx_transformation_part_group_by_clause"
    sql = sql & nwl & "     , tx_transformation_part_having_clause"
    sql = sql & nwl & "FROM dta_transformation_part"
    sql = sql & nwl & "WHERE id_transformation_part = '" & ip_id_transformation_part & "'"
    Dim dbs As DAO.Database:  Set dbs = CurrentDb
    Dim rst As DAO.Recordset: Set rst = dbs.OpenRecordset(sql)
    '
    ' Update Transformation Part SQL Clauses
    rst.Edit
    rst!tx_transformation_part_where_clause = source_attributes_to_placeholder(rst!tx_transformation_part_where_clause, ip_id_transformation_part)
    rst!tx_transformation_part_group_by_clause = source_attributes_to_placeholder(rst!tx_transformation_part_group_by_clause, ip_id_transformation_part)
    rst!tx_transformation_part_having_clause = source_attributes_to_placeholder(rst!tx_transformation_part_having_clause, ip_id_transformation_part)
    rst.Update
    rst.Close
    '
End Sub