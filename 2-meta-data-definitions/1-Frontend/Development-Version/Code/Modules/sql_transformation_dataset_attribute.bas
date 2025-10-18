Attribute VB_Name = "sql_transformation_dataset_attribute"
Option Compare Database
Option Explicit

Public Sub test_parse_transformation_dataset_attribute()

    Call parse_transformation_dataset_attribute( _
      "5f4a1942465c575a1f5a5a575d1e191c", _
      "2272de7c9edadb179dbc6932c0f725a3", _
      True, _
      True _
    )

End Sub

Public Sub parse_transformation_dataset_attribute( _
  ip_id_model As String, _
  ip_id_transformation_dataset As String, _
  Optional ip_is_debugging As Boolean = False, _
  Optional ip_is_testing As Boolean = False _
)
    '
    ' Local Variables for Building SQL
    Dim emp                As String: emp = ""
    Dim nwl                As String: nwl = vbNewLine
    Dim sql                As String: sql = ""
    '
    ' Delete existing attributes for this transformation column mapping attribute
    sql = "DELETE FROM dta_transformation_dataset_attribute " & _
          "WHERE id_transformation_dataset = '" & ip_id_transformation_dataset & "';"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    '
    ' Build SQL Statment to Insert Utilized Attributes is given column mapping
    sql = emp & emp & "INSERT INTO dta_transformation_dataset_attribute ("
    sql = sql & nwl & "  id_model, id_transformation_dataset, id_transformation_dataset_attribute,"
    sql = sql & nwl & "  cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "  tx_source_attribute"
    sql = sql & nwl & ")"
    sql = sql & nwl & "SELECT"
    sql = sql & nwl & "  id_model, id_transformation_dataset, id_transformation_dataset_attribute,"
    sql = sql & nwl & "  cd_source_alias, id_source_model, id_source_attribute,"
    sql = sql & nwl & "  tx_source_attribute"
    sql = sql & nwl & "FROM tmp_transformation_dataset_attribute"
    sql = sql & nwl & "WHERE id_transformation_dataset = '" & ip_id_transformation_dataset & "';"
    If ip_is_debugging Then Debug.Print sql
    If Not ip_is_testing Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True

End Sub