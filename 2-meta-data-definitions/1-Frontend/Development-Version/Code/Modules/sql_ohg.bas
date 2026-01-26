Attribute VB_Name = "sql_ohg"
Option Compare Database
Option Explicit

Public Sub add_group_if_not_exists(ip_id_model As String, ip_fn_group As String, ip_fd_group As String)
  '
  ' Local varibales
  Dim sql As String:       sql = "SELECT * FROM ohg_group " _
                                & "WHERE id_model = '" & ip_id_model & "' " _
                                & "AND   fn_group = '" & ip_fn_group & "' "
  Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
  '
  If (rst.EOF) Then ' Add the Group
    rst.AddNew
    rst!id_group = CreateMD5("|" & ip_id_model & "|" & ip_fn_group & "|")
    rst!id_model = ip_id_model
    rst!fn_group = ip_fn_group
    rst!fd_group = ip_fd_group
    rst.Update
    rst.Close
  End If
  '
End Sub

Public Function get_id_group_by_fn_group(ip_id_model As String, ip_fn_group As String) As String
  '
  ' Local varibales
  Dim sql As String:        sql = "SELECT * FROM ohg_group " _
                                & "WHERE id_model = '" & ip_id_model & "' " _
                                & "AND   fn_group = '" & ip_fn_group & "' "
  Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
  '
  If Not (rst.EOF) Then ' Add the Group
    get_id_group_by_fn_group = rst!id_group
  End If
  '
End Function


Public Sub add_hierarchy_if_not_exists(ip_id_model As String, ip_id_group As String, ip_id_hierarchy_parent As String)
  '
  ' Local varibales
  Dim sql As String:        sql = "SELECT * FROM ohg_hierachy " _
                                & "WHERE id_model            = '" & ip_id_model & "' " _
                                & "AND   id_group            = '" & ip_id_group & "' " _
                                & "AND   id_hierarchy_parent = '" & ip_id_hierarchy_parent & "'"
  Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
  '
  If (rst.EOF) Then ' Add the Group
    rst.AddNew
    rst!id_hierarchy = CreateMD5("|" & ip_id_model & "|" & ip_id_group & "|" & ip_id_hierarchy_parent & "|")
    rst!id_model = ip_id_model
    rst!id_group = ip_id_group
    rst!id_hierarchy_parent = ip_id_hierarchy_parent
    rst.Update
    rst.Close
  End If
  '
End Sub


Public Sub add_related_if_not_exists(ip_id_model As String, ip_id_group As String, ip_id_dataset As String)
  '
  ' Local varibales
  Dim sql As String:        sql = "SELECT * FROM ohg_related " _
                                & "WHERE id_model   = '" & ip_id_model & "' " _
                                & "AND   id_group   = '" & ip_id_group & "' " _
                                & "AND   id_dataset = '" & ip_id_dataset & "' "
  Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
  '
  If (rst.EOF) Then ' Add the Group
    rst.AddNew
    rst!id_group = CreateMD5("|" & ip_id_model & "|" & ip_id_group & "|" & ip_id_dataset & "|")
    rst!id_model = ip_id_model
    rst!id_group = ip_id_group
    rst!id_dataset = ip_id_dataset
    rst.Update
    rst.Close
  End If
  '
End Sub