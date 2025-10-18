Attribute VB_Name = "sql_transformation_part"
Option Compare Database
Option Explicit

Public Type typ_transformation_part
    id_transformation_part As String
    ni_transformation_part As Integer
    tx_transformation_part As String
    tx_transformation_part_where_clause    As String
    tx_transformation_part_group_by_clause As String
    tx_transformation_part_having_clause   As String
End Type
'
Public Sub test_parse_transformation_parts()
    parse_transformation_parts "5f4a1942465c575a1f5a5a575d1e191c", "06020d070f0d0a04030d0b0705021500", True, False
End Sub
'
Sub parse_transformation_parts(ip_id_model As String, ip_id_dataset As String, Optional is_debugging As Boolean = False, Optional is_testing As Boolean = False)
    '
    'Declare Local Variables
    Dim nwl   As String:        nwl = vbNewLine
    Dim emp   As String:        emp = ""
    Dim pos   As Integer:       pos = 1
    Dim sql   As String:        sql = "SELECT tx_source_query FROM dta_dataset WHERE id_dataset = '" & ip_id_dataset & "' AND id_model = '" & ip_id_model & "'"
    Dim rst   As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    Dim txt   As String:        txt = MinifySQL(rst.fields("tx_source_query"))
    Dim tpt   As typ_transformation_part
    '
    ' Delclare Local Variables for Extractio Where Clause
    Dim ni_pos_where_begin  As Integer
    Dim ni_pos_where_ended  As Integer
    Dim ni_pos_where_length As Integer
    '
    ' Delclare Local Variables for Extractio Group By Clause
    Dim ni_pos_group_by_begin  As Integer
    Dim ni_pos_group_by_ended  As Integer
    Dim ni_pos_group_by_length As Integer
    '
    ' Delclare Local Variables for Extractio Having Clause
    Dim ni_pos_having_begin  As Integer
    Dim ni_pos_having_ended  As Integer
    Dim ni_pos_having_length As Integer
    '
    If is_debugging Then Debug.Print vbNewLine & "Source Query: '" & txt & "'"
    '
    ' Decalre Local variable for processing
    tpt.ni_transformation_part = 0
    tpt.id_transformation_part = ""
    tpt.tx_transformation_part = ""
    tpt.tx_transformation_part_where_clause = ""
    tpt.tx_transformation_part_group_by_clause = ""
    tpt.tx_transformation_part_having_clause = ""
    '
    sql = "" ' delete existing transformation_part(s)
    sql = sql & emp & "DELETE *"
    sql = sql & nwl & "FROM dta_transformation_part"
    sql = sql & nwl & "WHERE id_model   = '" & ip_id_model & "'"
    sql = sql & nwl & "AND   id_dataset = '" & ip_id_dataset & "'"
    If (is_testing = False) Then
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
    End If

    '
    ' Initalize Variables
    Do Until (txt = "")
        '
        ' Find next "UNION"
        pos = InStr(1, txt, " UNION ", vbTextCompare): pos = IIf(pos = 0, Len(txt), pos)
        '
        ' Extract Transformation Part(s)
        tpt.id_transformation_part = CreateMD5("|" & ip_id_dataset & "|" & CStr(tpt.ni_transformation_part) & "|")
        tpt.ni_transformation_part = tpt.ni_transformation_part + 1
        tpt.tx_transformation_part = Mid(txt, 1, pos)
        sql = UCase(tpt.tx_transformation_part)
        '
        If (1 = 1) Then ' Extract WHERE Caluse if any
            '
            ' Determin Begin and End of WHERE Clause
            ni_pos_where_begin = InStr(1, sql, " WHERE ", vbTextCompare): ni_pos_where_ended = 0
            ni_pos_where_ended = IIf(ni_pos_where_begin = 0, 0, IIf(ni_pos_where_ended <> 0, ni_pos_where_ended, InStr(1, sql, " GROUP BY ", vbTextCompare)))
            ni_pos_where_ended = IIf(ni_pos_where_begin = 0, 0, IIf(ni_pos_where_ended <> 0, ni_pos_where_ended, InStr(1, sql, " HAVING ", vbTextCompare)))
            '
            tpt.tx_transformation_part_where_clause = ""
            If (ni_pos_where_begin <> 0) Then
                ni_pos_where_length = IIf(ni_pos_where_ended <> 0, ni_pos_where_ended, Len(sql)) - 1
                tpt.tx_transformation_part_where_clause = Mid(sql, ni_pos_where_begin, ni_pos_where_length)
            End If
            '
        End If
        '
        If (1 = 1) Then ' Extract GROUP BY Caluse if any
            '
            ' Determin Begin and End of WHERE Clause
            ni_pos_group_by_begin = InStr(1, sql, " GROUP BY ", vbTextCompare): ni_pos_group_by_ended = 0
            ni_pos_group_by_ended = IIf(ni_pos_group_by_begin = 0, 0, IIf(ni_pos_group_by_ended <> 0, ni_pos_group_by_ended, InStr(1, sql, " HAVING ", vbTextCompare)))
            '
            tpt.tx_transformation_part_group_by_clause = ""
            If (ni_pos_group_by_begin <> 0) Then
                ni_pos_group_by_length = IIf(ni_pos_group_by_ended <> 0, ni_pos_group_by_ended, Len(sql)) - 1
                tpt.tx_transformation_part_group_by_clause = Mid(sql, ni_pos_group_by_begin, ni_pos_group_by_length)
            End If
            '
        End If
        '
        If (1 = 1) Then ' Extract HAVING Caluse if any
            '
            ' Determin Begin and End of WHERE Clause
            ni_pos_having_begin = InStr(1, sql, " HAVING ", vbTextCompare): ni_pos_having_ended = 0
            '
            tpt.tx_transformation_part_having_clause = ""
            If (ni_pos_having_begin <> 0) Then
                ni_pos_having_length = IIf(ni_pos_having_ended <> 0, ni_pos_having_ended, Len(sql)) - 1
                tpt.tx_transformation_part_having_clause = Mid(sql, ni_pos_having_begin, ni_pos_having_length)
            End If
            '
        End If
        '
        ' If Debugging Shot Transformation Part
        If (is_debugging = True) Then
            Debug.Print String(80, "-")
            Debug.Print "id_transformation_part :                 '" & tpt.id_transformation_part & "'"
            Debug.Print "ni_transformation_part                 : '" & tpt.ni_transformation_part & "'"
            Debug.Print "tx_transformation_part                 : '" & tpt.tx_transformation_part & "'"
            Debug.Print "tx_transformation_part_where_clause    : '" & tpt.tx_transformation_part_where_clause & "'"
            Debug.Print "tx_transformation_part_group_by_clause : '" & tpt.tx_transformation_part_group_by_clause & "'"
            Debug.Print "tx_transformation_part_having_clause   : '" & tpt.tx_transformation_part_having_clause & "'"
        End If
        '
        If Not is_testing Then
            '
            ' Build SQL Statement for open "transformation_part"-recordset
            sql = "SELECT * FROM dta_transformation_part WHERE 1=2"
            Set rst = CurrentDb.OpenRecordset(sql)
            '
            ' Populate Transformation Part
            rst.AddNew
            rst.fields("id_model") = ip_id_model
            rst.fields("id_dataset") = ip_id_dataset
            rst.fields("id_transformation_part") = tpt.id_transformation_part
            rst.fields("ni_transformation_part") = tpt.ni_transformation_part
            rst.fields("tx_transformation_part") = tpt.tx_transformation_part
            rst.fields("tx_transformation_part_where_clause") = tpt.tx_transformation_part_where_clause
            rst.fields("tx_transformation_part_group_by_clause") = tpt.tx_transformation_part_group_by_clause
            rst.fields("tx_transformation_part_having_clause") = tpt.tx_transformation_part_having_clause
            rst.Update
            rst.Close
            '
        End If
        '
        ' Parse Transformation Dataset
        Call parse_transformation_dataset(ip_id_model, ip_id_dataset, tpt.id_transformation_part, tpt.tx_transformation_part, is_debugging, is_testing)
        '
        ' Parse Transformation Column Mapping
        Call parse_transformation_column_mapping(ip_id_model, ip_id_dataset, tpt.id_transformation_part, tpt.tx_transformation_part, is_debugging, is_testing)
        '
        ' Parse Transformation Part for Attributes Utilized
        Call parse_transformation_part_attribute(ip_id_model, tpt.id_transformation_part, is_debugging, is_testing)
        '
        ' Update the Metadata for the WHERE/GROUP BY/HAVING clauses with placeholder, referening id_source_attribute and id_transformation_dataset for the correct cd_alias
        ' to do some code
        '
        ' Remove processed part from the "Source Query"-text
        txt = IIf(Len(txt) - pos > 0, Mid(txt, pos + 7, Len(txt) - pos), "")
        '
    Loop
    '
End Sub
