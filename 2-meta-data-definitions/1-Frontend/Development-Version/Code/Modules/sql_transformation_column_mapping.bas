Attribute VB_Name = "sql_transformation_column_mapping"
Option Compare Database
Option Explicit

Public Type typ_transformation_column_mapping
    id_attribute              As String
    id_transformation_column_mapping As String
    tx_transformation_column_mapping As String
    is_in_group_by            As Boolean
End Type

Public Sub parse_transformation_column_mapping(ip_id_model As String, ip_id_dataset As String, ip_id_transformation_part As String, ip_tx_transformation_part As String, Optional ip_is_debugging As Boolean = False, Optional ip_is_testing As Boolean = False)
    '
    ' Declare Local Varaible
    Dim db               As DAO.Database
    Dim rst              As DAO.Recordset
    Dim tmp              As DAO.Recordset
    Dim tx_sql_statement As String
    Dim tx_sql_group_by  As String
    Dim ni_pos_begin     As Long
    Dim ni_pos_ended     As Long
    Dim ni_pos_length    As Long
    Dim id_model         As String
    Dim id_dataset       As String
    Dim sql              As String
    Dim map              As typ_transformation_column_mapping
    '
    If ip_is_debugging Then Debug.Print vbNewLine & String(40, "=") & vbNewLine & " Mappings:" & vbNewLine
    '
    Set db = CurrentDb
    '
    ' Minify and extract SELECT clause
    tx_sql_statement = ip_tx_transformation_part
    '
    ni_pos_begin = InStr(1, UCase(tx_sql_statement), "SELECT")
    ni_pos_ended = InStr(1, UCase(tx_sql_statement), "FROM")
    '
    If ni_pos_begin > 0 Then
        If ni_pos_ended = 0 Then ni_pos_ended = Len(tx_sql_statement)
        ni_pos_length = ni_pos_ended - ni_pos_begin
        tx_sql_statement = Trim(Mid(tx_sql_statement, ni_pos_begin + 7, ni_pos_length - 7))
        If ip_is_debugging Then Debug.Print "SELECT clause: " & tx_sql_statement
        If ip_is_debugging Then Debug.Print String(40, "-")
    End If
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
        ni_pos_length = ni_pos_ended - ni_pos_begin
        tx_sql_group_by = Mid(tx_sql_group_by, ni_pos_begin + 9, ni_pos_length - 9)
        '
        ' Show Group By text if Debugging Mode
        If ip_is_debugging Then Debug.Print "GROUP BY clause: " & tx_sql_group_by
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: at -> list expacted Attributes(s) */
        sql = "DELETE * FROM tmp_transformation_column_mapping_at"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        
        sql = "INSERT INTO tmp_transformation_column_mapping_at (id_attribute, ni_ordering, nm_target_column, tx_to_be_searched_1, tx_to_be_searched_2, tx_sql_statement) " _
            & "SELECT id_attribute, ni_ordering, nm_target_column, 'AS [' & nm_target_column & ']' AS tx_to_be_searched_1, 'AS ' & nm_target_column AS tx_to_be_searched_2, " _
            & "'" & Replace(tx_sql_statement, "'", "''") & "' AS tx_sql_statement " _
            & "FROM dta_attribute " _
            & "WHERE id_dataset = '" & ip_id_dataset & "' " _
            & "AND LEFT(nm_target_column, 4) <> 'meta' "
        If ip_is_debugging Then Debug.Print "SQL: " & vbNewLine & sql
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_column_mapping_at ORDER BY ni_ordering ASC"
            Set rst = db.OpenRecordset(sql): With rst: Do Until .EOF
                Debug.Print "--- at " & String(40, "-")
                Debug.Print "id_attribute        : '" & !id_attribute & "'"
                Debug.Print "ni_ordering         : " & !ni_ordering & ""
                Debug.Print "nm_target_column    : '" & !nm_target_column & "'"
                Debug.Print "tx_to_be_searched_1 : '" & !tx_to_be_searched_1 & "'"
                Debug.Print "tx_to_be_searched_2 : '" & !tx_to_be_searched_2 & "'"
            .MoveNext: Loop: End With
        
        End If
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: md -> Find de postions of "alias" of expected "target"-columns */
        sql = "DELETE * FROM tmp_transformation_column_mapping_md"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        
        sql = "INSERT INTO tmp_transformation_column_mapping_md (id_attribute, ni_ordering, nm_target_column, ni_position_till, ni_to_be_searched_length, tx_to_be_searched, dummy) " _
            & "SELECT id_attribute, ni_ordering, nm_target_column, " _
            & "       InStr(1, att.tx_sql_statement, att.tx_to_be_searched_1) + " _
            & "       InStr(1, att.tx_sql_statement, att.tx_to_be_searched_2) AS ni_position_till, " _
            & "       Iif(InStr(1, att.tx_sql_statement, att.tx_to_be_searched_1) > 0, Len(att.tx_to_be_searched_1), 0) + " _
            & "       Iif(InStr(1, att.tx_sql_statement, att.tx_to_be_searched_2) > 0, Len(att.tx_to_be_searched_2), 0) AS ni_to_be_searched_length," _
            & "       Iif(InStr(1, att.tx_sql_statement, att.tx_to_be_searched_1) > 0, att.tx_to_be_searched_1) + " _
            & "       Iif(InStr(1, att.tx_sql_statement, att.tx_to_be_searched_2) > 0, att.tx_to_be_searched_2) AS tx_to_be_searched," _
            & "       1 As ni_dummy " _
            & "FROM tmp_transformation_column_mapping_at AS att"
        If ip_is_debugging Then Debug.Print "SQL: " & vbNewLine & sql
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True

        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_column_mapping_md ORDER BY ni_ordering ASC"
            Set rst = db.OpenRecordset(sql): With rst: Do Until .EOF
                Debug.Print "--- md " & String(40, "-")
                Debug.Print "id_attribute             : '" & !id_attribute & "'"
                Debug.Print "ni_ordering              : " & !ni_ordering & ""
                Debug.Print "nm_target_column         : '" & !nm_target_column & "'"
                Debug.Print "ni_position_till         : " & !ni_position_till & ""
                Debug.Print "ni_to_be_searched_length : " & !ni_to_be_searched_length & ""
                Debug.Print "tx_to_be_searched        : " & !tx_to_be_searched & ""
            .MoveNext: Loop: End With
        
        End If
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: ls -> Last "ni_ordering" */
        sql = "DELETE * FROM tmp_transformation_column_mapping_ls"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        
        sql = "INSERT INTO tmp_transformation_column_mapping_ls (ni_ordering) " _
            & "SELECT MAX(md.ni_ordering) AS ni_ordering " _
            & "FROM tmp_transformation_column_mapping_md AS md"
        If ip_is_debugging Then Debug.Print "SQL: " & vbNewLine & sql
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True

        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_column_mapping_ls ORDER BY ni_ordering ASC"
            Set rst = db.OpenRecordset(sql): With rst: Do Until .EOF
                Debug.Print "--- ls " & String(40, "-")
                Debug.Print "ni_ordering : '" & !ni_ordering & "'"
            .MoveNext: Loop: End With
        
        End If
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: be -> Begin and End. */
        sql = "DELETE * FROM tmp_transformation_column_mapping_be"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        
        sql = "INSERT INTO tmp_transformation_column_mapping_be (id_attribute, ni_ordering, nm_target_column, ni_position_from, ni_position_till) " _
            & "SELECT md.id_attribute, md.ni_ordering, md.nm_target_column, " _
            & "       Iif(" _
            & "         Nz((SELECT MAX(lag.ni_position_till)         FROM tmp_transformation_column_mapping_md AS lag WHERE lag.ni_position_till < md.ni_position_till), 1) = 1, 1," _
            & "         Nz((SELECT MAX(lag.ni_position_till)         FROM tmp_transformation_column_mapping_md AS lag WHERE lag.ni_position_till < md.ni_position_till), 1) + " _
            & "         Nz((SELECT MAX(lag.ni_to_be_searched_length) FROM tmp_transformation_column_mapping_md AS lag WHERE lag.ni_position_till < md.ni_position_till), 1) + 2 " _
            & "       ) AS ni_position_from, " _
            & "       md.ni_position_till " _
            & "FROM tmp_transformation_column_mapping_md AS md"
        If ip_is_debugging Then Debug.Print "SQL: " & vbNewLine & sql
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True

        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_column_mapping_be ORDER BY ni_ordering ASC"
            Set rst = db.OpenRecordset(sql): With rst: Do Until .EOF
                Debug.Print "--- be " & String(40, "-")
                Debug.Print "id_attribute             : '" & !id_attribute & "'"
                Debug.Print "ni_ordering              : " & !ni_ordering & ""
                Debug.Print "nm_target_column         : '" & !nm_target_column & "'"
                Debug.Print "ni_position_from         : " & !ni_position_from & ""
                Debug.Print "ni_position_till         : " & !ni_position_till & ""
            .MoveNext: Loop: End With
        End If
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: be -> Mapping. */
        sql = "DELETE * FROM dta_transformation_column_mapping " _
            & "WHERE id_model               = '" & ip_id_model & "' " _
            & "AND   id_transformation_part = '" & ip_id_transformation_part & "'"
        If (ip_is_testing = False) Then DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        '
        ' Open Recordset Transformation Mapping to add records
        Set tmp = db.OpenRecordset("SELECT * FROM dta_transformation_column_mapping WHERE 1=2")
        '
        ' Loop thought the Begin-End dataset.
        Set rst = db.OpenRecordset("SELECT * FROM tmp_transformation_column_mapping_be ORDER BY ni_ordering ASC"): With rst: Do Until .EOF
            '
            ' Population Transformation Map record
            map.id_attribute = !id_attribute
            map.id_transformation_column_mapping = CreateMD5("|" & ip_id_model & "|" & !id_attribute & "|" & ip_id_transformation_part & "|")
            map.tx_transformation_column_mapping = Trim(IIf(!ni_position_from = 0, "", Mid(tx_sql_statement, !ni_position_from, (!ni_position_till - !ni_position_from))))
            map.is_in_group_by = IIf(!ni_position_till = 0, 0, IIf(InStr(1, tx_sql_group_by, map.tx_transformation_column_mapping, vbTextCompare), 1, 0))
            '
            ' if 1st character is "," then remove that and trim again
            If (Left(map.tx_transformation_column_mapping, 1) = ",") Then
                map.tx_transformation_column_mapping = Trim(Mid(map.tx_transformation_column_mapping, 2, Len(map.tx_transformation_column_mapping)))
            End If
            
            '
            ' Add record to dta_transformation_column_mapping.
            tmp.AddNew
            tmp.fields("id_model") = ip_id_model
            tmp.fields("id_attribute") = map.id_attribute
            tmp.fields("id_transformation_part") = ip_id_transformation_part
            tmp.fields("id_transformation_column_mapping") = map.id_transformation_column_mapping
            tmp.fields("tx_transformation_column_mapping") = map.tx_transformation_column_mapping
            tmp.fields("is_in_group_by") = map.is_in_group_by
            If ip_is_testing = False Then tmp.Update
            '
            If ip_is_debugging Then
                Debug.Print "--- be " & String(40, "-")
                Debug.Print "id_model                  : '" & ip_id_model & "'"
                Debug.Print "id_attribute              : '" & map.id_attribute & "'"
                Debug.Print "id_transformation_part    : '" & ip_id_transformation_part & "'"
                Debug.Print "id_transformation_column_mapping : '" & map.id_transformation_column_mapping & "'"
                Debug.Print "tx_transformation_column_mapping : '" & map.tx_transformation_column_mapping & "'"
                Debug.Print "is_in_group_by            : " & CStr(map.is_in_group_by) & ""
            End If
            '
            '
            ' Extract Attributes utilized in Column Mapping
            Call parse_transformation_column_mapping_attribute(ip_id_model, map.id_transformation_column_mapping, ip_is_debugging, ip_is_testing)
            '
        .MoveNext: Loop: End With
        '
    End If
    '
    Set db = Nothing
    '
End Sub