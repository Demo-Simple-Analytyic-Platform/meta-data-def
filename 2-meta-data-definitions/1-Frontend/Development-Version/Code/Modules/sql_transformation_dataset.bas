Attribute VB_Name = "sql_transformation_dataset"
Option Compare Database
Option Explicit

Public Type typ_transformation_dataset
    id_transformation_dataset As String
    ni_transformation_dataset As String
    cd_join_type              As String
    id_source_model           As String
    id_source_dataset         As String
    cd_alias                  As String
    tx_join_criteria          As String
End Type

Public Sub test_parse_transformation_dataset()
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset("SELECT * FROM dta_transformation_part where id_model = '5f4a1942465c575a1f5a5a575d1e191c' AND id_dataset = '06020d070f0d0a04030d0b0705021500'")
    Call parse_transformation_dataset(rst!id_model, rst!id_dataset, rst!id_transformation_part, rst!tx_transformation_part, True, False)
End Sub

Public Sub parse_transformation_dataset(ip_id_model As String, ip_id_dataset As String, ip_id_transformation_part As String, ip_tx_transformation_part As String, Optional ip_is_debugging As Boolean = False, Optional ip_is_testing As Boolean = False)
    '
    ' Example
    ' call parse_transformation_dataset ("5f4a1942465c575a1f5a5a575d1e191c", "06030d080305000102030a00040c1507", "a81fe3aa91717916ec30c16553953425", "SELECT 'ARR' AS [cd_stock]      , CONVERT(DATETIME, [src].[Date], 0) AS [dt_stock]      , [src].[Adj_Close_Adjusted_close_price_adjusted_for_splits_and_dividend_and_or_capital_gain_distributions] AS [tx_stock] FROM  [psa_yahoo_stock_info].[arr] AS [src] WHERE CONVERT(DATETIME, [src].[Date], 0) BETWEEN @dt_previous_stand AND @dt_current_stand AND   [src].[meta_is_active] = 1", True, False)
    '
    ' Local Variables database+recordset
    Dim tx As Variant
    Dim rs As DAO.Recordset
    Dim cr As DAO.Recordset
    Dim db As DAO.Database: Set db = CurrentDb
    '
    ' Local Variables
    Dim tx_sql_statement   As String:  tx_sql_statement = ip_tx_transformation_part
    Dim ni_position_begin  As Integer: ni_position_begin = 0
    Dim ni_position_end    As Integer: ni_position_end = 0
    Dim ni_position_length As Integer: ni_position_length = 0
    Dim ni_ordering        As Integer
    Dim vr_sql             As Variant
    '
    ' Local Variables for Building SQL
    Dim emp                As String: emp = ""
    Dim nwl                As String: nwl = vbNewLine
    Dim sql                As String: sql = ""
    '
    ' /* Temporary Variables for "Error handling" */
    Dim tx_error_message As String
    Dim tx_sql_execution As String
    '
    If (1 = 1) Then ' /* Extraction of "FROM/JOIN"-clauses of "Transformation"-part. */
        '
        If (Mid(tx_sql_statement, 1, Len("--- Warning ")) = "--- Warning ") Then
            Debug.Assert tx_sql_statement
            Call Err.Raise(16, "Error in SQL Statement. Please check 'Transformation'-part.", "Error")
        End If
        '
        ' /* Find " Beginning" of the "FROM/JOIN"-clause. */
        ni_position_begin = InStr(1, UCase(tx_sql_statement), "FROM", 1)
        '
        ' /* Find the "End" of the "FROM/JOIN"-clause. */
        ni_position_end = InStr(1, UCase(tx_sql_statement), "WHERE", 1)
        '
        ' /* If NO "WHERE"-clause found "search for "GROUP BY"-clause. */
        ni_position_end = IIf(ni_position_end = 0, InStr(1, tx_sql_statement, "GROUP BY"), ni_position_end)
        '
        ' /* If both the "Begin" and "End" have been found, determine the "Length". */
        ni_position_length = IIf(ni_position_end = 0, Len(tx_sql_statement), ni_position_end - ni_position_begin)
        '
        ' /* Extract only the "FROM/JOIN"-clause of the "Query". */
        tx_sql_statement = TRIM(Mid(tx_sql_statement, ni_position_begin, ni_position_length))
        '
        ' /* Show extracted "FROM/JOIN"-clause. */
        If ip_is_debugging Then Debug.Print tx_sql_statement
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: tx -> Convert "SQL"-statement to individual "words". */
        sql = "DELETE * FROM tmp_transformation_dataset_tx"
        DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
        tx = Split(tx_sql_statement): ni_ordering = 0
        Set rs = db.OpenRecordset("SELECT * FROM tmp_transformation_dataset_tx WHERE 1=2")
        For Each vr_sql In tx
            If Len(TRIM(vr_sql)) > 0 Then
                ni_ordering = ni_ordering + 1
                rs.AddNew
                rs.fields("id_model") = ip_id_model
                rs.fields("ni_ordering") = ni_ordering
                rs.fields("tx_sql") = vr_sql
                rs.Update:
            End If
        Next vr_sql
        '
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_dataset_tx ORDER BY ni_ordering ASC"
            Set rs = db.OpenRecordset(sql): With rs: Do Until .EOF
                Debug.Print "--- tmp_transformation_dataset_tx " & String(40, "-")
                Debug.Print "ni_ordering : '" & !ni_ordering & "'"
                Debug.Print "tx_sql      : '" & !tx_sql & ""
            .MoveNext: Loop: End With
        End If
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: md -> Match dataset(s) with "parts" of the "SQL"-statement and determine is alias is used. */)
        '
        If (1 = 1) Then ' /* Define Views (tmp_transformation_dataset_md_txt/prv/nxt) */
            Call CreateView("tmp_transformation_dataset_md_txt", "SELECT txt.id_model, CStr(txt.tx_sql) AS tx_sql, (txt.ni_ordering + 0) AS ni_ordering FROM tmp_transformation_dataset_tx AS txt;")
            Call CreateView("tmp_transformation_dataset_md_prv", "SELECT txt.id_model, CStr(txt.tx_sql) AS tx_sql_prev, (txt.ni_ordering + 1) AS ni_ordering FROM tmp_transformation_dataset_tx AS txt;")
            Call CreateView("tmp_transformation_dataset_md_nxt", "SELECT txt.id_model, CStr(txt.tx_sql) AS tx_sql_next, (txt.ni_ordering - 1) AS ni_ordering FROM tmp_transformation_dataset_tx AS txt;")
        End If
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_md_dst */
            Call CreateView("tmp_transformation_dataset_md_dst", "" _
            & vbNewLine & "SELECT " _
            & vbNewLine & "  dst.id_model, dst.id_dataset, dst.nm_target_schema, dst.nm_target_table, " _
            & vbNewLine & "  '[' & dst.nm_target_schema & '].'  & dst.nm_target_table & ''  AS tx_sql_match_1, " _
            & vbNewLine & "  ''  & dst.nm_target_schema &  '.[' & dst.nm_target_table & ']' AS tx_sql_match_2, " _
            & vbNewLine & "  '[' & dst.nm_target_schema & '].[' & dst.nm_target_table & ']' AS tx_sql_match_3,  " _
            & vbNewLine & "  ''  & dst.nm_target_schema &  '.'  & dst.nm_target_table & ''  AS tx_sql_match_4 " _
            & vbNewLine & "FROM  dta_dataset AS dst;")
        End If
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_md_select */
            Call CreateView("tmp_transformation_dataset_md_select", "" _
            & vbNewLine & "SELECT" _
            & vbNewLine & "  txt.ni_ordering," _
            & vbNewLine & "  txt.id_model AS id_source_model," _
            & vbNewLine & "  prv.tx_sql_prev," _
            & vbNewLine & "  txt.tx_sql," _
            & vbNewLine & "  nxt.tx_sql_next," _
            & vbNewLine & "  Nz([ds1].[id_dataset],       Nz([ds2].[id_dataset],       Nz([ds3].[id_dataset],       [ds4].[id_dataset])))       AS id_dataset," _
            & vbNewLine & "  Nz([ds1].[nm_target_schema], Nz([ds2].[nm_target_schema], Nz([ds3].[nm_target_schema], [ds4].[nm_target_schema]))) AS nm_target_schema," _
            & vbNewLine & "  Nz([ds1].[nm_target_table],  Nz([ds2].[nm_target_table],  Nz([ds3].[nm_target_table],  [ds4].[nm_target_table])))  AS nm_target_table," _
            & vbNewLine & "  IIf([nxt].[tx_sql_next] = 'AS', True, False) AS is_next_word_alias_keyword," _
            & vbNewLine & "  IIf([prv].[tx_sql_prev] IN ('FROM', 'JOIN'), True, False) AS is_prev_word_from_or_join_keyword," _
            & vbNewLine & "  IIf([txt].[tx_sql] IN ('INNER', 'CROSS', 'LEFT', 'RIGHT', 'FULL', 'JOIN', 'ON', 'FROM', 'WHERE', 'WHEN'), 1, 0) AS is_keyword" _
            & vbNewLine & "FROM (((((tmp_transformation_dataset_md_txt AS txt" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_nxt AS nxt ON (txt.ni_ordering = nxt.ni_ordering) AND (txt.id_model = nxt.id_model))" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_prv AS prv ON (txt.ni_ordering = prv.ni_ordering) AND (txt.id_model = prv.id_model))" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_dst AS ds1 ON (txt.id_model = ds1.id_model) AND (txt.tx_sql = ds1.tx_sql_match_1))" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_dst AS ds2 ON (txt.id_model = ds2.id_model) AND (txt.tx_sql = ds2.tx_sql_match_2))" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_dst AS ds3 ON (txt.id_model = ds3.id_model) AND (txt.tx_sql = ds3.tx_sql_match_3))" _
            & vbNewLine & "LEFT JOIN tmp_transformation_dataset_md_dst AS ds4 ON (txt.id_model = ds4.id_model) AND (txt.tx_sql = ds4.tx_sql_match_4);")
        End If
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_md_delete */
            Call CreateView("tmp_transformation_dataset_md_delete", "" _
            & vbNewLine & "DELETE tmp_transformation_dataset_md.*" _
            & vbNewLine & "FROM tmp_transformation_dataset_md;")
        End If
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_md_insert */
            Call CreateView("tmp_transformation_dataset_md_insert", "" _
            & vbNewLine & "INSERT INTO tmp_transformation_dataset_md ( ni_ordering, id_source_model, id_dataset, nm_target_schema, nm_target_table, tx_sql, tx_sql_prev, tx_sql_next, is_next_word_alias_keyword, is_prev_word_from_or_join_keyword, is_keyword )" _
            & vbNewLine & "SELECT tx.ni_ordering, tx.id_source_model, tx.id_dataset, tx.nm_target_schema, tx.nm_target_table, tx.tx_sql, tx.tx_sql_prev, tx.tx_sql_next, tx.is_next_word_alias_keyword, tx.is_prev_word_from_or_join_keyword, tx.is_keyword" _
            & vbNewLine & "FROM tmp_transformation_dataset_md_select AS tx;")
        End If
        '
        ' Truncate Table and then populate
        DoCmd.SetWarnings False ' Turn off Warnings
        DoCmd.OpenQuery ("tmp_transformation_dataset_md_delete")
        DoCmd.OpenQuery ("tmp_transformation_dataset_md_insert")
        DoCmd.SetWarnings True ' Turn on Warnings
        '
        ' Display Results
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_dataset_md ORDER BY ni_ordering ASC"
            Set rs = db.OpenRecordset(sql): With rs: Do Until .EOF
                Debug.Print "--- tmp_transformation_dataset_md " & String(40, "-")
                Debug.Print "ni_ordering                       : '" & !ni_ordering & "'"
                Debug.Print "id_source_model                   : '" & !id_source_model & "'"
                Debug.Print "id_dataset                        : '" & !id_dataset & "'"
                Debug.Print "nm_target_schema                  : '" & !nm_target_schema & "'"
                Debug.Print "nm_target_table                   : '" & !nm_target_table & "'"
                Debug.Print "tx_sql_prev                       : '" & !tx_sql_prev & "'"
                Debug.Print "tx_sql                            : '" & !tx_sql & "'"
                Debug.Print "tx_sql_next                       : '" & !tx_sql_next & "'"
                Debug.Print "is_next_word_alias_keyword        : '" & !is_next_word_alias_keyword & "'"
                Debug.Print "is_prev_word_from_or_join_keyword : '" & !is_prev_word_from_or_join_keyword & "'"
                Debug.Print "is_keyword                        : '" & !is_keyword & "'"
            .MoveNext: Loop: End With
        End If
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: ni -> determine "from/join"-type and "dataset" and its "alias". */)
        '
        If (1 = 1) Then '/* tmp_transformation_dataset_ni_als_foj_jtp */
            sql = emp & emp & "SELECT md.ni_ordering, md.id_source_model, md.id_dataset, md.nm_target_schema, md.nm_target_table,"
            sql = sql & nwl & "       md.tx_sql, md.tx_sql_prev, md.tx_sql_next,"
            sql = sql & nwl & "       md.is_next_word_alias_keyword, md.is_prev_word_from_or_join_keyword, md.is_keyword,"
            sql = sql & nwl & ""
            sql = sql & nwl & "  Iif(md.nm_target_schema IS NULL, -1,"
            sql = sql & nwl & "      Iif(md.is_next_word_alias_keyword = True,"
            sql = sql & nwl & "          md.ni_ordering + 2,"
            sql = sql & nwl & "          md.ni_ordering + 1)"
            sql = sql & nwl & "  ) AS ni_ordering_for_join_with_als,"
            sql = sql & nwl & ""
            sql = sql & nwl & "  Iif(md.nm_target_schema IS NULL, -1,"
            sql = sql & nwl & "      Iif(md.is_prev_word_from_or_join_keyword = True,"
            sql = sql & nwl & "          md.ni_ordering - 1,"
            sql = sql & nwl & "          -1)"
            sql = sql & nwl & "  ) AS ni_ordering_for_join_with_foj,"
            sql = sql & nwl & ""
            sql = sql & nwl & "  Iif(md.nm_target_schema IS NULL, -1,"
            sql = sql & nwl & "      Iif(md.is_prev_word_from_or_join_keyword = True,"
            sql = sql & nwl & "          Iif((SELECT CStr(tx_sql) FROM tmp_transformation_dataset_md AS tx WHERE tx.ni_ordering = (md.ni_ordering - 2)) IN ('INNER', 'CROSS', 'RIGHT', 'LEFT', 'FULL'),"
            sql = sql & nwl & "              md.ni_ordering - 2,"
            sql = sql & nwl & "              -1"
            sql = sql & nwl & "          ),"
            sql = sql & nwl & "          -1"
            sql = sql & nwl & "      )"
            sql = sql & nwl & "  ) AS ni_ordering_for_join_with_jtp"
            sql = sql & nwl & ""
            sql = sql & nwl & "FROM tmp_transformation_dataset_md AS md"
            sql = sql & nwl & ""
            sql = sql & nwl & "WHERE md.tx_sql NOT IN ('FROM', 'JOIN', 'INNER', 'CROSS', 'LEFT', 'RIGHT');"
            Call CreateView("tmp_transformation_dataset_ni_als_foj_jtp", sql)
        End If
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_ni_select */
            sql = emp & emp & "SELECT"
            sql = sql & nwl & "  md.tx_sql,"
            sql = sql & nwl & "  md.ni_ordering,"
            sql = sql & nwl & "  md.id_source_model,"
            sql = sql & nwl & "  md.id_dataset,"
            sql = sql & nwl & "  md.nm_target_schema,"
            sql = sql & nwl & "  md.nm_target_table,"
            sql = sql & nwl & "  "
            sql = sql & nwl & "  IIf(" ' /* cd_join_type */
            sql = sql & nwl & "      Nz(foj.tx_sql, 'n/a') NOT IN ('FROM', 'JOIN'), '',"
            sql = sql & nwl & "      Iif(Nz(jtp.tx_sql, '') IN ('INNER', 'CROSS', 'LEFT', 'RIGHT'), Nz(jtp.tx_sql, '') & ' ', '')"
            sql = sql & nwl & "  ) & Nz(foj.tx_sql, '') AS cd_join_type,"
            sql = sql & nwl & "  "
            sql = sql & nwl & "  Replace(Replace(Nz(als.tx_sql, '')," ' /* cd_alias */
            sql = sql & nwl & "    '[', ''), ']', ''"
            sql = sql & nwl & "  ) AS cd_alias,"
            sql = sql & nwl & "  "
            sql = sql & nwl & "  Iif("  ' /* tx_join_criteria */
            sql = sql & nwl & "    md.is_prev_word_from_or_join_keyword = 1 AND Nz(foj.tx_sql, '')  = 'FROM', NULL, Iif("
            sql = sql & nwl & "    md.is_prev_word_from_or_join_keyword = 1 AND Nz(foj.tx_sql, '') <> 'FROM', 'JOIN-CRIETERIA',"
            sql = sql & nwl & "    NULL)"
            sql = sql & nwl & "  ) As tx_join_criteria,"
            sql = sql & nwl & ""
            sql = sql & nwl & "  md.tx_sql_prev,"
            sql = sql & nwl & "  md.tx_sql_next,"
            sql = sql & nwl & "  md.ni_ordering_for_join_with_als"
            sql = sql & nwl & ""
            sql = sql & nwl & "FROM ((tmp_transformation_dataset_ni_als_foj_jtp AS md"
            sql = sql & nwl & ""
            sql = sql & nwl & "LEFT JOIN tmp_transformation_dataset_md AS als ON (als.ni_ordering = md.ni_ordering_for_join_with_als))"
            sql = sql & nwl & "LEFT JOIN tmp_transformation_dataset_md AS foj ON (foj.ni_ordering = md.ni_ordering_for_join_with_foj))"
            sql = sql & nwl & "LEFT JOIN tmp_transformation_dataset_md AS jtp ON (jtp.ni_ordering = md.ni_ordering_for_join_with_jtp);"
            Call CreateView("tmp_transformation_dataset_ni_select", sql)
        End If
        '
        If (1 = 1) Then ' /* Detele Query */
            Call CreateView("tmp_transformation_dataset_ni_delete", "" _
            & vbNewLine & "DELETE tmp_transformation_dataset_ni.*" _
            & vbNewLine & "FROM tmp_transformation_dataset_ni;")
        End If
        '
        If (1 = 1) Then ' /* Insert Query */
            Call CreateView("tmp_transformation_dataset_ni_insert", "" _
            & vbNewLine & "INSERT INTO tmp_transformation_dataset_ni (  ni_ordering, id_source_model, id_dataset, nm_target_schema, nm_target_table, cd_join_type, cd_alias, tx_join_criteria, tx_sql )" _
            & vbNewLine & "SELECT ni_ordering, id_source_model, id_dataset, nm_target_schema, nm_target_table, cd_join_type, cd_alias, tx_join_criteria, tx_sql" _
            & vbNewLine & "FROM tmp_transformation_dataset_ni_select;")
        End If
        '
        ' Truncate Table and then populate
        DoCmd.SetWarnings False ' Turn off Warnings
        DoCmd.OpenQuery ("tmp_transformation_dataset_ni_delete")
        DoCmd.OpenQuery ("tmp_transformation_dataset_ni_insert")
        DoCmd.SetWarnings True ' Turn on Warnings
        '
        ' Display Results
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_dataset_ni ORDER BY ni_ordering ASC"
            Set rs = db.OpenRecordset(sql): With rs: Do Until .EOF
                Debug.Print "--- tmp_transformation_dataset_ni " & String(40, "-")
                Debug.Print "ni_ordering                       : '" & !ni_ordering & "'"
                Debug.Print "id_source_model                   : '" & !id_source_model & "'"
                Debug.Print "id_dataset                        : '" & !id_dataset & "'"
                Debug.Print "nm_target_schema                  : '" & !nm_target_schema & "'"
                Debug.Print "nm_target_table                   : '" & !nm_target_table & "'"
                Debug.Print "cd_join_type                      : '" & !cd_join_type & "'"
                Debug.Print "cd_alias                          : '" & !cd_alias & "'"
                Debug.Print "tx_join_criteria                  : '" & !tx_join_criteria & "'"
                Debug.Print "tx_sql                            : '" & !tx_sql & "'"
            .MoveNext: Loop: End With
        End If
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: ls -> Last "ni_ordering" */
        '
        If (1 = 1) Then ' /* Detelet Query */
            Call CreateView("tmp_transformation_dataset_ls_delete", "" _
            & vbNewLine & "DELETE tmp_transformation_dataset_ls.*" _
            & vbNewLine & "FROM tmp_transformation_dataset_ls;")
        End If
        '
        If (1 = 1) Then ' /* Insert Query */
            Call CreateView("tmp_transformation_dataset_ls_insert", "" _
            & vbNewLine & "INSERT INTO tmp_transformation_dataset_ls (  ni_ordering )" _
            & vbNewLine & "SELECT MAX(ni.ni_ordering) AS ni_ordering" _
            & vbNewLine & "FROM tmp_transformation_dataset_ni AS ni;")
        End If
        '
        ' Truncate Table and then populate
        DoCmd.SetWarnings False ' Turn off Warnings
        DoCmd.OpenQuery ("tmp_transformation_dataset_ls_delete")
        DoCmd.OpenQuery ("tmp_transformation_dataset_ls_insert")
        DoCmd.SetWarnings True ' Turn on Warnings
        '
        ' Display Results
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_dataset_ls ORDER BY ni_ordering ASC"
            Set rs = db.OpenRecordset(sql): With rs: Do Until .EOF
                Debug.Print "--- tmp_transformation_dataset_ls " & String(40, "-")
                Debug.Print "ni_ordering                       : '" & !ni_ordering & "'"
            .MoveNext: Loop: End With
        End If
        '
    End If
    '
    If (1 = 1) Then ' /* "Temp"-table: ds -> Ordering for only "Datasets". */
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_ds_select */
            sql = emp & emp & "SELECT ni.cd_join_type,"
            sql = sql & nwl & "       ni.id_dataset,"
            sql = sql & nwl & "       ni.id_source_model,"
            sql = sql & nwl & "       ni.nm_target_schema,"
            sql = sql & nwl & "       ni.nm_target_table,"
            sql = sql & nwl & "       ni.cd_alias,"
            sql = sql & nwl & "       Iif(ni.cd_join_type = 'FROM', ni.ni_ordering, ni.ni_ordering_for_join_with_als+2) AS ni_ordering_from,"
            sql = sql & nwl & "       Iif(ni.cd_join_type = 'FROM', ni.ni_ordering, Nz(("
            sql = sql & nwl & "         SELECT MIN(lead.ni_ordering)"
            sql = sql & nwl & "         FROM [tmp_transformation_dataset_ni_select] AS lead"
            sql = sql & nwl & "         WHERE lead.ni_ordering > (ni.ni_ordering_for_join_with_als+1)"
            sql = sql & nwl & "         AND lead.tx_sql IN ('CROSS', 'LEFT', 'RIGT', 'FULL', 'JOIN')"
            sql = sql & nwl & "         AND ((lead.tx_sql <> 'JOIN' AND lead.tx_sql_next = 'JOIN') OR (lead.tx_sql = 'JOIN')) "
            sql = sql & nwl & "       ), [ls].[ni_ordering])) AS ni_ordering_till,"
            sql = sql & nwl & "       ni.tx_sql, ni.ni_ordering"
            sql = sql & nwl & "FROM tmp_transformation_dataset_ls AS ls, tmp_transformation_dataset_ni_select AS ni;"
            Call CreateView("tmp_transformation_dataset_ds_select", sql)
        End If
        '
        If (1 = 1) Then ' /* Detele Query */
            Call CreateView("tmp_transformation_dataset_ds_delete", "" _
            & vbNewLine & "DELETE tmp_transformation_dataset_ds.*" _
            & vbNewLine & "FROM tmp_transformation_dataset_ds;")
        End If
        '
        If (1 = 1) Then ' /* Insert Query */
            Call CreateView("tmp_transformation_dataset_ds_insert", "" _
            & vbNewLine & "INSERT INTO tmp_transformation_dataset_ds ( cd_join_type, id_dataset, id_source_model, nm_target_schema, nm_target_table, cd_alias, ni_ordering_from, ni_ordering_till )" _
            & vbNewLine & "SELECT cd_join_type, id_dataset, id_source_model, nm_target_schema, nm_target_table, cd_alias, ni_ordering_from, ni_ordering_till" _
            & vbNewLine & "FROM tmp_transformation_dataset_ds_select;")
        End If
        '
        ' Truncate Table and then populate
        DoCmd.SetWarnings False ' Turn off Warnings
        DoCmd.OpenQuery ("tmp_transformation_dataset_ds_delete")
        DoCmd.OpenQuery ("tmp_transformation_dataset_ds_insert")
        DoCmd.SetWarnings True ' Turn on Warnings
        '
        ' Display Results
        If ip_is_debugging Then
            sql = "SELECT * FROM tmp_transformation_dataset_ds ORDER BY ni_ordering_from ASC"
            Set rs = db.OpenRecordset(sql): With rs: Do Until .EOF
                Debug.Print "--- tmp_transformation_dataset_ds " & String(40, "-")
                Debug.Print "cd_join_type     : '" & !cd_join_type & "'"
                Debug.Print "id_dataset       : '" & !id_dataset & "'"
                Debug.Print "id_source_model  : '" & !id_source_model & "'"
                Debug.Print "nm_target_schema : '" & !nm_target_schema & "'"
                Debug.Print "nm_target_table  : '" & !nm_target_table & "'"
                Debug.Print "cd_alias         : '" & !cd_alias & "'"
                Debug.Print "ni_ordering_from : '" & !ni_ordering_from & "'"
                Debug.Print "ni_ordering_till : '" & !ni_ordering_till & "'"
            .MoveNext: Loop: End With
        End If
        '
    End If
    '
    If (1 = 1) Then '/* "Temp"-table: rs -> "Resultset". */
        '
        If (1 = 1) Then ' /* tmp_transformation_dataset_rs_md */
            sql = emp & emp & "SELECT ds.ni_ordering_from,"
            sql = sql & nwl & "       ds.ni_ordering_till,"
            sql = sql & nwl & "       md.ni_ordering,"
            sql = sql & nwl & "       md.tx_sql"
            sql = sql & nwl & "FROM tmp_transformation_dataset_ni_select AS md,"
            sql = sql & nwl & "     tmp_transformation_dataset_ds_select AS ds"
            sql = sql & nwl & "WHERE (((ds.cd_join_type)     <> '') "
            sql = sql & nwl & "AND    ((ds.ni_ordering_till) <> ds.ni_ordering_from) "
            sql = sql & nwl & "AND    ((md.ni_ordering)      >= ds.ni_ordering_from)"
            sql = sql & nwl & "AND    ((md.ni_ordering)      <= ds.ni_ordering_till));"
            Call CreateView("tmp_transformation_dataset_rs_md", sql)
        End If
        '
        If (1 = 1) Then ' /* ds.cd_join_type <> '' */
            sql = emp & emp & "SELECT"
            sql = sql & nwl & "    ds.cd_join_type,"
            sql = sql & nwl & "    ds.id_source_model,"
            sql = sql & nwl & "    ds.id_dataset AS id_source_dataset,"
            sql = sql & nwl & "    ds.nm_target_schema,"
            sql = sql & nwl & "    ds.nm_target_table,"
            sql = sql & nwl & "    ds.cd_alias,"
            sql = sql & nwl & "    md.ni_ordering AS ni_join_criteria,"
            sql = sql & nwl & "    IIf([ds].[cd_join_type] = 'FROM', NULL, [md].[tx_sql]) AS tx_join_criteria"
            sql = sql & nwl & "FROM tmp_transformation_dataset_ds AS ds"
            sql = sql & nwl & "    LEFT JOIN tmp_transformation_dataset_rs_md AS md"
            sql = sql & nwl & "    ON  (ds.ni_ordering_till = md.ni_ordering_till)"
            sql = sql & nwl & "    AND (ds.ni_ordering_from = md.ni_ordering_from)"
            sql = sql & nwl & "WHERE ds.cd_join_type <> ''"
            sql = sql & nwl & "ORDER BY md.ni_ordering_from ASC,"
            sql = sql & nwl & "         md.ni_ordering_till ASC,"
            sql = sql & nwl & "         md.ni_ordering ASC;"
            Call CreateView("tmp_transformantion_dataset_rs_select", sql)
        End If
        '
        If (1 = 1) Then ' /* ds.cd_join_type <> '' */
            sql = emp & emp & "SELECT DISTINCT"
            sql = sql & nwl & "    ds.cd_join_type,"
            sql = sql & nwl & "    ds.id_source_model,"
            sql = sql & nwl & "    ds.id_source_dataset,"
            sql = sql & nwl & "    ds.nm_target_schema,"
            sql = sql & nwl & "    ds.nm_target_table,"
            sql = sql & nwl & "    ds.cd_alias,"
            sql = sql & nwl & "    ds.cd_join_type & '|' & ds.nm_target_schema & '|' & ds.nm_target_table & '|' & ds.cd_alias AS id"
            sql = sql & nwl & "FROM tmp_transformantion_dataset_rs_select AS ds"
            Call CreateView("tmp_transformantion_dataset_rs_dist", sql)
        End If
        '
        If (1 = 1) Then ' /* ni/tx_join_criteria */
            sql = emp & emp & "SELECT '" & ip_id_model & "' AS id_model,"
            sql = sql & nwl & "       '" & ip_id_transformation_part & "' AS id_transformation_part,"
            sql = sql & nwl & "       CreateMD5("
            sql = sql & nwl & "       '|' & '" & ip_id_model & "' &"
            sql = sql & nwl & "       '|' & '" & ip_id_transformation_part & "' &"
            sql = sql & nwl & "       '|' & x.id_source_model &"
            sql = sql & nwl & "       '|' & x.id_source_dataset &"
            sql = sql & nwl & "       '|' & x.cd_alias &"
            sql = sql & nwl & "       '|') AS id_transformation_dataset,"
            sql = sql & nwl & "       (SELECT COUNT(*) FROM tmp_transformantion_dataset_rs_dist AS c WHERE c.id < x.id) AS ni_transformation_dataset,"
            sql = sql & nwl & "       x.cd_join_type,"
            sql = sql & nwl & "       x.id_source_model,"
            sql = sql & nwl & "       x.id_source_dataset,"
            sql = sql & nwl & "       x.cd_alias,"
            sql = sql & nwl & "       agg_list ("
            sql = sql & nwl & "         'tx_join_criteria', 'tmp_transformantion_dataset_rs_select',"
            sql = sql & nwl & "         'cd_join_type = ''' & x.cd_join_type & ''' AND nm_target_schema = ''' & x.nm_target_schema & ''' AND nm_target_table  = ''' & x.nm_target_table & ''' AND cd_alias = ''' & x.cd_alias & ''' ORDER BY ni_join_criteria ASC', ' '"
            sql = sql & nwl & "       ) AS tx_join_criteria"
            sql = sql & nwl & "FROM tmp_transformantion_dataset_rs_dist AS x"
            Call CreateView("tmp_transformantion_dataset_rs_ni_and_tx_join_criteria", sql)
        End If
        '
        If (1 = 1) Then ' /* String Aggregate */
            '
            sql = "DELETE * FROM dta_transformation_dataset WHERE id_transformation_part = '" & ip_id_transformation_part & "';"
            DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
            '
            sql = emp & emp & "INSERT INTO dta_transformation_dataset ("
            sql = sql & nwl & "  id_model, id_transformation_part, id_transformation_dataset, ni_transformation_dataset,"
            sql = sql & nwl & "  cd_join_type, id_source_model, id_source_dataset, cd_alias, tx_join_criteria"
            sql = sql & nwl & ")"
            sql = sql & nwl & "SELECT "
            sql = sql & nwl & "  id_model, id_transformation_part, id_transformation_dataset, ni_transformation_dataset,"
            sql = sql & nwl & "  cd_join_type, id_source_model, id_source_dataset, cd_alias, tx_join_criteria"
            sql = sql & nwl & "FROM tmp_transformantion_dataset_rs_ni_and_tx_join_criteria;"
            DoCmd.SetWarnings False: DoCmd.RunSQL sql: DoCmd.SetWarnings True
            '
        End If
    End If
End Sub





