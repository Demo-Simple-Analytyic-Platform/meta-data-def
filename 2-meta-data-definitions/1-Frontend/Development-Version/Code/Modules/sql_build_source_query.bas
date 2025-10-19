Attribute VB_Name = "sql_build_source_query"
Option Compare Database
Option Explicit
'
' Private Module Variables
Private p_tx_source_query As String
Private p_tx_sql_select   As String
Private p_tx_sql_from     As String
Private p_tx_sql_where    As String
Private p_tx_sql_group_by As String
Private p_tx_sql_having   As String
'
Public Function build_source_query(ip_id_model As String, ip_id_dataset As String, Optional ip_is_debugging As Boolean = False)
    '
    ' Example: ?build_source_query("5f4a1942465c575a1f5a5a575d1e191c", "06020d070f0d0a04030d0b0705021500", True)
    '
    ' Declare Local Variables
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    '
    ' Build SQL Query to loop throught the Transformation Parts
    sql = sql & emp & "SELECT ni_transformation_part, id_transformation_part"
    sql = sql & nwl & "FROM dta_transformation_part"
    sql = sql & nwl & "WHERE id_model   = '" & ip_id_model & "'"
    sql = sql & nwl & "AND   id_dataset = '" & ip_id_dataset & "'"
    sql = sql & nwl & "ORDER BY ni_transformation_part ASC;"
    If ip_is_debugging Then Debug.Print sql
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql): sql = ""
    '
    ' Loop throught the resultset
    Do While Not rst.EOF
        sql = sql & IIf(sql = "", "", nwl & "UNION ALL" & nwl) & build_sql_transformation_part(ip_id_model, rst.fields("id_transformation_part"), ip_is_debugging)
    rst.MoveNext: Loop
    '
    ' Return the Source Query build
    build_source_query = sql
    '
End Function

Public Function build_sql_clause_select(ip_id_model As String, ip_id_transformation_part As String, Optional ip_is_debugging As Boolean = False)
    '
    ' Example: ?build_sql_clause_select("5f4a1942465c575a1f5a5a575d1e191c", "6d2df191e0faa40455d3305cb20b28f2", True)
    '
    ' Declare Local Variables
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    '
    ' Build Mapping Attribute
    sql = sql & emp & "SELECT att.ni_ordering                      AS ni_ordering"
    sql = sql & nwl & "     , map.tx_transformation_column_mapping AS tx_mapping"
    sql = sql & nwl & "     , att.nm_target_column                 AS nm_column"
    sql = sql & nwl & ""
    sql = sql & nwl & "FROM dta_attribute AS att"
    sql = sql & nwl & ""
    sql = sql & nwl & "INNER JOIN dta_transformation_column_mapping AS map"
    sql = sql & nwl & " ON att.id_model     = map.id_model"
    sql = sql & nwl & "AND att.id_attribute = map.id_attribute"
    sql = sql & nwl & ""
    sql = sql & nwl & "WHERE map.id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & "AND   map.id_model               = '" & ip_id_model & "'"
    sql = sql & nwl & ""
    sql = sql & nwl & "ORDER BY att.ni_ordering"
    If ip_is_debugging Then Debug.Print sql
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql): sql = ""
    '
    ' Build SELECT-clause
    Do While Not rst.EOF: sql = sql & IIf(sql = "", "SELECT ", nwl & "     , ") & rst.fields("tx_mapping") & " AS " & rst.fields("nm_column"): rst.MoveNext: Loop
    If ip_is_debugging Then Debug.Print "SQL before replacing placeholders: " & sql
    '
    ' Replace the Meta-references for the Alias and Columns
    sql = placeholder_to_source_attributes(sql, ip_id_transformation_part)
    If ip_is_debugging Then Debug.Print "SQL After replacing placeholders: " & sql
    '
    ' Return "SELECT"-Clause
    build_sql_clause_select = sql
    '
End Function

Public Function build_sql_clause_from(ip_id_model As String, ip_id_transformation_part As String, Optional ip_is_debugging As Boolean = False)
    '
    ' Example: ?build_sql_clause_from("5f4a1942465c575a1f5a5a575d1e191c", "6d2df191e0faa40455d3305cb20b28f2", True)
    '
    ' Declare Local Variables
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    '
    ' Build Mapping Attribute
    sql = sql & emp & "SELECT tds.id_model"
    sql = sql & nwl & "     , tds.id_transformation_part"
    sql = sql & nwl & "     , tds.cd_join_type"
    sql = sql & nwl & "     , dst.nm_target_schema"
    sql = sql & nwl & "     , dst.nm_target_table"
    sql = sql & nwl & "     , tds.cd_alias"
    sql = sql & nwl & "     , tds.tx_join_criteria"
    sql = sql & nwl & "     , [tds].[cd_join_type] & ' [' & [dst].[nm_target_schema] & '.' & [dst].[nm_target_table] & '] AS [' & [tds].[cd_alias] & ']' & IIf("
    sql = sql & nwl & "         Len(Nz ([tds].[tx_join_criteria], '')) = 0,"
    sql = sql & nwl & "         '',"
    sql = sql & nwl & "         ' ' & Nz ([tds].[tx_join_criteria], '')"
    sql = sql & nwl & "    ) AS tx_sql"
    sql = sql & nwl & ""
    sql = sql & nwl & "FROM dta_transformation_dataset AS tds"
    sql = sql & nwl & ""
    sql = sql & nwl & "INNER JOIN dta_dataset AS dst"
    sql = sql & nwl & " ON (tds.id_source_dataset = dst.id_dataset)"
    sql = sql & nwl & "AND (tds.id_source_model = dst.id_model)"
    sql = sql & nwl & ""
    sql = sql & nwl & "WHERE tds.id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & "AND   tds.id_model               = '" & ip_id_model & "'"
    sql = sql & nwl & ""
    sql = sql & nwl & "ORDER BY tds.ni_transformation_dataset;"
    If ip_is_debugging Then Debug.Print sql
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql): sql = ""
    '
    ' Build SELECT-clause
    Do While Not rst.EOF: sql = sql & rst.fields("tx_sql") & vbNewLine: rst.MoveNext: Loop
    If ip_is_debugging Then Debug.Print "SQL before replacing placeholders: " & sql
    '
    ' Replace the Meta-references for the Alias and Columns
    sql = placeholder_to_source_attributes(sql, ip_id_transformation_part)
    If ip_is_debugging Then Debug.Print "SQL After replacing placeholders: " & sql
    '
    ' Return "FROM"-Clause
    build_sql_clause_from = sql
    '
End Function

Public Function build_sql_clause_where_group_by_having(ip_id_model As String, ip_id_transformation_part As String, Optional ip_is_debugging As Boolean = False)
    '
    ' Example 1: ?build_sql_clause_where_group_by_having("5f4a1942465c575a1f5a5a575d1e191c", "6d2df191e0faa40455d3305cb20b28f2", True)
    '
    ' Declare Local Variables
    Dim rst As DAO.Recordset
    Dim dbs As DAO.Database: Set dbs = CurrentDb
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    '
    ' Build Mapping Attribute
    sql = sql & emp & "SELECT Nz(tpt.tx_transformation_part_where_clause, '')    AS tx_sql_where"
    sql = sql & nwl & "     , Nz(tpt.tx_transformation_part_group_by_clause, '') AS tx_sql_group_by"
    sql = sql & nwl & "     , Nz(tpt.tx_transformation_part_having_clause, '')   AS tx_sql_having"
    sql = sql & nwl & ""
    sql = sql & nwl & "FROM dta_transformation_part AS tpt"
    sql = sql & nwl & ""
    sql = sql & nwl & "WHERE tpt.id_transformation_part = '" & ip_id_transformation_part & "'"
    sql = sql & nwl & "AND   tpt.id_model               = '" & ip_id_model & "';"
    If ip_is_debugging Then Debug.Print sql
    Set rst = CurrentDb.OpenRecordset(sql): sql = ""
    '
    ' Build WHERE/GROUP BY and HAVING-clause
    sql = sql & IIf(rst.fields("tx_sql_where") = "", "", IIf(sql = "", "", vbNewLine) & rst.fields("tx_sql_where"))
    sql = sql & IIf(rst.fields("tx_sql_group_by") = "", "", IIf(sql = "", "", vbNewLine) & rst.fields("tx_sql_group_by") & vbNewLine)
    sql = sql & IIf(rst.fields("tx_sql_having") = "", "", IIf(sql = "", "", vbNewLine) & rst.fields("tx_sql_having") & vbNewLine)
    If ip_is_debugging Then Debug.Print "SQL before replacing placeholders: " & sql
    '
    ' Done With recordset
    rst.Close
    '
    ' Replace the Meta-references for the Alias and Columns
    sql = placeholder_to_source_attributes(sql, ip_id_transformation_part)
    '
    ' Return "WHERE/GROUP BY and HAVING"-Clause after
    If ip_is_debugging Then Debug.Print "SQL After replacing placeholders: " & sql
    build_sql_clause_where_group_by_having = sql
    '
End Function

Public Function build_sql_transformation_part(ip_id_model As String, ip_id_transformation_part As String, Optional ip_is_debugging As Boolean = False)
    '
    ' Example: ?build_sql_transformation_part("5f4a1942465c575a1f5a5a575d1e191c", "6d2df191e0faa40455d3305cb20b28f2", True)
    '
    ' Declare Local Variables
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    '
    ' SQL for various clauses
    Dim sql_slct As String: sql_slct = build_sql_clause_select(ip_id_model, ip_id_transformation_part, ip_is_debugging)
    Dim sql_from As String: sql_from = build_sql_clause_from(ip_id_model, ip_id_transformation_part, ip_is_debugging)
    Dim sql_wgbh As String: sql_wgbh = build_sql_clause_where_group_by_having(ip_id_model, ip_id_transformation_part, ip_is_debugging)
    '
    ' Combining all SQL Clauses
    sql = sql & IIf(sql_slct = "", "", IIf(sql = "", "", vbNewLine) & sql_slct)
    sql = sql & IIf(sql_from = "", "", IIf(sql = "", "", vbNewLine) & sql_from)
    sql = sql & IIf(sql_wgbh = "", "", IIf(sql = "", "", vbNewLine) & sql_wgbh)
    '
    ' Return Transformation Part
    If ip_is_debugging Then Debug.Print sql
    build_sql_transformation_part = sql
    '
End Function