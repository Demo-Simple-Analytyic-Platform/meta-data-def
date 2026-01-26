Attribute VB_Name = "sql_parsing"
Option Compare Database
Option Explicit

Public Type SQLAnalysisResult
    HasSubqueries     As Boolean
    HasCTEs           As Boolean
    SubqueryCount     As Integer
    CTECount          As Integer
    SubqueryPositions As String
    CTENames          As String
    CTEPositions      As String
End Type
'
Public Sub parse_sql_statement_all_transfromations()
    '
    ' Declare Local Variables
    Dim sql As String:        sql = "SELECT id_model, id_dataset, fn_dataset FROM dta_dataset WHERE is_ingestion = False"
    Dim rst As Recordset: Set rst = CurrentDb.OpenRecordset(sql)
    '
    ' Extract Transformation Parts
    Do While Not rst.EOF
        '
        Debug.Print "ID Model : '" & rst!id_model & "' | ID Dataset : '" & rst!id_dataset & "'"
        Debug.Print "FN Dataset : '" & rst!fn_dataset & "'"
        Call parse_transformation_parts(rst!id_model, rst!id_dataset, False, False)
        '
    rst.MoveNext: Loop: rst.Close
    '
    ' Done
    Debug.Print "All done"
    '
End Sub
'
Public Sub parse_sql_statement(ip_id_model As String, ip_id_dataset As String)
    '
    ' Example:
    'Call parse_sql_statement("5f4a1942465c575a1f5a5a575d1e191c", "06030d080305000102030a00040c1507")
    '
    ' Extract Transformation Parts
    Call parse_transformation_parts(ip_id_model, ip_id_dataset)
    '
End Sub
'
Function agg_list(ConcatColumn As String, Tbl As String, _
             Optional Criteria As String = "", _
            Optional Delimiter As String = ", ") As String
                 
' ?agg_list("tx_join_criteria", "sometable_name", "cd_join_type = 'JOIN' AND nm_target_schema = 'dta_generic_utilities' AND nm_target_table = 'list_1_till_10000' AND cd_alias = 'ls' ORDER BY ni_join_criteria", " ")

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim result As String

    sql = "SELECT [" & ConcatColumn & "] FROM [" & Tbl & "]"
    If Criteria <> "" Then sql = sql & " WHERE " & Criteria

    Set rs = CurrentDb.OpenRecordset(sql)
    Do While Not rs.EOF
        result = result & rs.fields(0).Value & Delimiter
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Len(result) > 0 Then result = Left(result, Len(result) - Len(Delimiter))
    agg_list = result

End Function
'


Function AnalyzeSQL(sqlString As String) As SQLAnalysisResult
    ' Main function to analyze SQL string for subqueries and CTEs
    ' Returns a comprehensive analysis result
    
    Dim result As SQLAnalysisResult
    Dim cleanSQL As String
    
    ' Clean the SQL string (remove extra spaces, normalize)
    cleanSQL = CleanSQLString(sqlString)
    
    ' Initialize result
    result.HasSubqueries = False
    result.HasCTEs = False
    result.SubqueryCount = 0
    result.CTECount = 0
    result.SubqueryPositions = ""
    result.CTENames = ""
    result.CTEPositions = ""
    
    ' Detect CTEs
    DetectCTEs cleanSQL, result
    
    ' Detect Subqueries
    DetectSubqueries cleanSQL, result
    
    AnalyzeSQL = result
End Function

Function DetectCTEs(sqlString As String, ByRef result As SQLAnalysisResult)
    ' Detects Common Table Expressions in SQL string
    
    Dim upperSQL As String
    Dim withPos As Integer
    Dim currentPos As Integer
    Dim cteStart As Integer
    Dim cteEnd As Integer
    Dim cteName As String
    Dim parenLevel As Integer
    Dim i As Integer
    Dim char As String
    Dim inQuotes As Boolean
    Dim quoteChar As String
    
    upperSQL = UCase(Trim(sqlString))
    
    ' Look for WITH keyword at the beginning or after whitespace
    currentPos = 1
    
    Do While currentPos <= Len(upperSQL)
        withPos = InStr(currentPos, upperSQL, "WITH")
        
        If withPos = 0 Then Exit Do
        
        ' Check if WITH is at the beginning or preceded by whitespace/newline
        If withPos = 1 Or IsWhitespace(Mid(upperSQL, IIf(withPos = 1, 1, withPos - 1), 1)) Then
            ' Verify it's not part of another word
            If withPos + 4 <= Len(upperSQL) Then
                If IsWhitespace(Mid(upperSQL, withPos + 4, 1)) Then
                    ' Found a CTE
                    result.HasCTEs = True
                    result.CTECount = result.CTECount + 1
                    
                    ' Extract CTE name and position
                    cteStart = withPos + 4
                    ' Skip whitespace
                    While cteStart <= Len(upperSQL) And IsWhitespace(Mid(upperSQL, cteStart, 1))
                        cteStart = cteStart + 1
                    Wend
                    
                    ' Find CTE name (until space or opening parenthesis)
                    cteEnd = cteStart
                    While cteEnd <= Len(upperSQL)
                        char = Mid(upperSQL, cteEnd, 1)
                        If IsWhitespace(char) Or char = "(" Then Exit Do
                        cteEnd = cteEnd + 1
                    Wend
                    
                    If cteEnd > cteStart Then
                        cteName = Mid(sqlString, cteStart, cteEnd - cteStart)
                        If result.CTENames <> "" Then result.CTENames = result.CTENames & ", "
                        result.CTENames = result.CTENames & cteName
                        
                        If result.CTEPositions <> "" Then result.CTEPositions = result.CTEPositions & ", "
                        result.CTEPositions = result.CTEPositions & CStr(withPos)
                    End If
                End If
            End If
        End If
        
        currentPos = withPos + 1
    Loop
End Function

Function DetectSubqueries(sqlString As String, ByRef result As SQLAnalysisResult)
    ' Detects subqueries in SQL string using regex patterns
    
    On Error GoTo RegexError
    
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim cleanSQL As String
    Dim i As Integer
    
    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Clean SQL: remove string literals to avoid false matches
    cleanSQL = RemoveStringLiterals(sqlString)
    
    ' Configure regex for subquery detection
    regex.IgnoreCase = True
    regex.Global = True
    
    ' Pattern explanation:
    ' \(\s* - Opening parenthesis followed by optional whitespace
    ' SELECT\b - SELECT keyword with word boundary
    ' (?:[^()]*|\([^()]*\))* - Match content that doesn't contain unmatched parentheses
    ' \) - Closing parenthesis
    regex.pattern = "\(\s*SELECT\b(?:[^()]*|\([^()]*\))*\)"
    
    ' Find all matches
    Set matches = regex.Execute(cleanSQL)
    
    ' Process matches
    result.SubqueryCount = matches.Count
    result.HasSubqueries = (matches.Count > 0)
    
    ' Collect positions
    For i = 0 To matches.Count - 1
        Set match = matches(i)
        If result.SubqueryPositions <> "" Then result.SubqueryPositions = result.SubqueryPositions & ", "
        result.SubqueryPositions = result.SubqueryPositions & CStr(match.FirstIndex + 1) ' VBA is 1-based
    Next i
    
    Exit Function
    
RegexError:
    ' Fallback to original method if regex fails
    'DetectSubqueriesOriginal cleanSQL, result
    Exit Function
    
End Function


Private Function RemoveStringLiterals(sqlString As String) As String
    ' Removes string literals from SQL to avoid false regex matches
    
    On Error GoTo SimpleReturn
    
    Dim regex As Object
    Dim result As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    
    ' Remove single-quoted strings
    regex.pattern = "'(?:[^'\\]|\\.)*'"
    result = regex.Replace(sqlString, "''")
    
    ' Remove double-quoted strings
    regex.pattern = """(?:[^""\\]|\\.)*"""
    result = regex.Replace(result, """""")
    
    RemoveStringLiterals = result
    Exit Function
    
SimpleReturn:
    ' If regex fails, return original string
    RemoveStringLiterals = sqlString
End Function

Function HasSubqueries(sqlString As String) As Boolean
    ' Simple function to check if SQL contains subqueries
    Dim result As SQLAnalysisResult
    result = AnalyzeSQL(sqlString)
    HasSubqueries = result.HasSubqueries
End Function

Function HasCTEs(sqlString As String) As Boolean
    ' Simple function to check if SQL contains CTEs
    Dim result As SQLAnalysisResult
    result = AnalyzeSQL(sqlString)
    HasCTEs = result.HasCTEs
End Function

Function GetCTENames(sqlString As String) As String
    ' Returns comma-separated list of CTE names
    Dim result As SQLAnalysisResult
    result = AnalyzeSQL(sqlString)
    GetCTENames = result.CTENames
End Function

Function CountSubqueries(sqlString As String) As Integer
    ' Returns the number of subqueries found
    Dim result As SQLAnalysisResult
    result = AnalyzeSQL(sqlString)
    CountSubqueries = result.SubqueryCount
End Function

Function CountCTEs(sqlString As String) As Integer
    ' Returns the number of CTEs found
    Dim result As SQLAnalysisResult
    result = AnalyzeSQL(sqlString)
    CountCTEs = result.CTECount
End Function

Function MinifySQL(ByVal inputText As String) As String

    Dim txMinified As String
    Dim cp As Long, np As Long, lp As Long
    Dim CC As String, nc As String
    Dim pd As Boolean
    Dim i As Long, m As Long
    Dim e As String

    ' Step 1: Strip comments (you must implement this separately)
    txMinified = Trim(StripSQLComments(inputText))

    ' Step 2: Replace newlines with spaces
    txMinified = Replace(txMinified, vbCrLf, " ")
    txMinified = Replace(txMinified, vbCr, " ")
    txMinified = Replace(txMinified, vbLf, " ")

    ' Step 3: Initialize variables
    cp = 1
    lp = Len(txMinified)
    i = 0
    m = 1000000

    ' Step 4: Loop through characters
    Do While cp < lp And i < m
        i = i + 1
        pd = False
        np = cp + 1

        If cp <= lp Then CC = Mid(txMinified, cp, 1)
        If np <= lp Then nc = Mid(txMinified, np, 1)

        If Not pd And CC = " " And nc = " " Then
            pd = True
            txMinified = Left(txMinified, cp) & Mid(txMinified, np + 1)
            
        ElseIf Not pd And CC = " " And nc = "[" Then
            pd = True
            cp = InStr(cp + 1, txMinified, "]") + 1
            
        ElseIf Not pd And CC = " " And nc = "'" Then
            pd = True
            cp = InStr(cp + 1, txMinified, "'") + 1
            
        ElseIf Not pd And CC <> " " And nc = "[" Then
            pd = True
            cp = InStr(cp + 1, txMinified, "]") + 1
            
        ElseIf Not pd And CC <> " " And nc = "'" Then
            pd = True
            cp = InStr(cp + 1, txMinified, "'") + 1
            
        ElseIf Not pd Then
            pd = True
            cp = cp + 1
            
        End If

        lp = Len(txMinified)
        If cp = 0 Then cp = lp
        
    Loop

    ' Step 5: Handle max iteration warning
    If i = m Then
        e = "--- Warning --------------------------------------------------------------" & vbCrLf & _
            "  Function MinifySQL reached maximum iterations (" & m & ")." & vbCrLf & _
            "  Please check input for irregularities." & vbCrLf & _
            "---------------------------------------------------------------------------"
        txMinified = e
    End If

    ' Step 6: Return result
    MinifySQL = txMinified
    
End Function

Private Function CleanSQLString(sqlString As String) As String
    ' Cleans SQL string by normalizing whitespace
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim prevChar As String
    
    result = ""
    prevChar = ""
    
    For i = 1 To Len(sqlString)
        char = Mid(sqlString, i, 1)
        
        ' Convert tabs and newlines to spaces
        If char = vbTab Or char = vbCr Or char = vbLf Then
            char = " "
        End If
        
        ' Avoid multiple consecutive spaces
        If char = " " And prevChar = " " Then
            ' Skip this space
        Else
            result = result & char
        End If
        
        prevChar = char
    Next i
    
    CleanSQLString = Trim(result)
End Function

Private Function IsWhitespace(char As String) As Boolean
    ' Checks if character is whitespace
    IsWhitespace = (char = " " Or char = vbTab Or char = vbCr Or char = vbLf)
End Function

Function GetSQLAnalysisReport(result As SQLAnalysisResult) As String
    ' Returns a formatted report of SQL analysis
    Dim report As String
    
    report = "SQL Analysis Report" & vbCrLf
    report = report & "===================" & vbCrLf
    report = report & "Has CTEs: " & IIf(result.HasCTEs, "Yes", "No") & vbCrLf
    report = report & "CTE Count: " & result.CTECount & vbCrLf
    If result.CTENames <> "" Then
        report = report & "CTE Names: " & result.CTENames & vbCrLf
    End If
    report = report & "Has Subqueries: " & IIf(result.HasSubqueries, "Yes", "No") & vbCrLf
    report = report & "Subquery Count: " & result.SubqueryCount & vbCrLf
    
    GetSQLAnalysisReport = report
End Function

' Example usage function
Sub TestSQLAnalysis()
    Dim testSQL As String
    Dim result As SQLAnalysisResult
    
    ' Test SQL with CTE and subquery
    testSQL = "WITH sales_summary AS (" & vbCrLf & _
              "    SELECT customer_id, SUM(amount) as total" & vbCrLf & _
              "    FROM sales" & vbCrLf & _
              "    WHERE sale_date > '2023-01-01'" & vbCrLf & _
              ")" & vbCrLf & _
              "SELECT s.customer_id, s.total, c.name" & vbCrLf & _
              "FROM sales_summary s" & vbCrLf & _
              "JOIN customers c ON s.customer_id = c.id" & vbCrLf & _
              "WHERE s.total > (SELECT AVG(total) FROM sales_summary)"
    
    result = AnalyzeSQL(testSQL)
    
    Debug.Print GetSQLAnalysisReport(result)
End Sub

Public Sub test_minify_sql()

    Dim n As String: n = vbNewLine
    Dim s As String: s = ""
    s = s & s & "SELECT [ed].[cd_dividend_symbol]        AS [cd_symbol]"
    s = s & n & "     , [ed].[nr_dividend_amount_median] AS [nr_expected_dividend]"
    s = s & n & "     , DATEADD("
    s = s & n & "         DAY,"
    s = s & n & "         [ls].[ni_index] * [ed].[ni_days_between_median],"
    s = s & n & "         [ed].[dt_last_dividend]"
    s = s & n & "        ) AS [dt_expected_dividend]"
    s = s & n & "     "
    s = s & n & "FROM  [dta_yahoo_stock].[dividend_median] AS [ed]"
    s = s & n & ""
    s = s & n & "LEFT JOIN [dta_generic_utilities].[list_1_till_10000] AS [ls] ON [ls].[meta_is_active] = 1 AND [ls].[ni_index] <= ((30 * 12) + 1)"
    s = s & n & ""
    s = s & n & "WHERE [ed].[meta_is_active] = 1"
    s = s & n & "AND   DATEADD(DAY, ([ls].[ni_index] * [ed].[ni_days_between_median]), [ed].[dt_last_dividend]) <= DATEADD(YEAR, 30, GETDATE())"
    s = MinifySQL(s)
    Debug.Print s
    
End Sub

Public Function source_attributes_to_placeholder(ip_tx_sql As String, ip_id_transformation_part As String) As String
    '
    ' Declare Local Variables
    Dim rst As DAO.Recordset
    Dim dbs As DAO.Database: Set dbs = CurrentDb
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    Dim out As String: out = ip_tx_sql
    '
    ' Buils SQL Statement for extraction of placeholders vs source_attributes
    sql = sql & emp & "SELECT tuc.tx_source_attribute_1 AS tx_attribute_1"
    sql = sql & nwl & "     , tuc.tx_source_attribute_2 AS tx_attribute_2"
    sql = sql & nwl & "     , tuc.tx_source_attribute_3 AS tx_attribute_3"
    sql = sql & nwl & "     , tuc.tx_source_attribute_4 AS tx_attribute_4"
    sql = sql & nwl & "     , tuc.tx_placeholder        AS tx_placeholder"
    sql = sql & nwl & "FROM tmp_transformation_utilized_columns AS tuc "
    sql = sql & nwl & "WHERE tuc.id_transformation_part = '" & ip_id_transformation_part & "'"
    '
    ' Replace the Meta-references for the Alias and Columns
    Set rst = dbs.OpenRecordset(sql): Do While Not rst.EOF:
        out = Replace(out, rst!tx_attribute_1, rst!tx_placeholder)
        out = Replace(out, rst!tx_attribute_2, rst!tx_placeholder)
        out = Replace(out, rst!tx_attribute_3, rst!tx_placeholder)
        out = Replace(out, rst!tx_attribute_4, rst!tx_placeholder)
    rst.MoveNext: Loop: rst.Close
    '
    ' Returnm out
    source_attributes_to_placeholder = out
    '
End Function

Public Function placeholder_to_source_attributes(ip_tx_sql As String, ip_id_transformation_part As String) As String
    '
    ' Declare Local Variables
    Dim rst As DAO.Recordset
    Dim dbs As DAO.Database: Set dbs = CurrentDb
    Dim sql As String: sql = ""
    Dim emp As String: emp = ""
    Dim nwl As String: nwl = vbNewLine
    Dim out As String: out = ip_tx_sql
    '
    ' Buils SQL Statement for extraction of placeholders vs source_attributes
    sql = sql & emp & "SELECT tuc.tx_source_attribute_3 AS tx_attribute"
    sql = sql & nwl & "     , tuc.tx_placeholder        AS tx_placeholder"
    sql = sql & nwl & "FROM tmp_transformation_utilized_columns AS tuc "
    sql = sql & nwl & "WHERE tuc.id_transformation_part = '" & ip_id_transformation_part & "'"
    '
    ' Replace the Meta-references for the Alias and Columns
    Set rst = dbs.OpenRecordset(sql): Do While Not rst.EOF: out = Replace(out, rst!tx_placeholder, rst!tx_attribute): rst.MoveNext: Loop: rst.Close
    '
    ' Returnm out
    placeholder_to_source_attributes = out
    '
End Function