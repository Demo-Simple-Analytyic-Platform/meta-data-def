Attribute VB_Name = "sql_strip_comments"
Option Compare Database
Option Explicit

Public Function StripSQLComments(inputTextWithComments As String) As String
    ' VBA conversion of svf_strip_comments SQL function
    ' Removes all comments from SQL text while preserving string literals
    
    Dim sqlCode As String
    Dim textWithoutComments As String
    Dim i As Integer
    Dim char1 As String
    Dim char2 As String
    Dim trailingComment As String
    Dim whackCounter As Integer
    Dim maxLength As Integer
    
    ' Initialize variables
    sqlCode = inputTextWithComments
    textWithoutComments = ""
    i = 1  ' VBA uses 1-based indexing
    trailingComment = "N"
    whackCounter = 0
    maxLength = Len(inputTextWithComments)
    
    ' Loop through every single character in the text
    Do While i <= maxLength
        
        ' Extract the next character
        char1 = Mid(sqlCode, i, 1)
        
        ' Determine if character is NOT comment-related
        If char1 <> "-" And char1 <> "/" And char1 <> "'" And char1 <> "*" Then
            
            ' Check if character is newline character (CR or LF)
            If char1 = Chr(13) Or char1 = Chr(10) Then
                trailingComment = "N"
            ' Check if NOT space or tab
            ElseIf char1 <> Chr(32) And char1 <> Chr(9) And trailingComment = "N" Then
                trailingComment = "Y"
            End If
            
            ' Add character to text without comments
            If whackCounter = 0 Then
                textWithoutComments = textWithoutComments & char1
            End If
            
            ' Move to next character
            i = i + 1
            
        Else
            ' If comment-related character
            
            ' Set char2 with current character
            char2 = char1
            
            ' Fetch next character (if exists)
            If i < maxLength Then
                char1 = Mid(sqlCode, i + 1, 1)
            Else
                char1 = ""
            End If
            
            ' Check if line comment characters are found (--)
            If char1 = "-" And char2 = "-" And whackCounter = 0 Then
                
                ' Loop through characters until newline character is found
                Do While i <= maxLength And Mid(sqlCode, i, 1) <> Chr(13) And Mid(sqlCode, i, 1) <> Chr(10)
                    i = i + 1
                Loop
                
                ' Check what type of newline character is found
                If i <= maxLength Then
                    If Mid(sqlCode, i, 1) = Chr(13) And trailingComment = "N" Then
                        i = i + 1
                    End If
                    If i <= maxLength And Mid(sqlCode, i, 1) = Chr(10) And trailingComment = "N" Then
                        i = i + 1
                    End If
                End If
                
            ' Check if block comment start is found (/*)
            ElseIf char1 = "*" And char2 = "/" Then
                
                whackCounter = whackCounter + 1
                
                ' Skip both characters
                i = i + 2
                
            ' Check if block comment end is found (*/)
            ElseIf char1 = "/" And char2 = "*" Then
                
                whackCounter = whackCounter - 1
                
                ' Skip both characters
                i = i + 2
                
            ' Check if single quote character is found (start of string literal)
            ElseIf char2 = "'" And whackCounter = 0 Then
                
                ' Add found character to output
                textWithoutComments = textWithoutComments & char2
                    
                ' Move to next character
                i = i + 1
                    
                ' Loop through text until closing single quote is found
                Do While i <= maxLength And Mid(sqlCode, i, 1) <> "'"
                    
                    ' Add found character to output
                    textWithoutComments = textWithoutComments & Mid(sqlCode, i, 1)
                    
                    ' Move to next character
                    i = i + 1
                    
                Loop
                
                ' Add the closing quote
                If i <= maxLength Then
                    textWithoutComments = textWithoutComments & Mid(sqlCode, i, 1)
                    i = i + 1
                End If
                
            Else
                
                ' Add found character to output (if not in block comment)
                If whackCounter = 0 Then
                    textWithoutComments = textWithoutComments & char2
                End If
                
                ' Move to next character
                i = i + 1
                
            End If
            
        End If
        
    Loop
    
    ' Return the result
    StripSQLComments = textWithoutComments
    
End Function

' Helper function for testing
Sub TestStripSQLComments()
    Dim testSQL As String
    Dim result As String
    
    ' Test case 1: Line comments
    testSQL = "SELECT * FROM table1 -- This is a comment" & vbCrLf & _
              "WHERE id = 1 -- Another comment"
    
    result = StripSQLComments(testSQL)
    Debug.Print "=== Test 1: Line Comments ==="
    Debug.Print "Original:"
    Debug.Print testSQL
    Debug.Print "Stripped:"
    Debug.Print result
    Debug.Print ""
    
    ' Test case 2: Block comments
    testSQL = "SELECT /* comment */ column1, column2 /* another comment */ FROM table1"
    
    result = StripSQLComments(testSQL)
    Debug.Print "=== Test 2: Block Comments ==="
    Debug.Print "Original:"
    Debug.Print testSQL
    Debug.Print "Stripped:"
    Debug.Print result
    Debug.Print ""
    
    ' Test case 3: String literals with comment-like content
    testSQL = "SELECT * FROM table1 WHERE description = 'This -- is not a comment' AND name = 'Value /* not a comment */'"
    
    result = StripSQLComments(testSQL)
    Debug.Print "=== Test 3: String Literals ==="
    Debug.Print "Original:"
    Debug.Print testSQL
    Debug.Print "Stripped:"
    Debug.Print result
    Debug.Print ""
    
    ' Test case 4: Mixed comments
    testSQL = "/* Header comment */" & vbCrLf & _
              "SELECT col1, -- inline comment" & vbCrLf & _
              "       col2 /* block comment */" & vbCrLf & _
              "FROM table1" & vbCrLf & _
              "WHERE value = 'test -- not a comment'" & vbCrLf & _
              "-- Final comment"
    
    result = StripSQLComments(testSQL)
    Debug.Print "=== Test 4: Mixed Comments ==="
    Debug.Print "Original:"
    Debug.Print testSQL
    Debug.Print "Stripped:"
    Debug.Print result
End Sub

' Integration function for use with SQL analysis
Function AnalyzeSQLWithoutComments(sqlString As String) As SQLAnalysisResult
    ' Combines comment stripping with SQL analysis
    Dim cleanSQL As String
    
    ' First strip comments
    cleanSQL = StripSQLComments(sqlString)
    
    ' Then analyze the clean SQL
    AnalyzeSQLWithoutComments = AnalyzeSQL(cleanSQL)
End Function