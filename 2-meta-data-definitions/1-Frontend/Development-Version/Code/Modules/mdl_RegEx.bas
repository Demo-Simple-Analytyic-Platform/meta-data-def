Attribute VB_Name = "mdl_RegEx"
Option Compare Database
Option Explicit

Public Function RegExMatch(ByVal inputString As String, ByVal pattern As String) As Boolean
    ' Function to check if a string matches a RegEx pattern
    ' Returns True if match found, False otherwise
    
    On Error GoTo ErrorHandler
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .pattern = pattern
        .IgnoreCase = True
        .Global = True
    End With
    
    ' Test if pattern matches
    RegExMatch = regEx.Test(inputString)
    
ExitFunction:
    Set regEx = Nothing
    Exit Function
    
ErrorHandler:
    RegExMatch = False
    Resume ExitFunction
End Function

Public Sub test_RegExMatch()

    ' Example 1: Check if string contains email pattern
    Dim result As Boolean
    result = RegExMatch("test@example.com", "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")
    Debug.Print result
    ' Returns: True
    
    ' Example 2: Check if string is a valid date format (YYYY-MM-DD)
    result = RegExMatch("2025-10-31", "^\d{4}-\d{2}-\d{2}$")
    Debug.Print result
    ' Returns: True
    
    ' Example 3: Check for specific SQL pattern
    result = RegExMatch("SELECT * FROM table", "^SELECT.*FROM.*$")
    Debug.Print result
    ' Returns: True
    
    ' Example 4: No match
    result = RegExMatch("Hello World", "^\d+$")
    Debug.Print result
    ' Returns: False


End Sub