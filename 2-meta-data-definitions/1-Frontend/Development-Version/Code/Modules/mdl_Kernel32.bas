Option Compare Database
Option Explicit
'
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' Declare the Windows API function
#If VBA7 Then
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If


Public Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Function CheckDecimalSeparator() As String
    '
    Const LOCALE_USER_DEFAULT As Long = &H400
    Const LOCALE_SDECIMAL As Long = &HE
    '
    Dim buffer As String
    Dim Length As Long
    '
    ' Allocate buffer
    buffer = String$(10, vbNullChar)
    '
    ' Retrieve the decimal separator setting
    Length = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, buffer, Len(buffer))
    '
    ' Set return value
    CheckDecimalSeparator = Left$(buffer, Length - 1)
    
    
End Function