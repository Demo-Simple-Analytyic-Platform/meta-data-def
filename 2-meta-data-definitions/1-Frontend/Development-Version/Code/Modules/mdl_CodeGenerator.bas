Attribute VB_Name = "mdl_CodeGenerator"
''
' (c) Mehmet (P.R.M.) Misset
'
' Code Generator
'
' Errors:
' n/a
'
' @module CodeGenerator
' @author mehmet.misset@misset-data-analytics.nl
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Compare Database
Option Explicit
'
'
' -----------------------------------------------------------------------------
' Function    : get_random_code
' Description : Function returns random character string with default lente of
'               32 characters.
' -----------------------------------------------------------------------------
Public Function get_random_code(Optional ip_length As Byte = 32) As String

'    Dim hash As String
'    Dim i As Integer
'
'    hash = CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
'    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
'    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
'    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
    
    get_random_code = LCase(GenerateRandomGUID)

End Function

Function GenerateRandomGUID() As String
    ' Generates a random GUID as a text string with only numbers and letters
    ' Returns a 32-character string (no hyphens)
    
    Dim guidString As String
    Dim i As Integer
    Dim randomChar As String
    Dim charSet As String
    
    ' Character set: 0-9 and A-F (hexadecimal characters)
    charSet = "0123456789ABCDEF"
    
    ' Initialize the GUID string
    guidString = ""
    
    ' Generate 32 random hexadecimal characters
    For i = 1 To 32
        ' Get a random number between 1 and 16 (length of charSet)
        randomChar = Mid(charSet, Int(Rnd() * 16) + 1, 1)
        guidString = guidString & randomChar
    Next i
    
    GenerateRandomGUID = guidString
End Function

Function GenerateRandomGUIDWithFormat() As String
    ' Alternative version that generates a GUID in standard format but returns only alphanumeric
    ' This version ensures better randomness distribution
    
    Dim guidString As String
    Dim i As Integer
    Dim randomByte As Integer
    Dim hexChar As String
    
    ' Initialize random number generator
    Randomize
    
    guidString = ""
    
    ' Generate 16 bytes (128 bits) and convert to hex
    For i = 1 To 16
        randomByte = Int(Rnd() * 256) ' Random number 0-255
        hexChar = Right("0" & Hex(randomByte), 2) ' Convert to 2-digit hex
        guidString = guidString & hexChar
    Next i
    
    GenerateRandomGUIDWithFormat = UCase(guidString)
End Function

Function GenerateRandomAlphanumericGUID() As String
    ' Generates a random GUID-like string using both letters and numbers (not just hex)
    ' Returns a 32-character string with 0-9, A-Z
    
    Dim guidString As String
    Dim i As Integer
    Dim randomChar As String
    Dim charSet As String
    
    ' Character set: 0-9 and A-Z (36 characters total)
    charSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    ' Initialize random number generator
    Randomize
    
    ' Initialize the GUID string
    guidString = ""
    
    ' Generate 32 random alphanumeric characters
    For i = 1 To 32
        ' Get a random number between 1 and 36 (length of charSet)
        randomChar = Mid(charSet, Int(Rnd() * 36) + 1, 1)
        guidString = guidString & randomChar
    Next i
    
    GenerateRandomAlphanumericGUID = guidString
End Function

' Usage examples (you can test these in VBA immediate window):
' Debug.Print GenerateRandomGUID()
' Debug.Print GenerateRandomGUIDWithFormat()
' Debug.Print GenerateRandomAlphanumericGUID()