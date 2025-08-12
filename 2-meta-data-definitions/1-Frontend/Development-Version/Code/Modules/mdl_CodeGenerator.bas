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

    Dim hash As String
    Dim i As Integer
   
    hash = CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
    hash = hash & CStr(Rnd() * CDec(Format(Now(), "yyyymmddhhmmss")))
    
    get_random_code = LCase(Mid(encrypt(hash), 1, ip_length))

End Function