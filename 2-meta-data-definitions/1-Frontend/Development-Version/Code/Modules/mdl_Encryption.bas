''
' (c) Mehmet (P.R.M.) Misset
'
' Encription
'
'
' Errors:
' n/a
'
' @module CodeGenerator
' @description module provides various functions for basic encription of string
'              values, the encryption / decryption only works for ths machine
'              is run on. The code utilizes "unique" identifiers from the
'              computer the code is run on.
' @author mehmet.misset@misset-data-analytics.nl
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' @Dependencies: wshom (Windows Script Host Object Model)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Compare Database
Option Explicit
'
'
' -----------------------------------------------------------------------------
' Function    : decrypt
' Description : Function decrypts, encrypted value.
' -----------------------------------------------------------------------------
Public Function decrypt(ip_encrypted_value As String) As String
    '
    ' Local Vaiables
    Dim ark_data          As Long
    Dim encryption_key    As Long
    Dim decrypted_value   As String
    Dim index_xor_value_1 As Integer
    Dim index_xor_value_2 As Integer
    '
    ' Get Encryption Key
    encryption_key = get_encryption_key
    '
    ' Start De-cryption
    For ark_data = 1 To (Len(ip_encrypted_value) / 2)
        '
        ' The first value to be XOr-ed comes from the data to be encrypted
        index_xor_value_1 = val("&H" & (Mid$(ip_encrypted_value, (2 * ark_data) - 1, 2)))
        '
        ' The second value comes from the code key
        index_xor_value_2 = Asc(Mid$(encryption_key, ((ark_data Mod Len(encryption_key)) + 1), 1))
        '
        ' Add to De-crypted result
        decrypted_value = decrypted_value + Chr(index_xor_value_1 Xor index_xor_value_2)
        '
    Next ark_data
    '
    ' Return De-crypted Result
    decrypt = decrypted_value
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : encrypt
' Description : Function encrypts, un-encrypted character string.
' -----------------------------------------------------------------------------
Public Function encrypt(ip_to_be_encrypted_value As String) As String
    '
    ' Local Vaiables
    Dim ark_data          As Long
    Dim encryption_key    As Long
    Dim encrypted_value   As String
    Dim index_xor_value_1 As Integer
    Dim index_xor_value_2 As Integer
    Dim temp              As Integer
    Dim tempstring        As String
    '
    ' Get Encryption Key
    encryption_key = get_encryption_key
    '
    ' Loop throught all characters of the to be encrypted value
    For ark_data = 1 To Len(ip_to_be_encrypted_value)
        '
        'The first value to be XOr-ed comes from the data to be encrypted
        index_xor_value_1 = Asc(Mid$(ip_to_be_encrypted_value, ark_data, 1))
        '
        'The second value comes from the code key
        index_xor_value_2 = Asc(Mid$(encryption_key, ((ark_data Mod Len(encryption_key)) + 1), 1))
        '
        ' Rememeber tempory result
        temp = (index_xor_value_1 Xor index_xor_value_2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        '
        ' Add to Encrypted result
        encrypted_value = encrypted_value + tempstring
        '
    Next ark_data
    '
    ' Return De-crypted Result
    encrypt = encrypted_value
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : get_encryption_key
' Description : Function return local encryption key, based on local available
'               "unique" identifier.
' -----------------------------------------------------------------------------
Public Function get_encryption_key() As String
    '
    ' Local Variable
    Dim fso As New FileSystemObject
    '
    ' Retrun "EncryptionKey" based on the serial number of the c-drive, any
    ' other value can chosen, also some input value that must be provided.
    get_encryption_key = str(fso.GetDrive("C:\").SerialNumber)
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : encrypt_with_fixed_key
' Description : Function encrypts, un-encrypted character string.
' -----------------------------------------------------------------------------
Public Function encrypt_with_fixed_key(ip_to_be_encrypted_value As String) As String
    '
    ' Local Vaiables
    Dim ark_data          As Long
    Dim encryption_key    As Long
    Dim encrypted_value   As String
    Dim index_xor_value_1 As Integer
    Dim index_xor_value_2 As Integer
    Dim temp              As Integer
    Dim tempstring        As String
    '
    ' Get Encryption Key
    encryption_key = 1234567890
    '
    ' Loop throught all characters of the to be encrypted value
    For ark_data = 1 To Len(ip_to_be_encrypted_value)
        '
        'The first value to be XOr-ed comes from the data to be encrypted
        index_xor_value_1 = Asc(Mid$(ip_to_be_encrypted_value, ark_data, 1))
        '
        'The second value comes from the code key
        index_xor_value_2 = Asc(Mid$(encryption_key, ((ark_data Mod Len(encryption_key)) + 1), 1))
        '
        ' Rememeber tempory result
        temp = (index_xor_value_1 Xor index_xor_value_2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        '
        ' Add to Encrypted result
        encrypted_value = encrypted_value + tempstring
        '
    Next ark_data
    '
    ' Return De-crypted Result
    encrypt_with_fixed_key = encrypted_value
    '
End Function