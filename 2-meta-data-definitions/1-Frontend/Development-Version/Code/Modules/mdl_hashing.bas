Attribute VB_Name = "mdl_hashing"
Option Compare Database
Option Explicit

' This module is based on Code originally published by Nouba on www.office-loesung.de
' Revised for 64bit and extended for SHA512 suppport and simpler extensibility by Philipp Stiefel (codekabinett.com/en)
' Comments/explanations in this module are all by Philipp Stiefel (codekabinett.com/en)

' Scroll down to the HashAlgorithmType Enum and the GetProviderInfo function for hints on extending
' this module with different hash algorithms!


Private Const CRYPT_VERIFYCONTEXT     As Long = &HF0000000
'-- for CryptGetHashParam
Private Const HP_HASHVAL              As Long = 2
Private Const HP_HASHSIZE             As Long = 4

#If VBA7 Then
   Private Declare PtrSafe Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
            ByRef phProv As LongPtr, ByVal pszContainer As LongPtr, _
            ByVal pszProvider As LongPtr, ByVal dwProvType As Long, _
            ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptReleaseContext Lib "Advapi32" ( _
              ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptCreateHash Lib "Advapi32" ( _
            ByVal hProv As LongPtr, ByVal AlgId As Long, _
            ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
  Private Declare PtrSafe Function CryptHashData Lib "Advapi32" ( _
          ByVal hHash As LongPtr, ByRef pbData As Any, ByVal dwDataLen As Long, _
          ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptDestroyHash Lib "Advapi32" ( _
          ByVal hHash As LongPtr) As Long
  Private Declare PtrSafe Function CryptGetHashParam Lib "Advapi32" ( _
          ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, _
          ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
#Else
  Private Declare Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
            ByRef phProv As Long, ByVal pszContainer As Long, _
            ByVal pszProvider As Long, ByVal dwProvType As Long, _
            ByVal dwFlags As Long) As Long
  Private Declare Function CryptReleaseContext Lib "Advapi32" ( _
          ByVal hProv As Long, ByVal dwFlags As Long) As Long
  Private Declare Function CryptCreateHash Lib "Advapi32" ( _
          ByVal hProv As Long, ByVal AlgId As Long, _
          ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
  Private Declare Function CryptHashData Lib "Advapi32" ( _
          ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, _
          ByVal dwFlags As Long) As Long
  Private Declare Function CryptDestroyHash Lib "Advapi32" ( _
          ByVal hHash As Long) As Long
  Private Declare Function CryptGetHashParam Lib "Advapi32" ( _
          ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, _
          ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
#End If

Public Enum HashAlgorithmType
    ' This enum is an *incomplete* list of algorithm identifiers (ALG_ID).
    ' See: https://learn.microsoft.com/en-us/windows/win32/seccrypto/alg-id
    ' If you want to use a different algorithm, you must add the ALG_ID here and modify the
    ' GetProviderInfo function in this module to return a suitable provider name and type.
    CALG_MD5 = &H8003&
    CALG_SHA1 = &H8004&
    CALG_SHA_512 = &H800E&
End Enum

Private Type ProviderInfo
    ProviderName As String
    ProviderType As Long
End Type

Public Function GetHashOfString(StringData As String, ByVal HashAlgorithm As HashAlgorithmType) As String
    Dim tempArray() As Byte
    tempArray = StrConv(StringData, vbFromUnicode)
    GetHashOfString = GetHashOfByteArray(tempArray, HashAlgorithm)
End Function

Public Function GetHashOfByteArray(ByteData() As Byte, ByVal HashAlgorithm As HashAlgorithmType) As String

  Dim hBaseProvider As LongPtr
  Dim hHash    As LongPtr
  Dim lSize    As Long
  Dim baBuffer() As Byte
  Dim lIdx     As Long
  
  Dim provInfo As ProviderInfo
  provInfo = GetProviderInfo(HashAlgorithm)
  
  If CryptAcquireContext(hBaseProvider, 0, StrPtr(provInfo.ProviderName), _
                         provInfo.ProviderType, CRYPT_VERIFYCONTEXT) <> 0 Then
    If CryptCreateHash(hBaseProvider, HashAlgorithm, 0, 0, hHash) <> 0 Then
      If CryptHashData(hHash, ByteData(0), UBound(ByteData) + 1, 0) <> 0 Then
        If CryptGetHashParam(hHash, HP_HASHSIZE, lSize, 4, 0) <> 0 Then
          ReDim baBuffer(0 To lSize - 1) As Byte
          If CryptGetHashParam(hHash, HP_HASHVAL, baBuffer(0), lSize, 0) <> 0 Then
            For lIdx = 0 To UBound(baBuffer)
              GetHashOfByteArray = GetHashOfByteArray & Right("0" & Hex(baBuffer(lIdx)), 2)
            Next lIdx
          End If
        End If
      End If
      Call CryptDestroyHash(hHash)
    End If
    Call CryptReleaseContext(hBaseProvider, 0)
  End If
  
End Function

Private Function GetProviderInfo(ByVal HashAlgorithm As HashAlgorithmType) As ProviderInfo
    
    ' **************************************************************************************************
    ' This function is meant for illustration. You can/should modify the code to suit your requirements!
    '   (Depending on your requirements the values provided hardcoded in this function should rather be
    '   retrieved from a configuration table/file or from user selection.)
    ' **************************************************************************************************
    
    ' The functions returns *one* of potentially *multiple* valid combinations of ProviderName and ProviderType.
    '   E.g. the MS_ENHANCED_PROVIDER with RSA_AES provider type can also compute MD5 and SHA1 hashes.
    '   E.g. the "Microsoft Strong Cryptographic Provider" also uses the provider type PROV_RSA_FULL
    
    ' If you modify this function, *you* must make sure the combination of ProviderName and ProviderType is
    ' valid and suitable for the hash algorithm you want to use. See the documentation linked below to
    ' find/verify the combination you need to use.
    
    ' List of provider names and their supported algorithms: https://learn.microsoft.com/en-us/windows/win32/seccertenroll/cryptoapi-cryptographic-service-providers
    ' List of provider types supplied with Windows: https://learn.microsoft.com/en-us/windows/win32/seccrypto/cryptographic-provider-types
    ' The numeric values for the provider types supplied with Windows can be found in the wincrypt.h file in the Windows SDK
    
    Const MS_BASE_PROVIDER      As String = "Microsoft Base Cryptographic Provider v1.0"
    Const MS_ENHANCED_PROVIDER  As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"

    Const PROV_RSA_FULL As Long = 1
    Const PROV_RSA_AES  As Long = 24

    Dim provInfo As ProviderInfo

    Select Case HashAlgorithm
        Case HashAlgorithmType.CALG_MD5, HashAlgorithmType.CALG_SHA1
            provInfo.ProviderName = MS_BASE_PROVIDER
            provInfo.ProviderType = PROV_RSA_FULL
        Case HashAlgorithmType.CALG_SHA_512
            provInfo.ProviderName = MS_ENHANCED_PROVIDER
            provInfo.ProviderType = PROV_RSA_AES
        Case Else
            Err.Raise vbObjectError Or 1, "GetProviderInfo", "HashAlgorithm not implemented"
    End Select
    
    GetProviderInfo = provInfo
    
End Function

Public Function CreateMD5(ip_text_to_hash As String) As String
    CreateMD5 = LCase(GetHashOfString(ip_text_to_hash, CALG_MD5))
End Function

Public Sub TestHash()
  
    ' This procedure is just meant to test/demonstrate the two VBA functions wrapping the hash computation
  
    Const dataString As String = "Sample text as hash input"
  
    Dim bytes() As Byte
    bytes = StrConv(dataString, vbFromUnicode)
    
    Debug.Print "-- Hashes computed from bytes --"
    Debug.Print "MD5:  " & vbTab & vbTab & GetHashOfByteArray(bytes, CALG_MD5)  '-- ==> MD5-Digest
    Debug.Print "SHA-1: " & vbTab & vbTab & GetHashOfByteArray(bytes, CALG_SHA1) '-- ==> SHA1-Digest
    Debug.Print "SHA-512: " & vbTab & GetHashOfByteArray(bytes, CALG_SHA_512) '-- ==> SHA1-Digest
    
    Debug.Print vbCrLf & "-- Hashes computed from string data --"
    Debug.Print "MD5: " & vbTab & vbTab & GetHashOfString(dataString, CALG_MD5)   '-- ==> MD5-Digest (from String)
    Debug.Print "SHA-1: " & vbTab & vbTab & GetHashOfString(dataString, CALG_SHA1) '-- ==> SHA1-Digest (from String)
    Debug.Print "SHA-512: " & vbTab & GetHashOfString(dataString, CALG_SHA_512) '-- ==> SHA512  (from String)"

End Sub