Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' (c) Mehmet (P.R.M.) Misset
'
' Encription
'
'
' Errors:
' n/a
'
' @class       Credentials
' @description Class provides various properties / functions storing credentials
'              for local usage, these are stored in "enctyped" files. hereby the
'              filename and content is scrambaled utilizing a low grade
'              encyption, see "Encritpion"-module for working.
' @author      mehmet.misset@misset-data-analytics.nl
' @license     MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Compare Database
Option Explicit
'
'
' -----------------------------------------------------------------------------
' Private Constants
' -----------------------------------------------------------------------------
Private Const p_username   As String = "1"
Private Const p_password   As String = "2"
Private Const p_remember   As String = "3"
Private Const p_enviromt   As String = "4"
Private Const p_git_folder As String = "GitFolderPath"
Private Const p_afs_folder As String = "AfsFolderPath"
'
'
' -----------------------------------------------------------------------------
' Public Properties: Enviroment
' -----------------------------------------------------------------------------
Public Property Let Enviroment(Value As String):    Call add_credential(p_enviromt, Value): End Property
Public Property Get Enviroment() As String: Enviroment = get_credential(p_enviromt):         End Property
'
'
' -----------------------------------------------------------------------------
' Public Properties: Username
' -----------------------------------------------------------------------------
Public Property Let Username(Value As String):  Call add_credential(p_username, Value): End Property
Public Property Get Username() As String: Username = get_credential(p_username):        End Property
'
'
' -----------------------------------------------------------------------------
' Public Properties: Password
' -----------------------------------------------------------------------------
Public Property Let Password(Value As String):  Call add_credential(p_password, Value): End Property
Public Property Get Password() As String: Password = get_credential(p_password):        End Property
'
'
' -----------------------------------------------------------------------------
' Public Properties: Remember
' -----------------------------------------------------------------------------
Public Property Let Remember(Value As String):  Call add_credential(p_remember, Value): End Property
Public Property Get Remember() As String: Remember = get_credential(p_remember):        End Property
'
'
' -----------------------------------------------------------------------------
' Public Properties: Git Repository Path
' -----------------------------------------------------------------------------
Public Property Let Git_Folder(Value As String):    Call add_credential(p_git_folder, Value): End Property
Public Property Get Git_Folder() As String: Git_Folder = get_credential(p_git_folder):        End Property
'
'
' -----------------------------------------------------------------------------
' Public Properties: Git Repository Path
' -----------------------------------------------------------------------------
Public Property Let Azure_File_Storage_folder(Value As String):    Call add_credential(p_afs_folder, Value): End Property
Public Property Get Azure_File_Storage_folder() As String: Azure_File_Storage_folder = get_credential(p_afs_folder):        End Property
'
'
' -----------------------------------------------------------------------------
' Function    : get_credentials
' Description : Subroutine gets the "credentials"
' -----------------------------------------------------------------------------
Public Function get_credential(ip_nm_attribute As String) As String
    '
    ' Local Vaiables
    Dim fso As New FileSystemObject
    Dim tsm As TextStream
    Dim pth As String
    '
    ' Set Filepath, based on the credentials folder + value (attribute)
    pth = credentials_folder & LCase(encrypt(ip_nm_attribute)) & ".sdf"
    '
    ' Check is File Exists with name of strNameAttribute
    If fso.FileExists(pth) Then
        '
        ' Open File
        Set tsm = fso.OpenTextFile(pth, ForReading, True, TristateFalse)
        '
        ' Read line from fiel and set return value
        get_credential = decrypt(tsm.ReadLine)
        '
        ' Close TextStream
        tsm.Close
        '
    End If
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : add_credentials
' Description : Subroutine adds the "credentials"
' -----------------------------------------------------------------------------
Public Sub add_credential(ip_nm_attribute As String, ip_ds_attribute As String)
    '
    ' Local Vaiables
    Dim fso As New FileSystemObject
    Dim tsm As TextStream
    Dim pth As String
    '
    ' Set Filepath, based on the credentials folder + value (attribute)
    pth = credentials_folder & LCase(encrypt(ip_nm_attribute)) & ".sdf"
    '
    ' Check is File Exists with name of strNameAttribute
    If fso.FileExists(pth) Then
        '
        ' Drop file
        Call fso.DeleteFile(pth)
        '
    End If
    '
    ' Open File
    Set tsm = fso.OpenTextFile(pth, ForWriting, True, TristateFalse)
    '
    ' Read line from fiel and set return value
    Call tsm.WriteLine(encrypt(ip_ds_attribute))
    '
    ' Close TextStream
    tsm.Close
    '
End Sub
'
'
' -----------------------------------------------------------------------------
' Function    : windows_user
' Description : Function returns windows username
' -----------------------------------------------------------------------------
Private Function windows_user() As String
    '
    ' Extract windows Username
    windows_user = Environ("Username")
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : get_windows_user_folder
' Description : Function returns user-folder
' -----------------------------------------------------------------------------
Private Function user_folder() As String
    '
    ' Generate "User"-folder path
    user_folder = mdl_Folders.tx_user_folder()
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : get_windows_user_folder
' Description : Function returns user-folder
' -----------------------------------------------------------------------------
Private Function nm_application() As String
    '
    ' Generate "User"-folder path
    nm_application = Replace(CurrentProject.Name, ".accdb", "")
    '
End Function
'
'
' -----------------------------------------------------------------------------
' Function    : credentials_folder
' Description : Function returns windows username
' -----------------------------------------------------------------------------
Private Function credentials_folder() As String
    
    ' Generate "User"-folder path
    credentials_folder = user_folder() & nm_application & "\.framework_credentials" & "\"
    '
    ' Local Vaiables
    Dim fso As New FileSystemObject
    '
    ' Check in Credential Folder exists
    If Not fso.FolderExists(credentials_folder) Then
        '
        ' Folder does NOT exist, so create it
        Call fso.CreateFolder(credentials_folder)
        '
    End If
    '
End Function