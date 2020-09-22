Attribute VB_Name = "modDeclarations"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Type SHFILEOPSTRUCT
    hwnd                                   As Long
    wFunc                                  As Long
    pFrom                                  As String
    pTo                                    As String
    fFlags                                 As Integer
    fAnyOperationsAborted                  As Boolean
    hNameMappings                          As Long
    lpszProgressTitle                      As String
End Type
Public Const FO_DELETE                   As Long = &H3
Public Type IconType
    cbSize                                 As Long
    picType                                As PictureTypeConstants
    hIcon                                  As Long
End Type
Public Type CLSIdType
    Id(16)                                 As Byte
End Type
Public Type ShellFileInfoType
    hIcon                                  As Long
    iIcon                                  As Long
    dwAttributes                           As Long
    szDisplayName                          As String * 260
    szTypeName                             As String * 80
End Type
Public Const Large                       As Long = &H100
Public Const VER_PLATFORM_WIN32_NT       As Integer = 2
Public Type OSVERSIONINFO
    dwOSVersionInfoSize                    As Long
    dwMajorVersion                         As Long
    dwMinorVersion                         As Long
    dwBuildNumber                          As Long
    dwPlatformId                           As Long
    szCSDVersion                           As String * 128
End Type
Private Type TypeSignature
    SignatureFilename                      As String
    SignatureDate                          As String
    SignatureOnlineFilename                As String
    SignatureCount                         As Integer
End Type
Public Enum RM
    Normal = 0
    TrayOnly = 1
    ScanFile = 3
End Enum
#If False Then
Private Normal, TrayOnly, ScanFile
#End If
Private Type AntiVirus
    AVname                                 As String
    Runmode                                As RM
    Signature                              As TypeSignature
End Type
Public AV                                As AntiVirus
Private Type SHItemID
    cb                                     As Long
    abID                                   As Byte
End Type
Public Type ItemIDList
    mkid                                   As SHItemID
End Type
Public Type BROWSEINFO
    hOwner                                 As Long
    pidlRoot                               As Long
    pszDisplayName                         As String
    lpszTitle                              As String
    ulFlags                                As Long
    lpCallbackProc                         As Long
    lParam                                 As Long
    iImage                                 As Long
End Type
#If Win16 Then
Public Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, _
                                            ByVal hWndInsertAfter As Integer, _
                                            ByVal X As Integer, _
                                            ByVal Y As Integer, _
                                            ByVal cx As Integer, _
                                            ByVal cy As Integer, _
                                            ByVal wFlags As Integer)
#Else
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long
#End If
Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, _
                                                                     riid As CLSIdType, _
                                                                     ByVal fown As Long, _
                                                                     lpUnk As Object) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                ByVal dwFileAttributes As Long, _
                                                                                psfi As ShellFileInfoType, _
                                                                                ByVal cbFileInfo As Long, _
                                                                                ByVal uFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As Long, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As Long, _
                                                                          ByVal lpstrDefExt As Long, _
                                                                          ByVal lpstrFilter As Long, _
                                                                          ByVal lpstrTitle As Long) As Long
Public Declare Function GetFileNameFromBrowseA Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As String, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As String, _
                                                                          ByVal lpstrDefExt As String, _
                                                                          ByVal lpstrFilter As String, _
                                                                          ByVal lpstrTitle As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
