Attribute VB_Name = "BackupRestoreFunc"
Option Explicit

Public Const MAX_PATH = 260

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONUP = &H202

Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_BROWSEINCLUDEFILES = &H4000
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_STATUSTEXT = &H4

Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_FILESONLY = &H80

Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public BackupFolderName As String
Public SourceFolder As String
Public Function fBrowseForFolder(hWndOwner As Long, sPrompt As String) As String
Dim iNull    As Integer
Dim lpIDList As Long
Dim lresult  As Long
Dim sPath    As String
Dim udtBI    As BrowseInfo

With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
End With

lpIDList = SHBrowseForFolder(udtBI)

If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lresult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    
    iNull = InStr(sPath, vbNullChar)
    If iNull Then sPath = Left$(sPath, iNull - 1)
End If

fBrowseForFolder = sPath

End Function

Public Sub DoBackup(strSourcePath As String, strDestinationPath As String)
On Error Resume Next
Dim lFileOp  As Long
Dim lresult  As Long
Dim lFlags   As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim strSourceDir As String
Dim strDestinationDir As String

    Screen.MousePointer = vbHourglass
    BackupFolderName = strDestinationPath
    MkDir BackupFolderName & "\Backup - " & Format(Date, "yyyy.mm.dd")
    lFileOp = FO_COPY
    
            lFlags = lFlags And Not FOF_SILENT
            lFlags = lFlags Or FOF_NOCONFIRMATION
            lFlags = lFlags Or FOF_NOCONFIRMMKDIR
            lFlags = lFlags Or FOF_FILESONLY
            
        With SHFileOp
            .wFunc = lFileOp
            .pFrom = strSourcePath & vbNullChar
            .pTo = strDestinationPath & "\Backup - " & Format(Date, "yyyy.mm.dd") & vbNullChar
            .fFlags = lFlags
        End With
        
            lresult = SHFileOperation(SHFileOp)
              Screen.MousePointer = vbDefault
End Sub
