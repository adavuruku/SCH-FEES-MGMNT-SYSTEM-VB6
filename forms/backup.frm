VERSION 5.00
Begin VB.Form FrmBackupDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "backup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   450
      Left            =   5565
      TabIndex        =   4
      Top             =   1785
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Backup &Now"
      Height          =   450
      Left            =   3525
      TabIndex        =   3
      Top             =   1785
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   330
      Left            =   6675
      TabIndex        =   2
      Top             =   450
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   465
      Width           =   6390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "#"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Backup Folder :"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   195
      Width           =   2175
   End
End
Attribute VB_Name = "FrmBackupDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Mr. Atanu Maity
'          Date : 21-Aug-2006
'*************************************
'     Backup the Database (data.mdb)
'      Used Table : NA
'Module to take a copy of data.mdb in
'diffrent location with timestramp
'*************************************

Option Explicit
'>>> Declare File System variable
Dim Fs As New FileSystemObject
Dim NewFile As String

Private Sub Command1_Click()
    '>>> open folder browser dialog
    '>>> select the folder path
    '>>> store the path in text box text1
    Dim s As String
    s = BrowseFolders(hWnd, "Select Folder for Creating Backup file ... ", BrowseForEverything, CSIDL_DESKTOP)
    If s = "" Then
        MsgBox "Select Valid Folder for Creating Dump File.", vbInformation, "Creating Dump"
        Command1.SetFocus
        Exit Sub
    Else
        If Fs.FolderExists(s) = False Then
            MsgBox "Invalid Folder,Select Valid Folder. ", vbInformation, "Creating Dump"
            Command1.SetFocus
            Exit Sub
        End If
        Text1.Text = s
    End If

End Sub

Private Sub Command2_Click()
    On Error GoTo myer1
    '>>> check the selected folder wheather
    '>>> it is exist or not
    If Fs.FolderExists(Text1) = False Then
        MsgBox "Invalid Folder,Select Valid Folder. ", vbInformation, "Creating Dump"
        Command1.SetFocus
        Exit Sub
    End If
    
    '>>> save the settings in registry
    SaveSetting "BILLING_SOFTWARE", "BACKUP_DATABASE", "BACKUP_PATH", Text1
    
    '>>> copy the database file in selected folder for backup
    Fs.CopyFile App.Path & "\data.mdb", Text1.Text & "\" & NewFile, False
    MsgBox "Backup Process Complete.", vbInformation
    Exit Sub
myer1:
    '>> check the folder if the file alreday exist warn for overwrite
    If Err.Number = 58 Then
        If MsgBox("File Allready exist in same name , do you want to overwrite the existing file ..", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
            Fs.CopyFile App.Path & "\data.mdb", Text1.Text & "\" & NewFile, True
            MsgBox "Backup Process Complete", vbInformation
        End If
    Else
        MsgBox "Can not complete backup following error occured : " & Err.Description, vbCritical
    End If
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    '>>> new backup file name like data_040108_1745.mdb
    NewFile = "Data_" & Format(Now, "ddnnyy_hhnn") & ".mdb"
    Label2.Caption = "Backup Database Name : " & NewFile
    
    '>>> load the last saved settings from registry
    Text1.Text = GetSetting("BILLING_SOFTWARE", "BACKUP_DATABASE", "BACKUP_PATH", "")
End Sub
