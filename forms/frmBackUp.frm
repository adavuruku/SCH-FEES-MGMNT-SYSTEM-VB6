VERSION 5.00
Begin VB.Form frmDBBackUp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BackUp Database"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ClipControls    =   0   'False
   HelpContextID   =   2330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5910
      Top             =   165
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1590
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFC0FF&
      Height          =   2280
      Left            =   0
      ScaleHeight     =   2220
      ScaleWidth      =   6390
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      Begin VB.CommandButton cmdDestination 
         Caption         =   "..."
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   1185
         Width           =   375
      End
      Begin VB.TextBox txtDestination 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   3975
      End
      Begin VB.PictureBox progStat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1950
         ScaleHeight     =   285
         ScaleWidth      =   4365
         TabIndex        =   4
         Top             =   1755
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   2115
         Left            =   30
         Picture         =   "frmBackUp.frx":0000
         Top             =   45
         Width           =   1905
      End
      Begin VB.Label lblInform 
         BackStyle       =   0  'Transparent
         Caption         =   "Creating backup......."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   1965
         TabIndex        =   10
         Top             =   1545
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         Caption         =   "(0%)..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3855
         TabIndex        =   9
         Top             =   1515
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Destination"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   6
         Top             =   975
         Width           =   2055
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   1965
         TabIndex        =   1
         Top             =   45
         Width           =   2445
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current size of Database"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2865
      Width           =   1740
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1950
      TabIndex        =   7
      Top             =   2820
      Width           =   2280
   End
End
Attribute VB_Name = "frmDBBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mintCount As Integer, mintPause As Integer
Dim dbasize As Long
Dim PathName As String
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub Form_Activate()
lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
End Sub
Private Sub cmdDestination_Click()
On Error Resume Next

Dim strTemp As String
    strTemp = fBrowseForFolder(Me.hwnd, "Select backup path")
    If strTemp <> "" Then
    txtDestination = strTemp
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next

con
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [sessions]", db, adOpenDynamic, adLockOptimistic
rs.MoveLast
mdbname1 = rs!dbname & ".mdb"
'createsession.Show
'createsession.Text1 = mdbname1    '& ".mdb"
'RS.Close
'Set RS = Nothing
'App.Path & "\database\" & mdbname1 & "

PathName = App.Path & "\database\" & mdbname1 & ""
dbasize = FileLen(PathName)
End Sub





Private Sub Timer1_Timer()
On Error Resume Next

If txtDestination <> "" Then
    DoBackup PathName, txtDestination
    lblCount.Visible = True
    lblInform.Visible = True
'    lblCBK.Visible = True
    progStat.Visible = True
    progStat.Value = progStat.Value + 2
    CountMe
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
If progStat.Value = 100 Then
    MsgBox "Backup database complete", vbInformation
'    Timer1.Interval = 0
'    progStat.Value = 0
'    lblCBK.Visible = False
'    progStat.Visible = False
    Unload Me
Else
    If txtDestination = "" Then
     progStat.Value = 0
     
       'Your function, can be anything. Open another form, frmMain.show... Ect.
    End If
    End If
    End If
End Sub
Private Sub CountMe()
On Error Resume Next

   mintPause = mintPause + 1
   
    If mintCount < 0 Then
        mintCount = mintCount + 1
        lblCount.Caption = "(" & mintCount & "%)..."
 '        frmSplash.Refresh
         
    ElseIf mintCount < 100 Then
        mintCount = mintCount + 2
        lblCount.Caption = "(" & mintCount & "%)..."
'        frmSplash.Refresh
        
    End If
    
    If mintPause = 100 Then
        lblCount.Caption = "App..."
        lblInform.Caption = "Starting"
    ElseIf mintPause > 180 Then

        Unload Me

   End If
End Sub
