VERSION 5.00
Begin VB.Form login2 
   Caption         =   "Validate User"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   3975
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1365
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1365
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1215
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   870
         Width           =   2325
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&User Name:"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Password:"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   885
         Width           =   1080
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Admin Login Click Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Existing User Enter Username and Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "login2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next

Dim UserName As String
Dim password As String
UserName = txtUserName.Text
password = txtPassword.Text
m1 = "[UserName]='" + UserName + "'"
m2 = "[password]='" + password + "'"
m3 = m1 & "AND" & m2
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [useraccount]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "wrong"
Exit Sub
Else
If rs!firsterm = "1" Then
MDIForm11.mnFirst.Enabled = True
MDIForm11.we.Enabled = True
End If
If rs!create_session = "1" Then
MDIForm11.mnusession.Enabled = True
End If
If rs!exam_reg = "1" Then
MDIForm11.Image2.Enabled = True

MDIForm11.mnujcese.Enabled = True
MDIForm11.mnud.Enabled = True
MDIForm11.MNENTRANCE.Enabled = True
MDIForm11.finalexamreg.Enabled = True
End If
If rs!instalpay = "1" Then
MDIForm11.we.Enabled = True

MDIForm11.instalpayS.Enabled = True
MDIForm11.mnureport.Enabled = False
MDIForm11.Label1.Enabled = True
End If

If rs!editfee = "1" Then
MDIForm11.mnueditfee.Enabled = True
MDIForm11.mnueditfee.Enabled = True
End If

If rs!EditAccount = "1" Then
MDIForm11.mnueua.Enabled = True
End If

If rs!seconterm = "1" Then
MDIForm11.we.Enabled = True

MDIForm11.mnsecond.Enabled = True
MDIForm11.Label1.Enabled = True

If rs!EditAccount = "1" Then
MDIForm11.mnueua.Enabled = True
MDIForm11.Label1.Enabled = True
MDIForm11.Image2.Enabled = True
MDIForm11.Image5.Enabled = True
'MDIForm11.finalexamreg.Enabled = True
MDIForm11.we.Enabled = True
'MDIForm11finalexams.Enabled = True
End If

If rs!Thirdterm = "1" Then
MDIForm11.we.Enabled = True

MDIForm11.mnthird = True

If rs!EditAccount = "1" Then
MDIForm11.Label1.Enabled = True
MDIForm11.gh.Enabled = True
End If

If rs!report = "1" Then
'mnureport
MDIForm11.mnureport.Enabled = True
MDIForm11.mnucommonentrance.Enabled = True
MDIForm11.mnuneco.Enabled = True
MDIForm11.mnuwaec.Enabled = True
MDIForm11.mnujcese.Enabled = True
MDIForm11.mnumock.Enabled = True
End If

End If
End If
End If

Unload Me
'Form7.Show

End Sub

Private Sub Label2_Click()
Unload Me
frmLogin.Show
End Sub
