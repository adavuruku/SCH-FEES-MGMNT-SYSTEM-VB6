VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3180
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   5820
   HelpContextID   =   1570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1878.849
   ScaleMode       =   0  'User
   ScaleWidth      =   5464.666
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1215
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1470
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   1140
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   3285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter User name and Password to have access to the system"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Password:"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1485
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&User Name:"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

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
End If
If rs!create_session = "1" Then
MDIForm11.mnusession.Enabled = True
End If
If rs!exam_reg = "1" Then
MDIForm11.mnujcese.Enabled = True
MDIForm11.mnud.Enabled = True
MDIForm11.MNENTRANCE.Enabled = True
MDIForm11.finalexamreg.Enabled = True
End If
If rs!instalpay = "1" Then
MDIForm11.instalpayS.Enabled = True
MDIForm11.mnureport.Enabled = False
MDIForm11.Label1.Enabled = True
End If
If rs!editfee = "1" Then
MDIForm11.mnueditfee.Enabled = True
End If

If rs!EditAccount = "1" Then
MDIForm11.mnueua.Enabled = True
End If

If rs!seconterm = "1" Then
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
MDIForm11.mnthird = True

If rs!EditAccount = "1" Then
MDIForm11.Label1.Enabled = True
End If

If rs!report = "1" Then
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
Form7.Show
End Sub


Private Sub Command1_Click()
Form8.Show
End Sub
