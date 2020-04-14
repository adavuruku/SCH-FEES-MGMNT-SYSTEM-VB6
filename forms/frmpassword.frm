VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "CREATE USER ACCOUNT"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   HelpContextID   =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame s 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit User Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   6495
      Begin VB.CheckBox exam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exam Registration"
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
         TabIndex        =   20
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CheckBox create 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create New Session"
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
         Left            =   3600
         TabIndex        =   19
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox second 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Second term"
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
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox install 
         BackColor       =   &H00FFFFFF&
         Caption         =   "InstallPay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox first 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox edit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit fee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox third 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Third Term"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox report 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
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
         TabIndex        =   12
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox EditUserAcc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit User Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   7320
      Width           =   6495
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000003&
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   2640
         Picture         =   "frmpassword.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         Picture         =   "frmpassword.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Picture         =   "frmpassword.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Users Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.TextBox conpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox pass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox names 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Confirm Password:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User name:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE USER ACCOUNT/PRIVILEGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)
Unload Me
'MDIForm11.Show
End Sub

Private Sub Command1_Click()

'Dim db As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [useraccount]", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!password = conpass
rs!UserName = names
rs!editfee = edit
rs!firsterm = first
rs!seconterm = second
rs!Thirdterm = third
rs!instalpay = install
rs!password = pass.Text
rs!report = report.Value
rs!EditAccount = EditUserAcc
rs!create_session = create
rs!exam_reg = exam
rs.Update
rs.Close
Set rs = Nothing
'db.Close
'Set db = Nothing
MsgBox "password created and privilleges assigned", vbInformation
names.Text = ""
pass.Text = ""
conpass.Text = ""
exam.Value = 0
edit.Value = 0
first.Value = 0
install.Value = 0
create.Value = 0
second.Value = 0
third.Value = 0
'debt.Value = 0
EditUserAcc.Value = 0
report.Value = 0
Unload Me
End Sub



Private Sub Command2_Click()
'Cmbuname.Text = ""
pass.Text = ""
conpass.Text = ""
edit.Value = 0
first.Value = 0
install.Value = 0
second.Value = 0
third.Value = 0
debt.Value = 0
EditUserAcc.Value = 0
exam.Value = 0
End Sub

Private Sub Command3_Click()

End Sub

Private Sub conpass_LostFocus()
If pass.Text = conpass.Text Then
Exit Sub
Else
MsgBox "password not correct", vbInformation
conpass.SetFocus
End If
End Sub



Private Sub names_LostFocus()
'Dim db As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [useraccount] where username='" & names.Text & "'", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
MsgBox "User name exist ", vbInformation
names.Text = ""
names.SetFocus
Else
Exit Sub
End If

End Sub
