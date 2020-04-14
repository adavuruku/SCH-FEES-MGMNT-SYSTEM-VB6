VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "&EDIT USER ACCOUNT"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   HelpContextID   =   1150
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   11
      Top             =   6840
      Width           =   6495
      Begin VB.CommandButton cmdButton1 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   2520
         Picture         =   "frmupdate.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   480
         Picture         =   "frmupdate.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
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
         Picture         =   "frmupdate.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Change password"
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Cmbuname 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox pass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox conpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
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
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Old Password:"
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
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Password:"
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
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame s 
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
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   6495
      Begin VB.CheckBox create 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create session"
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
         Left            =   3960
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox exam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exam registration"
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
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   1935
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
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Width           =   2055
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
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   1695
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
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   4
         Top             =   840
         Width           =   1215
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1335
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
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
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
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT USER PRIVILEGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cmbuname_LostFocus()
On Error Resume Next

If Cmbuname.Text = "" Then
Exit Sub
End If
'Dim db As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
 If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [useraccount] where username='" & Cmbuname.Text & "'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "select appropriate user Name", vbCritical
Else
pass = rs!password
edit = rs!editfee
first = rs!firsterm
second = rs!seconterm
install = rs!instalpay
EditUserAcc = rs!EditAccount
report = rs!report
third = rs!Thirdterm
exam = rs!exam_reg
create = rs!create_session
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next

If Cmbuname.Text = "" Then
Exit Sub
End If
'Dim db As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [useraccount] where username='" & Cmbuname.Text & "'", cn, adOpenDynamic, adLockOptimistic
'password
rs!password = conpass.Text
rs!UserName = Cmbuname
rs!editfee = edit
rs!firsterm = first
rs!seconterm = second
rs!Thirdterm = third
rs!instalpay = install
'rs!password = pass.Text
rs!report = report
rs!EditAccount = EditUserAcc
rs!create_session = create
rs!exam_reg = exam
rs.Update
Set rs = Nothing
'db.Close
'Set db = Nothing
MsgBox "User Account Updated successfully", vbInformation
Cmbuname.Text = ""
exam.Value = 0
create.Value = 0
pass.Text = ""
conpass.Text = ""
report.Value = 0
first.Value = 0
install.Value = 0
second.Value = 0
third.Value = 0
'debt.Value = 0
EditUserAcc.Value = 0
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdButton1_Click(Index As Integer)
Unload Me
'MDIForm11.Show
End Sub

Private Sub Command2_Click()
Cmbuname.Text = ""
pass.Text = ""
conpass.Text = ""
report.Value = 0
first.Value = 0
install.Value = 0
second.Value = 0
EditUserAcc.Value = 0
create.Value = 0
exam.Value = 0
edit.Value = 0
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
On Error Resume Next

Label2.Visible = True
Label3.Visible = True
pass.Visible = True
conpass.Visible = True
Command4.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next

 If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [useraccount]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
Cmbuname.AddItem rs!UserName
rs.MoveNext
Loop
rs.Close
Set rs = Nothing

End Sub

