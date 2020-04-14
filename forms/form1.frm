VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "CREATE USER ACCOUNT"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   HelpContextID   =   2130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6915
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
      Height          =   3495
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   6495
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
         TabIndex        =   17
         Top             =   1800
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
         Left            =   240
         TabIndex        =   16
         Top             =   1320
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
         Left            =   240
         TabIndex        =   15
         Top             =   840
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
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
         Left            =   240
         TabIndex        =   13
         Top             =   2520
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
         Left            =   240
         TabIndex        =   12
         Top             =   2880
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
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   6495
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000009&
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
         Picture         =   "form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
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
         Picture         =   "form1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
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
         Picture         =   "form1.frx":074C
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
      Top             =   120
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
 If RS.State = adStateOpen Then RS.Close
RS.Open "select * from [useraccount]", cn, adOpenDynamic, adLockOptimistic
RS.AddNew
RS!password = conpass
RS!UserName = names
RS!editfee = edit
RS!firsterm = First
RS!seconterm = Second
RS!Thirdterm = third
RS!instalpay = install
RS!password = pass.Text
RS!report = report
RS!EditAccount = EditUserAcc

RS.Update
RS.Close
Set RS = Nothing
'db.Close
'Set db = Nothing
MsgBox "password created and privilleges assigned", vbInformation
names.Text = ""
pass.Text = ""
conpass.Text = ""
edit.Value = 0
First.Value = 0
install.Value = 0
Second.Value = 0
third.Value = 0
'debt.Value = 0
EditUserAcc.Value = 0
report.Value = o

End Sub



Private Sub Command2_Click()
'Cmbuname.Text = ""
pass.Text = ""
conpass.Text = ""
edit.Value = 0
First.Value = 0
install.Value = 0
Second.Value = 0
third.Value = 0
debt.Value = 0
EditUserAcc.Value = 0
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
 If RS.State = adStateOpen Then RS.Close
RS.Open "select * from [useraccount] where username='" & names.Text & "'", cn, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
MsgBox "User name exist ", vbInformation
names.Text = ""
names.SetFocus
Else
Exit Sub
End If

End Sub
