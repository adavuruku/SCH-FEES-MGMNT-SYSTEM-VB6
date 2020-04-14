VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000003&
   Caption         =   "Change Administrator Acces code"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   4095
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
         Left            =   1440
         Picture         =   "Form7.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
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
         Picture         =   "Form7.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox conpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox pass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox Cmbuname 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Admin"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   30
         TabIndex        =   6
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   4
         Top             =   480
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmbuname_LostFocus()
If Cmbuname.Text = "" Then
Exit Sub
End If
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [useraccount] where username='" & Cmbuname.Text & "'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "select appropriate user Name", vbCritical
Else
'pass = rs!password
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
'rs!password = conpass
If pass.Text = rs!password Then
rs!password = conpass
Else
MsgBox "Enter the Previous Password Correctly"
conpass = ""
Exit Sub
End If
rs!UserName = Cmbuname
'rs!password = pass.Text
rs.Update
Set rs = Nothing
MsgBox "Account Updated successfully", vbInformation
'rs.Close
'Set rs = Nothing
Unload Me
login2.Show
End Sub

Private Sub Command4_Click()
Label2.Visible = True
Label3.Visible = True
pass.Visible = True
conpass.Visible = True
Command4.Visible = False
End Sub

Private Sub cmdButton1_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
'If rs.State = adStateOpen Then rs.Close

'rs.Open "select * from [useraccount]", cn, adOpenDynamic, adLockOptimistic
'rs.MoveFirst
'Do While Not rs.EOF
'Cmbuname.AddItem rs!UserName
'rs.MoveNext
'Loop
'rs.Close
'Set rs = Nothing

End Sub
