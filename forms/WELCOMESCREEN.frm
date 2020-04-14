VERSION 5.00
Begin VB.Form WELCOMESCREEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   5520
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1215
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "admin"
         Top             =   870
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1365
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1200
         TabIndex        =   11
         Text            =   "admin"
         Top             =   480
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Password:"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&User Name:"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Top             =   1200
         Width           =   720
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Computer Science, School of Applied Sciences"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4080
      TabIndex        =   9
      Top             =   7200
      Width           =   8595
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In partial fulfilment for the award of Higher National Diploma in Computer Scince"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2520
      TabIndex        =   8
      Top             =   6840
      Width           =   11520
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Polytechnic Nasarawa, Nasarawa State"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4920
      TabIndex        =   7
      Top             =   7560
      Width           =   6675
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MR. N.I. AKOSU"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   5640
      TabIndex        =   6
      Top             =   5520
      Width           =   3825
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised By:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Width           =   3165
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FPN/SO4/2007/HCOM/058"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   4635
      TabIndex        =   4
      Top             =   3720
      Width           =   5925
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AGBAKAGBA JOHN EMIAKPO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   4395
      TabIndex        =   3
      Top             =   3240
      Width           =   6465
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed by:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   6015
      TabIndex        =   2
      Top             =   2280
      Width           =   2745
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AND MANAGEMENT SYSTEM FOR UPLIFT PRIMARY SCHOOL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   900
      TabIndex        =   1
      Top             =   1080
      Width           =   13605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESIGN AND IMPLEMENTATION OF SCHOOL FEES COLLECTION"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   14655
   End
   Begin VB.Image Image1 
      Height          =   11055
      Left            =   120
      Picture         =   "WELCOMESCREEN.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "WELCOMESCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Frame1.Visible = False
End Sub

Private Sub cmdOK_Click()
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

End If

If rs!Thirdterm = "1" Then
MDIForm11.mnthird = True
If rs!EditAccount = "1" Then
MDIForm11.mnueua.Enabled = True
MDIForm11.Label1.Enabled = True
End If
If rs!report = "1" Then
MDIForm11.mnureport = True
End If
End If
End If
End If
MDIForm11.Show
Unload Me

End Sub

Private Sub Command1_Click()
Frame1.Visible = True
End Sub
