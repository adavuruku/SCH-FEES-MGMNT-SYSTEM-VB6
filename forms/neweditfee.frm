VERSION 5.00
Begin VB.Form Neweditfee 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   HelpContextID   =   1050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000003&
      Caption         =   "&Update"
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
      Left            =   2400
      Picture         =   "neweditfee.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   41
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   40
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   39
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   38
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   37
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   36
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   35
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   255
      Left            =   5760
      TabIndex        =   33
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Pfemale6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox Pmale6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox Pfemale5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox Pmale5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox Pfemale4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   28
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox Pmale4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox Pfemale3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Pmale3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Pfemale2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Pmale2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Pfemale1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Pmale1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Nfemale3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Nmale3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Nfemale2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Nmale2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Nfemale1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Nmale1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Text            =   "222222"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H80000003&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   3480
      Picture         =   "neweditfee.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FOR NEW STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   44
      Top             =   600
      Width           =   3015
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   0
      Y1              =   1560
      Y2              =   1200
   End
   Begin VB.Line Line9 
      X1              =   6360
      X2              =   6360
      Y1              =   1560
      Y2              =   1200
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   6360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   6360
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line6 
      BorderStyle     =   4  'Dash-Dot
      X1              =   0
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   1560
      Y2              =   7800
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      TabIndex        =   14
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6360
      Y1              =   1560
      Y2              =   7800
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   1560
      Y2              =   7800
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1440
      Y1              =   1560
      Y2              =   7800
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NURSERY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRIMARY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   1230
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   10
      Top             =   7320
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   6600
      Width           =   165
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   6000
      Width           =   165
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   7
      Top             =   5400
      Width           =   165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   165
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATING OF SCHOOL FEES PRICES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Neweditfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbclass_LostFocus()

End Sub

Private Sub cmdButton_Click(Index As Integer)
Unload Me
MDIForm11.Show

End Sub

Private Sub Command1_Click()
Nfemale1.Enabled = True
Nmale1.Enabled = True
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command11_Click()

If rs.State = adStateOpen Then rs.Close
rs.Open "select *from NewFee_Rating", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF

If rs!CLASS = "nur1" And rs!SEX = "male" Then
  rs!Fees_due = Nmale1.Text
ElseIf rs!CLASS = "nur1" And rs!SEX = "female" Then
 rs!Fees_due = Nfemale1.Text
End If

If rs!CLASS = "nur2" And rs!SEX = "male" Then
rs!Fees_due = Nmale2.Text
ElseIf rs!CLASS = "nur2" And rs!SEX = "female" Then
rs!Fees_due = Nfemale2.Text
End If

If rs!CLASS = "nur3" And rs!SEX = "male" Then
rs!Fees_due = Nmale3.Text
ElseIf rs!CLASS = "nur3" And rs!SEX = "female" Then
rs!Fees_due = Nfemale3.Text
End If

'PRIMARY=======================

If rs!CLASS = "prim1" And rs!SEX = "male" Then
rs!Fees_due = Pmale1.Text
ElseIf rs!CLASS = "prim1" And rs!SEX = "female" Then
rs!Fees_due = Pfemale1.Text
End If

If rs!CLASS = "prim2" And rs!SEX = "male" Then
 rs!Fees_due = Pmale2.Text
ElseIf rs!CLASS = "prim2" And rs!SEX = "female" Then
rs!Fees_due = Pfemale2.Text
End If

If rs!CLASS = "prim3" And rs!SEX = "male" Then
rs!Fees_due = Pmale3.Text
ElseIf rs!CLASS = "prim3" And rs!SEX = "female" Then
rs!Fees_due = Pfemale3.Text
End If

If rs!CLASS = "prim4" And rs!SEX = "male" Then
rs!Fees_due = Pmale4.Text
ElseIf rs!CLASS = "prim4" And rs!SEX = "female" Then
rs!Fees_due = Pfemale4.Text
End If

If rs!CLASS = "prim5" And rs!SEX = "male" Then
rs!Fees_due = Pmale5.Text
ElseIf rs!CLASS = "prim5" And rs!SEX = "female" Then
rs!Fees_due = Pfemale5.Text
End If

If rs!CLASS = "prim6" And rs!SEX = "male" Then
rs!Fees_due = Pmale6.Text
ElseIf rs!CLASS = "prim6" And rs!SEX = "female" Then
rs!Fees_due = Pfemale6.Text
End If

rs.MoveNext
Loop


'rs.Update
MsgBox "Update is Successfull", vbInformation, "Fees Change Completed"
Unload Me

End Sub

Private Sub Command2_Click()
Nfemale2.Enabled = True
Nmale2.Enabled = True
End Sub

Private Sub Command3_Click()
Nfemale3.Enabled = True
Nmale3.Enabled = True
End Sub

Private Sub Command4_Click()
 Pmale1.Enabled = True
 Pfemale1.Enabled = True
End Sub

Private Sub Command5_Click()
Pmale2.Enabled = True
Pfemale2.Enabled = True
End Sub

Private Sub Command6_Click()
Pmale3.Enabled = True
Pfemale3.Enabled = True
End Sub

Private Sub Command7_Click()
Pmale4.Enabled = True
Pfemale4.Enabled = True
End Sub

Private Sub Command8_Click()
Pmale5.Enabled = True
Pfemale5.Enabled = True
End Sub

Private Sub Command9_Click()
Pmale6.Enabled = True
Pfemale6.Enabled = True
End Sub

Private Sub Form_Load()

If rs.State = adStateOpen Then rs.Close
rs.Open "select *from NewFee_Rating", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF

If rs!CLASS = "nur1" And rs!SEX = "male" Then
Nmale1.Text = rs!Fees_due
ElseIf rs!CLASS = "nur1" And rs!SEX = "female" Then
Nfemale1.Text = rs!Fees_due
End If

If rs!CLASS = "nur2" And rs!SEX = "male" Then
Nmale2.Text = rs!Fees_due
ElseIf rs!CLASS = "nur2" And rs!SEX = "female" Then
Nfemale2.Text = rs!Fees_due
End If

If rs!CLASS = "nur3" And rs!SEX = "male" Then
Nmale3.Text = rs!Fees_due
ElseIf rs!CLASS = "nur3" And rs!SEX = "female" Then
Nfemale3.Text = rs!Fees_due
End If

'PRIMARY=======================

If rs!CLASS = "prim1" And rs!SEX = "male" Then
Pmale1.Text = rs!Fees_due
ElseIf rs!CLASS = "prim1" And rs!SEX = "female" Then
Pfemale1.Text = rs!Fees_due
End If

If rs!CLASS = "prim2" And rs!SEX = "male" Then
Pmale2.Text = rs!Fees_due
ElseIf rs!CLASS = "prim2" And rs!SEX = "female" Then
Pfemale2.Text = rs!Fees_due
End If

If rs!CLASS = "prim3" And rs!SEX = "male" Then
Pmale3.Text = rs!Fees_due
ElseIf rs!CLASS = "prim3" And rs!SEX = "female" Then
Pfemale3.Text = rs!Fees_due
End If

If rs!CLASS = "prim4" And rs!SEX = "male" Then
Pmale4.Text = rs!Fees_due
ElseIf rs!CLASS = "prim4" And rs!SEX = "female" Then
Pfemale4.Text = rs!Fees_due
End If

If rs!CLASS = "prim5" And rs!SEX = "male" Then
Pmale5.Text = rs!Fees_due
ElseIf rs!CLASS = "prim5" And rs!SEX = "female" Then
Pfemale5.Text = rs!Fees_due
End If

If rs!CLASS = "prim6" And rs!SEX = "male" Then
Pmale6.Text = rs!Fees_due
ElseIf rs!CLASS = "prim6" And rs!SEX = "female" Then
Pfemale6.Text = rs!Fees_due
End If

rs.MoveNext
Loop








End Sub

