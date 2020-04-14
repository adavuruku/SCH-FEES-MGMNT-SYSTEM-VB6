VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form commonentrance 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Common Entrance Registration"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   HelpContextID   =   4150
   LinkTopic       =   "Form4"
   ScaleHeight     =   7605
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Left            =   3480
      Picture         =   "commonentrance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000003&
      Caption         =   "&Reset"
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
      Left            =   4680
      Picture         =   "commonentrance.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   855
   End
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
      Left            =   5880
      Picture         =   "commonentrance.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   855
   End
   Begin VB.ComboBox category 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "commonentrance.frx":0B8E
      Left            =   3600
      List            =   "commonentrance.frx":0B9B
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.ComboBox examyear 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "commonentrance.frx":0BB4
      Left            =   3600
      List            =   "commonentrance.frx":0BB6
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.ComboBox SEX 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "commonentrance.frx":0BB8
      Left            =   3600
      List            =   "commonentrance.frx":0BC2
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox regno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox cand_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Nosub 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Amt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5760
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker datee 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16761024
      CalendarTitleBackColor=   16761024
      Format          =   87818241
      CurrentDate     =   40058
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   28
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   27
      Top             =   3360
      Width           =   2385
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   26
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   25
      Top             =   2640
      Width           =   2385
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   24
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   23
      Top             =   2040
      Width           =   2385
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Candidate Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   17
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Subject Registered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   840
      TabIndex        =   16
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   15
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   13
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7800
      X2              =   7800
      Y1              =   1080
      Y2              =   7440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   120
      X2              =   7800
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   120
      X2              =   7800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMMON ENTRANCES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   22
      Top             =   6000
      Width           =   2385
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   960
      TabIndex        =   20
      Top             =   4560
      Width           =   2385
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   19
      Top             =   5400
      Width           =   2385
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   18
      Top             =   1440
      Width           =   2385
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exam Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   21
      Top             =   4080
      Width           =   2385
   End
End
Attribute VB_Name = "commonentrance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cand_name_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[a-z,A-Z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBER ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
Call clear
Form_Load
End Sub

Private Sub Command3_Click()
Call clear
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If cand_name.Text = "" Or _
SEX.Text = "" Or _
regno.Text = "" Or _
datee.Value = "" Or _
Nosub.Text = "" Or _
category.Text = "" Or _
Amt.Text = "" Or _
examyear.Text = "" Then
MsgBox "Some field(s) are empty"
Exit Sub
End If

If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[COMMONENTRANCE]", cn, adOpenDynamic, adLockOptimistic

With rs
.AddNew
!cand_name = cand_name.Text
!SEX = SEX.Text
!examtype = "Entrance Exam"
!regno = regno.Text
!Reg_date = datee.Value
!No_Sub_Offered = Nosub.Text
!category = category.Text
!amount = Amt.Text
!examyear = examyear.Text
MsgBox "Student Sucessfully Registered"
End With
rs.Update
Call clear
'RS.Update
rs.Close
Set rs = Nothing
Form_Load
End Sub

Private Sub examyear_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
Amt.Text = rs!centrance

For i = 1990 To 2030
examyear.AddItem i
Next i
End Sub
Public Sub clear()
cand_name.Text = ""
SEX.Text = ""
'examtype.Text = ""
regno.Text = ""
'datee.Value = ""
Nosub.Text = ""
category.Text = ""
Amt.Text = ""
examyear.Text = ""
End Sub
Public Sub validate()

End Sub

Private Sub Nosub_KeyPress(KeyAscii As Integer)
Dim res As Boolean


res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
If Val(Nosub) > 9 Then
MsgBox "Maximum Number of subject is Nine (9)"
Exit Sub
End If
End Sub

Private Sub Nosub_LostFocus()
If Val(Nosub) > 9 Then
MsgBox "Maximum Number of subject is Nine (9)"
Nosub = ""
Nosub.SetFocus
Exit Sub
End If
End Sub

Private Sub SEX_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[a-z,A-Z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBER ARE NOT ALLOWED "
KeyAscii = 0
End If

End Sub
