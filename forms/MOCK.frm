VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MOCK 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form6"
   ScaleHeight     =   7440
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000003&
      Caption         =   "Register"
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
      Picture         =   "MOCK.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
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
      Picture         =   "MOCK.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
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
      Picture         =   "MOCK.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   855
   End
   Begin VB.ComboBox category 
      Height          =   315
      ItemData        =   "MOCK.frx":0B8E
      Left            =   3600
      List            =   "MOCK.frx":0B9B
      TabIndex        =   16
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Amt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      TabIndex        =   6
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox Nosub 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox cand_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox regno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.ComboBox SEX 
      Height          =   315
      ItemData        =   "MOCK.frx":0BAE
      Left            =   3600
      List            =   "MOCK.frx":0BB8
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ComboBox examyear 
      Height          =   315
      ItemData        =   "MOCK.frx":0BCA
      Left            =   3600
      List            =   "MOCK.frx":0BCC
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker datee 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   4680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   179240961
      CurrentDate     =   40068
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   15
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MOCK EXAMINATION REGISTRATION"
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
      Left            =   480
      TabIndex        =   14
      Top             =   0
      Width           =   6975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   120
      X2              =   7800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   120
      X2              =   7800
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   7800
      X2              =   7800
      Y1              =   600
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   6840
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   12
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "MOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cand_name_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
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

End Sub

Private Sub Command4_Click()

If cand_name.Text = "" Or _
SEX.Text = "" Or _
regno.Text = "" Or _
Nosub.Text = "" Or _
category.Text = "" Or _
Amt.Text = "" Or _
examyear.Text = "" Then
MsgBox "Some field(s) are empty"
Exit Sub
Else

If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[MOCK]", cn, adOpenDynamic, adLockOptimistic

With rs
.AddNew
!cand_name = cand_name.Text
!SEX = SEX.Text
!examtype = "MOCK Exam"
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
End If
Form_Load

End Sub

Private Sub Command5_Click()

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
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
Amt.Text = rs!MOCK

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
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
