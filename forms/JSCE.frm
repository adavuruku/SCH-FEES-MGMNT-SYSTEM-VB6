VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form JSCE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Common Entrance Registration"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   HelpContextID   =   3540
   LinkTopic       =   "Form4"
   ScaleHeight     =   8295
   ScaleWidth      =   8400
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
      Left            =   3840
      Picture         =   "JSCE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
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
      Left            =   5040
      Picture         =   "JSCE.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
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
      Left            =   6240
      Picture         =   "JSCE.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6000
      Width           =   855
   End
   Begin MSComCtl2.DTPicker datee 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   4800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   179240961
      CurrentDate     =   40068
   End
   Begin VB.ComboBox examyear 
      Height          =   315
      ItemData        =   "JSCE.frx":0B8E
      Left            =   3960
      List            =   "JSCE.frx":0B90
      TabIndex        =   4
      Top             =   3480
      Width           =   3135
   End
   Begin VB.ComboBox SEX 
      Height          =   315
      ItemData        =   "JSCE.frx":0B92
      Left            =   3960
      List            =   "JSCE.frx":0B9C
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox category 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox regno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox cand_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Nosub 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   3960
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5280
      Width           =   3135
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
      Left            =   1200
      TabIndex        =   24
      Top             =   2760
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
      Left            =   1320
      TabIndex        =   23
      Top             =   2880
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
      Left            =   1200
      TabIndex        =   22
      Top             =   2040
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
      Left            =   1320
      TabIndex        =   21
      Top             =   2160
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
      Left            =   1200
      TabIndex        =   20
      Top             =   1440
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
      Left            =   1320
      TabIndex        =   19
      Top             =   1560
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
      Left            =   1200
      TabIndex        =   13
      Top             =   840
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
      Left            =   1200
      TabIndex        =   12
      Top             =   3960
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
      Left            =   1200
      TabIndex        =   11
      Top             =   5400
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
      Left            =   1200
      TabIndex        =   9
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   480
      X2              =   480
      Y1              =   720
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   8160
      X2              =   8160
      Y1              =   720
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   480
      X2              =   8160
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      X1              =   480
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JUNIOR WAEC REGISTRATION"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   6015
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
      Left            =   1320
      TabIndex        =   18
      Top             =   5520
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
      Left            =   1320
      TabIndex        =   16
      Top             =   4080
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
      Left            =   1320
      TabIndex        =   15
      Top             =   4920
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
      Left            =   1320
      TabIndex        =   14
      Top             =   960
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
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
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
      Left            =   1320
      TabIndex        =   17
      Top             =   3600
      Width           =   2385
   End
End
Attribute VB_Name = "JSCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Form_Load
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
rs.Open "select * from[JSCE]", cn, adOpenDynamic, adLockOptimistic

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
End If
Form_Load

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
Amt.Text = rs!JSCE

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

