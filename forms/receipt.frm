VERSION 5.00
Begin VB.Form receipt 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Receipt of school fees payment"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   HelpContextID   =   2620
   LinkTopic       =   "Form4"
   ScaleHeight     =   8355
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "&Print"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   ":00 k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label BAL 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Left            =   1920
      TabIndex        =   14
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label CLASS 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   5760
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   7320
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   6240
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   7440
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   7080
      Picture         =   "receipt.frx":0000
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   690
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   240
      Picture         =   "receipt.frx":0531
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPLIFT PRIMARY SCHOOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "No. 18, Orafaga dem Street"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Adekaa, Gboko, Benue State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label fee 
      BackStyle       =   0  'Transparent
      Caption         =   "TERM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label sumof 
      BackStyle       =   0  'Transparent
      Caption         =   "AMT"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   4560
      Width           =   5655
   End
   Begin VB.Label amount 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Label datepaid 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label receipt 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPT NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "receipt.frx":3F61B
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   7935
   End
End
Attribute VB_Name = "receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdGo_Click()
'Dim num As Currency

    'On Error GoTo BadNumber

    num = CCur(Text1.Text)
    'lblWords.Caption = Words_1_all(num)
    sumof.Caption = Words_Money(num)
    Exit Sub

BadNumber:
    MsgBox "The value must be a numeric currency value", _
        vbCritical
    'term1.txtfeepaid.SetFocus
    
    
End Sub

Private Sub Command1_Click()
On Error Resume Next
Me.PrintForm  'Call AccessShowReport
End Sub

Private Sub Form_Load()
'cmdGo_Click
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
