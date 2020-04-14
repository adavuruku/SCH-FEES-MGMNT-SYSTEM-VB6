VERSION 5.00
Begin VB.Form frmmock 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form6"
   ScaleHeight     =   2460
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Select Year"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      Begin VB.ComboBox txtdate 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "Display"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label yer 
         BackColor       =   &H80000003&
         Caption         =   "YEAR OF EXAMS"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   4800
         Picture         =   "frmmock.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   570
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report of Registered Mock Students"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3765
   End
End
Attribute VB_Name = "frmmock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call MOCK1
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1990 To 2030
txtdate.AddItem i
Next i
End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub txtdate_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
