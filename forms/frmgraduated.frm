VERSION 5.00
Begin VB.Form frmgraduated 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5865
   LinkTopic       =   "Form6"
   ScaleHeight     =   2175
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Select Year"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "Display"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox txtdate 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   4800
         Picture         =   "frmgraduated.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   570
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
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report of Graduated Students"
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
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3090
   End
End
Attribute VB_Name = "frmgraduated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

Call graduated
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

'con
If rs.State = adStateOpen Then rs.Close
rs.Open "select Distinct year_grad from [graduated]", cn, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
txtdate.AddItem rs!year_grad
rs.MoveNext
Loop

'Dim i As Integer
'For i = 1990 To 2030
'txtdate.AddItem i
'Next i
End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub txtdate_KeyPress(KeyAscii As Integer)
On Error Resume Next

Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
