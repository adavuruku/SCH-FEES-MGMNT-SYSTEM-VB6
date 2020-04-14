VERSION 5.00
Begin VB.Form frmcommon 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   HelpContextID   =   3470
   LinkTopic       =   "Form4"
   ScaleHeight     =   2760
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Select Year"
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      Begin VB.ComboBox txtdate 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "Display"
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
         Height          =   255
         Left            =   4935
         Picture         =   "common.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report of Registered  Common Entrance Students"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   5145
   End
End
Attribute VB_Name = "frmcommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call code30
End Sub

Private Sub Command2_Click()
On Error Resume Next
code30
End Sub

Private Sub Form_Load()
If rs.State = adStateOpen Then rs.Close
rs.Open "select Distinct examyear from [COMMONENTRANCE]", cn, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
txtdate.AddItem rs!examyear
rs.MoveNext
Loop


'On Error GoTo handler
'Dim i As Integer
'For i = 1990 To 2009
'txtdate.AddItem i
'Next i
'handler:
'If Err.Number = 424 Then
'MsgBox "Control name is not found", vbInformation, "Contact system Admnistrator"
'Exit Sub
'End If
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
