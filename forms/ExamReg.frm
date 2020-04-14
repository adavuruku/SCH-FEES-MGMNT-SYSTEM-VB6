VERSION 5.00
Begin VB.Form ExamReg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Examination Registration"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   HelpContextID   =   4070
   LinkTopic       =   "Form6"
   ScaleHeight     =   3315
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000003&
         Caption         =   "MOCK Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000003&
         Caption         =   "JSCE Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000003&
         Caption         =   "Common Entrance Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "NECO Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "WAEC Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   6360
         Picture         =   "ExamReg.frx":0000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   690
      End
   End
End
Attribute VB_Name = "ExamReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

waecneco.Show SSTab1

'waecneco.SSTab2.
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next

waecneco.Show SSTab2
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next

commonentrance.Show
Unload Me
End Sub

Private Sub Command4_Click()

On Error Resume Next
JSCE.Show
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next

MOCK.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
