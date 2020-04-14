VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "SELECT TERM "
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   HelpContextID   =   3160
   LinkTopic       =   "Form5"
   ScaleHeight     =   2445
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Admitted Student Fees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Install Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Third term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Second term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   5400
         Picture         =   "Form5.frx":0000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   690
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
StudentType.Show
End Sub

Private Sub Command2_Click()
Unload Me
seconterm.Show
End Sub

Private Sub Command3_Click()
Unload Me
Thirdterm.Show
End Sub

Private Sub Command4_Click()
Unload Me
instalpay.Show
End Sub

Private Sub Command7_Click()
editfee.Show
Unload Me
End Sub

Private Sub Command8_Click()
examfee.Show
Unload Me
End Sub

Private Sub Command5_Click()
newstudent.Show
Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
