VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form8"
   ScaleHeight     =   3090
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   4800
         Picture         =   "installmental.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   570
      End
      Begin VB.Label yer 
         BackColor       =   &H80000003&
         Caption         =   "Admission Number"
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report of Student payment Details"
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
      TabIndex        =   3
      Top             =   0
      Width           =   3555
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

Call installmental
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
