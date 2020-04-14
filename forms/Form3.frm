VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   HelpContextID   =   1660
   LinkTopic       =   "Form3"
   ScaleHeight     =   3765
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdButton 
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
         Left            =   3840
         Picture         =   "Form3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enter"
         Height          =   397
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmstatus 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000008&
         Caption         =   "Enter"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C1AB7D&
         Caption         =   "Select Status"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cmstatus.Text = "Old Student" Then
Unload Me
Me.Hide
'MDIForm1.Picture1.Visible = False
firsterm.Show
Else
If cmstatus.Text = "New Student" Then
Unload Me
Me.Hide
term1.Show
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
cmstatus.AddItem "Old Student"
cmstatus.AddItem "New Student"

End Sub
