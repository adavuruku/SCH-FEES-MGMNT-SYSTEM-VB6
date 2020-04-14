VERSION 5.00
Begin VB.Form adminopt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Administrative Task"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   HelpContextID   =   3990
   LinkTopic       =   "Form6"
   ScaleHeight     =   2880
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000003&
         Caption         =   "Edit New Student Fees"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000003&
         Caption         =   "Edit Regular Student Fee"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H80000003&
         Caption         =   "Edit Exams Fee"
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "Create User Account"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "Edit User Previlege"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000003&
         Caption         =   "Create New Session"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000003&
         Caption         =   "Backup Database"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   5895
         Picture         =   "adminopt.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   450
      End
   End
End
Attribute VB_Name = "adminopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
form1.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command3_Click()
createsession.Show
Unload Me
End Sub

Private Sub Command4_Click()
frmDBBackUp.Show
Unload Me
End Sub


Private Sub Command5_Click()
Neweditfee.Show
Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
Private Sub Command7_Click()
editfee.Show
Unload Me
End Sub

Private Sub Command8_Click()
examfee.Show
Unload Me
End Sub
