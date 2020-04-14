VERSION 5.00
Begin VB.Form filter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "SECTION TO SHOW LIST OF DEPBTORS"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   HelpContextID   =   10
   LinkTopic       =   "Form3"
   ScaleHeight     =   2670
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Select Term and Class Level"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "filter.frx":0000
         Left            =   1560
         List            =   "filter.frx":001F
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "Generate Report"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbterm 
         Height          =   315
         ItemData        =   "filter.frx":005F
         Left            =   1560
         List            =   "filter.frx":006C
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   5400
         Picture         =   "filter.frx":008B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5400
         TabIndex        =   6
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "Report Generation for Debtors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)
Unload Me
MDIForm11.Show

End Sub

Private Sub Command1_Click()
On Error Resume Next

If cmbterm = "" Or cmbclass = "" Then
MsgBox "Select Term and Class Level"
Exit Sub
Else
Call DebtReport
End If
Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
