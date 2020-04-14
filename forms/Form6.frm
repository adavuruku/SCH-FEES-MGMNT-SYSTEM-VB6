VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000009&
   Caption         =   "EXAM REGISTRATION UPDATE"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form6"
   ScaleHeight     =   6465
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear"
         Height          =   735
         Left            =   2640
         Picture         =   "Form6.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   720
         Picture         =   "Form6.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtcommon 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox txtjsce 
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
         Left            =   1920
         TabIndex        =   11
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtnecoexternal 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtnecointernal 
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
         Left            =   1920
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtwaecexternal 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtwaecinternal 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "COMMONENTRANCE EXAM FEE"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "JSCE EXAM FEE"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "NECO  EXTERNAL FEE"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "NECO INTERNAL FEE"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "WAEC EXTERNAL FEE"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "WAEC INTERNAL FEE"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub t_Change()

End Sub

Private Sub cmdButton_Click(Index As Integer)
If RS.State = adStateOpen Then RS.Close
RS.Open "select *from examfee", cn, adOpenDynamic, adLockOptimistic
RS!waecfee = txtwaecinternal.Text
RS!Ewaecfee = txtwaecexternal.Text
RS!neco = txtnecointernal.Text
RS!eneco = txtnecoexternal.Text
RS!JSCE = txtjsce.Text
RS!centrance = txtcommon.Text
RS.Update
MsgBox "Update successfull", vbInformation
Call clear
End Sub

Private Sub Form_Load()
If RS.State = adStateOpen Then RS.Close
RS.Open "select *from examfee", cn, adOpenDynamic, adLockOptimistic
txtwaecinternal.Text = RS!waecfee
txtwaecexternal.Text = RS!Ewaecfee
txtnecointernal.Text = RS!neco
txtnecoexternal.Text = RS!eneco
txtjsce.Text = RS!JSCE
txtcommon.Text = RS!centrance

End Sub
Private Function clear()
txtwaecinternal.Text = ""
txtwaecexternal.Text = ""
txtnecointernal.Text = ""
txtnecoexternal.Text = ""
txtjsce.Text = ""
txtcommon.Text = ""
End Function
