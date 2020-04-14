VERSION 5.00
Begin VB.Form examfee 
   BackColor       =   &H80000009&
   Caption         =   "EXAM REGISTRATION UPDATE"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   HelpContextID   =   3830
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdclose 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3600
         Picture         =   "examfee.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         Picture         =   "examfee.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4560
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
         Left            =   2640
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
         Left            =   2640
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
         Left            =   2640
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
         Left            =   2640
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
         Left            =   2640
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
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
         Caption         =   "COMMONENTRANCE EXAM FEE"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "JSCE EXAM FEE"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "NECO  EXTERNAL FEE"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "NECO INTERNAL FEE"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "WAEC EXTERNAL FEE"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "WAEC INTERNAL FEE"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "examfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub t_Change()

End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next

'For Each txt In examfee
 '   If txt = " " Then
  '  MsgBox "Some field are missing"
   ' Exit Sub
    'End If
'Next
If rs.State = adStateOpen Then rs.Close
rs.Open "select *from examfee", cn, adOpenDynamic, adLockOptimistic
rs!waecfee = txtwaecinternal.Text
rs!Ewaecfee = txtwaecexternal.Text
rs!neco = txtnecointernal.Text
rs!eneco = txtnecoexternal.Text
rs!JSCE = txtjsce.Text
rs!centrance = txtcommon.Text
rs.Update
MsgBox "Update successfull", vbInformation
Call clear
Unload Me
End Sub

Private Sub cmdclose_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = adStateOpen Then rs.Close
rs.Open "select *from examfee", cn, adOpenDynamic, adLockOptimistic
txtwaecinternal.Text = rs!waecfee
txtwaecexternal.Text = rs!Ewaecfee
txtnecointernal.Text = rs!neco
txtnecoexternal.Text = rs!eneco
txtjsce.Text = rs!JSCE
txtcommon.Text = rs!centrance

End Sub
Private Function clear()
txtwaecinternal.Text = ""
txtwaecexternal.Text = ""
txtnecointernal.Text = ""
txtnecoexternal.Text = ""
txtjsce.Text = ""
txtcommon.Text = ""
End Function





Private Sub txtcommon_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtjsce_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtnecoexternal_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtnecointernal_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtwaecexternal_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtwaecinternal_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
Public Sub CKECK()


End Sub
