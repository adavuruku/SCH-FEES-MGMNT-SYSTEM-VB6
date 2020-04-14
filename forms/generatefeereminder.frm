VERSION 5.00
Begin VB.Form generatefeereminder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   HelpContextID   =   3060
   LinkTopic       =   "Form5"
   ScaleHeight     =   2895
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Select Term"
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      Begin VB.TextBox TXTADMINNUM 
         Height          =   405
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Please enter admision number"
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbterm 
         Height          =   315
         ItemData        =   "generatefeereminder.frx":0000
         Left            =   1560
         List            =   "generatefeereminder.frx":000D
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "Display"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "ADMISSION NUMBER:"
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
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label yer 
         BackColor       =   &H80000003&
         Caption         =   "TERM"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   4815
         Picture         =   "generatefeereminder.frx":002C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   570
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate Fee Reminder Letter"
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
      Width           =   3150
   End
End
Attribute VB_Name = "generatefeereminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End Sub

Private Sub cmbterm_KeyPress(KeyAscii As Integer)
Select Case B
Case "1STTERM"
Case "2NDTERM"
Case "3RDTERM"
Case Else
MsgBox "PLEASE MAKE THE RIGHT TERM SELECTION"
KeyAscii = 0
End Select

End Sub

Private Sub Command2_Click()
On Error Resume Next

If TXTADMINNUM = "" Or (cmbterm = "") Then
MsgBox "Term or Admission Number Missing"
Exit Sub
Else
End If
'Dim rs2 As New ADODB.Recordset
 If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [" & generatefeereminder.cmbterm & "] where Admin_num='" & TXTADMINNUM & "'", cn, adOpenDynamic, adLockOptimistic
'If RS1.State = adStateOpen Then RS1.Close

'RS1.Open "select * from [2ndterm] where Admin_num='" & TXTADMINNUM & "'", cn, adOpenDynamic, adLockOptimistic

If rs.EOF Then
TXTADMINNUM = ""
Else
feereminder.Show
feereminder.lblbal.Caption = rs!arrears_due
feereminder.lclass.Caption = rs!CLASS
feereminder.lname.Caption = rs!Name
feereminder.lregno.Caption = rs!admin_num
'arrearsdue = RS!arrears_due
'txtname = RS!Name
rs.Close
Set rs = Nothing
Unload Me
End If


End Sub

Private Sub Image4_Click()
Unload Me
End Sub

