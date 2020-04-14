VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form instalpay 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Installmental Payment"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   HelpContextID   =   1360
   LinkTopic       =   "Form3"
   ScaleHeight     =   5400
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INSTALLMENTAL PAYMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin VB.TextBox telNo 
         Height          =   375
         Left            =   6480
         TabIndex        =   22
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdsavee 
         BackColor       =   &H80000003&
         Caption         =   "&Save"
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
         Index           =   6
         Left            =   3000
         Picture         =   "frminstalpay.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H80000003&
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
         Left            =   4080
         Picture         =   "frminstalpay.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5040
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   179240961
         CurrentDate     =   39912
      End
      Begin VB.TextBox txtname 
         Height          =   405
         Left            =   6960
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cmbterm 
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox arrearsdue 
         Height          =   405
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox TXTNEWAREARDUE 
         Height          =   405
         Left            =   6960
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox TXTADMINNUM 
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         ToolTipText     =   "Please enter admision number"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TXTAMTPA 
         Height          =   405
         Left            =   6960
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         TabIndex        =   24
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Teller No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
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
         Left            =   6240
         TabIndex        =   17
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   6360
         TabIndex        =   16
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AREARS DUE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT NO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW_AREARSDUE:"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT PAID:"
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
         Left            =   5385
         TabIndex        =   12
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADMISSION NUMBER:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   2355
      End
   End
End
Attribute VB_Name = "instalpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rs As New ADODB.Recordset



Private Sub arrearsdue_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If

End Sub

Private Sub Check1_Click()
Dim Paymode As String
If Check1.Value = 1 Then
Check2.Value = 0
Paymode = "cash"
telNo.Text = "Nill"
Else
telNo.Text = ""

End If

End Sub

Private Sub Check2_Click()
Dim Paymode As String
If Check2.Value = 1 Then
Check1.Value = 0
Paymode = "bank"
telNo.Visible = True
Label18.Visible = True
Else
telNo.Visible = False
Label18.Visible = False
End If
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

Private Sub cmdButton_Click(Index As Integer)
Unload Me
'MDIForm11.Show
End Sub


Private Function clear()
TXTADMINNUM.Text = ""
TXTAMTPA.Text = ""
TXTNEWAREARDUE.Text = ""
txtname.Text = ""
arrearsdue = ""
Text3 = ""

End Function

Private Sub cmdsavee_Click(Index As Integer)
On Error Resume Next

Dim num As Double
If Check1.Value = 0 And Check2.Value = 0 Then
MsgBox " Select Cash or Bank", vbInformation, "Error"
Exit Sub
End If
If telNo = "" Then
MsgBox "Enter Teller No"
Exit Sub
End If 'On Error GoTo ERRORHANDLER
'Dim RS1 As New Recordset
For Each txt In instalpay
    If txt = "" Then
    MsgBox "Some field are missing"
    Exit Sub
    End If
Next
cn.BeginTrans
Text2 = Val(TXTAMTPA) + Val(TXTNEWAREARDUE)
Set rs = cn.Execute("update receipt set receiptno='" & Text3.Text & "'")

 If rs.State = adStateOpen Then rs.Close

rs.Open "SELECT * FROM [" & instalpay.cmbterm & "] where Admin_num='" & TXTADMINNUM.Text & "'", cn, adOpenDynamic, adLockOptimistic
RS1.Open "SELECT * FROM [installmentpay]", cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!Name = txtname.Text
RS1!admin_num = TXTADMINNUM.Text
RS1!Fees_paid = TXTAMTPA.Text
RS1!arrears_due = TXTNEWAREARDUE.Text
RS1!Date = DTPicker1.Value

RS1!Receipt_no = Text3
Text2 = Val(rs![Fees_paid]) + Val(TXTAMTPA)
rs![arrears_due] = TXTNEWAREARDUE
RS1![Receipt_no] = Text3.Text
RS1.Update
rs.Update
RS1.Close
Set RS1 = Nothing
'rs.Update
cn.CommitTrans
ERRORHANDLER:
MsgBox "TRANSACTION NOT COMPLETEED"
receipt.Show
'***********************************
'receipt.amount = UCase(term1.txtname)
num = CCur(instalpay.TXTAMTPA.Text)
receipt.sumof = Words_Money(num)

receipt.receipt = instalpay.Text3.Text
receipt.amount = UCase(instalpay.txtname)
receipt.datepaid = Now()
receipt.fee = instalpay.cmbterm.Text
receipt.bal = instalpay.TXTNEWAREARDUE
receipt.CLASS = UCase(cmbclass)

'cn.RollbackTrans
Call clear
'Exit Sub
'RS.Close
'Set RS = Nothing
'cn.Close
'Set cn = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Dim num As String
If rs.State = adStateOpen Then rs.Close
rs.Open "select receiptNo from [Receipt]", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
num = Val(rs!receiptNo) + 1
Text3 = "000" & num
End If

cmbterm.AddItem "1stterm"
cmbterm.AddItem "2ndterm"
cmbterm.AddItem "3rdterm"
'cmbterm.SetFocus
End Sub



Private Sub Text3_LostFocus()
On Error Resume Next
'Dim rs As New ADODB.Recordset
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [" & instalpay.cmbterm & "] where receipt_no='" & Text3.Text & "'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
Exit Sub
Else
MsgBox "Receipt No Already Exist"
Text3.Text = ""
Text3.SetFocus
'TXTADMINNUM.SetFocus

'rs.Close
'cn.Close
'TXTAMTPA.SetFocus
End If
End Sub
Private Sub TXTADMINNUM_LostFocus()
On Error Resume Next

If TXTADMINNUM = "" Or (cmbterm = "") Then
MsgBox "Admission Number or Term Missing"

Exit Sub

Else
End If
'Dim rs2 As New ADODB.Recordset
 If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [" & instalpay.cmbterm & "] where Admin_num='" & TXTADMINNUM & "'", cn, adOpenDynamic, adLockOptimistic
If RS1.State = adStateOpen Then RS1.Close

RS1.Open "select * from [2ndterm] where Admin_num='" & TXTADMINNUM & "'", cn, adOpenDynamic, adLockOptimistic

If rs.EOF Then
MsgBox "student admission not found ", vbInformation
TXTADMINNUM = ""
'cmbterm = ""
'arrearsdue.Text = ""
'cmbterm.SetFocus
Else

'Exit Sub


arrearsdue = rs!arrears_due
txtname = rs!Name
If Val(rs!arrears_due) <= 0 Then
MsgBox "Thank You, Your fee payment for the Selected Term is complete"
Unload Me
Exit Sub

End If
'End If
RS1.Close
Set RS1 = Nothing
rs.Update
rs.Close
Set rs = Nothing
End If
End Sub


Private Sub TXTAMTPA_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub TXTAMTPA_LostFocus()
Dim TEST As Double
TEST = Val(arrearsdue.Text)
If TEST <= 0 Then
'MsgBox "Thank You, Your fee payment for the Selected Term is complete"
Exit Sub
'Unload Me
Else
TXTNEWAREARDUE.Text = Val(arrearsdue) - Val(TXTAMTPA)

End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub TXTNEWAREARDUE_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If

End Sub
