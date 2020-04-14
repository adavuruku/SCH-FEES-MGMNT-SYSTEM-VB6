VERSION 5.00
Begin VB.Form seconterm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Second Term Fee Paying Form"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   HelpContextID   =   110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Second Term School Fees Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   9855
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
         Left            =   3600
         TabIndex        =   30
         Top             =   3480
         Width           =   975
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
         Left            =   1800
         TabIndex        =   29
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox telNo 
         Height          =   375
         Left            =   6600
         TabIndex        =   28
         Top             =   3480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         ToolTipText     =   "Select date"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtrno 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
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
         Height          =   375
         Left            =   6600
         TabIndex        =   21
         ToolTipText     =   "Select date"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         ToolTipText     =   "Select date"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtaccfee 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtfeedue 
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox arrears_due 
         Height          =   375
         Left            =   6600
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtbal 
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtadminno 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         ToolTipText     =   "Select date"
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         ItemData        =   "frmsec.frx":0000
         Left            =   1800
         List            =   "frmsec.frx":000A
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "frmsec.frx":001C
         Left            =   6600
         List            =   "frmsec.frx":0026
         TabIndex        =   1
         Top             =   840
         Width           =   2775
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
         Left            =   5280
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   1155
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
         Left            =   360
         TabIndex        =   31
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name.:"
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
         Left            =   5640
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status:"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fee Paid:"
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
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fee Due:"
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
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Balance:"
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
         Left            =   5400
         TabIndex        =   16
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reciept No.:"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   1800
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Acc. Fee:"
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
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Arrears Due:"
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
         Index           =   0
         Left            =   4920
         TabIndex        =   13
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sex:"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class:"
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
         Left            =   5760
         TabIndex        =   11
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin No.:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   9855
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear"
         Height          =   735
         Left            =   3000
         Picture         =   "frmsec.frx":0037
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   855
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
         Left            =   1800
         Picture         =   "frmsec.frx":0479
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
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
         Left            =   4080
         Picture         =   "frmsec.frx":0783
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT OF SECOND TERM SCHOOL FEES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   24
      Top             =   120
      Width           =   8280
   End
End
Attribute VB_Name = "seconterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub arrears_due_Change()
'txtaccfee.Text = Val(arrears_due) + Val(txtfeedue)
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

Private Sub cmbsex_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
Unload Me
'MDIForm11.Show
End Sub

Private Sub cmdclear_Click()

                        
End Sub

Private Sub cmdend_Click()

End Sub

Private Sub cmdsave_Click()

End Sub

Private Sub cmdsavee_Click(Index As Integer)
On Error Resume Next

Dim num As Double
'
If Check1.Value = 0 And Check2.Value = 0 Then
MsgBox " Select Cash or Bank", vbInformation, "Error"
Exit Sub
End If
If telNo = "" Then
MsgBox "Enter Teller No"
Exit Sub
End If


If txtadminno.Text = "" Or _
cmbclass.Text = "" Or _
cmbsex.Text = "" Or _
txtfeedue.Text = "" Or _
txtfeepaid.Text = "" Or _
arrears_due.Text = "" Or _
txtaccfee.Text = "" Or _
txtbal.Text = "" Or _
txtstatus.Text = "" Or _
txtfeepaid.Text = "" Or _
txtrno.Text = "" Then
MsgBox "some field are empty", vbInformation, "Fill missing field"
Exit Sub
End If

If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [2ndTerm]", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!Name = txtname.Text
rs!admin_num = txtadminno
rs!CLASS = cmbclass
rs!balcf = arrears_due
rs!Fees_paid = txtfeepaid
rs!arrears_due = txtbal
rs!Receipt_no = txtrno
rs!Fees_due = txtfeedue
rs!ACCFEE = txtaccfee
rs.Update
rs.Close
Set rs = Nothing
Set rs = cn.Execute("update receipt set receiptno='" & txtrno.Text & "'")

'GENERATING SCHOOL FEE RECEIPT
receipt.Show

receipt.amount = UCase(term1.txtname)
num = CCur(seconterm.txtfeepaid)
receipt.sumof = Words_Money(num)

receipt.receipt = seconterm.txtrno.Text
receipt.amount = UCase(seconterm.txtname)
receipt.datepaid = Now()
receipt.fee = "Second Term"
receipt.bal = txtbal
receipt.CLASS = UCase(cmbclass)

Call clear
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next

If RS1.State = adStateOpen Then RS1.Close

RS1.Open "select * from [Fee_Rating]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If RS1.EOF Then
message = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
cmbsex.Text = ""
cmbclass.Text = ""
cmbclass.SetFocus
Exit Sub
Else
feedue = RS1![Fees_due]
txtfeedue.Text = RS1![Fees_due]
txtfeedue.SetFocus

End If
RS1.Close
Set RS1 = Nothing

End Sub

Private Sub Command2_Click()
Call clear
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim num As Double
If rs.State = adStateOpen Then rs.Close
rs.Open "select receiptNo from [Receipt]", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
'rs.MoveFirst
num = Val(rs!receiptNo) + 1
txtrno = "000" & num
End If
End Sub

Private Sub txtadminno_LostFocus()
On Error GoTo handler:
If txtadminno = "" Then
Exit Sub
End If
Dim SEX As String
Dim CLASS As String
Dim arears As String
'******************************************
Dim hold As String
admin_num = txtadminno.Text
m1 = "[admin_num]='" + admin_num + "'"

Set rs = cn.Execute("select * from [1stterm]" & "where" & m1)
Set RS1 = cn.Execute("select * from [student_details]" & "where" & m1)
Set RS3 = cn.Execute("select * from [2ndterm]" & "where" & m1)
'**************************************
If Not RS3.EOF Then
MsgBox "Any further payment must be done with the installment option"
 txtadminno.Text = ""
 txtadminno.SetFocus
Else

'*************************************
If rs.EOF Then
message = MsgBox("invalid regestration no", vbCritical, "ERROR")
txtadminno.Text = ""
Else
'arrears_due.Text = rs!arrears_due
hold = rs!arrears_due
arrears_due.Text = Val(hold)
cmbclass = rs!CLASS
cmbsex = RS1!SEX
txtname.Text = RS1!Name
 If rs.State = adStateOpen Then rs.Close
'****************************************
SEX = cmbsex.Text
CLASS = cmbclass.Text
m1 = "[sex]='" + SEX + "'"
m2 = "[class]='" + CLASS + "'"
m3 = m1 & "AND" & m2
 If RS1.State = adStateOpen Then RS1.Close
RS1.Open "select * from [Fee_Rating]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If RS1.EOF Then
message = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
cmbsex.Text = ""
cmbclass.Text = ""
cmbclass.SetFocus
Else

txtfeedue.Text = RS1![Fees_due]
txtaccfee = Val(txtfeedue) + Val(arrears_due) 'accumulated fees

End If
End If
End If
handler:
If Err.Number = 3021 Then
MsgBox "Record not found", vbInformation, "NO Record"
Exit Sub
End If
End Sub
Private Function clear()
On Error Resume Next

txtadminno.Text = ""
cmbclass.Text = ""
cmbsex.Text = ""
txtfeedue.Text = ""
txtfeepaid.Text = ""
arrears_due.Text = ""
txtaccfee.Text = ""
txtbal.Text = ""
txtstatus.Text = ""
txtfeepaid.Text = ""
txtrno.Text = ""
txtname.Text = ""
Check1.Value = 0
Check2.Value = 0
telNo = ""

End Function


Private Sub txtfeedue_LostFocus()
txtaccfee.Text = Val(arrears_due) + Val(txtfeedue)
txtfeepaid.SetFocus
End Sub

Private Sub txtfeepaid_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtfeepaid_LostFocus()
On Error Resume Next

Dim PAID As Double
Dim DUE As Double
PAID = Val(txtfeepaid.Text)
DUE = Val(txtaccfee.Text)
ST = DUE - PAID
If ST = 0 Then
txtstatus.Text = "COMPLETE PAYMENT"
txtbal.Text = 0#
Exit Sub
Else
txtstatus.Text = "PART PAYMENT"
txtbal.Text = ST
End If
'txtrno.SetFocus
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtrno_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
