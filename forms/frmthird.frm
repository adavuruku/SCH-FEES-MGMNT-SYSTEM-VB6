VERSION 5.00
Begin VB.Form Thirdterm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Third Term Fee Paying Form"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10590
   HelpContextID   =   750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Third Term Fee Paying Interface"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   9975
      Begin VB.TextBox telNo 
         Height          =   375
         Left            =   6600
         TabIndex        =   30
         Top             =   4080
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
         Left            =   1800
         TabIndex        =   29
         Top             =   4080
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
         Left            =   3600
         TabIndex        =   28
         Top             =   4080
         Width           =   975
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
         TabIndex        =   27
         ToolTipText     =   "Select date"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         Height          =   405
         Left            =   6600
         TabIndex        =   21
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
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Select date"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "frmthird.frx":0000
         Left            =   6600
         List            =   "frmthird.frx":000A
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         ItemData        =   "frmthird.frx":001B
         Left            =   1800
         List            =   "frmthird.frx":0025
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtbal 
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox arrears_due 
         Height          =   375
         Left            =   6600
         TabIndex        =   8
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtfeedue 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtaccfee 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
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
         TabIndex        =   32
         Top             =   3840
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
         Left            =   5280
         TabIndex        =   31
         Top             =   4080
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name:"
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
         Index           =   2
         Left            =   5760
         TabIndex        =   22
         Top             =   360
         Width           =   780
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
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   2760
         Width           =   1500
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
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1320
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
         TabIndex        =   18
         Top             =   960
         Width           =   750
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
         Left            =   1170
         TabIndex        =   17
         Top             =   840
         Width           =   540
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
         Left            =   4965
         TabIndex        =   16
         Top             =   1560
         Width           =   1545
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
         Left            =   600
         TabIndex        =   15
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label8 
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
         Index           =   0
         Left            =   840
         TabIndex        =   14
         Top             =   3240
         Width           =   885
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
         Left            =   5520
         TabIndex        =   13
         Top             =   2640
         Width           =   1065
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
         Left            =   585
         TabIndex        =   11
         Top             =   1440
         Width           =   1125
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
         Left            =   5400
         TabIndex        =   9
         Top             =   2040
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   5760
      Width           =   9975
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
         Left            =   3360
         Picture         =   "frmthird.frx":0037
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
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
         Left            =   720
         Picture         =   "frmthird.frx":0341
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "Clear"
         Height          =   735
         Left            =   2040
         Picture         =   "frmthird.frx":064B
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT OF THIRD TERM SCHOOL FEES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   10245
   End
End
Attribute VB_Name = "Thirdterm"
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
Dim num As Double
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

rs.Open "select * from [3rdterm]", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!Name = txtname
rs!admin_num = txtadminno
rs!CLASS = cmbclass
rs!balcf = arrears_due
rs!Fees_paid = txtfeepaid
rs!arrears_due = txtbal
rs!Receipt_no = txtrno
rs!Fees_due = txtfeedue
rs!acc_fee = txtaccfee
rs.Update
rs.Close
Set rs = Nothing

'increament receipt number by 1
Set rs = cn.Execute("update receipt set receiptno='" & txtrno.Text & "'")

receipt.Show
'receipt.receipt = Thirdterm.txtrno.Text
'receipt.amount = Thirdterm.txtname
'receipt.sumof = Thirdterm.txtfeepaid
'receipt.datepaid = Now()
'receipt.fee = "Third Term"

receipt.amount = UCase(term1.txtname)
num = CCur(Thirdterm.txtfeepaid)
receipt.sumof = Words_Money(num)

receipt.receipt = Thirdterm.txtrno.Text
receipt.amount = UCase(Thirdterm.txtname)
receipt.datepaid = Now()
receipt.fee = "Second Term"
receipt.bal = txtbal
receipt.CLASS = UCase(cmbclass)

Call clear
Unload Me
End Sub

Private Sub Command1_Click()
'Dim db As New ADODB.Connection
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\FeesMgtSystem.mdb;Persist Security Info=False"

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
Dim num As String
If rs.State = adStateOpen Then rs.Close
rs.Open "select receiptNo from [Receipt]", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
'rs.MoveFirst
num = Val(rs!receiptNo) + 1
txtrno = "000" & num
End If
End Sub

Private Sub txtaccfee_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub txtadminno_LostFocus()
If txtadminno = "" Then
Exit Sub
End If
'Dim db As New ADODB.Connection
'Dim rs  As New ADODB.Recordset
'Dim rs3 As New ADODB.Recordset
Dim SEX As String
Dim CLASS As String
Dim arears As String
'******************************************
Dim hold As String
admin_num = txtadminno.Text
m1 = "[admin_num]='" + admin_num + "'"
Set rs = cn.Execute("select * from [2ndterm]" & "where" & m1)
Set RS1 = cn.Execute("select * from [student_details]" & "where" & m1)
Set RS3 = cn.Execute("select * from [3rdterm]" & "where" & m1)
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
'arrears_due.Text = Val(hold)
'txtfeepaid.SetFocus

End If
End If
End If


End Sub
Private Function clear()
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
txtname = ""
Check1.Value = 0
Check2.Value = 0

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

End Sub


Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
