VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form newstudent 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   LinkTopic       =   "Form6"
   ScaleHeight     =   7635
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      BackColor       =   &H00808000&
      Height          =   375
      Left            =   9360
      TabIndex        =   40
      Top             =   3960
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   35
      Top             =   6480
      Width           =   9975
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H80000003&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         Picture         =   "newstudent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H80000003&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Picture         =   "newstudent.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
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
         Left            =   3840
         Picture         =   "newstudent.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Admission of New Student"
      Height          =   5415
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   9975
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808000&
         Height          =   375
         Left            =   9120
         TabIndex        =   39
         Top             =   2160
         Width           =   375
      End
      Begin VB.ComboBox term 
         Height          =   315
         ItemData        =   "newstudent.frx":0B8E
         Left            =   1800
         List            =   "newstudent.frx":0B98
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtadminno 
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtfeedue 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
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
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtbal 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
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
         Left            =   6600
         TabIndex        =   10
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtsweater 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtuniform 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   2415
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
         TabIndex        =   14
         Top             =   4560
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
         TabIndex        =   13
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox telNo 
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   4560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker txtdateadmit 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   105381891
         CurrentDate     =   39787
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Term:"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admission No.:"
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
         Left            =   4800
         TabIndex        =   34
         Top             =   960
         Width           =   1800
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
         TabIndex        =   33
         Top             =   1560
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
         Left            =   1080
         TabIndex        =   32
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Admitted:"
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
         Left            =   0
         TabIndex        =   31
         Top             =   1080
         Width           =   1830
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
         Left            =   480
         TabIndex        =   30
         Top             =   2640
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
         Left            =   360
         TabIndex        =   29
         Top             =   3240
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
         Left            =   5520
         TabIndex        =   28
         Top             =   3840
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
         Left            =   480
         TabIndex        =   27
         Top             =   3720
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reciept No"
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
         Left            =   5280
         TabIndex        =   26
         Top             =   3360
         Width           =   1185
      End
      Begin VB.Label Label11 
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
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sweater Fee"
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
         Left            =   5160
         TabIndex        =   24
         Top             =   2280
         Width           =   1320
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Uniform Fee"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   1500
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
         TabIndex        =   22
         Top             =   4560
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
         TabIndex        =   21
         Top             =   4320
         Width           =   1335
      End
   End
   Begin VB.TextBox bal 
      Height          =   285
      Left            =   8280
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Admission No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5340
      TabIndex        =   37
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADMISSION OF NEW STUDENT"
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
      Left            =   1680
      TabIndex        =   36
      Top             =   480
      Width           =   7545
   End
End
Attribute VB_Name = "newstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sta As Integer



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

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        txtfeedue.Text = Val(txtfeedue) + (Val(txtsweater.Text))
    Else
        txtfeedue.Text = Val(txtfeedue) - (Val(txtsweater.Text))
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        txtfeedue.Text = Val(txtfeedue) + (Val(txtuniform.Text))
    Else
        txtfeedue.Text = Val(txtfeedue) - (Val(txtuniform.Text))
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

Private Sub cmbsex_LostFocus()
On Error Resume Next

'Dim rs As New ADODB.Recordset
If cmbsex = "" And (cmbclass = "") Then
Exit Sub
Else
End If
Dim SEX As String
Dim CLASS As String
SEX = cmbsex.Text
CLASS = cmbclass.Text
m1 = "[sex]='" + SEX + "'"
m2 = "[class]='" + CLASS + "'"
m3 = m1 & "AND" & m2
If RS1.State = adStateOpen Then RS1.Close

RS1.Open "select * from  [NewFee_Rating]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If RS1.EOF Then
message = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
cmbsex.Text = ""
cmbclass.Text = ""
'cmbclass.SetFocus
Else

txtfeedue = RS1![Fees_due]

'cn.Close
'Set cn = Nothing
End If
RS1.Close
Set RS1 = Nothing
If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [price] where class='" & cmbclass.Text & "'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "select Appropriate class", vbInformation
Else
txtuniform.Text = rs!uniform_price
txtsweater.Text = rs!sweater_price

End If

End Sub

Private Sub cmdButton_Click(Index As Integer)
Unload Me

End Sub

Private Sub cmdclear_Click()
Call clear
End Sub

Private Sub cmdsave_Click()
On Error Resume Next

'End Sub
Dim num As Double
Dim result As Label
Dim CODE As String
Dim CLASS As String
Dim CCLASS As String
Dim term1 As String
If Len(cmbclass) = 4 Then
CODE = Left$(cmbclass, 3)
CLASS = Right$(cmbclass, 1)
Else
CODE = Left$(cmbclass, 4)
CLASS = Right$(cmbclass, 1)
End If
CCLASS = cmbclass.Text

term1 = term.Text
If Check1.Value = 0 And Check2.Value = 0 Then
MsgBox " Select Cash or Bank", vbInformation, "Error"
Exit Sub
End If
If telNo = "" Then
MsgBox "Enter Teller No"
Exit Sub
End If
 
If txtname.Text = "" Or _
txtstatus.Text = "" Or _
txtadminno = "" Or _
cmbclass = "" Or _
txtfeedue = "" Or _
txtfeepaid = "" Or _
txtbal = "" Or _
Text1.Text = "" Or _
cmbsex.Text = "" Then
MsgBox "Some fields are Empty", vbInformation, "fill missing fields"
Exit Sub
End If


'Set RS1 = cn.Execute("Insert into temp1 values('" & txtname.Text & "','" & txtstatus.Text & "','" & txtadminno & "','" & cmbclass & "','" & txtfeedue & "','" & txtfeepaid & "','" & txtbal & "','" & Text1.Text & "','" & txtdateadmit.Value & "','" & cmbsex.Text & "')")

Set rs = cn.Execute("update receipt set receiptno='" & Text1.Text & "'")


Set RS1 = cn.Execute("Insert into temp1 values('" & txtname.Text & "','" & txtstatus.Text & "','" & txtadminno & "','" & CLASS & "','" & txtfeedue & "','" & txtfeepaid & "','" & txtbal & "','" & Text1.Text & "','" & txtdateadmit.Value & "','" & cmbsex.Text & "','" & CODE & "','" & CCLASS & "')")

sql1 = cn.Execute("insert into Student_Details (admin_num,name,Status,sex,class,CODE,Date_admitted)" & "select admin_num,name,status,sex,class,CODE,Date_admitted From temp1")

'sql3 = "insert into 1stterm (name,status,admin_num,CLASS,Fees_due,fees_paid,Arrears_due,receipt_no )" & "select name,status,admin_num,CCLASS,Fees_due,fees_paid,Arrears_due,receipt_no From temp1"

sql3 = "insert into " & term1 & " (name,status,admin_num,CLASS,Fees_due,fees_paid,Arrears_due,receipt_no )" & "select name,status,admin_num,CCLASS,Fees_due,fees_paid,Arrears_due,receipt_no From temp1"
MsgBox "RECORD SAVED SUCCESSFULLY"
'cn.Execute sql1
cn.Execute sql3
sql2 = "delete*from temp1"
cn.Execute sql2


receipt.Show
receipt.receipt = Text1.Text
' ASSIGN NAME
receipt.amount = UCase(txtname)
num = CCur(txtfeepaid.Text)
receipt.sumof = Words_Money(num)
receipt.datepaid = Now()
receipt.fee = UCase(term.Text)   '"FIRST TERM"
receipt.bal = txtbal
receipt.CLASS = UCase(cmbclass)

Call clear
Unload Me


End Sub
Private Function clear()
txtadminno.Text = ""
'txtdateadmit.Value = ""
cmbclass.Text = ""
cmbsex.Text = ""
txtfeedue.Text = ""
txtfeepaid.Text = ""
txtbal.Text = ""
Text1.Text = ""

End Function

Private Sub Form_Load()
sta = 0

Dim num As String
If rs.State = adStateOpen Then rs.Close
rs.Open "select receiptNo from [Receipt]", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
'rs.MoveFirst
num = Val(rs!receiptNo) + 1
Text1 = "000" & num
End If

cmbclass.AddItem "NUR1"
cmbclass.AddItem "NUR2"
cmbclass.AddItem "NUR3"
cmbclass.AddItem "PRIM1"
cmbclass.AddItem "PRIM2"
cmbclass.AddItem "PRIM3"
cmbclass.AddItem "PRIM4"
cmbclass.AddItem "PRIM5"
cmbclass.AddItem "PRIM6"

cmbsex.AddItem "MALE"
cmbsex.AddItem "FEMALE"

End Sub

Private Sub Option1_Click()
'Dim store As Double
'store = Val(txtfeedue.Text)
'If Option1.Value = True Then
'Option2.Value = False
'txtfeedue.Text = Val(txtfeedue) + Val(txtsweater.Text)
Else

End If
End Sub

Private Sub Option2_Click()
'If Option2.Value = True Then
'txtfeedue.Text = Val(txtfeedue) - (Val(txtsweater.Text))
End If
End Sub

Private Sub Option3_Click()
'If Option3.Value = True Then
'Option4.Value = False
'txtfeedue.Text = Val(txtfeedue) - (Val(txtuniform.Text))
End If

End Sub

Private Sub Option3_LostFocus()
'Dim bal As String
If txtfeepaid = "" Then
Exit Sub
Else
End If
Dim PAID As Double
Dim DUE As Double
PAID = Val(txtfeepaid.Text)
DUE = Val(txtfeedue.Text)
ST = DUE - PAID
If ST = 0 Then
txtstatus.Text = "COMPLETE PAYMENT"
txtbal.Text = 0#
bal.Text = "0000"
Exit Sub
Else
txtstatus.Text = "PART PAYMENT"
txtbal.Text = ST
    bal.Text = ST
End If
Text1.SetFocus

End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Option3.Value = False
txtfeedue.Text = Val(txtfeedue) + (Val(txtuniform.Text))
End If

End Sub

Private Sub txtadminno_LostFocus()
'Dim rs As New ADODB.Recordset
 If RS1.State = adStateOpen Then RS1.Close
RS1.Open "select * from [Student_Details] where Admin_num='" & txtadminno & "'", cn, adOpenDynamic, adLockOptimistic
If Not RS1.EOF Then
MsgBox "Admission Number Already Exist "
txtadminno.Text = ""
txtadminno.SetFocus
'TXTADMINNUM.SetFocus
Else
Exit Sub
'rs.Close
'cn.Close
'TXTAMTPA.SetFocus
End If
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

'Dim bal As String
If txtfeepaid = "" Then
Exit Sub
Else
End If
Dim PAID As Double
Dim DUE As Double
PAID = Val(txtfeepaid.Text)
DUE = Val(txtfeedue.Text)
ST = DUE - PAID
If ST = 0 Then
txtstatus.Text = "COMPLETE PAYMENT"
txtbal.Text = 0#
bal.Text = "0000"
Exit Sub
Else
txtstatus.Text = "PART PAYMENT"
txtbal.Text = ST
    bal.Text = ST
End If
Text1.SetFocus

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub
