VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form firsterm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "First Term Fee Payment For Old Student"
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   10065
   HelpContextID   =   410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Clear"
      Height          =   735
      Left            =   1320
      Picture         =   "frmfees.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdsavee 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   120
      Picture         =   "frmfees.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add"
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
      Index           =   0
      Left            =   6240
      Picture         =   "frmfees.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   7080
      Picture         =   "frmfees.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   4440
      Picture         =   "frmfees.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
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
      Left            =   2640
      Picture         =   "frmfees.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student's Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtsttus 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Text            =   "Old Student"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Fee"
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   6600
         TabIndex        =   20
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtbal 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6600
         TabIndex        =   6
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtfeedue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   5
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtdateadmit 
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
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Select date"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbsex 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmfees.frx":1374
         Left            =   6600
         List            =   "frmfees.frx":137E
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "frmfees.frx":1390
         Left            =   1800
         List            =   "frmfees.frx":139A
         TabIndex        =   2
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtadminno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
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
         Left            =   9120
         TabIndex        =   9
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59965443
         CurrentDate     =   39787
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5280
         TabIndex        =   19
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   600
         TabIndex        =   17
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   720
         TabIndex        =   16
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5400
         TabIndex        =   15
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   600
         TabIndex        =   14
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4800
         TabIndex        =   13
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5880
         TabIndex        =   12
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5520
      Top             =   8880
      Visible         =   0   'False
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\FeesMgtSystem.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   4860
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "firsterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbsex_LostFocus()
If (cmbclass = "") Or (cmbsex = "") Then
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
RS1.Open "select * from  [Fee_Rating]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If RS1.EOF Then
message = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
cmbsex.Text = ""
cmbclass.Text = ""
Else

txtfeedue = RS1![Fees_due]
'cmbsex.Enabled = False
'cmbclass.Enabled = False
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
Unload Me
MDIForm11.Show
End Sub

Private Sub cmdclear_Click()
Call clear
cmbsex.Enabled = True
cmbclass.Enabled = True

End Sub

Private Sub cmdend_Click()

End Sub

'Set RS = cn.Execute("Insert

Private Sub cmdsave_Click()



End Sub

Private Sub cmdsavee_Click(Index As Integer)
'Dim RS1 As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
'Dim db As New ADODB.Connection
If (txtadminno = "") Or (txtdateadmit = "") Or (txtname = "") Or (cmbclass = "") Or (cmbsex = "") Or (txtfeedue = "") Or (txtfeepaid = "") Or (txtstatus = "") Or (txtbal = "") Or (Text1 = "") Then
MsgBox ("Ensure that No field(s) is Empty")
Exit Sub
Else

Set rs = cn.Execute("Insert into temp1 values('" & txtname.Text & "','" & txtsttus.Text & "','" & txtadminno & "','" & cmbclass & "','" & txtfeedue & "','" & txtfeepaid & "','" & txtbal & "','" & Text1.Text & "','" & txtdateadmit.Text & "','" & cmbsex.Text & "')")
sql1 = "insert into Student_Details (admin_num,name,Status,sex,class,Date_admitted)" & "select admin_num,name,sex,class,Date_admitted,Status From temp1"
sql3 = "insert into 1stterm (name,status,admin_num,class,Fees_due,fees_paid,Arrears_due,receipt_no )" & "select name,status,admin_num,class,Fees_due,fees_paid,Arrears_due,receipt_no From temp1"
cn.Execute sql1
cn.Execute sql3
sql2 = "delete*from temp1"
cn.Execute sql2
'cn.Close
'Set cn = Nothing



Call clear
End If
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
db.Close
Set db = Nothing
End Sub

Private Sub DTPicker1_Change()
txtdateadmit = DTPicker1
End Sub
Public Function validate()

End Function
Private Function clear()
txtadminno.Text = ""
txtdateadmit.Text = ""
cmbclass.Text = ""
cmbsex.Text = ""
txtfeedue.Text = ""
txtfeepaid.Text = ""
txtbal.Text = ""
Text1.Text = ""
txtstatus.Text = ""
txtname.Text = ""
End Function


Private Sub txtadminno_LostFocus()
If txtadminno = "" Then
Exit Sub
Else
'Dim SEX As String
'Dim CLASS As String
'Dim rs As New ADODB.Recordset
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [Student_Details] where Admin_num='" & txtadminno & "'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "invalid Admission number ", vbCritical
txtadminno.Text = ""
txtadminno.SetFocus
Else
txtname.Text = rs!Name
txtdateadmit.Text = rs!Date_admitted
cmbclass = rs!CLASS
cmbsex = rs!SEX
End If
End If
'*******************************
cmbsex.SetFocus
rs.Close
Set rs = Nothing

End Sub

Private Sub txtfeepaid_LostFocus()
Dim PAID As Double
Dim DUE As Double
PAID = Val(txtfeepaid.Text)
DUE = Val(txtfeedue.Text)
ST = Val(DUE) - Val(PAID)
If ST = 0 Then
txtstatus.Text = "COMPLETE PAYMENT"
txtbal.Text = 0#
Exit Sub
Else
txtstatus.Text = "PART PAYMENT"
txtbal.Text = ST
End If
Text1.SetFocus
End Sub
