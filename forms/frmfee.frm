VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form firsterm 
   BackColor       =   &H00FFC0FF&
   Caption         =   "First Term Fee Payment"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Student's Details"
      Height          =   5415
      Left            =   4920
      TabIndex        =   4
      Top             =   1560
      Width           =   6375
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Fee"
         Height          =   375
         Left            =   4800
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         TabIndex        =   24
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox txtbal 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   3360
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
         TabIndex        =   9
         Top             =   2880
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Select date"
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         ItemData        =   "frmfee.frx":0000
         Left            =   1800
         List            =   "frmfee.frx":000A
         TabIndex        =   7
         Top             =   2400
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "frmfee.frx":001C
         Left            =   1800
         List            =   "frmfee.frx":0026
         TabIndex        =   6
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtadminno 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
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
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59703299
         CurrentDate     =   39787
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Name:"
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
         TabIndex        =   27
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Reciept No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Balance:"
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
         TabIndex        =   21
         Top             =   4200
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   20
         Top             =   3840
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fee Paid:"
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
         TabIndex        =   19
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fee Due:"
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
         TabIndex        =   18
         Top             =   2880
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Date Admitted:"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Sex:"
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
         TabIndex        =   16
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         Caption         =   "Class:"
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
         TabIndex        =   15
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
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
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   6960
      Width           =   6375
      Begin VB.CommandButton cmdend 
         BackColor       =   &H00FFC0FF&
         Caption         =   "End"
         Height          =   255
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Clear"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Save"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5760
      Top             =   7680
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
      Left            =   6300
      TabIndex        =   22
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "firsterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbsex_LostFocus()
Dim SEX As String
Dim CLASS As String
SEX = cmbsex.Text
CLASS = cmbclass.Text
m1 = "[sex]='" + SEX + "'"
m2 = "[class]='" + CLASS + "'"
m3 = m1 & "AND" & m2
RS1.Open "select * from  [Fee_Rating]" & "where" & m3, cn, adOpenDynamic, adLockOptimistic
If RS1.EOF Then
message = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
cmbsex.Text = ""
cmbclass.Text = ""
'cmbclass.SetFocus
Else

txtfeedue = RS1![Fees_due]
cmbsex.Enabled = False
cmbclass.Enabled = False
'cn.Close
'Set cn = Nothing
End If
RS1.Close
Set RS1 = Nothing
End Sub

Private Sub cmdclear_Click()
Call clear
cmbsex.Enabled = True
cmbclass.Enabled = True

End Sub

Private Sub cmdend_Click()
Unload Me
MDIForm1.Show
End Sub

'Set RS = cn.Execute("Insert
Private Sub cmdsave_Click()
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim db As New ADODB.Connection

Set rs = cn.Execute("Insert into temp1 values('" & txtadminno & "','" & cmbsex & "','" & cmbclass & "','" & txtdateadmit & "','" & txtfeepaid & "','" & txtbal & "','" & Text1.Text & "','" & txtname.Text & "' )")
sql1 = "insert into Student_Details (admin_num,name,sex,class,Date_admitted)" & "select admin_num,name,sex,class,Date_admitted From temp1"
sql3 = "insert into 1stterm (name,admin_num,class,fees_paid,Arrears_due,receipt_no )" & "select name,admin_num,class,fees_paid,Arrears_due,receipt_no From temp1"
cn.Execute sql1
cn.Execute sql3
sql2 = "delete*from temp1"
cn.Execute sql2
cn.Close
Set cn = Nothing



Call clear

End Sub

Private Sub Command1_Click()
Dim db As New ADODB.Connection
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\FeesMgtSystem.mdb;Persist Security Info=False"
RS1.Open "select * from [Fee_Rating]" & "where" & m3, db, adOpenDynamic, adLockOptimistic
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

End Function


Private Sub txtadminno_LostFocus()
Dim rs As New ADODB.Recordset
rs.Open "select * from [1stterm] where Admin_num='" & txtadminno & "'", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
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

Private Sub txtfeepaid_LostFocus()
Dim PAID As Integer
Dim DUE As Integer
PAID = Val(txtfeepaid.Text)
DUE = Val(txtfeedue.Text)
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
