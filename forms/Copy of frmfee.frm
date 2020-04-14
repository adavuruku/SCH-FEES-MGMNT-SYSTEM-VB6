VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5160
      Top             =   7680
      Width           =   4920
      _ExtentX        =   8678
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5160
      TabIndex        =   19
      Top             =   6960
      Width           =   4935
      Begin VB.CommandButton cmdend 
         Caption         =   "End"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student's Details"
      Height          =   4935
      Left            =   5160
      TabIndex        =   7
      Top             =   1920
      Width           =   4935
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
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Format          =   59703299
         CurrentDate     =   39787
      End
      Begin VB.TextBox txtbal 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtstatus 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtfeepaid 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   3120
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
         TabIndex        =   16
         Top             =   2640
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
         TabIndex        =   0
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox cmbsex 
         Height          =   315
         ItemData        =   "Copy of frmfee.frx":0000
         Left            =   1800
         List            =   "Copy of frmfee.frx":000A
         TabIndex        =   3
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ComboBox cmbclass 
         Height          =   315
         ItemData        =   "Copy of frmfee.frx":001C
         Left            =   1800
         List            =   "Copy of frmfee.frx":001E
         TabIndex        =   2
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtadminno 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         TabIndex        =   15
         Top             =   3960
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         TabIndex        =   14
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         TabIndex        =   13
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         TabIndex        =   12
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         TabIndex        =   11
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         TabIndex        =   10
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         TabIndex        =   9
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
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
      Left            =   5640
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmbsex_LostFocus()
Dim SEX As String
Dim CLASS As String
txtdateadmit.Text = DTPicker1

SEX = cmbsex.Text
CLASS = cmbclass.Text
m1 = "[sex]='" + SEX + "'"
m2 = "[class]='" + CLASS + "'"
m3 = m1 & "AND" & m2
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\FeesMgtSystem.mdb;Persist Security Info=False"
rs.Open "select * from [Fee_Rating]" & "where" & m3, db, adOpenDynamic, adLockOptimistic
If rs.EOF Then
aj = MsgBox("select Appropriate Class And Sex", vbCritical, "ERROR")
Else
txtfeedue.Text = rs![Fees_due]
rs.Close
Set rs = Nothing
db.Close
Set db = Nothing
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo errorhandler
cn.BeginTrans
Set rs = db.Execute("Insert into Student_Details values('" & txtadminno & "','" & cmbsex & "','" & cmbclass & "','" & txtdateadmit & "')")
cn.CommitTrans
MsgBox "good"
Exit Sub
errorhandler:
'cn.RollbackTrans
MsgBox "dgghjjjjjkj"
'cn.Close
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
