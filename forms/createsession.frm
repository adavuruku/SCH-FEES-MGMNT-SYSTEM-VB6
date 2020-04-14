VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form createsession 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CREATE NEW SESSION"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   HelpContextID   =   2500
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton createtable 
      BackColor       =   &H80000003&
      Caption         =   "Create Tables "
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\database\sessions.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sch Fees Mgt System\database\sessions.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "New session to be created"
      Top             =   1440
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create New Session"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   7200
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7200
      X2              =   7200
      Y1              =   360
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   2880
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " NEW SESSION"
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
      Left            =   705
      TabIndex        =   5
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " CURRENT SESSION"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT NEW SESSION TO USE AS CURRENT SESSION"
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
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "createsession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next

Dim FileSystemObject As Object

Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
If Combo1.Text = "" Then
MsgBox "Please Select new Session"
Exit Sub
ElseIf Val(Combo1.Text) <= Val(Text1.Text) Then
MsgBox "New session must greater than previous session"
Exit Sub
Else
FileSystemObject.CopyFile App.Path & "\database\" & Text1.Text & ".mdb", App.Path & "\database\" & Combo1.Text & ".mdb"
MsgBox Combo1.Text & "New session is sucessuflly Created"

If rs.State = adStateOpen Then rs.Close
con
rs.Open "select * from[sessions]", db, adOpenDynamic, adLockOptimistic
rs.AddNew
rs!dbname = Combo1.Text
rs.Update
Combo1.Text = ""
End If
defs
Set rs = Nothing
createtable_Click
MsgBox "Restart the System to Effect New changes", vbInformation
End
End Sub
Public Sub defs()
On Error Resume Next

Dim TEST As String
create
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [Student_Details]", db, adOpenDynamic, adLockOptimistic
'Set RS1 = db.Execute("insert into graduated (name,sex)" & "select name,sex From Student_Details where class='" & "prim3" & "'")
rs.Close
Set rs = Nothing
End Sub
Private Sub Command2_Click()
With CommonDialog1
.filter = "Database files (*.mdb)|*.mdb"
.ShowOpen
Text1.Text = .FileName
End With
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next


FileCopy App.Path & "\test.html", App.Path & "\look.html"

End Sub

Private Sub Command5_Click()

End Sub

Private Sub createtable_Click()
On Error Resume Next

create
Dim CODE As String
Dim CLASS As String
Dim yearr As Integer

If RS4.State = adStateOpen Then RS4.Close
RS4.Open "select * from [Student_Details]", db, adOpenDynamic, adLockOptimistic
yearr = Year(Date)
CLASS = "6"
CODE = "PRIM"
m1 = "[CODE]='" + CODE + "'"
m2 = "[class]='" + CLASS + "'"
m3 = m1 & "AND" & m2

Set rs = db.Execute("Delete * From [1stterm]")
Set rs = db.Execute("Delete * From [2ndterm]")
Set rs = db.Execute("Delete * From [3rdterm]")
Set rs = db.Execute("Delete * From [commonentrance]")
Set rs = db.Execute("Delete * From [neco]")
Set rs = db.Execute("Delete * From [waec]")
Set rs = db.Execute("Delete * From [jsce]")
Set rs = db.Execute("Delete * From [installmentpay]")
Set rs = db.Execute("Delete * From [MOCK]")

'Set RS4 = db.Execute("insert into graduated (name,sex,class,Admin_num,Date_admited)" & "select name,sex,class,Admin_num,Date_admitted From [Student_Details]" & "where" & m3)
'Set RS4 = db.Execute("delete * from[Student_Details]" & "where" & m3)
'Set RS4 = db.Execute("update Student_Details set class = val(Student_Details.class) + 1 ")
'Set RS4 = db.Execute("insert into graduated (name,sex,class,Admin_num,Date_admited)" & "select name,sex,class,Admin_num,Date_admitted From [Student_Details]" & "where" & m3)
'Set RS4 = db.Execute("delete * from[Student_Details]" & "where" & m3)
'Set RS4 = db.Execute("update Student_Details set class = val(Student_Details.class) + 1 ")

Set RS4 = db.Execute("update Student_Details set year_grad=('" & yearr & "')" & "where" & m3)
Set RS4 = db.Execute("insert into graduated (name,sex,class,Admin_num,Date_admited,year_grad)" & "select name,sex,class,Admin_num,Date_admitted,year_grad From [Student_Details]" & "where" & m3)
Set RS4 = db.Execute("delete * from[Student_Details]" & "where" & m3)
Set RS4 = db.Execute("update Student_Details set class = val(Student_Details.class) + 1 ")
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim RIGHT1 As String
Dim LEFT1 As String
con
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [sessions]", db, adOpenDynamic, adLockOptimistic
rs.MoveLast
mdbname1 = rs!dbname
'createsession.Show
createsession.Text1 = mdbname1    '& ".mdb"
RIGHT1 = Val(Right$(mdbname1, 4)) + 1
LEFT1 = Val(Left$(mdbname1, 4)) + 1
Combo1.Text = LEFT1 & RIGHT1
End Sub
