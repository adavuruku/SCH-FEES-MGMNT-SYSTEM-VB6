Attribute VB_Name = "Module1"
Global cn As New ADODB.Connection
Global sql, sql2, sqll, sql3 As String
Global sql4, sql5, sql6 As String
Global rs As New ADODB.Recordset
Global RS1 As New ADODB.Recordset
Global RS2 As New ADODB.Recordset
Global RS3 As New ADODB.Recordset
Global RS4 As New ADODB.Recordset
Global RS5 As New ADODB.Recordset
Global RS6 As New ADODB.Recordset
Global RS7 As New ADODB.Recordset
Global RS8 As New ADODB.Recordset
Global RS9 As New ADODB.Recordset
Global RS10 As New ADODB.Recordset
Global RS11 As New ADODB.Recordset
Global m1, m2, m3, m4 As String
Global mdbname1 As String
Global mdbname As String
Global SESSIONAME As String
Global create1 As String
Global db As New ADODB.Connection

Sub Main()
con
rs.Open "select * from [sessions]", db, adOpenDynamic, adLockOptimistic
rs.MoveLast
mdbname1 = rs!dbname & ".mdb"
SESSIONAME = rs!dbname
With cn
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\database\" & mdbname1 & ";Persist Security Info=False"
.Open
End With

'put form name here to test
'********************************************************
'WELCOMESCREEN.Show
MDIForm11.Show
MDIForm11.Label6.Caption = SESSIONAME
login2.Show
'frmLogin.Show
'newstudent.Show
'createsession.Show
'receipt.Show
'filter.Show
'editfee.Show
'commonentrance.Show
'generatefeereminder.Show
End Sub
Public Sub con()
If db.State = adStateOpen Then db.Close
With db
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\sessions.mdb;Persist Security Info=False"
.Open
End With
End Sub

Public Sub create()
con

If rs.State = adStateOpen Then rs.Close

rs.Open "select * from [sessions]", db, adOpenDynamic, adLockOptimistic
rs.MoveLast
create1 = rs!dbname & ".mdb"
If db.State = adStateOpen Then db.Close
With db
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\database\" & create1 & ";Persist Security Info=False"
.Open
End With
End Sub

Public Sub DebtReport()
If rs.State = adStateOpen Then rs.Close

rs.Open "Select * From [" & filter.cmbterm & "] Where " & " class='" & filter.cmbclass.Text & "'" & "and" & "  (arrears_due > 0)", cn, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
'If rs.RecordCount > 0 Then
With DataReport1.Sections("section1").Controls
.Item("name").DataField = rs("name").Name
.Item("admin_num").DataField = rs("Admin_num").Name
.Item("Arrears_due").DataField = rs("Arrears_due").Name
End With
DataReport1.Sections("section4").Controls.Item("class").Caption = rs("class").Value
'DataReport1.Sections("section4").Controls.Item("class").Caption = " Prim1"

DataReport1.Sections("section5").Controls.Item("fxtotal").DataField = rs("Arrears_due").Name

Set DataReport1.DataSource = rs
DataReport1.Show
Set rs = Nothing
Else
MsgBox "No Record Found"
Exit Sub
End If
'End If
End Sub
'complete payment modules
Public Sub fullpayment()
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [" & MDIForm11.mnuist.Caption & "] where (Arrears_due = 0)ORDER BY CLASS ASC", cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then

With DataReport2.Sections("section1").Controls
.Item("a").DataField = rs("name").Name
.Item("b").DataField = rs("Admin_num").Name
.Item("c").DataField = rs("Fees_due").Name
.Item("Class").DataField = rs("Class").Name
End With
DataReport2.Sections("section5").Controls.Item("fxamount").DataField = rs("Fees_due").Name

Set DataReport2.DataSource = rs
DataReport2.Show
Set rs = Nothing

Else
MsgBox "No record Found"
Exit Sub
End If
End Sub

Public Sub fullpayment2()
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [" & MDIForm11.mnusecondterm.Caption & "] where (Arrears_due = 0)", cn, adOpenKeyset, adLockOptimistic
'If rs.RecordCount > 1 Then
If Not rs.EOF Then

With DataReport2.Sections("section1").Controls
.Item("a").DataField = rs("name").Name
.Item("b").DataField = rs("Admin_num").Name
.Item("c").DataField = rs("Fees_due").Name
.Item("Class").DataField = rs("Class").Name
End With
DataReport2.Sections("section5").Controls.Item("fxamount").DataField = rs("Fees_due").Name
Set DataReport2.DataSource = rs
DataReport2.Show
Set rs = Nothing
Else
MsgBox "No record Found"
Exit Sub
End If
End Sub
Public Sub fullpayment3()
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [" & MDIForm11.mnuthirdterm.Caption & "] where (Arrears_due = 0)", cn, adOpenKeyset, adLockOptimistic
'If rs.RecordCount > 1 Then
If Not rs.EOF Then
With DataReport2.Sections("section1").Controls
.Item("a").DataField = rs("name").Name
.Item("b").DataField = rs("Admin_num").Name
.Item("c").DataField = rs("Fees_due").Name
.Item("Class").DataField = rs("Class").Name
End With
DataReport2.Sections("section5").Controls.Item("fxamount").DataField = rs("Fees_due").Name

Set DataReport2.DataSource = rs
DataReport2.Show

Set rs = Nothing
Else
MsgBox "No record Found"
Exit Sub
End If
End Sub

Function AccessShowReport(sAccessDBPath As String, sReportName As String) As String
    Dim oAccess As Object 'Access.Application
    
    On Error GoTo ErrFailed
    AccessShowReport = ""
    'Create Access
    Set oAccess = CreateObject("Access.Application")
    'Open Database
    oAccess.OpenCurrentDatabase sAccessDBPath
    'Open report
    oAccess.DoCmd.OpenReport sReportName, 0 'acViewNormal
    'Show Access Report
    oAccess.Visible = True
    oAccess.CloseCurrentDatabase
    Set oAccess = Nothing
    Exit Function

ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    AccessShowReport = Err.Description
    If oAccess Is Nothing = False Then
        oAccess.CloseCurrentDatabase
        Set oAccess = Nothing
    End If
End Function
Public Sub waecreport()
If rs.State = adStateOpen Then rs.Close

rs.Open ("select * from waec where examyear ='" & frmwaec.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If Not rs.EOF Then
 
With waeclist.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Cand_name").Name
    .Item("text2").DataField = rs("sex").Name
    .Item("text3").DataField = rs("Reg_date").Name
    .Item("text4").DataField = rs("No_Sub_Offered").Name
    .Item("text5").DataField = rs("regno").Name
    .Item("text6").DataField = rs("Amount").Name
    End With
 waeclist.Sections("SECTION4").Controls.Item("Label1").Caption = rs("examyear").Value
 waeclist.Sections("SECTION4").Controls.Item("Label8").Caption = rs("Examtype").Value
     
  waeclist.Sections("SECTION5").Controls.Item("fxtotal").DataField = rs("Amount").Name
    Set waeclist.DataSource = rs
    waeclist.Show
    Else
    MsgBox "No Record Found For The Selected" & "-" & frmwaec.txtdate.Text, vbInformation
    End If
End Sub
Public Sub code2()
If rs.State = adStateOpen Then rs.Close
rs.Open ("select * from Neco where examyear ='" & frmneco.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If rs.EOF Then
 MsgBox "No Student Registered For The Selected " & "," & frmneco.txtdate.Text, vbInformation
 Else
 necoreport.Sections("SECTION4").Controls.Item("Label6").Caption = rs("examyear").Value
 necoreport.Sections("SECTION4").Controls.Item("Label8").Caption = rs("Examtype").Value
With necoreport.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Cand_name").Name
    .Item("text2").DataField = rs("sex").Name
    .Item("text3").DataField = rs("Reg_date").Name
    .Item("text4").DataField = rs("No_Sub_Offered").Name
    .Item("text6").DataField = rs("regno").Name
    .Item("text5").DataField = rs("Amount").Name
         End With
       necoreport.Sections("SECTION5").Controls.Item("fxtotal").DataField = rs("Amount").Name
    Set necoreport.DataSource = rs
    necoreport.Show vbModal
    End If
End Sub
Public Sub code20()
If rs.State = adStateOpen Then rs.Close
rs.Open ("select * from JSCE where examyear ='" & frmjsce.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If rs.EOF Then
 MsgBox "No student Registered For The selected year" & "," & frmjsce.txtdate.Text
 frmjsce.txtdate = ""
frmjsce.txtdate.SetFocus
 Else
 jscereport.Sections("SECTION4").Controls.Item("Label6").Caption = rs("examyear").Value
 jscereport.Sections("SECTION4").Controls.Item("Label8").Caption = rs("examtype").Value
With jscereport.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Cand_name").Name
    .Item("text2").DataField = rs("sex").Name
    .Item("text3").DataField = rs("Reg_date").Name
    .Item("text4").DataField = rs("No_Sub_Offered").Name
    .Item("text5").DataField = rs("regno").Name
    .Item("text6").DataField = rs("Amount").Name
         End With
    jscereport.Sections("SECTION5").Controls.Item("fxtotal").DataField = rs("Amount").Name

    Set jscereport.DataSource = rs
    jscereport.Show vbModal
    End If
End Sub
Public Sub code30()
If rs.State = adStateOpen Then rs.Close
rs.Open ("select * from COMMONENTRANCE where examyear ='" & frmcommon.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If rs.EOF Then
 MsgBox "No student Registered For The selected year" & "," & frmcommon.txtdate.Text
 frmcommon.txtdate.Text = ""
 frmcommon.txtdate.SetFocus
 Else
 commonentra.Sections("SECTION4").Controls.Item("Label6").Caption = rs("examyear").Value
  commonentra.Sections("SECTION4").Controls.Item("Label8").Caption = rs("examtype").Value
With commonentra.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Cand_name").Name
    .Item("text2").DataField = rs("sex").Name
    .Item("text3").DataField = rs("Reg_date").Name
    .Item("text4").DataField = rs("No_Sub_Offered").Name
        .Item("text5").DataField = rs("regno").Name
        .Item("text6").DataField = rs("Amount").Name
         End With
   commonentra.Sections("SECTION5").Controls.Item("fxtotal").DataField = rs("Amount").Name
    Set commonentra.DataSource = rs
     commonentra.Show vbModal
    End If
End Sub
Public Sub MOCK1()
If rs.State = adStateOpen Then rs.Close

rs.Open ("select * from MOCK where examyear ='" & frmmock.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If Not rs.EOF Then
 
With DataReport4.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Cand_name").Name
    .Item("text2").DataField = rs("sex").Name
    .Item("text3").DataField = rs("Reg_date").Name
    .Item("text4").DataField = rs("No_Sub_Offered").Name
    .Item("text5").DataField = rs("regno").Name
    .Item("text6").DataField = rs("Amount").Name
    End With
DataReport4.Sections("SECTION4").Controls.Item("Label1").Caption = rs("examyear").Value
DataReport4.Sections("SECTION4").Controls.Item("Label8").Caption = rs("Examtype").Value
DataReport4.Sections("SECTION5").Controls.Item("fxtotal").DataField = rs("Amount").Name
    Set DataReport4.DataSource = rs
    DataReport4.Show
    Else
    MsgBox "No Record Found For The Selected" & "-" & frmwaec.txtdate.Text, vbInformation
    End If
End Sub

Public Sub graduated()
If rs.State = adStateOpen Then rs.Close

rs.Open ("select * from graduated where year_grad ='" & frmgraduated.txtdate.Text & "'"), cn, adOpenDynamic, adLockOptimistic
 If Not rs.EOF Then
 
With DataReport5.Sections("SECTION1").Controls
    .Item("Text1").DataField = rs("Admin_num").Name
    .Item("text2").DataField = rs("name").Name
    .Item("text3").DataField = rs("sex").Name
    .Item("text4").DataField = rs("date_admited").Name
        End With
DataReport5.Sections("SECTION4").Controls.Item("Label1").Caption = rs("year_grad").Value
    Set DataReport5.DataSource = rs
    DataReport5.Show
    Else
    MsgBox "No Record Found For The Selected" & "-" & frmgraduated.txtdate.Text, vbInformation
    End If
End Sub

Public Sub installmental()
 If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [installmentpay] where Admin_num='" & Form8.Text1 & "'", cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then

With DataReport6.Sections("section1").Controls
'.Item("name").DataField = rs("name").Name
.Item("amt").DataField = rs("Fees_paid").Name
.Item("Arrears_due").DataField = rs("Arrears_due").Name
.Item("pdate").DataField = rs("DATE").Name
.Item("Text1").DataField = rs("receipt_no").Name

End With
'DataReport6.Sections("section5").Controls.Item("fxamount").DataField = rs("Fees_due").Name
DataReport6.Sections("section4").Controls.Item("lname").Caption = rs("name").Value
DataReport6.Sections("section4").Controls.Item("Label6").Caption = rs("Admin_num").Value

Set DataReport6.DataSource = rs
DataReport6.Show
Set rs = Nothing

Else
MsgBox "No Record Found"
Exit Sub
End If
End Sub
