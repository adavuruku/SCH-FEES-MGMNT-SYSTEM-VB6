Attribute VB_Name = "Module2"
Public Sub comprehensive()
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [Reportable] order by class asc", cn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
MsgBox ("Record Not Found")
Exit Sub
End If
With DataReport3.Sections("section1").Controls
.Item("class").DataField = rs("class").Name
.Item("RollNo").DataField = rs("RollNo").Name
.Item("No_PaidFull").DataField = rs("NoPaid").Name
.Item("No_NotpaidFull").DataField = rs("No_NotPaid").Name
.Item("TFeePaid").DataField = rs("TFeePaid").Name
.Item("TArrearsDue").DataField = rs("TArrearsDue").Name
End With
DataReport3.Sections("section5").Controls.Item("fxfeepaid").DataField = rs("TFeePaid").Name
DataReport3.Sections("section5").Controls.Item("fxArreardue").DataField = rs("TArrearsDue").Name
Set DataReport3.DataSource = rs
DataReport3.Show
Set rs = Nothing
'Set db = Nothing
End Sub
Public Sub computerreport()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "NUR1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur1"
End With
RS1.Update
handler:
errodata
'If Err.Number = 3021 Then
'MsgBox "NO RECORD FOUND"
'Exit Sub

End Sub

Public Sub prm1()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim1"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm2()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm3()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM4" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm4()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM4" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim4"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm5()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM5" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim5"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm6()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "PRIM6" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim6"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub Nur2()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "NUR2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub nur3()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnufirst.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "NUR3" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub reportdelete()
Dim sql2 As String

sql2 = "delete*from Reportable"
cn.Execute sql2

End Sub
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'comprehensive report for second term
Public Sub comprehensive2()
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [Reportable]order by class ", cn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
MsgBox ("Record Not Found")
Exit Sub
End If
With DataReport3.Sections("section1").Controls
.Item("class").DataField = rs("class").Name
.Item("RollNo").DataField = rs("RollNo").Name
.Item("No_PaidFull").DataField = rs("NoPaid").Name
.Item("No_NotpaidFull").DataField = rs("No_NotPaid").Name
.Item("TFeePaid").DataField = rs("TFeePaid").Name
.Item("TArrearsDue").DataField = rs("TArrearsDue").Name
End With
DataReport3.Sections("section5").Controls.Item("fxfeepaid").DataField = rs("TFeePaid").Name
DataReport3.Sections("section5").Controls.Item("fxArreardue").DataField = rs("TArrearsDue").Name
Set DataReport3.DataSource = rs
DataReport3.Show
Set rs = Nothing
'Set db = Nothing
End Sub
Public Sub computerreport2()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]  ", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur1"
End With
RS1.Update
handler:
errodata
End Sub

Public Sub prm11()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim1"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm22()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm33()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim3" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm44()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim4" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim4"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm55()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim5" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim5"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm66()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim6" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim6"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub Nur22()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub nur33()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnusecondterms.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur3" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub reportdelete1()
Dim sql2 As String

sql2 = "delete*from Reportable"
cn.Execute sql2

End Sub
'*****************************************************************************
'comprehensive report for third term
Public Sub comprehensive22()
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from [Reportable]order by class ", cn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
MsgBox ("Record Not Found")
Exit Sub
End If
With DataReport3.Sections("section1").Controls
.Item("class").DataField = rs("class").Name
.Item("RollNo").DataField = rs("RollNo").Name
.Item("No_PaidFull").DataField = rs("NoPaid").Name
.Item("No_NotpaidFull").DataField = rs("No_NotPaid").Name
.Item("TFeePaid").DataField = rs("TFeePaid").Name
.Item("TArrearsDue").DataField = rs("TArrearsDue").Name
End With
DataReport3.Sections("section5").Controls.Item("fxfeepaid").DataField = rs("TFeePaid").Name
DataReport3.Sections("section5").Controls.Item("fxArreardue").DataField = rs("TArrearsDue").Name
Set DataReport3.DataSource = rs
DataReport3.Show
Set rs = Nothing
'Set db = Nothing
End Sub
Public Sub computerreport22()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]  ", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur1"
End With
RS1.Update
handler:
errodata
End Sub

Public Sub prm111()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim1" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim1"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm222()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm333()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim3" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm444()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim4" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim4"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm555()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim5" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim5"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub prm666()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "prim6" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "prim6"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub Nur222()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur2" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur2"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub nur333()
On Error GoTo handler:
Dim nclass As Integer
Dim Tpaid As Double
Dim Ndue As Integer
Dim TDue As Double
Dim Nopaid As Integer
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[" & MDIForm11.mnuthird.Caption & "]", cn, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
If rs!CLASS = "nur3" Then
nclass = nclass + 1
Tpaid = Tpaid + Val(rs![Fees_paid])
If Val(rs![arrears_due]) > 0 Then
Ndue = Ndue + 1
TDue = TDue + Val(rs![arrears_due])
Else
Nopaid = Nopaid + 1
End If
End If
rs.MoveNext
Loop
sql = "select * from [Reportable]"
Set RS1 = New ADODB.Recordset
With RS1
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
RS1.AddNew
RS1!RollNo = nclass
RS1!Nopaid = Nopaid
RS1!TfeePaid = Tpaid
RS1!TarrearsDue = TDue
RS1!No_NotPaid = Ndue
RS1!CLASS = "nur3"
End With
RS1.Update
handler:
errodata
End Sub
Public Sub reportdelete2()
Dim sql2 As String
sql2 = "delete*from Reportable"
cn.Execute sql2

End Sub
Public Sub errodata()
If Err.Number = 3021 Then
'MsgBox "NO RECORD FOUND"
Exit Sub
End If
End Sub
