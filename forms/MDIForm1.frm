VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1455
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   10080
      Left            =   0
      ScaleHeight     =   10020
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   615
      Width           =   15240
      Begin VB.Image Image1 
         Height          =   3435
         Left            =   5520
         Picture         =   "MDIForm1.frx":0000
         Top             =   2160
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
   End
   Begin VB.Menu we 
      Caption         =   "Payment"
      WindowList      =   -1  'True
      Begin VB.Menu mnFirst 
         Caption         =   "First Term fee"
      End
      Begin VB.Menu mnsecond 
         Caption         =   "Second Term Fees"
      End
      Begin VB.Menu mnthird 
         Caption         =   "Third Term Fees"
      End
   End
   Begin VB.Menu gh 
      Caption         =   ""
   End
   Begin VB.Menu mnueditfee 
      Caption         =   "Edit School Fees"
   End
   Begin VB.Menu kk 
      Caption         =   ""
   End
   Begin VB.Menu axz 
      Caption         =   ""
   End
   Begin VB.Menu instalpayS 
      Caption         =   "Installmental payment"
   End
   Begin VB.Menu cc 
      Caption         =   ""
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu dep 
         Caption         =   "List of Deptors"
      End
   End
   Begin VB.Menu mnueua 
      Caption         =   "Edit User Account"
      Begin VB.Menu mnucre 
         Caption         =   "Create user account"
      End
      Begin VB.Menu mnuedit 
         Caption         =   "edit user account"
      End
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
      Begin VB.Menu enda 
         Caption         =   "End Application"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub dep_Click()
'sql = "select admin_num,arrears_due from 2009_1stterm where (arrears_due > 0)"
filter.Show

End Sub

Private Sub enda_Click()
Dim a As String
a = MsgBox("do you want to end application", vbYesNo, "CONFIRMATION")
Select Case a
  Case vbNo
  Unload Me
  MDIForm1.Show
  Case vbYes
  End
  End Select
  
End Sub

Private Sub instalpayS_Click()
instalpay.Show
End Sub

Private Sub MDIForm_Load()
mnFirst.Enabled = False
mnsecond.Enabled = False
mnueditfee.Enabled = False
instalpayS.Enabled = False
mnureport.Enabled = False
mnueua.Enabled = False
mnueua.Enabled = False

End Sub

Private Sub mnthird_Click()
Thirdterm.Show
End Sub

Private Sub mnucre_Click()
Form1.Show
End Sub

Private Sub mnuedit_Click()
Form2.Show
End Sub

Private Sub mnueditfee_Click()
editfee.Show
End Sub
Private Sub mnFirst_Click()
Form3.Show
'firsterm.Show
End Sub

Private Sub mnsecond_Click()
seconterm.Show
End Sub



