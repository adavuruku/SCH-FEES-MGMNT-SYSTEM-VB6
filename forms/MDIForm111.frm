VERSION 5.00
Begin VB.MDIForm MDIForm11 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0FF&
   Caption         =   "School Fees Collection Management System"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4680
   Icon            =   "MDIForm111.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   10920
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   6
         Top             =   240
         Width           =   3615
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   9
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Session"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2760
            TabIndex        =   7
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Image Image10 
         Height          =   735
         Left            =   720
         Picture         =   "MDIForm111.frx":0442
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   735
      End
      Begin VB.Image Image9 
         Height          =   375
         Left            =   14760
         Picture         =   "MDIForm111.frx":0A28
         Stretch         =   -1  'True
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "(F1)"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   8400
         Width           =   615
      End
      Begin VB.Image Image7 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":0F59
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Final Exam Registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About the System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   375
         TabIndex        =   3
         Top             =   7200
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrative task"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   4800
         Width           =   1620
      End
      Begin VB.Image Image8 
         Height          =   4695
         Left            =   10440
         Picture         =   "MDIForm111.frx":1875
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Fee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   720
         TabIndex        =   1
         Top             =   2640
         Width           =   690
      End
      Begin VB.Image Image6 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":33D1
         Top             =   6480
         Width           =   810
      End
      Begin VB.Image Image5 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":3801
         Top             =   3960
         Width           =   810
      End
      Begin VB.Image Image4 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":3E66
         Top             =   8880
         Width           =   810
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":4397
         ToolTipText     =   "Fee payment"
         Top             =   2640
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   720
         Picture         =   "MDIForm111.frx":48E6
         Top             =   960
         Width           =   810
      End
      Begin VB.Image Image1 
         Height          =   10935
         Left            =   0
         Picture         =   "MDIForm111.frx":4E7B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15855
      End
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
      Begin VB.Menu instalpayS 
         Caption         =   "Installmental payment"
      End
      Begin VB.Menu kk 
         Caption         =   "-"
      End
   End
   Begin VB.Menu axz 
      Caption         =   ""
   End
   Begin VB.Menu cc 
      Caption         =   ""
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnudetails 
         Caption         =   "Student payment Details"
      End
      Begin VB.Menu dash101 
         Caption         =   "-"
      End
      Begin VB.Menu dep 
         Caption         =   "List of Deptors"
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnufullpayment 
         Caption         =   "Complete Payment"
         Begin VB.Menu mnuist 
            Caption         =   "1stterm"
         End
         Begin VB.Menu mnudashs 
            Caption         =   "-"
         End
         Begin VB.Menu mnusecondterm 
            Caption         =   "2ndTerm"
         End
         Begin VB.Menu HHJJJJ 
            Caption         =   "-"
         End
         Begin VB.Menu mnuthirdterm 
            Caption         =   "3rdterm"
         End
      End
      Begin VB.Menu mnudass44 
         Caption         =   "-"
      End
      Begin VB.Menu mnugra 
         Caption         =   "Graduation Report"
         Begin VB.Menu mnuprisch 
            Caption         =   "Primary school"
         End
         Begin VB.Menu mnusecsch 
            Caption         =   "Secondary school"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnucomp 
         Caption         =   "Comprehensive Report"
         Begin VB.Menu dash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnufirst 
            Caption         =   "1stterm"
         End
         Begin VB.Menu mnusecondterms 
            Caption         =   "2ndTerm"
         End
         Begin VB.Menu mnuthird 
            Caption         =   "3rdTerm"
         End
      End
      Begin VB.Menu HTYR 
         Caption         =   "-"
      End
      Begin VB.Menu finalexamreport 
         Caption         =   "FINAL EXAMINATION REPORT"
         Begin VB.Menu mnuwaec 
            Caption         =   "Registered Waec Student"
            Visible         =   0   'False
         End
         Begin VB.Menu mnucommonentrance 
            Caption         =   "Registered Commonrntrance Student"
         End
         Begin VB.Menu mnudash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuneco 
            Caption         =   "Registered Neco Student"
            Visible         =   0   'False
         End
         Begin VB.Menu dash5 
            Caption         =   "-"
         End
         Begin VB.Menu mnujsce 
            Caption         =   "Registered Jsce Student"
            Visible         =   0   'False
         End
         Begin VB.Menu dash6 
            Caption         =   "-"
         End
         Begin VB.Menu mnumock 
            Caption         =   "Registered Mock Student"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudashQR 
            Caption         =   "-"
         End
      End
      Begin VB.Menu WWQQQ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuletter 
         Caption         =   "Generate Fee Reminder Letter"
      End
   End
   Begin VB.Menu mnueua 
      Caption         =   "Administrative Operation"
      Begin VB.Menu mnucre 
         Caption         =   "Create user account"
      End
      Begin VB.Menu mnuedit 
         Caption         =   "edit user account"
      End
      Begin VB.Menu mnbkp 
         Caption         =   "Backup Entire Databse"
      End
      Begin VB.Menu mnusession 
         Caption         =   "Create new Session"
      End
      Begin VB.Menu mnueditfee 
         Caption         =   "Edit School Fees"
      End
      Begin VB.Menu gh 
         Caption         =   "Edit Exam Fee"
      End
      Begin VB.Menu UUUUUUU 
         Caption         =   "-"
      End
   End
   Begin VB.Menu finalexamreg 
      Caption         =   "Examination Registration"
      Begin VB.Menu MNENTRANCE 
         Caption         =   "COMMON ENTRANCE"
      End
      Begin VB.Menu mnujcese 
         Caption         =   "JSCE EXAM"
         Visible         =   0   'False
      End
      Begin VB.Menu mnud 
         Caption         =   "WAEC & NECO EXAMS"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu enda 
      Caption         =   "End Application"
   End
End
Attribute VB_Name = "MDIForm11"
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
  'Unload Me
  'MDIForm11.Show
  Exit Sub
  Case vbYes
  End
  End Select
  
End Sub

Private Sub fisrtstterm_Click()

End Sub

Private Sub gh_Click()
examfee.Show
End Sub




Private Sub Image10_Click()
frmDBBackUp.Show
End Sub

Private Sub Image2_Click()
ExamReg.Show
End Sub

Private Sub Image4_Click()
Dim a As String
a = MsgBox("Do you want to Quit application", vbYesNo, "CONFIRMATION")
Select Case a
  Case vbNo
  Exit Sub
  'Unload Me
  'MDIForm11.Show
  Case vbYes
  MsgBox "Thanks for Using the Software", vbInformation
  End
  End Select

End Sub

Private Sub Image5_Click()
adminopt.Show
End Sub

Private Sub Image6_Click()
frmAbout.Show
End Sub

Private Sub Image9_Click()
End
End Sub

Private Sub instalpayS_Click()
instalpay.Show
End Sub

Private Sub Label1_Click()
Form5.Show
End Sub

Private Sub MDIForm_Load()
Image1.Height = Picture1.Height
Image1.Width = Screen.Width
Image9.Left = Screen.Width - Image9.Width
Picture2.Left = Screen.Width - (Picture2.Width + 2 * Image9.Width + 10)
Image8.Left = Screen.Width / 2 - (Image8.Width / 2)
Image8.Top = Screen.Height / 2 - (Image8.Height / 2)
mnFirst.Enabled = False
mnsecond.Enabled = False
mnueditfee.Enabled = False
instalpayS.Enabled = False
mnureport.Enabled = False
mnueua.Enabled = False
mnueua.Enabled = False
mnueditfee.Enabled = False
Label1.Enabled = False
Image5.Enabled = False
Image3.Enabled = False
Image2.Enabled = False
gh.Enabled = False
 mnthird.Enabled = False
'mnFirst.Enabled = False
 mnujcese.Enabled = False
 mnud.Enabled = False
 MNENTRANCE.Enabled = False
 'finalexamreport.Enabled = False
 we.Enabled = False
 finalexamreg.Enabled = False
End Sub

Private Sub mnbkp_Click()
frmDBBackUp.Show
End Sub

Private Sub MNENTRANCE_Click()
commonentrance.Show
End Sub

Private Sub mnthird_Click()
Thirdterm.Show
End Sub

Private Sub mnucommonentrance_Click()
frmcommon.Show
End Sub

Private Sub mnucre_Click()
form1.Show
End Sub

Private Sub MNUD_Click()
waecneco.Show
End Sub

Private Sub mnudetails_Click()
Form8.Show
End Sub

Private Sub mnuedit_Click()
Form2.Show
End Sub

Private Sub mnueditfee_Click()
editfee.Show
End Sub
Private Sub mnFirst_Click()
StudentType.Show
'firsterm.Show
End Sub

Private Sub mnsecond_Click()
seconterm.Show
End Sub



Private Sub mnufirst_Click()
Call reportdelete
Call computerreport
Call prm1
Call prm3
Call nur3
Call Nur2
Call prm4
Call prm5
Call prm6
Call prm2
Call comprehensive
End Sub

Private Sub mnuist_Click()
Call fullpayment


End Sub

Private Sub mnujcese_Click()
JSCE.Show
End Sub

Private Sub mnujsce_Click()
frmjsce.Show
End Sub

Private Sub mnuletter_Click()
generatefeereminder.Show
End Sub

Private Sub mnumock_Click()
frmmock.Show
End Sub

Private Sub mnuneco_Click()
frmneco.Show
End Sub

Private Sub mnuprisch_Click()
frmgraduated.Show
End Sub

Private Sub mnusecondterm_Click()
Call fullpayment2
End Sub

Private Sub mnusecondterms_Click()
Call reportdelete1
Call computerreport2
Call prm11
Call prm33
Call nur33
Call Nur22
Call prm44
Call prm55
Call prm66
Call prm22
Call comprehensive2
End Sub

Private Sub mnusession_Click()
createsession.Show
End Sub

Private Sub mnuthird_Click()
Call reportdelete2
Call computerreport22
Call prm111
Call prm333
Call nur333
Call Nur222
Call prm444
Call prm555
Call prm666
Call prm222
Call comprehensive22
End Sub

Private Sub mnuthirdterm_Click()
Call fullpayment3
End Sub

Private Sub mnuwaec_Click()
frmwaec.Show
End Sub

