VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User's Login"
   ClientHeight    =   2640
   ClientLeft      =   2835
   ClientTop       =   3435
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1559.799
   ScaleMode       =   0  'User
   ScaleWidth      =   5859.022
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   -75
         Top             =   2640
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   330
         Left            =   4320
         TabIndex        =   7
         Top             =   1740
         Width           =   945
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   330
         Left            =   3360
         TabIndex        =   6
         Top             =   1740
         Width           =   930
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3375
         PasswordChar    =   "X"
         TabIndex        =   4
         Top             =   1320
         Width           =   1890
      End
      Begin VB.ComboBox cbUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   930
         Width           =   2670
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © Boggyman TM  2006 - 2007  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   2295
         TabIndex        =   8
         Top             =   2355
         Width           =   3045
      End
      Begin VB.Image Image2 
         Height          =   75
         Left            =   30
         Picture         =   "frmLogin1.frx":0000
         Stretch         =   -1  'True
         Top             =   2550
         Width           =   6975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2340
         TabIndex        =   5
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2340
         TabIndex        =   3
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   2490
         Left            =   60
         Picture         =   "frmLogin1.frx":2060
         Top             =   45
         Width           =   2145
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select your username and enter your password in the space provided bellow."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   2340
         TabIndex        =   1
         Top             =   195
         Width           =   3690
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   28.168
      X2              =   3704.141
      Y1              =   2313.111
      Y2              =   2313.111
   End
   Begin VB.Line Line1 
      X1              =   28.168
      X2              =   3746.394
      Y1              =   2384.011
      Y2              =   2384.011
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer
Public LoginSucceeded As Boolean
Private Sub cbUser_Click()
    txtpassword.Text = ""
End Sub

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    'Unload Me
    End
End Sub
Private Sub cmdOK_Click()
Dim t As String
Dim rs As New ADODB.Recordset
Dim cConn As String
Set rs = New ADODB.Recordset

  With rs
    .Open "SELECT * FROM dbLogin WHERE Username='" & cbUser.Text & "' And Password='" & txtpassword.Text & "'", _
     cConnect, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then
            MsgBox "Unrecognized User", vbOKOnly + vbCritical, _
            "Warning: End User"
            txtpassword.SetFocus
            SendKeys "{Home}+{End}"
            ctr = ctr + 1
        If ctr = 3 Then
                MsgBox "You have exceeded the number of attempts"
                End
        End If
        Else
          Unload Me
          fMainForm.lblUser.Caption = rs!Username
          frmMessage.Show 'vbModal
          End If
        End With
End Sub
'
Private Sub Form_Load()
Dim vntemp As Variant
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset

  With rs
    .Open "SELECT * FROM dbLogin", cConnect, adOpenDynamic, adLockOptimistic
  End With
  
 '*******Combolist********
 
    Do While Not rs.EOF
     vntemp = rs!Username
      If IsNull(vntemp) Then vntemp = ""
       cbUser.AddItem CStr(vntemp)
        rs.MoveNext
    Loop
    txtpassword.MaxLength = 20 'Text field max. number of character input.
End Sub


Private Sub Timer1_Timer()
If Image2.Left > 5225 Then
    Image2.Left = -3945
    Else
    Image2.Left = Image2.Left + 100
    End If
End Sub


