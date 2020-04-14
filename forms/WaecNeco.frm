VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form waecneco 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Examinations Registration"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   HelpContextID   =   2460
   LinkTopic       =   "Form4"
   ScaleHeight     =   5625
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "WAEC REGISTRATION"
      TabPicture(0)   =   "WaecNeco.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "category"
      Tab(0).Control(1)=   "Option2"
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(3)=   "cname"
      Tab(0).Control(4)=   "SubNo"
      Tab(0).Control(5)=   "examyear"
      Tab(0).Control(6)=   "regno"
      Tab(0).Control(7)=   "examtype"
      Tab(0).Control(8)=   "amt"
      Tab(0).Control(9)=   "sex"
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(12)=   "Command5"
      Tab(0).Control(13)=   "Adodc1"
      Tab(0).Control(14)=   "datee"
      Tab(0).Control(15)=   "Command4"
      Tab(0).Control(16)=   "Command6"
      Tab(0).Control(17)=   "Command7"
      Tab(0).Control(18)=   "Label28"
      Tab(0).Control(19)=   "Label27"
      Tab(0).Control(20)=   "Label26"
      Tab(0).Control(21)=   "Label33"
      Tab(0).Control(22)=   "Label5"
      Tab(0).Control(23)=   "Label15"
      Tab(0).Control(24)=   "Label10"
      Tab(0).Control(25)=   "Label6"
      Tab(0).Control(26)=   "Label1"
      Tab(0).Control(27)=   "Label2"
      Tab(0).Control(28)=   "Label3"
      Tab(0).Control(29)=   "Label8"
      Tab(0).Control(30)=   "Line1"
      Tab(0).Control(31)=   "Line2"
      Tab(0).Control(32)=   "Line3"
      Tab(0).Control(33)=   "Line4"
      Tab(0).Control(34)=   "Label12"
      Tab(0).Control(35)=   "Label13"
      Tab(0).Control(36)=   "Label9"
      Tab(0).Control(37)=   "Label4"
      Tab(0).Control(38)=   "Label7"
      Tab(0).Control(39)=   "Label14"
      Tab(0).Control(40)=   "Label16"
      Tab(0).Control(41)=   "Label11"
      Tab(0).Control(42)=   "Label34"
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "NECO REGISTRATION"
      TabPicture(1)   =   "WaecNeco.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label23"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label25"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Line5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Line6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label29"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label30"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label31"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command12"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command11"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Command10"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "datee1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "examyear1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "regno1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "examtype1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "amt1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "SubNo1"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "sex1"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Command1"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command8"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command9"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cname1"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Option3"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Option4"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "category1"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).ControlCount=   33
      Begin VB.TextBox category1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2520
         Width           =   3135
      End
      Begin VB.OptionButton Option4 
         Height          =   375
         Left            =   1920
         TabIndex        =   63
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         Height          =   375
         Left            =   2640
         TabIndex        =   62
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox category 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2640
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         Height          =   375
         Left            =   -72360
         TabIndex        =   59
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Left            =   -73080
         TabIndex        =   56
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox cname1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   3000
         TabIndex        =   55
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox cname 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   -72000
         TabIndex        =   1
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox SubNo 
         Height          =   315
         ItemData        =   "WaecNeco.frx":0038
         Left            =   -72000
         List            =   "WaecNeco.frx":003A
         TabIndex        =   6
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ComboBox sex1 
         Height          =   315
         ItemData        =   "WaecNeco.frx":003C
         Left            =   3000
         List            =   "WaecNeco.frx":0046
         TabIndex        =   38
         Top             =   1920
         Width           =   3135
      End
      Begin VB.ComboBox SubNo1 
         Height          =   315
         ItemData        =   "WaecNeco.frx":0058
         Left            =   3000
         List            =   "WaecNeco.frx":005A
         TabIndex        =   37
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox amt1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   36
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox examtype1 
         Height          =   315
         ItemData        =   "WaecNeco.frx":005C
         Left            =   8160
         List            =   "WaecNeco.frx":0066
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "NECO"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox regno1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   8160
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox examyear1 
         Height          =   315
         Left            =   8160
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox examyear 
         Height          =   315
         Left            =   -66840
         TabIndex        =   0
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox regno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   -66840
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox examtype 
         Height          =   315
         ItemData        =   "WaecNeco.frx":0076
         Left            =   -66840
         List            =   "WaecNeco.frx":0080
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "WAEC"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox amt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
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
         Height          =   495
         Left            =   -66840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ComboBox sex 
         Height          =   315
         ItemData        =   "WaecNeco.frx":0090
         Left            =   -72000
         List            =   "WaecNeco.frx":009A
         TabIndex        =   3
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4080
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -74040
         Top             =   3720
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\COURTINFOSYS\database\courtdatabase.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\COURTINFOSYS\database\courtdatabase.mdb;Persist Security Info=False"
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
      Begin MSComCtl2.DTPicker datee 
         Height          =   375
         Left            =   -66840
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   105250817
         CurrentDate     =   40058
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   -72000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   -70440
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   -69000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker datee1 
         Height          =   375
         Left            =   8160
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   105250817
         CurrentDate     =   40058
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H80000008&
         Caption         =   "Command2"
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Candidate Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "External"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   65
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Internal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   64
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Internal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72360
         TabIndex        =   60
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "External"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         TabIndex        =   58
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "Candidate Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   57
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   9840
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line7 
         X1              =   9840
         X2              =   9840
         Y1              =   480
         Y2              =   4440
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   120
         Y1              =   480
         Y2              =   4440
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   9840
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         TabIndex        =   54
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   53
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Candidate Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year of Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         TabIndex        =   50
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         TabIndex        =   49
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Subject Sit For:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   48
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         TabIndex        =   47
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         TabIndex        =   46
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68760
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68760
         TabIndex        =   29
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Subject Sit For:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   27
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68760
         TabIndex        =   25
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Year of Exams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68760
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   21
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Candidate Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   16
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68760
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   -74880
         X2              =   -74880
         Y1              =   720
         Y2              =   4680
      End
      Begin VB.Line Line2 
         X1              =   -65160
         X2              =   -65160
         Y1              =   720
         Y2              =   4680
      End
      Begin VB.Line Line3 
         X1              =   -74880
         X2              =   -65160
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line4 
         X1              =   -74880
         X2              =   -65160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74640
         TabIndex        =   20
         Top             =   2760
         Width           =   2385
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68640
         TabIndex        =   19
         Top             =   2160
         Width           =   1785
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74640
         TabIndex        =   18
         Top             =   1560
         Width           =   2385
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74640
         TabIndex        =   22
         Top             =   2160
         Width           =   2385
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68640
         TabIndex        =   24
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68640
         TabIndex        =   26
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74640
         TabIndex        =   28
         Top             =   3480
         Width           =   2385
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68640
         TabIndex        =   30
         Top             =   3600
         Width           =   1785
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68640
         TabIndex        =   32
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C1AB7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   2415
      End
   End
End
Attribute VB_Name = "waecneco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_LostFocus()
If Check1.Value = True Then
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic

Amt.Text = rs!Ewaecfee
ElseIf Check1.Value = False Then

Amt.Text = rs!waecfee
End If
'RS.Close
'Set RS = Nothing
End Sub

Private Sub cname_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub cname1_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub Command1_Click()
If cname1.Text = "" Or _
sex1.Text = "" Or _
examtype1.Text = "" Or _
regno1.Text = "" Or _
SubNo1.Text = "" Or _
category1.Text = "" Or _
amt1.Text = "" Or _
examyear1.Text = "" Or _
Option3.Value = False Or _
Option4.Value = False Then
MsgBox "some Fields are missing"
Exit Sub
End If
sql = "select * from [neco]"
Set rs = New ADODB.Recordset
With rs
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
End With
With rs
.AddNew
!cand_name = cname1.Text
!SEX = sex1.Text
!examtype = examtype1.Text
!regno = regno1.Text
!Reg_date = datee1.Value
!No_Sub_Offered = SubNo1.Text
!category = category1.Text
!amount = amt1.Text
!examyear = examyear1.Text
MsgBox "Student Sucessfully Registered"
End With
rs.Update
Call clear1
rs.Update
rs.Close
Set rs = Nothing

End Sub

Private Sub Command2_Click()
If cname.Text = "" Or _
SEX.Text = "" Or _
examtype.Text = "" Or _
regno.Text = "" Or _
SubNo.Text = "" Or _
category.Text = "" Or _
Amt.Text = "" Or _
examyear.Text = "" Or _
Option1.Value = False Or _
Option2.Value = False Then
MsgBox "some Fields are missing"
Exit Sub
End If


sql = "select * from [waec]"
Set rs = New ADODB.Recordset
With rs
If .State = adStateOpen Then .Close
.Open sql, cn, adOpenDynamic, adLockOptimistic
End With
With rs
.AddNew
!cand_name = cname.Text
!SEX = SEX.Text
!examtype = examtype.Text
!regno = regno.Text
!Reg_date = datee.Value
!No_Sub_Offered = SubNo.Text
!category = category.Text
!amount = Amt.Text
!examyear = examyear.Text
MsgBox "Student Sucessfully Registered"
End With
rs.Update
Call clear
rs.Update
rs.Close
Set rs = Nothing
End Sub

Public Sub clear()
cname.Text = ""
SEX.Text = ""
examtype.Text = ""
regno.Text = ""
'datee.Text = ""
SubNo.Text = ""
category.Text = ""
Amt.Text = ""
examyear.Text = ""
Option1.Value = False
Option2.Value = False

End Sub

Public Sub clear1()
cname1.Text = ""
sex1.Text = ""
examtype1.Text = ""
regno.Text = ""
'datee.Text = ""
SubNo1.Text = ""
category1.Text = ""
amt1.Text = ""
examyear1.Text = ""
Option3.Value = False
Option4.Value = False
End Sub

Private Sub Command3_Click()
Call clear
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command8_Click()
Call clear1
End Sub

Private Sub Command9_Click()
Unload Me
End Sub

Private Sub examyear_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub examyear1_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
For i = 1990 To 2030
examyear1.AddItem i
examyear.AddItem i
Next i
For i = 1 To 7
SubNo.AddItem i
SubNo1.AddItem i

Next i
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
category.Text = "External Student"
Option2.Value = False
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
Amt.Text = rs!Ewaecfee
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
category.Text = "Internal Student"
Option1.Value = False
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
Amt.Text = rs!waecfee
End If
End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
category1.Text = "Internal Student"
Option4.Value = False
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
amt1.Text = rs!neco
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
category1.Text = "External Student"
Option3.Value = False
If rs.State = adStateOpen Then rs.Close
rs.Open "select * from[examfee]", cn, adOpenDynamic, adLockOptimistic
amt1.Text = rs!eneco
End If
End Sub

Private Sub SEX_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub sex1_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[A-Z,a-z]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "NUMBERS ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub SubNo_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub SubNo_LostFocus()
If Val(SubNo) > 9 Then
MsgBox "Maximum Number of subject is Nine (9)"
SubNo = ""
SubNo.SetFocus
Exit Sub
End If
End Sub

Private Sub SubNo1_KeyPress(KeyAscii As Integer)
Dim res As Boolean
res = Chr(KeyAscii) Like "[0-9]"
If res = False And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 13 Then
MsgBox "ALPHABET ARE NOT ALLOWED "
KeyAscii = 0
End If
End Sub

Private Sub SubNo1_LostFocus()
If Val(SubNo1) > 9 Then
MsgBox "Maximum Number of subject is Nine (9)"
SubNo1 = ""
SubNo1.SetFocus
Exit Sub
End If
End Sub
