VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Student_Entry 
   Caption         =   "Student Data Entry"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form6"
   ScaleHeight     =   10950
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Guardians Personal Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   26
      Top             =   7440
      Width           =   13335
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   8640
         TabIndex        =   33
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   8640
         TabIndex        =   32
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         Height          =   975
         Left            =   2400
         TabIndex        =   31
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   8640
         TabIndex        =   30
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2400
         TabIndex        =   29
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   38
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   35
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox EContactPhone1 
      Height          =   495
      Left            =   9120
      TabIndex        =   22
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emergency Contact Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7080
      TabIndex        =   19
      Top             =   3960
      Width           =   6495
      Begin VB.TextBox EContactPhone2 
         Height          =   495
         Left            =   2040
         TabIndex        =   28
         Top             =   2040
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Text            =   "Contact Relationship"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox EContactName 
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.ComboBox cmbstate 
      Height          =   315
      Left            =   3000
      TabIndex        =   17
      Text            =   "Select State"
      Top             =   7080
      Width           =   3735
   End
   Begin VB.TextBox address 
      Height          =   1095
      Left            =   3000
      TabIndex        =   15
      Top             =   5040
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   10200
      ScaleHeight     =   2475
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox phone 
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   6240
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   4440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Format          =   84017153
      CurrentDate     =   42466
   End
   Begin VB.ComboBox cmblevel 
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Text            =   "Select Level"
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox lname 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox studID 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox fname 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "State of Origin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Student_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
