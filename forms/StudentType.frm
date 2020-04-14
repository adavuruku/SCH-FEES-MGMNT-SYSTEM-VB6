VERSION 5.00
Begin VB.Form StudentType 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Option"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000003&
         Caption         =   "Returning Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "New Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   4440
         Picture         =   "StudentType.frx":0000
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   570
      End
   End
End
Attribute VB_Name = "StudentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
newterm1.Show
Unload Me
End Sub

Private Sub Command2_Click()
term1.Show
Unload Me
End Sub
