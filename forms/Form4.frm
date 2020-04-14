VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   HelpContextID   =   2760
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'If rs.State = adStateOpen Then rs.Close
RS.Open "select * from[table1]", cn, adOpenDynamic, adLockOptimistic
RS.AddNew
RS!SEX = Text1.Text
RS!Name = Text2.Text
RS.Update
RS.Close
Set RS = Nothing
Text1 = ""
Text2 = ""
End Sub
