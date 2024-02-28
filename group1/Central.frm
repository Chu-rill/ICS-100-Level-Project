VERSION 5.00
Begin VB.Form Central 
   BackColor       =   &H00400000&
   Caption         =   "Central"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Bill Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Customer List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "Central"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
customerGrid.Show

End Sub

Private Sub Command2_Click()
billGrid.Show

End Sub

