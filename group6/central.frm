VERSION 5.00
Begin VB.Form central 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Register Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton C 
      Caption         =   "Existing Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "central"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C_Click()
existing.Show

End Sub

Private Sub Command1_Click()
Register.Show

End Sub

