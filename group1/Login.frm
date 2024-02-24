VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00800080&
   Caption         =   "Login"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "&"
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Login"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "UserName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox ("Login Successsful")
Central.Show
Login.Hide

End Sub

