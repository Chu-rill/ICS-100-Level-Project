VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6480
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6480
      TabIndex        =   1
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Serial Number:"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "ID Number:"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Login to Vote"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Empty field not allowed", vbExclamation
Else
    MsgBox ("Login Successful " & vbNewLine & _
    "You can now vote")
    center.Show
    Login.Hide
    
    End If
End Sub
