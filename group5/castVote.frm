VERSION 5.00
Begin VB.Form castVote 
   BackColor       =   &H80000010&
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3720
      Width           =   3615
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Ibrahim Quazim Adebayo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   7
      Top             =   2520
      Width           =   3975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Atiku Abubarkar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Bola Ahmed Tinubu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Vote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Peter Obi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Serial No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cast Vote"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "castVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Dim selectedOption As String
If Option1.Value = True Then
    selectedOption = Option1.Caption
ElseIf Option2.Value = True Then
    selectedOption = Option2.Caption
ElseIf Option3.Value = True Then
    selectedOption = Option3.Caption
ElseIf Option4.Value = True Then
    selectedOption = Option4.Caption
End If

 center.AddDataToFlexGrid Text1.Text, Text2.Text, selectedOption
MsgBox ("Vote Casted")
Text1.Text = ""
Text2.Text = ""


castVote.Hide

End Sub

