VERSION 5.00
Begin VB.Form removephar 
   BackColor       =   &H8000000A&
   Caption         =   "Remove Product"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   6
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   6000
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "Product ID:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Quantity:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "removephar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim name As String
Dim quantity As Double
Dim ID As String

name = Text1.Text
quantity = Val(Text3.Text)
ID = Text5.Text



If Text1.Text = "" Or Text3.Text = "" Or Text5.Text = "" Then
    MsgBox "Input field can not be empty", vbExclamation
Else
    MsgBox ("Product Removed" & vbNewLine & _
    name & vbNewLine & _
    quantity & vbNewLine & _
    ID)
    removephar.Hide

    Text1.Text = ""
    Text3.Text = ""
    Text5.Text = ""
End If
End Sub
