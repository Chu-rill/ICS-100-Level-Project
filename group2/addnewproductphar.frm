VERSION 5.00
Begin VB.Form addnewproductphar 
   BackColor       =   &H80000004&
   Caption         =   "Add New Product"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Add"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   2055
   End
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
      Height          =   405
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   5
      Top             =   5640
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   4
      Top             =   4920
      Width           =   3375
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
      Height          =   1125
      Left            =   3720
      TabIndex        =   3
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Add New Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   960
      Width           =   2520
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000004&
      Caption         =   "Product ID:"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Description:"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "addnewproductphar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
addnewproductphar.Hide

End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Or Text5.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "Input field can not be empty", vbExclamation
    Else
      MsgBox ("Item Added")
      pharmacy.AddDataToFlexGrid Text1.Text, Text5.Text, Text2.Text, Val(Text3.Text), Val(Text4.Text)
  
      addnewproductphar.Hide
    
        Text1.Text = ""
        Text5.Text = ""
        Text2.Text = ""
         Text3.Text = ""
        Text4.Text = ""
    End If

End Sub



Private Function GenerateRandomID() As String
    Dim id As String
    Dim i As Integer
    
    Randomize ' Initialize random number generator
    
    ' Generate 10-character ID
    For i = 1 To 10
        ' Append a random uppercase letter or digit to the ID
        id = id & Chr(Asc("A") + Int(Rnd * 26)) ' Uppercase letter (A-Z)
        ' Alternatively, you can use: id = id & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomID = id
End Function
Private Sub Form_Load()
Dim id As String

id = GenerateRandomID()

Text5.Text = id
End Sub
