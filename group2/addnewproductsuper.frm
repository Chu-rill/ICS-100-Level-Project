VERSION 5.00
Begin VB.Form addnewproductsuper 
   BackColor       =   &H00404000&
   Caption         =   "Add New Product"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   2295
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
      Left            =   5400
      TabIndex        =   9
      Top             =   5400
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
      Left            =   5400
      TabIndex        =   7
      Top             =   4680
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
      Height          =   1365
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
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
      Left            =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
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
      Left            =   5400
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   360
      Width           =   2520
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      Caption         =   "Price:"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3840
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "addnewproductsuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
addnewproductsuper.Hide


End Sub

Private Sub Command2_Click()
       If Text1.Text = "" Or Text5.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "Input field can not be empty", vbExclamation
    Else
     ' MsgBox ("Item Added")
     MsgBox ("Item Added " & vbNewLine & _
    "Product Name: " & Text1.Text & vbNewLine & _
    "Product ID: " & Text5.Text & vbNewLine & _
    "Description: " & Text2.Text & vbNewLine & _
    "Quantity: " & Text3.Text & vbNewLine & _
    "Price: " & Text4.Text)
      'supermarket.AddDataToFlexGrid Text1.Text, Text5.Text, Text2.Text, Val(Text3.Text), Val(Text4.Text)
      rs.AddNew
      rs.Fields(0).Value = Text1.Text
      rs.Fields(1).Value = Text5.Text
      rs.Fields(2).Value = Text2.Text
      rs.Fields(3).Value = Text3.Text
      rs.Fields(4).Value = Text4.Text
    
      rs.Update
      'MsgBox ("succesful")
      addnewproductsuper.Hide
    
        Text1.Text = ""
        Text5.Text = ""
        Text2.Text = ""
         Text3.Text = ""
        Text4.Text = ""
    End If

End Sub

Private Function GenerateRandomID() As String
    Dim ID As String
    Dim i As Integer
    
    Randomize ' Initialize random number generator
    
    ' Generate 10-character ID
    For i = 1 To 3
        ' Append a random uppercase letter or digit to the ID
       
        ID = ID & Int(Rnd * 10)
        ' Alternatively, you can use: id = id & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomID = ID
End Function
Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
Set rs = db.OpenRecordset("Select * from supermarket")

Dim ID As String

ID = GenerateRandomID()

Text5.Text = ID
End Sub
