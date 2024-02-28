VERSION 5.00
Begin VB.Form addnewproductphar 
   BackColor       =   &H00404000&
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
      Left            =   3960
      TabIndex        =   12
      Top             =   960
      Width           =   2520
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
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
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   5640
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
      Left            =   2280
      TabIndex        =   1
      Top             =   4920
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
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
addnewproductphar.Hide

End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Or Text5.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "Input field can not be empty", vbExclamation
    Else
      'MsgBox ("Item Added")
      MsgBox ("Item Added " & vbNewLine & _
    "Product Name: " & Text1.Text & vbNewLine & _
    "Product ID: " & Text5.Text & vbNewLine & _
    "Description: " & Text2.Text & vbNewLine & _
    "Quantity: " & Text3.Text & vbNewLine & _
    "Price: " & Text4.Text) 'vbNewLine & _
   ' "Receipt No: " & ID)
    '  pharmacy.AddDataToFlexGrid Text1.Text, Text5.Text, Text2.Text, Val(Text3.Text), Val(Text4.Text)
      rs.AddNew
      rs.Fields(0).Value = Text1.Text
      rs.Fields(1).Value = Text5.Text
      rs.Fields(2).Value = Text2.Text
      rs.Fields(3).Value = Text3.Text
      rs.Fields(4).Value = Text4.Text
    
      rs.Update
     ' MsgBox ("succesful")
  
      addnewproductphar.Hide
    
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
 ' Uppercase letter (A-Z)

        ' Alternatively, you can use: id = id & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomID = ID
End Function
Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
Set rs = db.OpenRecordset("Select * from pharmcy")

Dim ID As String

ID = GenerateRandomID()

Text5.Text = ID
End Sub
