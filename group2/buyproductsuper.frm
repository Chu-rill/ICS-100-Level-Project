VERSION 5.00
Begin VB.Form buyproductsuper 
   BackColor       =   &H00404000&
   Caption         =   "Buy Product"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   2655
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
      Height          =   1125
      Left            =   4560
      TabIndex        =   9
      Top             =   4080
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
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
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
      Height          =   405
      Left            =   4560
      TabIndex        =   3
      Top             =   1560
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
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      Caption         =   "Customer's Adress:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "buyproductsuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Function GenerateRandomID() As String
    Dim ID As String
    Dim i As Integer
    
    Randomize ' Initialize random number generator
    

 For i = 1 To 10
 ID = ID & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
        ' Alternatively, you can use: id = id & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomID = ID
End Function
Private Sub Command1_Click()

Dim productName As String
Dim customerName As String
Dim quantity As Integer
Dim price As Integer
Dim address As String
Dim ID As String

ID = GenerateRandomID()

productName = Text1.Text
customerName = Text2.Text
quantity = Val(Text3.Text)
price = Val(Text4.Text)
address = Text5.Text




If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Input field can not be empty", vbExclamation
Else

   MsgBox ("Purchase Successful " & vbNewLine & _
    "Product Name: " & productName & vbNewLine & _
    "Customer Name: " & customerName & vbNewLine & _
    "Quantity: " & quantity & vbNewLine & _
    "Price: " & price & vbNewLine & _
    "Address: " & address & vbNewLine & _
    "Receipt No: " & ID)

     rs.AddNew
      rs.Fields(0).Value = Text1.Text
      rs.Fields(1).Value = Text2.Text
      rs.Fields(2).Value = Text3.Text
      rs.Fields(3).Value = Text4.Text
      rs.Fields(4).Value = Text5.Text
      rs.Fields(5).Value = ID
      
      
      rs.Update
    buyproductphar.Hide

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End If
  
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
Set rs = db.OpenRecordset("Select * from salessuper")
End Sub
