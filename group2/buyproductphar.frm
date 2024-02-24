VERSION 5.00
Begin VB.Form buyproductphar 
   BackColor       =   &H8000000A&
   Caption         =   "Buy Product"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   885
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4200
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
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
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
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
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   855
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
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
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
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
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
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "buyproductphar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim productName As String
Dim customerName As String
Dim quantity As Double
Dim price As Integer
Dim address As String

productName = Text1.Text
customerName = Text2.Text
quantity = Val(Text3.Text)
price = Val(Text4.Text)
address = Text5.Text


If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Input field can not be empty", vbExclamation
Else

    MsgBox ("Purchase Successful " & vbNewLine & _
    productName & vbNewLine & _
    customerName & vbNewLine & _
    quantity & vbNewLine & _
    price & vbNewLine & _
    address)

    buyproductphar.Hide

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End If
End Sub

