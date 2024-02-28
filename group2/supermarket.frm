VERSION 5.00
Begin VB.Form supermarket 
   BackColor       =   &H00404000&
   Caption         =   "Supermarket"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Purchase Product"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Product"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Product"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "supermarket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
addnewproductsuper.Show

End Sub

Private Sub Command2_Click()
removesuper.Show

End Sub

Private Sub Command3_Click()
buyproductsuper.Show

End Sub


'Public Sub AddDataToFlexGrid(Text1 As String, Text5 As String, Text2 As String, Text3 As Integer, Text4 As Double)
    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
 '   Dim newRow As Integer
  '  newRow = supermarket.FlexGrid1.Rows + 1
   ' supermarket.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
  '  supermarket.FlexGrid1.TextMatrix(newRow - 1, 0) = Text1
  '  supermarket.FlexGrid1.TextMatrix(newRow - 1, 1) = Text5
  '  supermarket.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
  '  supermarket.FlexGrid1.TextMatrix(newRow - 1, 3) = Text3
  '  supermarket.FlexGrid1.TextMatrix(newRow - 1, 4) = Text4
'End Sub

Private Sub Command4_Click()
inventorysuper.Show

End Sub

Private Sub Command5_Click()
salessuper.Show

End Sub
