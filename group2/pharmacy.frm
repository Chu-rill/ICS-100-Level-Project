VERSION 5.00
Begin VB.Form pharmacy 
   BackColor       =   &H00404000&
   Caption         =   "Pharmacy"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   14745
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
      Left            =   1080
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
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
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
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
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
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
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "pharmacy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset


Private Sub Command1_Click()
addnewproductphar.Show

End Sub

Private Sub Command2_Click()
removephar.Show

End Sub

Private Sub Command3_Click()
buyproductphar.Show

End Sub
'Public Sub AddDataToFlexGrid(Text1 As String, Text5 As String, Text2 As String, Text3 As Integer, Text4 As Double)
    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
 '   Dim newRow As Integer
  '  newRow = pharmacy.FlexGrid1.Rows + 1
   ' pharmacy.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
   ' pharmacy.FlexGrid1.TextMatrix(newRow - 1, 0) = Text1
   ' pharmacy.FlexGrid1.TextMatrix(newRow - 1, 1) = Text5
   ' pharmacy.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
   ' pharmacy.FlexGrid1.TextMatrix(newRow - 1, 3) = Text3
   ' pharmacy.FlexGrid1.TextMatrix(newRow - 1, 4) = Text4
'End Sub

Private Sub Command4_Click()
inventoryphar.Show

End Sub

Private Sub Command5_Click()
salesphar.Show

End Sub
