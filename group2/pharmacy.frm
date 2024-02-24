VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form pharmacy 
   Caption         =   "Pharmacy"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   4575
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   12632256
      AllowUserResizing=   3
      GridLineWidth   =   2
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
      Top             =   3720
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
      Top             =   2640
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
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "pharmacy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
addnewproductphar.Show

End Sub

Private Sub Command2_Click()
removephar.Show

End Sub

Private Sub Command3_Click()
buyproductphar.Show

End Sub

Public Sub AddDataToFlexGrid(Text1 As String, Text5 As String, Text2 As String, Text3 As Integer, Text4 As Double)
    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
    Dim newRow As Integer
    newRow = pharmacy.FlexGrid1.Rows + 1
    pharmacy.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
    pharmacy.FlexGrid1.TextMatrix(newRow - 1, 0) = Text1
    pharmacy.FlexGrid1.TextMatrix(newRow - 1, 1) = Text5
    pharmacy.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
    pharmacy.FlexGrid1.TextMatrix(newRow - 1, 3) = Text3
    pharmacy.FlexGrid1.TextMatrix(newRow - 1, 4) = Text4
End Sub



Private Sub Form_Load()
    ' Populate FlexGrid with default data
    With FlexGrid1
        .Rows = 6 ' Set the number of rows including header row
        .Cols = 5 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Quantity"
        .TextMatrix(0, 4) = "Price"
        
        
        
        ' Default data
         .TextMatrix(1, 0) = "Product A"
        .TextMatrix(1, 1) = "001"
        .TextMatrix(1, 2) = "Description of Product A"
        .TextMatrix(1, 3) = "10"
        .TextMatrix(1, 4) = "10"
        
        .TextMatrix(2, 0) = "Product B"
        .TextMatrix(2, 1) = "002"
        .TextMatrix(2, 2) = "Description of Product B"
        .TextMatrix(2, 3) = "20"
        .TextMatrix(2, 4) = "15"
        
        .TextMatrix(3, 0) = "Product C"
        .TextMatrix(3, 1) = "003"
        .TextMatrix(3, 2) = "Description of Product C"
        .TextMatrix(3, 3) = "15"
        .TextMatrix(3, 4) = "20"
        
        .TextMatrix(4, 0) = "Product D"
        .TextMatrix(4, 1) = "004"
        .TextMatrix(4, 2) = "Description of Product D"
        .TextMatrix(4, 3) = "25"
        .TextMatrix(4, 4) = "25"
        
        .TextMatrix(5, 0) = "Product E"
        .TextMatrix(5, 1) = "005"
        .TextMatrix(5, 2) = "Description of Product E"
        .TextMatrix(5, 3) = "30"
        .TextMatrix(5, 4) = "30"
        
            .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        
         Dim col As Integer
        For col = 0 To .Cols - 1
            .ColWidth(col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next col
        
    End With






' Set the column width for the first column to be 1000 twips
'FlexGrid1.ColWidth(2) = 2000

' Set the row height for the first row to be 1000 twips
'FlexGrid1.RowHeight(0) = 1000

End Sub
