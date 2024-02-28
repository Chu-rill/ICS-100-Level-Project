VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form inventorysuper 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   14850
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   3615
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "inventorysuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Form_Load()
    ' Open the database
    Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
    Set rs = db.OpenRecordset("Select * from supermarket")
    
    ' Populate FlexGrid with default dummy data
    With FlexGrid1
        .Rows = 6 ' Set the initial number of rows including header row
        .Cols = 5 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Quantity"
        .TextMatrix(0, 4) = "Price"
        
        ' Populate FlexGrid with default dummy data
        '.TextMatrix(1, 0) = "Product A"
        '.TextMatrix(1, 1) = "001"
        '.TextMatrix(1, 2) = "Description of Product A"
        '.TextMatrix(1, 3) = "10"
        '.TextMatrix(1, 4) = "10"
        
        '.TextMatrix(2, 0) = "Product B"
        '.TextMatrix(2, 1) = "002"
        '.TextMatrix(2, 2) = "Description of Product B"
        '.TextMatrix(2, 3) = "20"
        '.TextMatrix(2, 4) = "15"
        
        ' Adjust column alignment and width
        For Col = 0 To .Cols - 1
            .ColAlignment(Col) = flexAlignCenterCenter
            .ColWidth(Col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next Col
    End With
    
    ' Close the recordset and database
    rs.Close
    
    ' Open the recordset again to append data from the database
    Set rs = db.OpenRecordset("Select * from supermarket")
    
    ' Append additional data from the database
    Dim rowIndex As Integer
    rowIndex = 1 ' Start appending from the fourth row (assuming dummy data occupies first three rows)
    Do While Not rs.EOF
        If rowIndex > FlexGrid1.Rows - 1 Then
            ' Add additional rows if needed
            FlexGrid1.Rows = FlexGrid1.Rows + 1
        End If
        FlexGrid1.TextMatrix(rowIndex, 0) = rs.Fields("Name").Value
        FlexGrid1.TextMatrix(rowIndex, 1) = rs.Fields("ID").Value
        FlexGrid1.TextMatrix(rowIndex, 2) = rs.Fields("Description").Value
        FlexGrid1.TextMatrix(rowIndex, 3) = rs.Fields("Quantity").Value
        FlexGrid1.TextMatrix(rowIndex, 4) = rs.Fields("Price").Value
        rs.MoveNext
        rowIndex = rowIndex + 1
    Loop
    
    ' Close the recordset and database
    rs.Close
    db.Close
End Sub


