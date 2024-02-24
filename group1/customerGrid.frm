VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form customerGrid 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Add Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   4335
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   8
      Cols            =   7
   End
End
Attribute VB_Name = "customerGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Customer.Show

End Sub

Private Sub Form_Load()
    ' Populate FlexGrid with default data
    With FlexGrid1
        .Rows = 6 ' Set the number of rows including header row
        .Cols = 8 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "First Name"
        .TextMatrix(0, 3) = "Last Name"
        .TextMatrix(0, 4) = "Address"
        .TextMatrix(0, 5) = "Phone"
        .TextMatrix(0, 6) = "Meter No"
        .TextMatrix(0, 7) = "Connection Date"
        
     ' Default data
    .TextMatrix(1, 0) = "Customer A"
    .TextMatrix(1, 1) = "001"
    .TextMatrix(1, 2) = "John"
    .TextMatrix(1, 3) = "Doe"
    .TextMatrix(1, 4) = "123 Main St"
    .TextMatrix(1, 5) = "08011234567"
    .TextMatrix(1, 6) = "M12345"
    .TextMatrix(1, 7) = "2022-01-01"
    
    .TextMatrix(2, 0) = "Customer B"
    .TextMatrix(2, 1) = "002"
    .TextMatrix(2, 2) = "Jane"
    .TextMatrix(2, 3) = "Smith"
    .TextMatrix(2, 4) = "456 Elm St"
    .TextMatrix(2, 5) = "08125678901"
    .TextMatrix(2, 6) = "M54321"
    .TextMatrix(2, 7) = "2022-01-15"
    
    .TextMatrix(3, 0) = "Customer C"
    .TextMatrix(3, 1) = "003"
    .TextMatrix(3, 2) = "Alice"
    .TextMatrix(3, 3) = "Johnson"
    .TextMatrix(3, 4) = "789 Oak St"
    .TextMatrix(3, 5) = "08039012345"
    .TextMatrix(3, 6) = "M98765"
    .TextMatrix(3, 7) = "2022-02-01"
    
    .TextMatrix(4, 0) = "Customer D"
    .TextMatrix(4, 1) = "004"
    .TextMatrix(4, 2) = "Bob"
    .TextMatrix(4, 3) = "Williams"
    .TextMatrix(4, 4) = "321 Pine St"
    .TextMatrix(4, 5) = "08153456789"
    .TextMatrix(4, 6) = "M67890"
    .TextMatrix(4, 7) = "2022-02-15"
    
    .TextMatrix(5, 0) = "Customer E"
    .TextMatrix(5, 1) = "005"
    .TextMatrix(5, 2) = "Eva"
    .TextMatrix(5, 3) = "Brown"
    .TextMatrix(5, 4) = "555 Cedar St"
    .TextMatrix(5, 5) = "08097890123"
    .TextMatrix(5, 6) = "M23456"
    .TextMatrix(5, 7) = "2022-03-01"
        
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
          Dim col As Integer
        For col = 0 To .Cols - 1
            .ColWidth(col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next col
    End With

    ' Set the column width for the third column to be 2000 twips
    
End Sub
Public Sub AddDataToFlexGrid(Text1 As String, Text2 As String, Text3 As String, Text4 As String, Text5 As String, Text6 As String, Text7 As String)


    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
    Dim newRow As Integer
    newRow = customerGrid.FlexGrid1.Rows + 1
    customerGrid.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
    ' customerGrid.FlexGrid1.TextMatrix(newRow - 1, 0) = Text1
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 1) = Text1
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 3) = Text3
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 4) = Text4
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 5) = Text5
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 6) = Text6
    customerGrid.FlexGrid1.TextMatrix(newRow - 1, 7) = Text7
End Sub
