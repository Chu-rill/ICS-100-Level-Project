VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form billGrid 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H000080FF&
      Caption         =   "Add Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   4095
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   4095
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   8
      Cols            =   6
   End
End
Attribute VB_Name = "billGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Bill.Show
End Sub

Private Sub Form_Load()
    ' Populate FlexGrid with default data
    With FlexGrid1
        .Rows = 6 ' Set the number of rows including header row
        .Cols = 7 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Serial  No"
        .TextMatrix(0, 2) = "Ciustomer Name"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Phone No"
        .TextMatrix(0, 5) = "Meter No"
        .TextMatrix(0, 6) = "Status"

        ' Default data
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = "SN001"
    .TextMatrix(1, 2) = "John Doe"
    .TextMatrix(1, 3) = "A1"
    .TextMatrix(1, 4) = "08011234567"
    .TextMatrix(1, 5) = "M12345"
    .TextMatrix(1, 6) = "Pending"
    
    .TextMatrix(2, 0) = ""
    .TextMatrix(2, 1) = "SN002"
    .TextMatrix(2, 2) = "Jane Smith"
    .TextMatrix(2, 3) = "B2"
    .TextMatrix(2, 4) = "08125678901"
    .TextMatrix(2, 5) = "M54321"
    .TextMatrix(2, 6) = "Approved"
    
    .TextMatrix(3, 0) = ""
    .TextMatrix(3, 1) = "SN003"
    .TextMatrix(3, 2) = "Alice Johnson"
    .TextMatrix(3, 3) = "C3"
    .TextMatrix(3, 4) = "08039012345"
    .TextMatrix(3, 5) = "M98765"
    .TextMatrix(3, 6) = "Pending"
    
    .TextMatrix(4, 0) = ""
    .TextMatrix(4, 1) = "SN004"
    .TextMatrix(4, 2) = "Bob Williams"
    .TextMatrix(4, 3) = "D4"
    .TextMatrix(4, 4) = "08153456789"
    .TextMatrix(4, 5) = "M67890"
    .TextMatrix(4, 6) = "Approved"
    
    .TextMatrix(5, 0) = ""
    .TextMatrix(5, 1) = "SN005"
    .TextMatrix(5, 2) = "Eva Brown"
    .TextMatrix(5, 3) = "E5"
    .TextMatrix(5, 4) = "08097890123"
    .TextMatrix(5, 5) = "M23456"
    .TextMatrix(5, 6) = "Pending"
        
    
     .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
    
           Dim col As Integer
        For col = 0 To .Cols - 1
            .ColWidth(col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next col
    End With

  
End Sub
Public Sub AddDataToFlexGrid(Text1 As String, Text2 As String, Text3 As String, Text4 As String, Text5 As String, Text6 As String)


    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
    Dim newRow As Integer
    newRow = billGrid.FlexGrid1.Rows + 1
    billGrid.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
  '  billGrid.FlexGrid1.TextMatrix(newRow - 1, 0) = Text1
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 1) = Text1
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 3) = Text3
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 4) = Text4
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 5) = Text5
    billGrid.FlexGrid1.TextMatrix(newRow - 1, 6) = Text6

End Sub

