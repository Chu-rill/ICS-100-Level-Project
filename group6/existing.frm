VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form existing 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8070
      _Version        =   393216
   End
End
Attribute VB_Name = "existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Form_Load()
    ' Open the database
    Set db = OpenDatabase("C:\Users\Churchill\Desktop\group6\table.mdb")
    Set rs = db.OpenRecordset("Select * from table1")
    
    ' Populate FlexGrid with default dummy data
    With FlexGrid1
        .Rows = 3 ' Set the number of rows including header row
        .Cols = 13 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = "Customer Full Name"
        .TextMatrix(0, 1) = "Customer Address"
        .TextMatrix(0, 2) = "Customer Phone No:"
        .TextMatrix(0, 3) = "Customer Occupation"
        .TextMatrix(0, 4) = "Customer Age"
        .TextMatrix(0, 5) = "Guarantor Name"
        .TextMatrix(0, 6) = "Guarantor Phone No:"
        .TextMatrix(0, 7) = "Guarantor Occupation"
        .TextMatrix(0, 8) = "Amount Borrowed"
        .TextMatrix(0, 9) = "Rate"
        .TextMatrix(0, 10) = "Time"
         .TextMatrix(0, 11) = "SI"
          .TextMatrix(0, 12) = "Total"
        
         For Col = 0 To .Cols - 1
            .ColAlignment(Col) = flexAlignCenterCenter
            .ColWidth(Col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next Col
        
        ' Populate FlexGrid with data from the recordset
     '   Dim rowIndex As Integer
      '  rowIndex = 1
       ' Do While Not rs.EOF
        '    .TextMatrix(rowIndex, 0) = rs.Fields("Customers's Full Name").Value
         '   .TextMatrix(rowIndex, 1) = rs.Fields("Customer Address").Value
          '  .TextMatrix(rowIndex, 2) = rs.Fields("Customer's Phone Number").Value
          '  .TextMatrix(rowIndex, 3) = rs.Fields("Customer's Occupation").Value
          '  .TextMatrix(rowIndex, 4) = rs.Fields("Customer Age").Value
          '  .TextMatrix(rowIndex, 5) = rs.Fields("Guarantor's Name").Value
          '  .TextMatrix(rowIndex, 6) = rs.Fields("Guarantor's Phone Number").Value
          '  .TextMatrix(rowIndex, 7) = rs.Fields("Guarantors Occupation").Value
          '  .TextMatrix(rowIndex, 8) = rs.Fields("Amount borrowed").Value
          '  .TextMatrix(rowIndex, 9) = rs.Fields("Rate").Value
           ' .TextMatrix(rowIndex, 10) = rs.Fields("Time").Value
          '  .TextMatrix(rowIndex, 11) = rs.Fields("SI").Value
          '  .TextMatrix(rowIndex, 12) = rs.Fields("Total").Value

           ' rs.MoveNext
            'rowIndex = rowIndex + 1
        'Loop
    End With
    
    
    ' Close the recordset and database
    rs.Close
    
    ' Open the recordset again to append data from the database
    Set rs = db.OpenRecordset("Select * from table1")
    
    ' Append additional data from the database
   Dim rowIndex As Integer
    rowIndex = 1 ' Start appending from the fourth row (assuming dummy data occupies first three rows)
    Do While Not rs.EOF
        If rowIndex > FlexGrid1.Rows - 1 Then
            ' Add additional rows if needed
            FlexGrid1.Rows = FlexGrid1.Rows + 1
        End If
        
             FlexGrid1.TextMatrix(rowIndex, 0) = rs.Fields("Customers's Full Name").Value
             FlexGrid1.TextMatrix(rowIndex, 1) = rs.Fields("Customer Address").Value
             FlexGrid1.TextMatrix(rowIndex, 2) = rs.Fields("Customer's Phone Number").Value
             FlexGrid1.TextMatrix(rowIndex, 3) = rs.Fields("Customer's Occupation").Value
             FlexGrid1.TextMatrix(rowIndex, 4) = rs.Fields("Customer Age").Value
             FlexGrid1.TextMatrix(rowIndex, 5) = rs.Fields("Guarantor's Name").Value
             FlexGrid1.TextMatrix(rowIndex, 6) = rs.Fields("Guarantor's Phone Number").Value
             FlexGrid1.TextMatrix(rowIndex, 7) = rs.Fields("Guarantors Occupation").Value
             FlexGrid1.TextMatrix(rowIndex, 8) = rs.Fields("Amount borrowed").Value
             FlexGrid1.TextMatrix(rowIndex, 9) = rs.Fields("Rate").Value
             FlexGrid1.TextMatrix(rowIndex, 10) = rs.Fields("Time").Value
             FlexGrid1.TextMatrix(rowIndex, 11) = rs.Fields("SI").Value
             FlexGrid1.TextMatrix(rowIndex, 12) = rs.Fields("Total").Value
        rs.MoveNext
        rowIndex = rowIndex + 1
    Loop
    
    ' Close the recordset and database
    rs.Close
    db.Close
End Sub


