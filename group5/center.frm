VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form center 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
      Height          =   4095
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Cast Vote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "center"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
castVote.Show

End Sub

Private Sub Form_Load()
    With FlexGrid1
        .Rows = 6 ' Set the number of rows including header row
        .Cols = 4 ' Set the number of columns
        
        ' Set headers
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Serial No"
        .TextMatrix(0, 2) = "ID"
        .TextMatrix(0, 3) = "Candite"

        
        
        
        ' Default data
         .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = "sdfdfdsfsdf"
        .TextMatrix(1, 2) = "4234343434"
        .TextMatrix(1, 3) = "Peter Obi"

        
        .TextMatrix(2, 0) = ""
        .TextMatrix(2, 1) = "002wefwef"
        .TextMatrix(2, 2) = "234354545235"
        .TextMatrix(2, 3) = "Bola Ahmed Tinubu"

        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "003343432423"
        .TextMatrix(3, 2) = "4324324234"
        .TextMatrix(3, 3) = "Ibrahim Quazim Adebayo "

        
        .TextMatrix(4, 0) = ""
        .TextMatrix(4, 1) = "453534"
        .TextMatrix(4, 2) = "dsgdsfdfdsf"
        .TextMatrix(4, 3) = "Atiku Abubarkar"
        
        .TextMatrix(5, 0) = ""
        .TextMatrix(5, 1) = "00543423"
        .TextMatrix(5, 2) = "werewrewrwer"
        .TextMatrix(5, 3) = "Ibrahim Quazim Adebayo"
        
          .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
         Dim col As Integer
        For col = 0 To .Cols - 1
            .ColWidth(col) = 2000 ' Set the width as desired (2000 twips in this case)
        Next col
        'selectedOption
    End With

End Sub
Public Sub AddDataToFlexGrid(Text1 As String, Text2 As String, optionCaption As String)
    ' Assuming FlexGrid1 is the name of the FlexGrid control on your form
    ' Add a new row to the FlexGrid
    Dim newRow As Integer
    newRow = center.FlexGrid1.Rows + 1
    center.FlexGrid1.Rows = newRow
    
    ' Set values for each column in the new row
    center.FlexGrid1.TextMatrix(newRow - 1, 1) = Text1
    center.FlexGrid1.TextMatrix(newRow - 1, 2) = Text2
    center.FlexGrid1.TextMatrix(newRow - 1, 3) = optionCaption
End Sub

