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
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Calculate Votes"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
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
'Dim FlexGrid1 As New MSFlexGrid

Private Sub Command1_Click()
castVote.Show

End Sub

Private Sub Command2_Click()
     Dim nameCounts As Scripting.Dictionary
    Dim winner As String

    ' Calculate occurrences of each name
    Set nameCounts = CountOccurrencesInFlexGrid(FlexGrid1)

    ' Find the winner
    winner = GetMaxOccurrences(nameCounts)

    ' Construct message for displaying total votes
    Dim msg As String
    Dim name As Variant
    For Each name In nameCounts.Keys
        msg = msg & name & ": " & nameCounts(name) & vbCrLf
    Next name

    ' Display the total votes for each candidate
    MsgBox "Total votes for each candidate:" & vbCrLf & msg, vbInformation, "Candidate Totals"

    ' Display the winner
    MsgBox "The winner is: " & winner, vbInformation, "Winner"

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
        .TextMatrix(1, 1) = "HEWYVB"
        .TextMatrix(1, 2) = "4234"
        .TextMatrix(1, 3) = "Peter Obi"

        
        .TextMatrix(2, 0) = ""
        .TextMatrix(2, 1) = "ALOKMW"
        .TextMatrix(2, 2) = "2343"
        .TextMatrix(2, 3) = "Bola Ahmed Tinubu"

        
        .TextMatrix(3, 0) = ""
        .TextMatrix(3, 1) = "POUMVZ"
        .TextMatrix(3, 2) = "4324"
        .TextMatrix(3, 3) = "Ibrahim Quazim Adebayo "

        
        .TextMatrix(4, 0) = ""
        .TextMatrix(4, 1) = "PLSMAI"
        .TextMatrix(4, 2) = "0195"
        .TextMatrix(4, 3) = "Atiku Abubarkar"
        
        .TextMatrix(5, 0) = ""
        .TextMatrix(5, 1) = "MWLJDM"
        .TextMatrix(5, 2) = "3985"
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
Private Function CountOccurrencesInFlexGrid(ByVal flexGrid As MSFlexGrid) As Scripting.Dictionary
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim i As Integer, j As Integer
    Dim nameCounts As New Scripting.Dictionary
    
    rowCount = flexGrid.Rows - 1 ' Exclude the fixed row
    colCount = flexGrid.Cols
    
    ' Count occurrences of each name
    For i = 1 To rowCount
        Dim name As String
        name = flexGrid.TextMatrix(i, 3) ' Assuming the names are in the first column
        
        If nameCounts.Exists(name) Then
            nameCounts(name) = nameCounts(name) + 1
        Else
            nameCounts.Add name, 1
        End If
    Next i
    
    Set CountOccurrencesInFlexGrid = nameCounts
End Function
Private Function GetMaxOccurrences(ByVal counts As Scripting.Dictionary) As String
    Dim maxCount As Integer
    Dim maxName As String
    Dim name As Variant
    
    maxCount = 0
    
    ' Find the name(s) with the maximum occurrence count
    For Each name In counts.Keys
        If counts(name) > maxCount Then
            maxCount = counts(name)
            maxName = name
        ElseIf counts(name) = maxCount Then
            ' In case of a tie, select the candidate whose name comes first alphabetically
            If StrComp(name, maxName, vbTextCompare) < 0 Then
                maxName = name
            End If
        End If
    Next name
    
    GetMaxOccurrences = maxName
End Function


