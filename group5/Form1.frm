VERSION 5.00
Begin VB.Form Register 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   975
      Left            =   5880
      TabIndex        =   19
      Top             =   7680
      Width           =   4695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   7080
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   6240
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Voter Registration Form"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   18
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Full Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "D.O.B :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Phone NO :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "L.G.A of Residency :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Nationality :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "State of Residency :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "State of Origin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label11 
      Caption         =   "Gender :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   7080
      Width           =   2535
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim serial As String
'Dim ID As Double

'serial = Text1.Text
'ID = Val(Text2.Text)

'If Text1.Text = "" Or Text2.Text = "" Then
 '   MsgBox "Empty field not allowed", vbExclamation
'Else
 '   MsgBox ("Login Successful " & vbNewLine & _
  '  "You can now vote")
   ' Login.Show
    'Register.Hide
    
   ' End If
   
    Dim serialNumber As String
    Dim id As String
    
    serialNumber = GenerateRandomSerialNumber()
    id = GenerateRandomID()
    
    MsgBox ("write down serial and id number")
   
    MsgBox "Serial Number: " & serialNumber & vbCrLf & _
           " ID: " & id, vbInformation
   
   Login.Show
   Register.Hide
   
End Sub

Private Function GenerateRandomSerialNumber() As String
    Dim serialNumber As String
    Dim i As Integer
    
    Randomize ' Initialize random number generator
    
    ' Generate 16-character serial number
    For i = 1 To 16
        ' Append a random uppercase letter or digit to the serial number
        serialNumber = serialNumber & Chr(Asc("A") + Int(Rnd * 26)) ' Uppercase letter (A-Z)
        ' Alternatively, you can use: serialNumber = serialNumber & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomSerialNumber = serialNumber
End Function

Private Function GenerateRandomID() As String
    Dim id As String
    Dim i As Integer
    
    Randomize ' Initialize random number generator
    
    ' Generate 10-character ID
    For i = 1 To 10
        ' Append a random uppercase letter or digit to the ID
        id = id & Chr(Asc("A") + Int(Rnd * 26)) ' Uppercase letter (A-Z)
        ' Alternatively, you can use: id = id & Chr(48 + Int(Rnd * 10)) ' Digit (0-9)
    Next i
    
    GenerateRandomID = id
End Function

Private Sub Form_Load()
    ' Usage example:

 
    

        

End Sub


Private Sub serialNumber_Change()

End Sub
