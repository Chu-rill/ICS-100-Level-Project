VERSION 5.00
Begin VB.Form removesuper 
   BackColor       =   &H00404000&
   Caption         =   "Remove Product"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      Caption         =   "Product ID:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "removesuper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()

Dim deleteSQL As String
Dim name As String
Dim ID As String

name = Text1.Text

ID = Text5.Text




If Text1.Text = "" Or Text5.Text = "" Then
    MsgBox "Input field can not be empty", vbExclamation
Else
    MsgBox ("Product Removed" & vbNewLine & _
    "Product Name: " & name & vbNewLine & _
    "Product ID: " & ID)
    removephar.Hide

    Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
    
     'deleteSQL = "DELETE * FROM supermarket WHERE ID = " & ID
     deleteSQL = "DELETE FROM supermarket WHERE ID = '" & ID & "'"
     
    ' Execute the delete query
    db.Execute deleteSQL
    
    ' Close the database
    db.Close
    
    ' Reload the data in FlexGrid after deletion
    'Form_Load

    removesuper.Hide
    

    Text1.Text = ""
    Text5.Text = ""
End If

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Churchill\Desktop\group2\table.mdb")
Set rs = db.OpenRecordset("Select * from supermarket")
End Sub
