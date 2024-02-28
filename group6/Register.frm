VERSION 5.00
Begin VB.Form Register 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   9840
      TabIndex        =   23
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      MaskColor       =   &H8000000B&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Text10 
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
      Left            =   9840
      TabIndex        =   20
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text9 
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
      Left            =   9840
      TabIndex        =   19
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text8 
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
      Left            =   9840
      TabIndex        =   18
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text7 
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
      Left            =   9840
      TabIndex        =   17
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text6 
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
      Left            =   3240
      TabIndex        =   16
      Top             =   4800
      Width           =   2655
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
      Left            =   3240
      TabIndex        =   10
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text4 
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
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox Text3 
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
      Left            =   3240
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text2 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
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
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000040&
      Caption         =   "Time(M):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000040&
      Caption         =   "Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000040&
      Caption         =   "Amount Borrowed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7800
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000040&
      Caption         =   "Guarantor's Occupation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000040&
      Caption         =   "Guarantor's Phone Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7080
      TabIndex        =   12
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000040&
      Caption         =   " Guarantor's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000040&
      Caption         =   "Customer's  Age:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000040&
      Caption         =   "Cuustomer's Occupation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "Customer's Phone Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "Customer's Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "Customer's  Full Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "REGISTER NEW USERS"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Private Sub Command1_Click()
Dim P As Double
Dim R As Double
Dim T As Double
Dim SI As Double

P = Val(Text9.Text)
R = Val(Text10.Text)
T = Val(Text11.Text)

SI = (P * R * T) / 100

rs.AddNew
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = Text4.Text
rs.Fields(4).Value = Text5.Text
rs.Fields(5).Value = Text6.Text
rs.Fields(6).Value = Text7.Text
rs.Fields(7).Value = Text8.Text
rs.Fields(8).Value = Text9.Text
rs.Fields(9).Value = Text10.Text
rs.Fields(10).Value = Text11.Text
rs.Fields(11).Value = SI
rs.Fields(12).Value = Val(Text9.Text) + SI

    
rs.Update
MsgBox ("succesful")
      
      Register.Hide
      
MsgBox ("User Sucessfully Added.")
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Churchill\Desktop\group6\table.mdb")
Set rs = db.OpenRecordset("Select * from table1")
End Sub


