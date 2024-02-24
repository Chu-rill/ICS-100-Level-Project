VERSION 5.00
Begin VB.Form Bill 
   BackColor       =   &H00800080&
   Caption         =   "Bill"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      Left            =   6960
      TabIndex        =   7
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox Text2 
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
      Left            =   6960
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text3 
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
      Left            =   6960
      TabIndex        =   5
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox Text4 
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
      Left            =   6960
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.TextBox Text5 
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
      Left            =   6960
      TabIndex        =   3
      Top             =   4560
      Width           =   4095
   End
   Begin VB.TextBox Text6 
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
      Left            =   6960
      TabIndex        =   2
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Add Bill"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   14
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "Serial No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800080&
      Caption         =   "Customer Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800080&
      Caption         =   "Unit :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Caption         =   "Phone No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800080&
      Caption         =   "Meter No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800080&
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "Bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 MsgBox ("Item Added")
      billGrid.AddDataToFlexGrid Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text
  
     Bill.Hide
    
        Text1.Text = ""
        Text2.Text = ""
         Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

