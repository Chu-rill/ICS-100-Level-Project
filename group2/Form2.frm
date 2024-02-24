VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Home"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Supermarket"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6720
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Pharmacy"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to Pharmart, click to continue..."
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
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   5295
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
pharmacy.Show

End Sub

Private Sub Command2_Click()
supermarket.Show

End Sub
