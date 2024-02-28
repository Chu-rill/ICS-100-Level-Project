VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00404000&
   Caption         =   "Login"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      MaskColor       =   &H80000005&
      TabIndex        =   3
      Top             =   5280
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "   Show password"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "#"
      TabIndex        =   5
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
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
      Left            =   4800
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "  Login"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 If Check1.Value = vbChecked Then
        ' If the checkbox is checked, show the password character
        Text2.PasswordChar = ""
    Else
        ' If the checkbox is unchecked, hide the password character
        Text2.PasswordChar = "#"
End If
End Sub

Private Sub Command1_Click()

If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Empty value not allowed.", vbExclamation
ElseIf Text1.Text = "Group 2" And Text2.Text = "admin" Then
    MsgBox ("Login Successful")
    Home.Show
    Login.Hide
End If


End Sub

