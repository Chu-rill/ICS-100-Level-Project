VERSION 5.00
Begin VB.Form customer 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   ForeColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Caption         =   "FirstName"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Surname :"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
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
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
