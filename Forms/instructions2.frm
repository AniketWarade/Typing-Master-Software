VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "instructions2.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   5160
      Picture         =   "instructions2.frx":164A
      ScaleHeight     =   4875
      ScaleWidth      =   10035
      TabIndex        =   1
      Top             =   5280
      Width           =   10095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   $"instructions2.frx":A0E84
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   15960
      TabIndex        =   6
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   $"instructions2.frx":A0F26
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   600
      TabIndex        =   5
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "First Character Line from Keyboard:-"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   4200
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "There are some finger positions used while typing.  Following this finger positions we can increase our speed while typing."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   19215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Typing Techniques"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7200
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form5
Form5.Show
Unload Me
End Sub

Private Sub Command2_Click()
Load Form3
Form3.Show
Unload Me
End Sub
