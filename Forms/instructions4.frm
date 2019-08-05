VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "instructions4.frx":0000
   LinkTopic       =   "Form12"
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
      TabIndex        =   6
      Top             =   9600
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   5040
      Picture         =   "instructions4.frx":164A
      ScaleHeight     =   4875
      ScaleWidth      =   10035
      TabIndex        =   1
      Top             =   5040
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   $"instructions4.frx":A0E84
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
      Left            =   15720
      TabIndex        =   5
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   $"instructions4.frx":A0F3F
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
      TabIndex        =   4
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Third Character Line from Keyboard:-"
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
      TabIndex        =   3
      Top             =   3720
      Width           =   7965
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form7
Form7.Show
Unload Me
End Sub

Private Sub Command2_Click()
Load Form5
Form5.Show
Unload Me
End Sub

