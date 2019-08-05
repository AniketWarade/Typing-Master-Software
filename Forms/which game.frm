VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "which game.frx":0000
   LinkTopic       =   "Form15"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "Game 2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   10920
      TabIndex        =   3
      Top             =   2520
      Width           =   5655
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Play Word Game 2"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   $"which game.frx":164A
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Game 1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   4080
      TabIndex        =   1
      Top             =   2520
      Width           =   5655
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Play Word Game 1"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   $"which game.frx":173D
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "!!!!!  PLAY A WORD GAME  !!!!!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6840
      TabIndex        =   7
      Top             =   1200
      Width           =   7050
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form13
Form13.Show
Unload Me
End Sub

Private Sub Command2_Click()
Load Form16
Form16.Show
Unload Me
End Sub

Private Sub Command3_Click()
Load Form8
Form8.Show
Unload Me
End Sub

