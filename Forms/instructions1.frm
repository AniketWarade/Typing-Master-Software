VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   Icon            =   "instructions1.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "instructions1.frx":164A
   ScaleHeight     =   9000
   ScaleWidth      =   15765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   675
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Skip"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click on skip if you want to skip the instructions."
      Top             =   9840
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form8
Form8.Show
Unload Me
End Sub

Private Sub Command2_Click()
Load Form4
Form4.Show
Unload Me
End Sub
