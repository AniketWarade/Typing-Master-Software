VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "selecttest.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "selecttest.frx":164A
      Left            =   11040
      List            =   "selecttest.frx":165A
      TabIndex        =   13
      Text            =   "2 Min."
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      ItemData        =   "selecttest.frx":1680
      Left            =   1320
      List            =   "selecttest.frx":16B4
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Width           =   13455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Extras"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   15360
      TabIndex        =   4
      Top             =   360
      Width           =   4695
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Play a word game"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Learn More"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   7920
         Width           =   2430
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   "Want to double your speed !! Then practice atleast 6 hours in a week and get the better result."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   720
         TabIndex        =   7
         Top             =   4680
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   708
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9360
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   120
      Picture         =   "selecttest.frx":1858
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Duration:-"
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
      Left            =   8400
      TabIndex        =   12
      Top             =   3120
      Width           =   2130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
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
      Left            =   17160
      TabIndex        =   10
      Top             =   9480
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Typing Test"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Choose Test :-"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1440
      TabIndex        =   2
      Top             =   3360
      Width           =   3660
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
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
      Height          =   555
      Left            =   14040
      TabIndex        =   1
      Top             =   9480
      Width           =   1065
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If List1.Text = "Aesob's Fables" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Astronauts" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Behind The Scene: Movie Credits" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "DNA Research -- The Human Genome Project" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "History of Photography" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Hubble Space Telescope" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Netiquette" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Rules of Baseball (from 1889)" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Speeding up the strategy process" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Stinging Insects" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Eight Tools for Creating New Value" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Light Match Girl (adapted)" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Tale of Peter Rabbit" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Thoughts of Benjamin Franklin" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "What is a Cast?" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Yosemite National Park" Then
Load Form9
Form9.Show
Unload Me
End If
End Sub

Private Sub Command2_Click()
Load Form12
Form12.Show
Unload Me
End Sub

Private Sub Command3_Click()
Load Form11
Form11.Show
Unload Me
End Sub

Private Sub Command4_Click()
confirm = MsgBox("Are you sure you want to exit", vbYesNo, "!!! Alert !!!")
If confirm = vbYes Then
End
Else
Load Form8
Form8.Show
End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontUnderline = False
End Sub

Private Sub Label15_Click()
Shell "Explorer C:\Typing Master Software\image.html"
End Sub

Private Sub Label16_Click()
Shell "Explorer C:\Users\Aniket\Desktop\Typing Master Software\Lessons\typing.html"
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontUnderline = True
End Sub

Private Sub List1_DblClick()
If List1.Text = "Aesob's Fables" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Astronauts" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Behind The Scene: Movie Credits" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "DNA Research -- The Human Genome Project" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "History of Photography" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Hubble Space Telescope" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Netiquette" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Observations of Father of Computing (1791 - 1871)" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Rules of Baseball (from 1889)" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Speeding up the strategy process" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Stinging Insects" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Eight Tools for Creating New Value" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Light Match Girl (adapted)" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "The Tale of Peter Rabbit" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Thoughts of Benjamin Franklin" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "What is a Cast?" Then
Load Form9
Form9.Show
Unload Me
ElseIf List1.Text = "Yosemite National Park" Then
Load Form9
Form9.Show
Unload Me
End If
End Sub



