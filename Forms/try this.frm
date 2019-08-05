VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "try this.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   5400
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   7
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   8
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   9
      Top             =   5280
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   10
      Top             =   6240
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3120
      TabIndex        =   11
      Top             =   7200
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Begin"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   2535
   End
   Begin VB.OptionButton timepractice 
      BackColor       =   &H8000000D&
      Caption         =   "Practice Mode - Timer counts up instead of down.  "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13800
      TabIndex        =   6
      Top             =   8280
      Width           =   3255
   End
   Begin VB.OptionButton Time40 
      BackColor       =   &H8000000D&
      Caption         =   "40 seconds(very easy)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13800
      TabIndex        =   1
      Top             =   3360
      Width           =   3255
   End
   Begin VB.OptionButton time5 
      BackColor       =   &H8000000D&
      Caption         =   "5 seconds(very hard)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13800
      TabIndex        =   5
      Top             =   7320
      Width           =   3375
   End
   Begin VB.OptionButton time10 
      BackColor       =   &H8000000D&
      Caption         =   "10 seconds(hard)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      TabIndex        =   4
      Top             =   6240
      Width           =   3255
   End
   Begin VB.OptionButton Time15 
      BackColor       =   &H8000000D&
      Caption         =   "15 seconds(medium)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   3
      Top             =   5280
      Width           =   3255
   End
   Begin VB.OptionButton time25 
      BackColor       =   &H8000000D&
      Caption         =   "25 seconds(easy)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   2
      Top             =   4320
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   23
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   21
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   20
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   19
      Top             =   7320
      Width           =   4455
   End
   Begin VB.Label win 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2040
      TabIndex        =   18
      Top             =   1080
      Width           =   60
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   $"try this.frx":164A
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3360
      TabIndex        =   15
      Top             =   8160
      Width           =   9615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select an Amount of Time:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13440
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label time 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9600
      TabIndex        =   17
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Time remaining:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8040
      TabIndex        =   13
      Top             =   960
      Width           =   3150
   End
   Begin VB.Label winner 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3960
      TabIndex        =   16
      Top             =   1080
      Width           =   60
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Load Form16
Form16.Show
Unload Me
End Sub

Private Sub Form_load()
Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Randomize
Dim aString
winner.Caption = ""
win.Caption = ""
Timer1.Enabled = True
    aString = Array("hello", "hi", "love", "life", "like", "live", "book", "reserve", "crazy", "stupid", "lazy", "menu", "visual", "basic", "quick", "fast", _
                    "hard", "internet", "interest", "interval", "tag", "mail", "post", "left", "right", "angle", "triangle", "rectangle", "interface", _
                    "application", "programming", "chat", "room", "message", "Jack", "John", "Liza", "Pizza", "pizza", "Eric", "Tony", "George", "Harry", _
                    "Kate", "Marlon Brando", "Audry Hepburn", "Katherin Hepburn", "actor", "actoress", "soft", "rough", "Tehran", "Iran", "USA", "Russia", _
                    "Spain", "Poland", "Iceland", "island", "keyboard", "keypad", "mailbox", "scroll", "command", "authority", "obligation", "warfare", _
                    "peace", "singer", "pop", "rock", "hip hop", "country", "rap", "classic", "Dally", "football", "soccer", "FIFA", "UEFA", "Asia", _
                    "Africa", "America", "Australia", "Austria", "other", "staircase", "ahead", "bump", "remark", "board", "insist", "accuse", _
                    "Ku-Klux-Klan", "sawmill", "fussy", "rough", "straighten", "cross", "mercy", "God", "mother", "father", "brother", "sister", "law", _
                    "lawyer", "whistle", "towel", "pass", "take", "forfiet", "sound", "suicide", "operator", "set", "surgery", "effort", "drop", _
                    "drag", "bold", "fusion", "razor", "turn", "cliff", "hitch", "urban", "canoe", "roller", "surf", "paddle", "carpetpiller", _
                    "threat", "consist", "bandit", "terror", "damp", "urgent", "supersonic", "tram", "silkworm", "praise", "compound", _
                    "hurricane", "gallop", "talent", "pandant", "ranch", "silicon", "jigsaw", "airplane", "caffeine", "starry", "vanish", _
                    "bloom", "accomplish", "miracle", "recommend", "ceilling", "sew", "weave", "hum", "rent", "lick", "puzzle", "insualte", _
                    "issue", "budget", "military", "influence", "howl", "garlic", "terrify", "bat", "keep", "sieze", "grab", "rare", "loosen", _
                    "blossom", "fun", "locksmith", "experiment", "sharp", "wound", "sled", "slope", "steep", "abandon", "book", "dull", "passerby", _
                    "homely", "blood", "stare", "carefree", "desperate", "despair", "trial", "clerk", "trap", "weep", "shut", "eager", "boathouse", "hire", "wage", _
                    "yell", "push", "maze", "through", "aground", "landlord", "landlady", "exhaust", "nest", "stove", "meanwhile", "suspect", "turn", _
                    "various", "passage", "sheet", "ordinary", "sleepyhead", "wipe", "neat", "narrow", "cough", "kneel", "knife", "fork", "spoon", "plate", _
                    "dish", "heap", "pile", "lunge", "federation", "bureau", "investigation", "pour", "upside down", "plead", "division", "parking lot", "FBI", _
                    "victim", "medical", "garbage", "defend", "nasty", "evidence", "gluttony", "envy", "wrath", "fierce", "homicide", "freeway", "humanity", _
                    "sloth", "pride", "lust", "retire", "revolve", "offence", "fool", "buttonhole", "tollboth", "dive", "ankle", "stiff", "greed", "lower", _
                    "cord", "collar", "hothead", "reveng", "underworld", "calm", "mixture", "godfather", "bet", "respect", "brandy", "approval", "drug", _
                    "sticky", "glue", "candle", "violent", "persuade", "fatty", "convict", "shelter", "enclosure", "loyal", "wax", "indoors", "honor", _
                    "fox", "volume", "function", "subject", "telephone", "zippo")
                    
   
Label1.Caption = aString(CInt(Rnd * 305))
Label2.Caption = aString(CInt(Rnd * 305))
Label3.Caption = aString(CInt(Rnd * 305))
Label4.Caption = aString(CInt(Rnd * 305))
Label5.Caption = aString(CInt(Rnd * 305))
If Time40.Value = True Then
time.Caption = "40"
End If
If time25.Value = True Then
time.Caption = "25"
End If
If Time15.Value = True Then
time.Caption = "15"
End If
If time10.Value = True Then
time.Caption = "10"
End If
If time5.Value = True Then
time.Caption = "5"
End If
If timepractice.Value = True Then
time.Caption = "1"
End If

End Sub



Private Sub Timer1_timer()
If Text1.Text = Label1.Caption And Text2.Text = Label2.Caption And Text3.Text = Label3.Caption And Text4.Text = Label4.Caption And Text5.Text = Label5.Caption Then
winner.Caption = "You Win!"
Command1.Caption = "New Game"
If Command1.Caption = "New Game" Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End If
Timer1.Enabled = False
End If
If Time40.Value = True Then
time.Caption = time
time = time - 1
End If
If time25.Value = True Then
time.Caption = time
time = time - 1
End If
If Time15.Value = True Then
time.Caption = time
time = time - 1
End If
If time10.Value = True Then
time.Caption = time
time = time - 1
End If
If time5.Value = True Then
time.Caption = time
time = time - 1
End If
If timepractice.Value = True Then
time.Caption = time
time = time + 1
End If
If time.Caption <= 0 Then
Timer1.Enabled = False
win.Caption = "Game Over!  You ran out of time!"
Timer1.Enabled = False
End If
End Sub


