VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19335
   Icon            =   "wordgame.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   8940
   ScaleWidth      =   19335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   6480
   End
   Begin VB.Timer WholeGame 
      Interval        =   60000
      Left            =   2640
      Top             =   6480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11760
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Letters have been typed as a record."
      Top             =   -240
      Width           =   15255
      Begin VB.TextBox txtText 
         Height          =   495
         Left            =   480
         TabIndex        =   9
         ToolTipText     =   "Type the word in the box on top, here."
         Top             =   7200
         Width           =   8775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Typed Letters"
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
         Left            =   8160
         TabIndex        =   13
         Top             =   9960
         Width           =   4665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Typed Words"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   9960
         Width           =   4530
      End
      Begin VB.Label lblScoreLet 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         TabIndex        =   8
         ToolTipText     =   "Correct typed letters"
         Top             =   9120
         Width           =   345
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   7
         ToolTipText     =   "Correct typed words"
         Top             =   9120
         Width           =   345
      End
      Begin VB.Label lblRecLet 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   6
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label lblRec 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   5
         ToolTipText     =   "Word have been typed as a record."
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Letters:"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   4
         Top             =   5400
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Words:"
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
         Index           =   2
         Left            =   4560
         TabIndex        =   3
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Records:"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   2880
         Width           =   2325
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   720
         TabIndex        =   1
         ToolTipText     =   "Type the words appeared here in the white box below."
         Top             =   840
         Width           =   11595
      End
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
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
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "Exit"
      Top             =   4200
      Width           =   780
   End
   Begin VB.Label Game 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Play New Game"
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
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Click here to start the game."
      Top             =   2640
      Width           =   3300
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim iSec As Integer
Dim iSSec As String

Private Sub Form_Load()
    On Error Resume Next
    Me.Move GetSetting(App.Title, "FormState", "Left"), GetSetting(App.Title, "FormState", "Top")
    If GetSetting(App.Title, "THErecordOFtypingINtheGAME", "THISforMAKINGsureIFthereHASbeenANYbreakingINrecord") = "True" Then
        lblRec = GetSetting(App.Title, "THErecordOFtypingINtheGAME", "THISisTHErecordOFgame")
    End If
    If GetSetting(App.Title, "THErecordOFtypingINtheGAME", "THISforMAKINGsureIFthereHASbeenANYbreakingINrecordFORletters") = "True" Then
        lblRecLet = GetSetting(App.Title, "THErecordOFtypingINtheGAME", "THISisTHErecordOFgameLetters")
    End If
End Sub


Private Sub Game_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Handle
       Timer1.Enabled = False
    iSSec = InputBox("Enter how many seconds a word to be held:")
    If iSSec <> "" Then
        iSec = CInt(iSSec)
        Timer1.Interval = iSec * 1000
        WholeGame.Enabled = True
        txtText.Locked = False
        txtText.Text = ""
        lblScore = 0
        lblScoreLet = 0
        Show_Next_Word
    End If
    Exit Sub
Handle:
    If (Err.Number = 13) And (iSSec <> "") Then
        MsgBox "Please type a number!!!", vbCritical
    Else
        'Do nothing
    End If
End Sub

Private Sub lblExit_click()
    On Error Resume Next
    Load Form13
    Form13.Show
    Unload Me
End Sub



Private Sub Timer1_timer()
    On Error Resume Next
    txtText.Text = ""
    txtText.SetFocus
    Show_Next_Word
End Sub

Private Sub txtText_Change()
    On Error Resume Next
    If txtText.Text = lblText.Caption Then
        Timer1.Enabled = False
        Timer1.Enabled = True
        lblScore.Caption = lblScore.Caption + 1
        lblScoreLet.Caption = lblScoreLet + Len(txtText.Text)
        Timer1_timer
        txtText.Text = ""
    End If
End Sub

Private Sub Show_Next_Word()
    On Error Resume Next
    Randomize
    Dim aString
    Beep
    If lblRec = "" Then lblRec = "0"
    If lblRecLet = "" Then lblRecLet = "0"
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
                    
    lblText = aString(CInt(Rnd * 305))
End Sub


Private Sub WholeGame_Timer()
    On Error Resume Next
    Timer1.Enabled = False
    txtText.Locked = True
    If CInt(lblScore.Caption) > CInt(lblRec.Caption) Then
        MsgBox "Congratulations!, You typed more than last record. You Broke The Record!!!" & vbLf & "Number of words types in 1 minute: " & lblScore
        lblRec.Caption = lblScore.Caption
        lblRecLet.Caption = lblScoreLet.Caption
        SaveSetting App.Title, "THErecordOFtypingINtheGAME", "THISisTHErecordOFgame", lblScore.Caption
        SaveSetting App.Title, "THErecordOFtypingINtheGAME", "THISforMAKINGsureIFthereHASbeenANYbreakingINrecord", "True"
    Else
        MsgBox "You typed " & lblScore.Caption & " word(s) in 1 minute."
    End If
    If CInt(lblScoreLet.Caption) > CInt(lblRecLet.Caption) Then
        MsgBox "Congratulations!, You typed more letters than last record. You Broke The Record!!!" & vbLf & "Number of letters types in 1 minute: " & lblScoreLet
        lblRecLet.Caption = lblScoreLet.Caption
        SaveSetting App.Title, "THErecordOFtypingINtheGAME", "THISisTHErecordOFgameLetters", lblRecLet.Caption
        SaveSetting App.Title, "THErecordOFtypingINtheGAME", "THISforMAKINGsureIFthereHASbeenANYbreakingINrecordFORletters", "True"
    Else
        MsgBox "You typed " & lblScoreLet.Caption & " letters in 1 minute."
    End If
    WholeGame.Enabled = False
End Sub


