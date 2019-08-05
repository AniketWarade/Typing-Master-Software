VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form9 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14175
   Icon            =   "typingtest.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   8565
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   480
      Top             =   7920
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   3600
      TabIndex        =   13
      Text            =   "Type over here!"
      Top             =   9600
      Width           =   9375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8520
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4935
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"typingtest.frx":164A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   17280
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
      Begin VB.Label Label8 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   18360
      Top             =   4080
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      Caption         =   "Please press Enter after each and every word you type in the below textbox..."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   3240
      TabIndex        =   19
      Top             =   8640
      Width           =   10155
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   15960
      TabIndex        =   18
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Words Typed Incorrect:-"
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
      Left            =   12000
      TabIndex        =   17
      Top             =   1200
      Width           =   3690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Words Typed Correctly:-"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   3705
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      Caption         =   "Currently typing word from the given text..."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   1440
      TabIndex        =   15
      Top             =   6960
      Width           =   5835
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4320
      TabIndex        =   14
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   7800
      TabIndex        =   12
      Top             =   7080
      Width           =   7335
   End
   Begin VB.Label Label3 
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
      Left            =   18120
      TabIndex        =   11
      Top             =   9720
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "!!! Start Typing. All the Best !!!"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   6330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   17640
      TabIndex        =   8
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Cancel"
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
      Left            =   15840
      TabIndex        =   7
      Top             =   9720
      Width           =   1635
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, wordstyped, dur, err_word As Integer
Dim astronauts
Dim rabbits
Dim strebel
Dim yosemite
Dim fables
Dim hubble
Dim netiquette
Dim baseball
Dim franklin
Dim dna
Dim cast
Dim photo
Dim girl
Dim eighttools
Dim credits
Dim insects

Private Sub Command1_Click()
Load Form8
Form8.Show
Unload Me
End Sub


Private Sub Command2_Click()
On Error Resume Next
wordstyped = Label10.Caption
err_word = Label14.Caption


Dim gross, net, acc As Integer

gross = wordstyped / dur
gross = CInt(gross)
net = gross - (err_word * 2)
net = CInt(net)
acc = (net / gross) * 100
acc = CInt(acc)

If gross <= 0 Or net <= 0 Or acc <= 0 Then
gross = 0
net = 0
acc = 0
Form10.Label5.Caption = dur & " min"
Form10.Label6.Caption = gross & " wpm"
Form10.Label7.Caption = net & " wpm"
Form10.Label8.Caption = acc & " %"
Else
Form10.Label5.Caption = dur & " min"
Form10.Label6.Caption = gross & " wpm"
Form10.Label7.Caption = net & " wpm"
Form10.Label8.Caption = acc & " %"
End If


If acc > 79 And acc <= 100 Then
Form10.Label11.Caption = "Excellent !!!"
ElseIf acc > 59 And acc <= 79 Then
Form10.Label11.Caption = "Very Good !!!"
ElseIf acc > 39 And acc <= 59 Then
Form10.Label11.Caption = "Good !!!"
ElseIf acc > 29 And acc <= 39 Then
Form10.Label11.Caption = "Average !!!"
ElseIf acc > 19 And acc <= 29 Then
Form10.Label11.Caption = "Poor !!!"
ElseIf acc > 0 And acc <= 19 Then
Form10.Label11.Caption = "Very Poor !!!"
End If

Load Form10
Form10.Show
Unload Me
End Sub



Private Sub Form_Load()

i = -1
Label10.Caption = 0
Label14.Caption = 0



'For Setting Timer
Label6.Caption = "2"
If Form8.Combo1.ListIndex = 0 Then
Label6.Caption = "2"
ElseIf Form8.Combo1.ListIndex = 1 Then
Label6.Caption = "5"
ElseIf Form8.Combo1.ListIndex = 2 Then
Label6.Caption = "10"
ElseIf Form8.Combo1.ListIndex = 3 Then
Label6.Caption = "15"
End If
Timer1.Enabled = False
Timer2.Enabled = False
dur = Label6.Caption



'For selecting which paragraph to type
If Form8.List1.Text = "Aesob's Fables" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\fables.txt"
Fables_Next
ElseIf Form8.List1.Text = "Astronauts" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\astronauts.txt"
Astronauts_Next
ElseIf Form8.List1.Text = "Behind The Scene: Movie Credits" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\credits.txt"
Credits_Next
ElseIf Form8.List1.Text = "DNA Research -- The Human Genome Project" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\dna.txt"
Dna_Next
ElseIf Form8.List1.Text = "History of Photography" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\photo.txt"
Photo_Next
ElseIf Form8.List1.Text = "Hubble Space Telescope" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\hubble.txt"
Hubble_Next
ElseIf Form8.List1.Text = "Netiquette" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\netiquette.txt"
Netiquette_Next
ElseIf Form8.List1.Text = "Rules of Baseball (from 1889)" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\baseball.txt"
Baseball_Next
ElseIf Form8.List1.Text = "Speeding up the strategy process" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\strebel.txt"
Strebel_Next
ElseIf Form8.List1.Text = "Stinging Insects" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\insects.txt"
Insects_Next
ElseIf Form8.List1.Text = "The Eight Tools for Creating New Value" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\eight tools.txt"
Eighttools_Next
ElseIf Form8.List1.Text = "The Light Match Girl (adapted)" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\girl.txt"
Girl_Next
ElseIf Form8.List1.Text = "The Tale of Peter Rabbit" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\rabbits.txt"
Rabbits_Next
ElseIf Form8.List1.Text = "Thoughts of Benjamin Franklin" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\franklin.txt"
Franklin_Next
ElseIf Form8.List1.Text = "What is a Cast?" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\cast.txt"
Cast_Next
ElseIf Form8.List1.Text = "Yosemite National Park" Then
RichTextBox1.FileName = "C:\Typing Master Software\Lessons\yosemite.txt"
Yosemite_Next
End If


'Default timer
'Text1.Locked = True
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Text1.Text = KeyPress Then
Timer1.Enabled = True
End If
If KeyAscii = 13 Then
If Text1.Text = Label9.Caption Then
  RichTextBox1.SelColor = vbBlue
    If RichTextBox1.FileName = "C:\Typing Master Software\Lessons\astronauts.txt" Then
    Astronauts_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\rabbits.txt" Then
    Rabbits_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\strebel.txt" Then
    Strebel_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\yosemite.txt" Then
    Yosemite_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\netiquette.txt" Then
    Netiquette_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\hubble.txt" Then
    Hubble_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\baseball.txt" Then
    Baseball_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\fables.txt" Then
    Fables_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\franklin.txt" Then
    Franklin_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\dna.txt" Then
    Dna_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\cast.txt" Then
    Cast_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\photo.txt" Then
    Photo_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\girl.txt" Then
    Girl_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\eight tools.txt" Then
    Eighttools_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\credits.txt" Then
    Credits_Next
    ElseIf RichTextBox1.FileName = "C:\Typing Master Software\Lessons\insects.txt" Then
    Insects_Next
    End If
    Label10.Caption = Label10.Caption + 1
    Text1 = ""
Else
MsgBox "You have typed incorrect word!!! Please type correct word.", vbInformation, "!!! Error !!!"
Label14.Caption = Label14.Caption + 1
End If
End If
End Sub


Private Sub Timer1_timer()
On Error Resume Next
If Label8.Caption = 0 Then
Label6.Caption = Label6.Caption - 1
End If
If Label8.Caption = 0 And Label6.Caption >= 0 Then
Label8.Caption = "59"
Else
Label8.Caption = Label8.Caption - 1
End If
If Label8.Caption = 1 Then
Label8.Caption = "59"
Label6.Caption = Label6.Caption - 1
End If
If Label6.Caption = 0 And Label8.Caption = 2 Then
Label8.Caption = "0"
Label6.Caption = "0"
End If
If Label8.Caption = 0 And Label6.Caption = 0 Then
Timer1.Enabled = False
MsgBox "Your Time is Up!! Please press Enter to continue.", vbInformation, "!!!Time Up!!!"


wordstyped = Label10.Caption
err_word = Label14.Caption


Dim gross, net, acc As Integer

gross = wordstyped / dur
net = gross - (err_word * 2)
acc = (net / gross) * 100

If gross <= 0 Or net <= 0 Or acc <= 0 Then
gross = 0
net = 0
acc = 0
Form10.Label5.Caption = dur & " min"
Form10.Label6.Caption = gross & " wpm"
Form10.Label7.Caption = net & " wpm"
Form10.Label8.Caption = acc & " %"
Else
Form10.Label5.Caption = dur & " min"
Form10.Label6.Caption = gross & " wpm"
Form10.Label7.Caption = net & " wpm"
Form10.Label8.Caption = acc & " %"
End If



If acc > 79 And acc <= 100 Then
Form10.Label11.Caption = "Excellent !!!"
ElseIf acc > 59 And acc <= 79 Then
Form10.Label11.Caption = "Very Good !!!"
ElseIf acc > 39 And acc <= 59 Then
Form10.Label11.Caption = "Good !!!"
ElseIf acc > 29 And acc <= 39 Then
Form10.Label11.Caption = "Average !!!"
ElseIf acc > 19 And acc <= 29 Then
Form10.Label11.Caption = "Poor !!!"
ElseIf acc > 0 And acc <= 19 Then
Form10.Label11.Caption = "Very Poor !!!"
End If

Load Form10
Form10.Show
Unload Me
End If
End Sub



Private Sub Text1_Click()
Text1 = ""
End Sub



Private Sub Astronauts_Next()
    On Error Resume Next
    Beep
    i = i + 1
    astronauts = Array("The", "word", "'astronaut'", "derives", "from", "the", "Greek", "words", "meaning", "'star'", "and", "'sailor.'", "These", "men", "and", "women", "--Alan", "Shepherd,", "John", "Glenn,", "Jr.,", "Sally", "Ride,", "to", "name", "a", "few", "--conjure", "up", "images", "of", "bravery", "and", "adventure.", "They", "are", "modern", "heroes,", "helping", "humanity", "reach", "for", "the", "stars.", "When", "the", "space", "program", "began", "in", "1959,", "there", "were", "only", "seven", "such", "people", "in", "the", "entire", "country.", "They", "all", "were", "then", "--or", "had", "been", "in", "the", "past", "--in", "the", "armed", "forces.", "Most", "of", "them", "were", "test", "pilots,", "used", "to", "the", "dangers", "that", "came", "from", "'pushing the envelope.'", "That", "was", "only", "36", _
"years", "ago,", "and", "since", "then", "much", "has", "changed.", "Today", "The", "Americans", "who", "make", "up", "the", "shuttle", "crews", "are", "comprised", "of", "every", "race,", "creed,", "color,", "and", "gender.", "As", "of", "May", "2", "1993,", "180", "men", "and", "21", "women", "astronauts", "were", "Caucasian,", "six", "men", "and", "one", "woman", "were", "African-American,", "three", "men", "and", "one", "woman", "were", "Hispanic,", "and", "two", "men", "were", "of", "Asian", "descent.", "NASA,", "the", "National", "Aeronautics", "and", "Space", "Administration,", "chooses", "its", "astronauts", "from", "an", "increasingly", "diverse", "pool", "of", "applicants", "that", "'looks like America.'", "Thousands", "of", "applications", "come", "in", "from", "all", "over", "the", "world;", "from", "these,", "approximately", "100", "men", "and", "women", "are", "chosen", "for", "an", "intensive", "astronaut", "candidate", _
"training", "program", "every", "two", "years.", "I", "cannot", "imagine", "a", "better", "career.", "I've", "done", "more", "than", "I", "could", "ever", "have", "imagined.", "I'm", "thankful", "that", "I've", "been", "at", "the", "right", "place", "'at", "the", "right", "time,'", "says", "astronaut", "Kenneth", "S.", "Reightler.", "The", "training", "is", "demanding,", "but", "the", "study", "time", "involved", "is", "no", "longer", "than", "that", "of", "any", "other", "professional", "career", "requiring", "graduate/post-graduate", "study.", "If", "becoming", "an", "astronaut", "is", "a", "dream,", "held", "long", "and", "steadfast,", "then", "this", "labor", "will", "be", "one", "of", "love.", "The", "preparation", "for", "becoming", "an", "astronaut", "actually", "begins", "in", "elementary", "school.", _
"'It", "is", "here", "that", "the", "foundations", "are", "laid", "down", "and", "then", "built", "upon,'", "comments", "Colonel", "Charlie", "Bolden,", "Deputy", "Commandant", "of", "Midshipmen", "at", "the", "U.S.", "Naval", "Academy.", "'Start", "with", "the", "basics", "and", "get", "them", "down", "first...", "you", "can't", "do", "anything", "without", "math", "and", "science.'", "At", "this", "level,", "students", "should", "read", "everything", "they", "can", "get", "their", "hands", "on", "about", "astronauts", "and", "space", "in", "general.", "Later,", "once", "they", "have", "found", "a", "specific", "field", "of", "interest,", "they", "can", "begin", "to", "focus", "on", "that.", "But", "there's", "more", "than", "book", "learning", "involved.", _
"Other", "skills", "Bolden", "feels", "are", "integral", "to", "becoming", "an", "astronaut", "are:", "being", "able", "to", "work", "as", "a", "team", "player;", "understanding", "and", "appreciating", "your", "ethnic,", "cultural,", "and", "American", "history;", "and", "maintaining", "a", "grasp", "on", "current", "events.", "Bolden", "does", "not", "hold", "the", "opinion", "that", "America's", "young", "people", "are", "less", "equipped", "than", "their", "foreign", "counterparts.", "I", "listen", "to", "people", "say", "that", "American", "students", "can't", "do", "this", "and", "can't", "do", "that.", "I", "don't", "buy", "into", "vor", "accept", "that", "at", "all.", "Kids", "are", "just", "as", "sharp", "as", "when", "I", "came", "up.", "It's", "just", "motivation", "that", "may", "be", "lacking,", "Bolden", "said.", _
"It's", "a", "good", "idea", "to", "look", "for", "out-of-school", "opportunities", "like", "internships", "or", "apprenticeships,", "summer", "programs,", "and", "fieldwork,", "to", "help", "balance", "textbook", "learning.", "After all,", "going", "into", "space", "is", "a", "real-world", "--", "not", "a", "virtual", "--experience;", "the", "more", "experience", "you", "bring", "to", "it,", "the", "better", "you'll", "be", "able", "to", "deal", "with", "the", "demands", "of", "the", "job.", "Internships", "are", "also", "a", "good", "way", "of", "finding", "out,", "if", "your", "career", "choice", "is", "the", "right", "one", "for", "you.", "Many", "students,", "when", "faced", "with", "the", "realities", "of", "a", "particular", "field,", "decide", "it's", "not", "what", "they", "wanted,", "after", "all.", "It's", "better", "to", "decide", "early", "on", "that", "this", "is", "the", "case,", "before", "you", "devote", "a", "lot", "of", "time", "and", "effort", "to", "something", "you", "don't", _
"enjoy.", "But", "choose", "your", "activities", "wisely.", "'If", "you", "do", "things", "thinking", "that,", "'This", "will", " look", "good", "on", "a", "resume,", "or", "'I'm", "not", "going", "to", "like", "it", "but", "it'll", "help", "me", "get", "selected", "someday,", "you", "will", "do", "yourself", "a", "disservice.", "You're", "not", "going", "to", "do", "as", "well", "as", "you", "would", "at", "something", "in", "which", "you're", "interested,", "said", "Reightler.", "The", "'minimum degree requirement'", "for", "an", "astronaut", "is", "a", "Bachelor's", "degree", "from", "an", "accredited", "institution.", "Three", "years", "of", "related", "increasingly", "responsible", "professional", "experience", "must", "follow", "that", "degree.", "Most", "astronauts", "to", "date,", "however,", "have", "continued", "with", "their", "education", "to", "the", "post-graduate (Master's and Ph.D.)", "levels;", "they", "have", "been", "able", "to", _
"substitute", "that", "education", "for", "all", "or", "part", "of", "their", "work", "experience", "requirement.", "NASA", "contributes", "funds", "to", "51", "colleges", "and", "universities", "through", "its", "Space", "Grant", _
"Consortia.", "If", "you", "attend", "one", "of", "these", "institutions,", "you", "are", "ensured", "that", "the", "curriculum", "for", "space", "programs", "will", "conform", "with", "the", "guidelines", "NASA", "finds", "acceptable.", "However,", "you", "can", "attend", "any", "college", "or", "university", "that", "you", "wish,", "since", "almost", "all", "offer", "degrees", "in", "technical", "fields,", "math,", "and", "science.", "You", "should", "check", "with", "a", "guidance/college", "counselor", "--and", "NASA--", "before", "you", "go", "to", "confirm", "that", "the", "school", "will", "offer", "the", "subjects", "you", "need.", _
"Regardless", "of", "which", "school", "you", " do", "attend,", "you", "once", "again", "need", "to", "go", "for", "the", "grades.", "For", "one", "thing,", "you", "will", "need", "them", "to", "be", "accepted", "into", "a", "good", "Master's", "program.", "And", "for", "another,", "NASA", "needs", "applicants", "who", "can", "demonstrate", "success", "at", "the", "higher-education", "levels.", "Math", "and", "science", "are", "obviously", "critical", "to", "becoming", "an", "astronaut,", "but", "they", "are", "not", "the", "only", "subjects", "you", "will", "need", "to", "get", "where", "you're", "going.", "English", "-- an", "ability", "to", "communicate,", "both", "written", "and", "verbally", "-- is", "also", "vital.", "So", "is", "history,", "not", "just", "of", "the", "space", _
"program,", "but", "of", "the", "world.", "And", "Bolden,", "who", "had", "to", "wake", "up", "extra", "early", "to", "attend", "a", "class", "in", "basic", "Russian", "language", "to", "prepare", "for", "his", "mission", "with", "a", "cosmonaut,", "suggests", "that", "every", "American", "should", "be--", "at", "the", "very", "least--", "bilingual.", "A", "significant", "number", "of", "applicants", "do", "not", "meet", "medical", "standards,", "and", "still", "others", "withdraw", "after", "gaining", "a", "complete", "understanding", "of", "the", "very", "demanding", "job.", "Based", "on", "the", "information", "collected", "during", "this", "investigation,", "ASB", "will", "choose", "its", "final", "candidates", "and", " pass", "that", "recommendation", "on", "to", "the", "NASA", "Administrator;", "he", "or", "she", "will", "make", "the", "final", "choices.", _
"And", "if", "an", "applicant", "has", "hopes", "for", "a", "large", "salary", "-- say,", "comparable", "to", "those", "in", "the", "private", "sector", "-- they", "should", "look", "elsewhere.", "Astronauts", "begin", "their", "salary", "in", "accordance", "with", "the", "U.S.", "Government", "pay", "scale", "at", "GS-11", "status", "(approximately $39,000)", "and", "top", "off", "at", "GS-14", "status", "(approximately $78,000).")
  Label9 = astronauts(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find astronauts(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub


Private Sub Rabbits_Next()
    On Error Resume Next
    Beep
    i = i + 1
    rabbits = Array("Once", "upon", "a", "time", "there", "were", "four", "little", "rabbits.", "Their", "names", "were:", "Flopsy,", "Mopsy,", "Cottontail,", "and", "Peter.", "They", "lived", "with", "their", "Mother", "under", "the", "root", "of", "a", "very", "big", "fir", "tree.", "'Now,", "children,'", "said", "old", "Mrs.", "Rabbit", "one", "morning,", "'you", "may", "go", "into", "the", "fields", "or", "down", "the", "road.", "But", "don't", "go", "into", "Mr.", "McGregor's", "garden.", "Your", "father", "had", "an", "accident", "there.", "Mrs.", "McGregor", "put", "him", "in", "a", "pie.'", "The", "children", "gasped.", _
"Mrs.", "Rabbit", "picked", "up", "her", "basket", "and", "her", "umbrella.", "She", "gave", "them", "all", "a", "big", "smile.", "'Now run along, children, and don't get into mischief. I am going out.'", "Then", "she", "went", "out", "the", "door.", "She", "went", "to", "the", "baker's.", "She", "bought", "a", "loaf", "of", "brown", "bread", "and", "five", "sweet", "buns.", "They", "would", "have", "a", "nice", "dinner.", _
"Flopsy,", "Mopsy,", "and", "Cottontail", "were", "good", "little", "bunnies.", "They", "went", "down", "the", "lane", "to", "find", "some", "blackberries.", "But", "Peter", "was", "very", "naughty.", "He", "ran", "right", "away", "to", "Mr.", "McGregor's", "garden", "and", "squeezed", "under", "the", "gate!", "First", "he", "ate", "some", "lettuce", "and", "some", "carrots.", "Then", "he", "nibbled", "on", "some", "radishes.", "And", "then,", "feeling", "rather", "sick,", "he", "went", "to", "look", "for", "some", "parsley.", "But", "when", "he", "came", "to", "the", "end", "of", "a", "cucumber", "frame,", "whom", "should", "he", "meet", "but", "Mr.", "McGregor!", _
"Mr.", "McGregor", "was", "on", "his", "hands", "and", "knees", "planting", "young", "cabbages,", "but", "he", "jumped", "up", "and", "ran", "after", "Peter,", "waving", "a", "rake", "and", "calling", "out,", "'Stop thief!'", "Peter", "was", "very", "scared.", "He", "rushed", "all", "over", "the", "garden.", "He", "had", "forgotten", "the", "way", "back", "to", "the", "gate.", "He", "lost", "one", "of", "his", "shoes", "among", "the", "cabbages,", "and", "the", "other", "shoe", "near", "the", "potatoes.", "Without", "shoes,", "he", "dropped", "to", "all", "four", "legs", "so", "he", "could", "run", "faster.", "He", "might", "have", "gotten", "completely", "away", "if", "he", "had", "not", "unfortunately", "run", "into", "a", "gooseberry", "fence.", "He", "was", "wearing", "a", "blue", "jacket", "with", "brass", "buttons,", "and", "the", "buttons", "got", "caught", "in", "the", "net.", _
"Peter", "gave", "himself", "up", "for", "lost,", "and", "he", "shed", "big", "tears.", "But", "some", "friendly", "sparrows", "heard", "him", "crying.", "They", "flew", "down", "to", "him", "and", "begged", "him", "to", "keep", "trying", "to", "escape.", "'You can do it, Peter,'", "they", "said.", "Mr.", "McGregor", "had", "a", "big", "butterfly", "net", "in", "his", "hands.", "He", "came", "up", "behind", "Peter", "and", "tried", "to", "scoop", "him", "up.", "But", "Peter", "wriggled", "out", "just", "in", "time,", "leaving", "his", "jacket", "behind", "him.", "He", "rushed", "into", "the", "toolshed", "and", "saw", "a", "large", "watering", "can.", "It", "looked", "big", "and", "beautiful", "and", "a", "great", "place", "in", "which", "to", "hide.", "He", "jumped", "inside.", "It", "wasn't", "until", "he", "had", "that", "he", "realized", "it", "had", "water", "in", "it.", _
"Mr.", "McGregor", "was", "quite", "sure", "Peter", "was", "somewhere", "in", "the", "shed.", "Perhaps", "he", "was", "hidden", "underneath", "a", "flower", "pot.", "He", "turned", "each", "one", "over", "carefully,", "looking", "under", "them", "one", "by", "one,", "but", "Peter", "wasn't", "there.", "It", "was", "then", "that", "Peter", "sneezed,", "'Ahchoo!'", "Mr.", "McGregor", "was", "after", "him", "in", "no", "time.", "The", "net", "was", "no", "good", "in", "such", "a", "small", "space,", "so", "he", "tried", "to", "step", "on", "Peter.", "Peter", "was", "fast,", "though,", "and", "he", "jumped", "out", "of", "a", "window,", "upsetting", "three", "plants.", "The", "window", "was", "too", "small", "for", "Mr.", "McGregor", "to", "fit", "through.", "And", "he", "was", "tired", "of", "running", "after", "Peter.", "He", "went", "back", "to", "his", "work.", _
"Outside,", "Peter", "sat", "down", "to", "rest.", "He", "was", "out", "of", "breath", "and", "trembling", "with", "fright.", "And", "he", "still", "didn't", "have", "any", "idea", "which", "way", "to", "go.", "Also,", "he", "was", "very", "damp", "from", "sitting", "in", "that", "can.", "His", "fur", "was", "matted,", "and", "he", "was", "cold.", "He", "sat", "for", "a", "while,", "then", "he", "began", "to", "wander", "about,", "going", "hop-hop-hop.", "He", "didn't", "go", "very", "fast,", "and", "he", "looked", "all", "around", "to", "see", "where", "he", "was.", "He", "found", "a", "door", "in", "a", "wall,", "but", "it", "was", "locked.", "There", "was", "no", "room", "for", "a", "fat", "little", "rabbit", "to", "squeeze", "underneath.", _
"An", "old", "mouse", "was", "running", "in", "and", "out", "over", "the", "stone", "doorstep,", "carrying", "peas", "and", "beans", "to", "her", "family", "in", "the", "woods.", "Peter", "asked", "her", "the", "way", "to", "the", "gate.", "But", "she", "had", "such", "a", "large", "pea", "in", "her", "mouth", "that", "she", "could", "not", "answer.", "She", "only", "shook", "her", "head", "at", "him.", "Peter", "began", "to", "cry", "again.", _
"Then", "he", "tried", "to", "find", "his", "way", "straight", "across", "the", "garden.", "But", "he", "became", "more", "and", "more", "confused.", "He", "didn't", "know", "this", "garden,", "and", "he", "didn't", "know", "which", "trail", "led", "out.", "After", "a", "while,", "he", "came", "to", "a", "pond", "where", "Mr.", "McGregor", "filled", "his", "watering", "cans.", "A", "white", "cat", "was", "staring", "at", "some", "goldfish", "in", "the", "pond.", "She", "sat", "very,", "very", "still.", "But", "now", "and", "then", "the", "tip", "of", "her", "tail", "twitched", "as", "if", "it", "were", "alive.", "Peter", "thought", "it", "best", "to", "go", "away", "without", "speaking", "to", "her.", "He", "had", "heard", "about", "cats", "from", "his", "cousin,", "little", "Benjamin", "Bunny.", "He", "didn't", "think", "she'd", "be", "very", "friendly.", _
"He", "went", "back", "the", "way", "he", "came,", "towards", "the", "toolshed.", "Suddenly,", "quite", "close", "to", "him,", "he", "heard", "the", "noise", "of", "a", "hoe--scr-r-ritch,", "scr-r-ratch.", "Peter", "dove", "under", "the", "bushes,", "shaking", "with", "fear.", "He", "waited", "and", "waited,", "but", "nothing", "happened.", "He", "poked", "his", "nose", "out,", "then", "his", "front", "paws,", "then", "his", "back", "paws.", "He", "climbed", "on", "top", "of", "a", "wheelbarrow", "and", "peeped", "over.", "The", "first", "thing", "he", "saw", "was", "Mr.", "McGregor", "hoeing", "openions.", "His", "back", "was", "turned", "towards", "Peter", "and", "beyond", "him", "was", "the", "gate!", "Oh,", "how", "wonderful!", _
"Peter", "got", "down", "very", "quietly", "off", "the", "wheelbarrow.", "He", "started", "running", "as", "fast", "as", "he", "could", "go,", "along", "a", "straight", "walk", "behind", "some", "berry", "bushes.", "Mr.", "McGregor", "caught", "sight", "of", "him", "at", "the", "corner,", "but", "Peter", "did", "not", "care.", "He", "slipped", "underneath", "the", "gate", "and", "out", "into", "the", "wood", "outside", "the", "garden.", "Safe", "at", "last!", _
"Peter", "never", "stopped", "running", "or", "looked", "behind", "him", "till", "he", "got", "home", "to", "the", "big", "fir", "tree.", "He", "was", "so", "tired", "that", "he", "flopped", "down", "on", "the", "nice", "soft", "sand", "on", "the", "floor", "of", "the", "rabbit", "hole", "and", "shut", "his", "eyes.", "His", "mother", "was", "busy", "cooking.", "She", "wondered", "what", "he", "had", "done", "with", "his", "clothes.", "She", "didn't", "know", "that", "Mr.", "McGregor", "hung", "up", "the", "little", "jacket", "and", "the", "shoes", "as", "a", "scarecrow", "to", "frighten", "the", "birds.", _
"Peter", "did", "not", "feel", "very", "well", "during", "the", "rest", "of", "the", "evening.", "His", "mother", "put", "him", "to", "bed,", "made", "him", "a", "hot", "cut", "of", "tea,", "and", "gave", "him", "a", "spoonful", "of", "bad-tasting", "medicine.", "'Drink", "it", "up,", "Peter.'", "Then", "she", "kissed", "him", "goodnight.", "But", "Flopsy,", "Mopsy,", "and", "Cottontail", "felt", "just", "fine.", "They", "had", "bread,", "and", "milk,", "and", "blackberries", "for", "supper.")
  Label9 = rabbits(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find rabbits(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Strebel_Next()
    On Error Resume Next
    Beep
    i = i + 1
    strebel = Array("To", "adapt,", "companies", "have", "to", "speed", "up", "their", "strategy", "processes,", "but", "in", "so", "doing,", "large", "companies", "in", "rapidly", "changing", "environments", "face", "major", "challenges.", "These", "challenges", "mainly", "have", "to", "do", "with", "size,", "in", "particular", "the", "distance", "between", "the", "top", "of", "the", "organization", "and", "the", "front-line,", "as", "well", "as", "the", "many", "different", "industry", "segments", "and", "initiatives", "present", "in", "large", "companies", "This", "'Perspectives", "for", "Managers'", "describes", "the", "challenges", "and", "suggests", "a", "pattern", "in", "the", "way", "certain", "successful,", "large,", "fast-moving", "companies", "are", "dealing", "with", "them,", "as", "well", "as", "certain", "advantages", "they", "have", "over", "smaller", "companies.", _
"The", "first", "challenge", "for", "large", "companies", "is", "to", "produce", "the", "bottom-up", "experimentation", "needed", "for", "adapting", "to", "a", "fast", "moving", "environment.", "It", "is", "well", "known", "that", "the", "mindset,", "politics,", "culture", "and", "systems", "that", "supported", "past", "success", "are", "a", "heavy", "legacy", "in", "large", "companies.", "Unless", "radically", "modified,", "this", "legacy", "quickly", "suffocates", "new", "rule-breaking", "initiatives.", "The", "literature", "on", "innovation", "points", "to", "the", "importance", "of", "creating", "a", "culture", "that", "tolerates", "honest", "failure,", "the", "open", "exchange", "of", "business", "ideas,", "the", "rotation", "of", "talent", "through", "project", "teams,", "and", "multi-channel", "access", "to", "resources.", "Not", "only", "is", "this", "difficult", "to", "put", "in", "place,", "but", "it", "is", "often", "not", "enough", "to", "produce", "a", _
"rapid", "flow", "of", "initiatives", "in", "the", "face", "of", "a", "conservative", "mainstream", "culture.", "The", "second", "challenge", "is", "the", "distance", "between", "the", "top", "and", "the", "front-line", "in", "large", "companies.", "Getting", "feedback", "into", "strategic", "thinking", "at", "the", "top", "from", "front-line", "experimentation,", "customers", "and", "suppliers", "is", "no", "easy", "matter,", "nor", "is", "the", "communication", "of", "the", _
 "strategic", "roadmap", "to", "the", "front-line", "and", "the", "shaping", "of", "incentives", "and", "support", "systems.", "All", "of", "this", "is", "complicated", "in", "global", "companies", "by", "language", "and", "differences", "in", "national", "culture.", "The", "time", "lag", "between", "a", "change", "in", "direction", "at", "the", "top", "and", "its", "full", "comprehension", "throughout", "the", "company", "makes", "it", "impossible", "to", "shift", "direction", "rapidly", "and", "frequently", "without", "disorienting", "people.", _
"This", "is", "part", "of", "a", "third", "challenge,", "the", "difficulty", "of", "reorganizing", "rapidly", "to", "exploit", "new", "opportunities.", "Large", "companies", "are", "commonly", "organized", "around", "business", "units", "with", "bottom", "line", "responsibility.", "The", "push", "for,", "and", "the", "rewards", "based", "on,", "business", "unit", "performance", "result", "in", "competing", "business", "'silos',", "between", "which", "there", "is", "little", "cooperation.", "Reorganization", "usually", "involves", "shifting", "the", "focus", "of", "bottom", "line", "responsibility", "to", "a", "dimension", "that", "cuts", "across", "the", "original", "organizational", "'silos'.", "The", "resistance", "in", "the", "old", "'silos'", "to", "managing", "in", "the", "new", "dimension,", "plus", "the", "communication", "difficulties", "already", "mentioned,", "make", "rapid", "reorganization", "in", "large", "companies", "very", "difficult.", _
"A", "fourth", "challenge", "is", "to", "achieve", "focus", "amid", "a", "multiplicity", "of", "competing", "initiatives", "that", "emerge", "once", "the", "organization", "is", "energized.", "Large", "companies", "have", "a", "larger", "portfolio", "of", "businesses,", "plus", "the", "resources", "to", "support", "many", "more", "initiatives", "than", "smaller", "companies.", "In", "addition,", "there", "are", "more", "top", "managers,", "each", "with", "their", "own", "favorite", "projects.", "This", "often", "results", "in", "an", "overload", "of", "change", "and", "growth", "initiatives,", "causing", "confusion", "at", "the", "front-line.", "On", "top", "of", "this,", "the", "politics", "between", "senior", "managers", "from", "different", "businesses", "with", "different", "perspectives", "delay", "critical", "choices", "and", "further", "aggravate", "the", "lack", "of", "focus.")
  Label9 = strebel(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find strebel(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
  End Sub
  
  

Private Sub Yosemite_Next()
    On Error Resume Next
    Beep
    i = i + 1
    yosemite = Array("Yosemite", "National", "Park", "is", "located", "in", "the", "Sierra", "Nevada", "Mountains", "of", "central", "California,", "a", "setting", "so", "spectacular", "and", "awe-inspiring", "that", "John", "Muir,", "a", "naturalist", "and", "the", "Sierra", "Club", "founder,", "called", "it", "'a", "landscape...", "that", "after", "all", "my", "wanderings", "still", "appears", "as", "the", "most", "beautiful", "I", "have", "ever", "beheld'.", _
"Over", "millions", "of", "years,", "earthquakes,", "glaciers,", "and", "other", "forces", "of", "nature", "have", "left", "indelible", "footprints", "on", "this", "landscape,", "footprints", "still", "visible", "in", "the", "towering", "granite", "cliffs,", "thundering", "waterfalls,", "steep", "mountains,", "and", "deep", "alpine", "lakes,", "and", "they", "all", "beckon", "visitors", "to", "explore,", "climb,", "photograph,", "and", "experience", "their", "grandeur.", _
"Yosemite's", "geological", "history", "has", "been", "evolving", "for", "some", "500", "million", "years,", "from", "the", "area's", "initial", "position", "on", "the", "ocean", "floor,", "to", "its", "later", "incarnations", "as", "gentle,", "rolling", "hills,", "and", "then", "the", "steep", "Sierra", "Nevada", "mountain", "range", "replete", "with", "deep", "river", "canyons.", _
"Three", "million", "years", "ago,", "the", "ice", "age", "brought", "glaciers", "that", "scraped", "and", "carved", "the", "valleys", "and", "canyons", "with", "such", "force", "that", "the", "remaining", "granite", "still", "shows", "the", "direction", "of", "the", "glacial", "movement,.", "Those", "same", "glaciers", "created", "massive", "rock", "formations", "that", "present", "some", "of", "the", "most", "difficult", "and", "popular", "climbing", "challenges", "in", "the", "world", "for", "today's", "enthusiastic", "rock", "climbers.", _
"When", "the", "last", "glacier", "finally", "melted", "10,000", "years", "ago,", "rock", "debris", "dammed", "the", "valley", "and", "created", "Lake", "Yosemite,", "while", "tributary", "creeks", "plummeted", "off", "sheer", "cliffs", "and", "gave", "birth", "to", "the", "park's", "famed", "waterfalls.", "Sediment", "continued", "to", "fill", "the", "lake", "until", "it", "eventually", "formed", "the", "present", "valley", "floor.", _
"Yosemite's", "geological", "evolution", "continues", "today", "as", "Mirror", "Lake", "slowly", "fills", "with", "sediment", "in", "much", "the", "same", "way", "as", "Lake", "Yosemite", "did.", "In", "1996,", "the", "Happy", "Isles", "rockslide", "sent", "80,000", "tons", "of", "rock", "to", "the", "valley", "floor", "at", "160", "mph", "--", "proof", "that", "powerful", "forces", "are", "still", "at", "work.", _
"The", "park", "has", "a", "rich", "organizational", "history,", "as", "well.", "The", "area's", "first", "residents", "were", "Native", "Americans,", "who", "inhabited", "the", "region", "perhaps", "as", "far", "back", "as", "10,000", "years", "ago.", "The", "most", "recent", "tribe", "was", "the", "Miwok,", "who", "called", "Yosemite", "Valley", "Ahwahnee,", "the", "place", "of", "the", "gaping", "mouth,", "and", "who", "called", "themselves", "Ahwahneechee.", _
"The", "discovery", "of", "gold", "in", "California's", "foothills", "brought", "the", "first", "non-native", "settlers", "to", "the", "region", "around", "1850;", "dazzled", "by", "the", "beauty", "and", "grandeur", "of", "the", "landscape,", "they", "quickly", "spread", "the", "word", "of", "its", "existence,", "and", "in", "1855,", "the", "first", "tourists", "came", "to", "visit.", "About", "the", "same", "time,", "Galen", "Clark", "homesteaded", "to", " Yosemite", "Valley", "and", "began", "his", "crusade", "to", "protect", "the", "area.", "Clark", "so", "fiercely", "and", "loyally", "protected", "Yosemite", "until", "his", "death", "at", "age", "96", "that", "he", "became", "known", "as", "the", "'Guardian", "of", "Yosemite.'", _
"In", "1854,", "Abraham", "Lincoln", "signed", "the", "Yosemite", "Grant", "that", "set", "aside", "Yosemite", "Valley", "and", "the", "Mariposa", "Grove", "of", "Giant", "Sequoias", "as", "the", "first", "state", "park.", "In", "1864,", "Yosemite", "became", "the", "first", "territory", "ever", "set", "aside", "by", "the", "United", "States", "Congress", "for", "public", "use", "and", "protection;", "however,", "it", "took", "another", "thirty-six", "years", "and", "the", "efforts", "of", "Muir,", "Theodore", "Roosevelt,", "and", "other", "influential", "conservationists", "to", "bring", "all", "the", "lands", "of", "today's", "National", "Park", "under", "federal", "protection.", _
"In", "1889,", "Robert", "Underwood", "Johnson,", "editor", "of", "Century", "Magazine,", "and", "John", "Muir", "had", "growing", "concerns", "about", "the", "devastating", "effects", "of", "sheep", "grazing", "in", "the", "high", "country,", "and", "they", "launched", "a", "successful", "campaign", "to", "persuade", "Congress", "to", "set", "aside", "this", "area", "as", "a", "national", "park.", "On", "October", "1,", "1890,", "Congress", "set", "aside", "more", "than", "1,500", "square", "miles", "of", "reserved", "forest", "lands,", "and", "today", "94.5%", "of", "the", "park's", "1,170", "square", "mile", "area", "is", "dedicated,", "as", "true", "wilderness", "for", "all", "who", "visit", "to", "enjoy.", _
"At", "Muir's", "invitation,", "Theodore", "Roosevelt", "visited", "Yosemite", "to", "experience", "its", "splendor", "first-hand.", "By", "all", "accounts,", "Roosevelt", "was", "so", "impressed", "he", "became", "instrumental", "in", "the", "return", "of", "Yosemite", "Valley", "and", "Mariposa", "Grove", "to", "federal", "protection", "in", "1906.", "In", "1932,", "the", "Wawona", "Basin,", "including", "the", "Wawona", "Hotel", "and", "golf", "course,", "were", "added", "to", "the", "National", "Park", "to", "bring", "it", "to", "its", "current", "1,170", "square", "miles.", "The", "classic", "beauty", "and", "fascinating", "ecosystems", "of", "the", "Sierra", "Nevada", "mountain", "range", "are", "well", _
"represented", "in", "Yosemite", "National", "Park,", "where", "almost", "95%", "of", "the", "park", "is", "designated", "wilderness.", "A", "place", "retaining", "its", "primeval", "character", "and", "influence,", "this", "wilderness", "area", "provides", "opportunities", "for", "solitude", "and", "unconfined", "types", "of", "experience.", "The", "lands", "within", "the", "wilderness", "range", "in", "elevation", "from", "290", "feet", "to", "13,114", "feet,", "the", "latter", "at", "the", "summit", "of", "Mt.", "Lyell.", "And", "with", "over", "800", "miles", "of", "trails", "to", "hike,", "visitors", "will", "find", "a", "spectrum", "of", "opportunities,", "of", "natural", "systems,", "and", "of", "levels", _
"of", "use.", "There", "are", "four", "regions", "within", "the", "park.", "The", "first", "region", "is", "Yosemite", "Valley.", "Beautiful", "and", "majestic,", "Yosemite", "Valley", "boasts", "some", "of", "the", "most", "popular", "trailheads", "in", "the", "park.", "These", "strenuous", "trails", "lead", "the", "hiker", "up", "the", "seemingly", "sheer", "granite", "walls", "which", "form", "the", "Valley,.", "To", "reach", "the", "top", "of", "such", "spectacles", "as", "Half", "Dome,", "Yosemite", "Falls,", "Glacier", "Point,", "and", "El", "Capitan,", "the", "hiker", "will", "encounter", "steep", "terrain,", "switchbacks,", "and", "rapidly", "changing", "weather", "conditions,.", "Spectacular", "views", _
"from", "the", "rim", "of", "the", "valley", "reward", "the", "hiker", "and", "quickly", "erase", "memories", "of", "any", "hardships", "suffered", "en", "route.", "In", "winter", "and", "early", "spring,", "nearly", "all", "of", "the", "trails", "in", "Yosemite", "Valley", "remain", "open,", "although", "some", "have", "winter", "routes", "designated.", "In", "addition", "to", "serving", "the", "hardy", "day-hiker,", "these", "trailheads", "provide", "access", "for", "seasoned", "winter", "enthusiasts", "into", "Yosemite's", "snow-covered", "wilderness.", _
"The", "second", "region", "consists", "of", "Tuolumne", "Meadows", "and", "Tioga", "Road.", "Tuolumne", "Meadows", "is", "located", "at", "over", "8,600", "feet", "in", "elevation.", "From", "Tuolumne,", "numerous", "trails", "lead", "hikers", "to", "lakes,", "meadows,", "and", "beautiful", "river", "canyons.", "More", "strenuous", "overnight", "backpacking", "trips", "allow", "the", "adventurer", "access", "to", "the", "northernmost", "reaches", "of", "the", "Park,", "the", "area", "between", "Tuolumne", "Meadows", "and", "Yosemite", "Valley,", "or", "along", "extended", "wilderness", "routes", "such", "as", "the", "John", "Muir", "Trail", "or", "the", "Pacific", "Crest", "Trail.", _
"Because", "of", "its", "elevation,", "temperatures", "in", "Tuolumne", "average", "15", "to", "20", "degrees", "cooler", "than", "in", "Yosemite", "Valley", "itself,.", "Snow", "is", "not", "uncommon", "as", "late", "as", "June", "or", "as", "early", "as", "September.", "Peaks,", "such", "as", "Mt,.", "Dana", "and", "Mt.", "Lyell,", "retain", "snow", "throughout", "the", "summer", "and", "provide", "breathtaking", "views", "from", "their", "bases", "and", "summits,.", "Once", "the", "Tioga", "Road", "closes", "for", "the", "season", "(generally in early November),", "this", "area", "is", "accessible", "only", "to", "well-experienced", "winter", "enthusiasts", "with", "snow", "shoes", "or", "skis.", _
"Tioga", "Road", "is", "a", "scenic", "45", "mile", "drive", "that", "covers", "almost", "4,000", "feet", "of", "elevation", "change.", "It", "begins", "at", "Crane", "Flat,", "travels", "through", "Tuolumne", "Meadows,", "and", "then", "climbs", "over", "Tioga", "Pass.", "The,", " road", "is", "open", "to", "vehicles", "from", "late", "May", "or", "June", "(weather permitting)", "until", "the", "first", "major", "snow", "storm", "after", "November", "1.", "During", "this", "1", "1/2", "hour", "drive,", "visitors", "pass", "through", "meadows", "and", "forests,", "lakes,", "and", "granite", "domes.", _
"Day", "hikes", "from", "the", "Tioga", "Road", "are", "abundant.", "For", "backpackers,", "there", "are", "also", "numerous", "trailheads.", "Some", "trails", "provide", "one", "way", "trips", "into", "Yosemite", "Valley,.", "Others,", "less", "traveled", "but", "no", "less", "scenic,", "head", "into", "the", "northern", "part", "of", "the", "park.", "These", "trails", "lead", "through", "forested", "areas,", "deep", "canyons,", "past", "lakes,", "and", "finally", "above", "the", "tree-line.", "In", "winter", "and", "early", "spring,", "the", "snow-covered", "Tioga", "Road", "serves", "as", "an", "ungroomed", "cross-country", "ski", "route", "for", "the", "adventurous", "and", "seasoned", "winter", "camper.", _
"The", "third", "region", "is", "the", "Hetch", "Hetchy", "Reservoir,", "located", "in", "the", "northwest", "part", "of", "the", "park;", "it", "serves", "as", "the", "portal", "to", "many", "beautiful,", "but", "less-well-traveled", "areas", "in", "Yosemite.", "The", "Hetch", "Hetchy", "Road", "normally", "stays", "open", "year-round,", "and", "the", "reservoir", "itself", "lies", "at", "a", "relatively", "low", "3,900", "feet,", "making", "this", "a", "good", "area", "for", "spring", "and", "fall", "wilderness", "travel.", _
"High", "temperatures", "prevail", "along", "the", "trail", "during", "the", "summer", "months,", "but", "this", "is", "a", "small", "price", "to", "pay", "for", "the", "breathtaking", "areas", "that", "can", "be", "reached", "from", "here.", "Several", "lakes", "and", "popular", "valleys", "are", "all", "located", "within", "15", "miles", "of", "the", "reservoir,", "and", "hikers", "may", "elect", "to", "begin", "longer", "trips", "at", "this", "point,", "either", "toward", "Tuolumne", "Meadows", "or", "the", "northernmost", "reaches", "of", "the", "park.", _
"The", "fourth", "region", "contains", "Wawona", "and", "Glacier", "Point,.")
  Label9 = yosemite(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find yosemite(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
  
End Sub









Private Sub Netiquette_Next()
    On Error Resume Next
    Beep
    i = i + 1
netiquette = Array("Do", "you", "ever", "send", "or", "receive", "e-mails?", "Are", "you", "on", "the", "Internet", "a", "lot?", "Do", "you", "go", "to", "chat", "rooms?", "Did", "you", "know", "there", "are", "rules", "of", "behavior", "for", "all", "of", "these?", "The", "rules", "are", "called", "Netiquette.", "They", "can", "be", "applied", "to", "almost", "every", "situation--", "real", "life,", "as", "well", "as", "cyberspace.", "These", "rules", "might", "seem", "kind", "of", "obvious", "to", "you,", "but", "they're", "important.", "They", "will", "help", "you", "and", "your", "friends", "Internet", "experience", "be", "productive,", "as", "well", "as", "fun.", _
"Don't", "break", "any", "laws.", "Cyberspace", "is", "an", "extension", "of", "the", "real", "world.", "Just", "because", "you're", "dealing", "with", "computers", "and", "not", "people", "(at least not directly),", "it", "doesn't", "mean", "the", "law", "no", "longer", "applies.", "If", "it's", "illegal", "in", "the", "real", "world,", "it's", "probably", "illegal", "in", "cyberspace.", _
"Be", "polite.", "You've", "certainly", "heard", "the", "saying", "'Do", "unto", "others", "as", "you", "would", "have", "them", "do", "unto", "you.'", "It's", "true", "not", "only", "when", "you're", "talking", "to", "someone", "face-to-face,", "but", "also", "when", "you're", "talking", "to", "them", "on", "the", "Web.", "Being", "polite", "doesn't", "just", "include", "what", "you", "say.", "It", "also", "covers", "how", "you", "say", "it,", "and", "to", "whom.", "If", "you", "use", "humor", "or", "sarcasm", "in", "your", "writing,", "be", "sure", "the", "recipient", "is", "going", "to", "like", "your", "sense", "of", "humor.", "Otherwise,", "you", "might", "offend", "the", "person.", "And", "never", "send", "an", "e-mail", "when", "you're", "angry", "or", "upset!", "It", "will", "probably", "show", "in", "your", "message,", "and", "you", "may", "hurt", "someone", "else's", "feelings.", _
"Be", "careful", "of", "flames", "and", "flame", "wars.", "A", "'flame'", "refers", "to", "any", "insulting", "message.", "Nearly", "everyone", "receives", "a", "flame", "sooner", "or", "later.", "The", "question", "becomes", "what", "to", "do", "with", "them.", "The", "best", "move", "is", "to", "just", "ignore", "them.", "If,", "however,", "you", "feel", "the", "flame", "is", "worth", "a", "response,", "your", "reply", "may", "start", "a", "flame", "war,", "where", "the", "insults", "continue", "back", "and", "forth.", "In", "these", "situations,", "everybody", "loses.", _
"Be", "patient.", "Everybody", "was", "new", "to", "the", "Internet", "at", "one", "time", "or", "another!", "Be", "patient", "with", "those", "who", "are", "new.", "If", "someone", "shows", "poor", "Netiquette,", "don't", "get", "angry.", "Politely", "respond", "and", "tell", "them", "of", "their", "error.", "Usually,", "the", "other", "person", "will", "be", "thankful", "for", "the", "advice.", "You", "may", "even", "find", "that", "you", "have", "made", "a", "new", "friend.", "Sometimes", "you'll", "bump", "into", "someone", "that", "you", "just", "don't", "see", "eye", "to", "eye", "with.", "In", "those", "cases,", "all", "you", "can", "do", "is", "agree", "to", "disagree.", _
"Be", "concise", "and", "accurate.", "E-mail", "wasn't", "intended", "for", "carrying", "on", "long", "talks.", "E-mail", "is", "about", "convenience.", "Keep", "your", "messages", "brief", "and", "to", "the", "point.", "If", "you", "write", "long", "messages", "that", "drift", "from", "one", "subject", "to", "the", "next,", "your", "reader's", "attention", "will", "drift", "as", "well.", "Also,", "take", "the", "time", "to", "delete", "long", "header", "text", "when", "you're", "forwarding", "or", "replying", "to", "an", "e-mail.", "This", "shows", "the", "recipient", "that", "you", "know", "their", "time", "is", "valuable,", "too.", "And", "above", "all,", "proofread", "your", "message", "for", "spelling", "and", "grammar", "before", "you", "send", "it.", "In", "cyberspace,", "people", "only", "know", "you", "from", "what", "you", "type.", "If", "you're", "sloppy", "with", "the", "basics,", "how", "can", "anyone", "trust", "you", "when", "it's", "really", "important?", _
"Always", "fill-in", "the", "subject", "line!", "When", "you", "start", "receiving", "dozens", "of", "e-mails", "a", "day,", "you'll", "come", "to", "appreciate", "people", "who", "use", "subject", "lines.", "The", "subject", "line", "is", "the", "only", "label", "you", "have", "to", "identify", "each", "of", "the", "messages", "in", "your", "mailbox.", "Use", "something", "appropriate,", "such", "as", "the", "topic", "sentence", "or", "the", "main", "idea.", "Be", "sure", "to", "keep", "it", "brief,", "though,", "since", "most", "e-mail", "clients", "only", "show", "the", "first", "20-30", "characters", "of", "the", "subject", "line", "anyway.", _
"Be", "careful", "with", "formatting.", "Remember", "that", "not", "everyone", "uses", "the", "same", "software", "to", "read", "their", "e-mail.", "Just", "because", "you", "can", "center", "your", "text", "and", "make", "it", "green", "and", "boldface", "doesn't", "mean", "your", "recipient", "can", "too!", "Some", "e-mail", "clients", "make", "all", "messages", "left-justified", "plain", "text,", "no", "matter", "how", "they", "were", "sent.", "If", "it's", "important", "to", "use", "formatting,", "make", "sure", "the", "recipient", "can", "read", "it", "first.", _
"Think", "about", "signature", "files.", "These", "files", "are", "text", "files", "that", "you", "can", "include", "with", "each", "message.", "They", "provide", "information", "about", "you,", "the", "sender.", "Keep", "them", "short,", "less", "than", "10", "lines.", "Long", "ones", "make", "threads", "hard", "to", "read", "because", "of", "the", "extra", "text", "separating", "each", "message.", "Signature", "files", "usually", "include", "the", "following", "information:", _
"Name", _
"E-mail", "Address", _
"Homepage", "Address", _
"Character", "Quote", _
"A", "character", "quote", "is", "that", "cute", "little", "phrase", "or", "saying", "you", "often", "see", "at", "the", "end", "of", "a", "signature", "file.", "It", "can", "be", "funny", "or", "serious.", "It", "can", "be", "a", "favorite", "saying", "of", "yours.", "Perhaps", "you", "have", "a", "line", "from", "a", "movie", "that", "you", "really", "like.", "It", "is", "meant", "to", "let", "your", "reader", "know", "a", "little", "something", "extra", "about", "you.", _
"Use", "care", "when", "attaching", "files.", "Attachments", "are", "an", "easy", "way", "to", "share", "programs,", "graphics,", "sounds,", "or", "any", "other", "kind", "of", "file", "you", "find", "on", "your", "computer.", "If", "you", "intend", "to", "send", "someone", "a", "large", "file", "(say, greater than 20K)", "you", "should", "contact", "them", "and", "ask", "their", "permission", "first.", "There", "are", "three", "reasons", "for", "this.", "First,", "you", "don't", "even", "know", "if", "they", "have", "e-mail", "software", "that", "can", "handle", "attached", "files.", "Second,", "large", "files", "(we're", "talking", "megabytes", "here)", "can:", _
"completely", "fill", "someone's", "e-mail", "box.", "Then", "they", "can't", "receive", "any", "more", "e-mail.", "Third,", "some", "e-mail", "programs", "limit", "the", "size", "of", "file", "attachments.", "Your", "e-mail", "message", "may", "not", "get", "sent", "if", "the", "file", "is", "too", "large,", "or", "it", "may", "bounce", "back", "to", "you.", "Graphics", "and", "sound", "files", "tend", _
"to", "be", "quite", "large,", "so", "send", "these", "with", "care.", _
"Think", "about", "distribution", "lists.", "You", "may", "want", "to", "send", "an", "e-mail", "to", "several", "people", "at", "once.", "The", "best", "way", "to", "do", "this", "is", "to", "use", "a", "list.", "Start", "a", "blank", "text", "file.", "Start", "listing", "all", "the", "e-mail", "addresses", "you", "want", "included", "in", "the", "list,", "separated", "by", "a", "comma", "and", "a", "blank", "space.", "Cut", "and", "paste", "the", "distribution", "list", "into", "the", "'To,'", "the", "'CC,'", "or", "the", "'BCC'", "field", "in", "your", "e-mail", "client.", "Be", "sure", "to", "use", "your", "own", "e-mail", "address", "in", "the", "'To'", "field.", "That", "way", "you", "will", "receive", "an", "e-mail", "to", "verify", "that", "it", "was", "sent!", "The", "advantage", "to", "the", "BCC", "field", "is", "that", "everyone", "on", "your", "list", "will", "receive", "a", "copy", "of", "your", "message", "without", "the", "entire", "list", "appearing", "in", "the", "header.", _
"Distribution", "lists", "are", "not", "intended", "for", "spamming", "purposes.", "It's", "often", "the", "practice", "to", "include", "instructions", "on", "how", "to", "be", "removed", "from", "a", "distribution", "list", "for", "those", "who", "don't", "wish", "to", "receive", "future", "mailings.", _
"Do", "not", "spam!", "Spam", "is", "the", "e-mail", "equivalent", "of", "junk", "mail,", "and", "spamming", "is", "the", "practice", "of", "sending", "out", "unsolicited", "bulk", "e-mail.", "Some", "e-mail", "systems", "let", "you", "filter", "out", "unwanted", "mail,", "but", "they're", "not", "perfect.", "And", "nothing", "gets", "people", "angrier", "than", "seeing", "their", "inbox", "fill", "up", "with", "junk", "every", "day.", "Do", "your", "part", "to", "help", "by", "not", "spamming", "others.", "And", "report", "any", "spam", "you", "get", "to", "your", "Internet", "Service", "Provider--", "ISP.", _
"Group", "communication", "is", "no", "different", "from", "sending", "out", "a", "single", "e-mail.", "But", "sometimes", "it's", "hard", "to", "tell", "what", "good", "Netiquette", "is", "when", "joining", "a", "new", "online", "group.", "As", "always,", "it's", "important", "to", "follow", "the", "basic", "rules", "above,", "but", "each", "group", "has", "its", "own", "quirks", "and", "customs.", "Here's", "some", "general", "advice", "for", "dealing", "with", "new", "group", "Netiquette", "situations.", _
"When", "you're", "new", "to", "the", "area,", "it's", "best", "to", "observe", "for", "a", "moment", "before", "diving", "into", "the", "mix.", "When", "you", "enter", "a", "new", "chat", "room,", "or", "when", "you", "want", "to", "post", "something", "to", "a", "bulletin", "board", "for", "the", "first", "time,", "stop", "and", "look", "around.", "The", "best", "way", "to", "get", "a", "feel", "for", "the", "correct", "group", "Netiquette", "is", "to", "see", "what", "the", "people", "around", "you", "are", "doing.", "Another", "good", "way", "to", "find", "out", "what's", "going", "on", "is", "to", "look", "for", "a", "FAQ,", "About,", "or", "Help", "page.", _
"Privacy", "is", "another", "issue", "on", "the", "Internet.", "The", "truth", "is,", "there", "is", "no", "such", "thing", "as", "complete", "privacy.", "If", "you", "send", "out", "an", "e-mail,", "there's", "always", "a", "chance", "someone", "else", "could", "intercept", "it.", "And", "with", "most", "e-mail", "systems,", "the", "E-mail", "Administrator", "has", "access", "to", "your", "messages.", "Some", "companies", "even", "monitor", "their", "employee", "e-mail.", "As", "a", "result,", "you", "need", "to", "think", "about", "security.", "The", "following", "are", "a", "few", "tips", "to", "get", "you", "started.")
 Label9 = netiquette(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find netiquette(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub




Private Sub Hubble_Next()
    On Error Resume Next
    Beep
    i = i + 1
  hubble = Array("The", "Hubble", "Space", "Telescope,", "also", "called", "the", "HST,", "is", "a", "telescopic", "observatory", "that", "orbits", "600", "kilometers", "above", "the", "Earth.", "It", "is", "the", "first", "scientific", "mission", "of", "any", "kind", "designed", "for", "routine", "service", "by", "astronauts.", "It", "is", "one", "of", "the", "largest", "and", "most", "complex", "satellites", "ever", "built", "and", "was", "deployed", "April", "25,", "1990", "from", "the", "space", "shuttle", "Discovery.", "When", "Hubble", "was", "designed", "in", "the", "1970s,", "NASA", "stated", "that", "its", "main", "scientific", "objectives", "were", "to", "determine:", _
"*", "The", "constitution,", "physical", "characteristics,", "and", "dynamics", "of", "celestial", "bodies.", _
"*", "The", "nature", "of", "processes", "that", "occur", "in", "the", "extreme", "physical", "conditions", "existing", "in", "and", "between", "astronomical", "objects.", _
"*", "The", "history", "and", "evolution", "of", "the", "universe.", _
"*", "Whether", "the", "laws", "of", "nature", "are", "universal", "in", "the", "space-time", "continuum.", _
"That's", "a", "pretty", "tall", "order", "for", "a", "relatively", "small", "object!", _
"The", "Hubble", "Space", "Telescope", "is", "as", "large", "as", "a", "tractor-trailer", "(43.5", "feet", "long", "and", "14", "feet", "in", "diameter)", "and", "weighs", "12", "tons.", "It", "looks", "like", "nothing", "more", "than", "a", "five-story", "tower", "of", "stacked", "silver", "canisters.", "But", "each", "canister", "houses", "important", "telescope", "equipment.", "The", "telescope", "has", "focusing", "mirrors,", "computers,", "imaging", "instruments,", "and", "pointing", "and", "control", "mechanisms.", "Extending", "from", "the", "telescope", "are", "solar", "panels", "for", "generating", "electricity", "and", "antennae", "for", "communicating", "with", "operators", "on", "the", "ground.", _
"Scientists", "can", "see", "better", "with", "Hubble", "than", "with", "any", "telescope", "on", "the", "ground", "because", "Hubble", "travels", "above", "the", "atmosphere.", "On", "Earth,", "the", "atmosphere", "makes", "pictures", "taken", "by", "visible-light", "telescopes", "look", "smeary", "–", "clouds,", "precipitation,", "and", "atmospheric", _
"temperature", "changes", "get", "in", "the", "way.", "But", "in", "the", "near", "vacuum", "of", "space,", "Hubble", "can", "take", "very", "clear,", "crisp", "pictures", "--", "no", "clouds", "or", "atmosphere", "exist", "between", "it", "and", "the", "star", "or", "galaxy", "it's", "observing.", _
"Every", "day", "Hubble", "archives", "between", "three", "and", "five", "gigabytes", "of", "data", "and", "sends", "10", "to", "15", "gigabytes", "to", "astronomers", "around", "the", "world.", "Once", "the", "telescope", "gathers", "the", "data,", "its", "computers", "turn", "the", "information", "into", "long", "strings", "of", "numbers", "that", "are", "beamed", "to", "Earth", "as", "radio", "signals.", "This", "information", "streams", "through", "a", "series", "of", "satellite", "relays", "to", "the", "Goddard", "Space", "Flight", "Center", "in", "Greenbelt,", "Maryland.", "Then", "it", "travels", "by", "telephone", "line", "to", "the", "Space", "Telescope", "Science", "Institute", "in", "Baltimore.", "It", "is", "there", "that", "the", "numbers", "are", "converted", "back", "into", "pictures", "and", "useful", "data", "for", "study.", _
"The", "Hubble", "Space", "Telescope", "is", "named", "for", "astronomer", "Edwin", "P.", "Hubble.", "As", "a", "young", "boy,", "Hubble", "read", "tales", "of", "traveling", "to", "undersea", "cities", "and", "journeying", "to", "the", "center", "of", "the", "Earth.", "These", "stories", "by", "adventure", "novelists", "stoked", "young", "Hubble's", "imagination", "of", "faraway", "places.", "He", "fulfilled", "those", "childhood", "dreams", "by", "becoming", "an", "astronomer,", "exploring", "distant", "galaxies", "with", "telescopes,", "and", "developing", "theories", "that", "changed", "the", "field", "of", "astronomy.", _
"But", "Hubble", "didn't", "settle", "immediately", "on", "astronomy", "as", "a", "career.", "He", "first", "studied", "law", "as", "a", "Rhodes", "Scholar", "at", "Queens", "College", "in", "Oxford,", "England.", "A", "year", "after", "passing", "the", "bar", "exam,", "Hubble", "realized", "that", "his", "love", "of", "exploring", "the", "stars", "was", "greater", "than", "his", "attraction", "to", "law.", "He", "therefore", "abandoned", "law", "and", "studied", "astronomy", "at", "the", "University", "of", "Chicago,", "where", "he", "completed", "his", "doctoral", "thesis", "in", "1917.", _
"Using", "the", "largest", "telescope", "of", "its", "day,", "Hubble", "looked", "far", "into", "space", "to", "study", "the", "Andromeda", "nebula", "(which", "we", "now", "call", "the", "Andromeda", "Galaxy", "or", "M31).", "With", "this", "powerful", "telescope,", "he", "saw", "stars", "near", "the", "edge", "of", "Andromeda,", "and", "he", "reasoned", "that", "Andromeda", "was", "a", "galaxy", "in", "its", "own", "right.", "He", "proved", "that", "galaxies,", "composed", "of", "stars", "and", "planets,", "exist", "beyond", "our", "own", "Milky", "Way.", "He", "also", "proved", "that", "they", "are", "distributed", "almost", "uniformly", "in", "every", "direction.", _
"As", "Hubble", "continued", "his", "study,", "he", "discovered", "that", "the", "universe", "is", "expanding.", "The", "more", "distant", "a", "galaxy", "is", "from", "Earth,", "the", "more", "its", "light", "is", "shifted", "toward", "the", "red", "end", "of", "the", "spectrum", "--", "redshifted.", "In", "1929", "he", "determined", "that", "the", "more", "distant", "the", "galaxy", "from", "Earth,", "the", "greater", "its", "redshift,", "and", "the", "faster", "it", "appears", "to", "move", "away.", "Known", "as", "Hubble's", "Law,", "this", "discovery", "is", "the", "foundation", "of", "the", "Big", "Bang", "theory.", "The", "theory", "states", "that", "the", "universe", "began", "after", "a", "cataclysmic", "event", "in", "which", "it", "underwent", "a", "rapid", "period", "of", "inflation.", "The", "movement", "of", "the", "galaxies", "away", "from", "us", "is", "the", "continuation", "of", "that", "motion.", "Hubble's", "discovery", "is", "considered", "one", "of", "the", "greatest", _
"triumphs", "of", "20th", "century", "astronomy.", "However,", "confirmation", "was", "difficult,", "until", "the", "launch", "of", "the", "HST.", _
"With", "the", "telescope", "in", "orbit,", "astronomers", "are", "getting", "a", "picture", "of", "the", "universe", "clearer", "than", "anything", "they", "could", "previously", "obtain.", "The", "telescope's", "stunning", "photos", "are", "showing", "the", "world", "about", "the", "wonders", "of", "space.", "Many", "of", "the", "world's", "foremost", "astronomers", "are", "using", "Hubble", "to", "probe", "the", "horizons", "of", "space", "and", "time.", "Designed", "to", "last", "at", "least", "15", "years,", "Hubble", "is", "providing", "intriguing", "new", "data", "about", "black", "holes,", "the", "birth", "of", "galaxies,", "planetary", "systems", "around", "other", "stars,", "and,", "of", "course,", "the", "expansion", "of", "the", "universe.")
Label9 = hubble(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find hubble(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Baseball_Next()
    On Error Resume Next
    Beep
    i = i + 1
baseball = Array("Baseball", "has", "been", "the", "national", "pastime", "in", "the", "United", "States", "since", "the", "middle", "of", "the", "1800s.", "Each", "period", "has", "had", "a", "unique", "flavor,", "and", "it", "is", _
"sometimes", "useful", "to", "read", "about", "a", "period", "from", "the", "perspective", "of", "the", "time", "in", "which", "it", "took", "place.", "The", "following", "comes", "from", "'Spalding's", "Baseball", "Guide", "and", "Official", "League", "Book", "for", "1889,", "written", "in", "that", "same", "year;", "the", "language", "has", "not", "been", "edited", "to", "reflect", "modern", "expressions", "or", "word", "usage.", _
"The", "Joint", "Rules", "Committee", "and", "their", "work", "The", "argument", "brought", "to", "bear", "in", "favor", "of", "the", "elimination", "of", "outs", "from", "foul", "balls", "from", "the", "code", "was", "in", "the", "main", "as", "follows:", _
"When", "the", "batsman", "hits", "a", "fair", "ball,", "while", "at", "the", "same", "time", "that", "he", "gives", "the", "fielders", "a", "chance'", "to", "put", "him", "out,", "he", "himself", "is", "also", "given", "an", "equal", "chance", "of", "making", "a", "base", "or", "of", "scoring", "a", "run;", "but", "when", "he", "hits", "a", "foul", "ball,", "while", "he", "affords", "the", "fielders", "an", "opportunity", "to", "catch", "him", "out,", "no", "such", "compensating", "advantage", "is", "given", "him", "in", "the", "way", "of", "earning", "a", "base", "or", "a", "run", "as", "in", "the", "case", "of", "a", "fair", "hit", "ball;", "and", "it", "is", "in", "this", "that", "the", "working", "of", "the", "foul", "ball", "rule", "becomes", "so", "palpably", "unjust.", "It", "is", "sufficient", "punishment", "for", "hitting", "a", "foul", "ball", "that", "he,", "as", "batsman,", "be", "deprived", "of", "making", "a", "base,", "without", "adding", "the", "unjust", "penalty", "of", "an", "out.", _
"This", "one", "sided", "condition", "of", "things,", "too,", "is", "increased", "when", "a", "double", "play", "is", "made", "on", "the", "catch", "of", "a", "foul", "ball,", "for", "not", "only", "is", "the", "batsman", "unfairly", "punished,", "but", "also", "the", "base", "runner", "who", "may", "have", "made", "the", "base", "by", "a", "clean", "hit.", _
"Twenty", "odd", "years", "ago", "George", "Wright", "suggested", "to", "the", "Chairman", "of", "the", "old", "National", "Association's", "Committee", "of", "Rules", "that", "it", "would", "be", "a", " good", "plan", "to", "allow", "base", "runners", "to", "overrun", "first", "base,", "giving", "them", "the", "privilege", "to", "return", "and", "touch", "the", "base", "again", "without", "being", "put", "out,", "before", "attempting", "to", "make", "another", "base.", "The", "suggestion", "was", "adopted,", "and", "the", "rule", "went", "into", "effect", "in", "1870,", " and", "it", "has", "been", "in", "operation", "ever", "since.", "When", "the", "amendment", "was", "presented", "at", "the", "convention", "of", "1869,", "a", "delegate", "wanted", "the", "rule", "applied", "to", "all", "bases,", "but", "the", "majority", "preferred", "to", "test", "the", "experiment", "as", "proposed", "at", "first", "base.", _
"The", "rule", "of", "extending", "the", "over-running", "to", "all", "the", "bases", "was", "advocated", "at", "the", "last", "meeting", "in", "1888", "of", "the", "Joint", "Committee", "of", "Rules,", "but", "it", "was", "not", "adopted.", _
"The", "rule", "is", "worthy", "of", "consideration,", "in", "view", "of", "the", "constant", "sprains", "and", "injuries", "of", "one", "kind", "and", "another", "arising", "from", "sliding", "to", "bases.", "There", "has", "not", "been", "a", "single", "instance", "of", "an", "injury", "occurring", "from", "the", "working", "of", "the", "rule", "of", "overrunning", "first", "base", "since", "the", "rule", "was", "adopted,", "while", "serious", "injuries", "are", "of", "daily", "occurrence", "in", "match", "games,", "arising", "from", "collisions", "at", "other", "bases", "than", "first,", "and", "these", "are", "due", "entirely", "to", "the", "absence", "of", "the", "overrunning", "rule.", "The", "most", "irritating", "disputes", "caused", "by ", "questions", "involved", "in", "sliding", "to", "bases", "and", "in", "running", "up", "against", "base", "players,", "are", "also", _
"due", "to", "the", "same", "cause.", "Why", "not", "put", "a", "stop", "to", "these", "injuries", "and", "these", "disputes", "by", "giving", "the", "base", "runner", "the", "same", "privileges", "in", "overrunning", "second,", "third,", "and", "home", "bases", "that", "he", "now", "has", "in", "overrunning", "first", "base?", "In", "every", "way", "will", "the", "adoption", "of", "the", "rule", "suggested", "be", "an", "improvement,", "and", "not", "the", "least", "of", "its", "advantages", "will", "be", "its", "gain", "to", "base", "running,", "which", "is,", "next", "to", "fielding,", "the", "most", "attractive", "feature", "of", "our", "game.")
Label9 = baseball(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find baseball(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Fables_Next()
    On Error Resume Next
    Beep
    i = i + 1
fables = Array("Aesop", "was", "one", "of", "the", "great", "Greek", "writers.", "He", "is", "best", "known", "for", "his", "fables,", "stories", "that", "have", "a", "moral.", "They", "teach", "us", "something", "about", "how", "we", "should", "live", "our", "lives.", "Aesop", "wrote", "thousands", "of", "these", "stories.", "Here", "are", "a", "few.", "The", "Wolf", "in", "Sheep's", "Clothing", "Once", "upon", "a", "time,", "a", "Wolf", "decided", "to", "disguise", "the", "way", "he", "looked.", _
"He", "thought", "it", "would", "help", "him", "get", "food", "more", "easily.", "He", "put", "on", "the", "skin", "of", "a", "sheep,", "then", "he", "went", "out", "with", "the", "flock", "into", "the", "pasture.", "Even", "the", "shepherd", "was", "fooled", "by", "his", "clever", "costume.", "In", "the", "evening,", "the", "shepherd", "put", "him", "in", "with", "the", "rest", "of", "the", "sheep.", "He", "closed", "the", "gate", "and", "made", "sure", "it", "was", "secure", "before", "he", "went", "to", "bed.", "In", "the", "middle", "of", "the", "night,", "he", "came", "back", "to", "the", "fold", "to", "get", "some", "meat", "for", "the", "next", "day.", "Instead", "of", "a", "sheep,", "though,", "he", "grabbed", "the", "Wolf,", "killing", "him", "instantly.", _
"Those", "who", "look", "to", "harm", "others", "will", "be", "harmed", "themselves.", "The", "Bat", "and", "the", "Weasel.", _
"A", "Bat", "fell", "on", "the", "ground", "and", "was", "caught", "by", "a", "Weasel.", "It", "begged", "the", "Weasel", "to", "spare", "its", "life,", "but", "the", "Weasel", "refused.", "It", "told", "the", "Bat", "that", "birds,", "by", "nature,", "were", "its", "enemy.", "The", "Bat", "assured", "him", "that", "it", "was", "not", "a", "bird,", "it", "was", "a", "mouse.", "The", "Weasel", "thought", "a", "moment,", "then", "set", "it", "free.", "A", "while", "later,", "the", "Bat", "fell", "again", "to", "the", "ground,", "and", "it", "was", "caught", "by", "another", "Weasel.", "It", "begged", "this", "Weasel", "not", "to", "eat", "him,", "either.", "The", "Weasel,", "though,", "said", "it", "did", "not", "like", "mice", "at", "all", "and", "would", "eat", "it.", "The", "Bat", "told", "the", "Weasel", "that", "it", "was", "not", "a", "mouse,", "but", "a", "bat.", "The", "second", "Weasel", "had", "no", "good", "answer,", "so", "he", "let", "it", "go.", _
"The", "Bat", "knew", "it", "is", "always", "wise", "to", "turn", "events", "to", "your", "advantage.", "The", "Lion", "and", "the", "Mouse", _
"A", "sleeping", "Lion", "was", "woken", "up", "by", "a", "Mouse", "running", "over", "his", "face.", "He", "got", "up", "angrily", "and", "caught", "the", "scared", "little", "Mouse.", "He", "was", "about", "to", "kill", "the", "Mouse,", "but", "it", "said", "in", "its", "squeaky", "little", "voice,", "'If", "you", "would", "only", "spare", "my", "life,", "I", "would", "be", "sure", "to", "repay", "your", "kindness.,'", "The", "Lion", "laughed", "at", "such", "nonsense,", "but", "he", "let", "him", "go.", "A", "short", "time", "later,", "though,", "the", "Lion", "was", "caught", "by", "some", "hunters.", "They", "bound", "him", "by", "ropes", "to", "the", "ground.", "The", "Mouse", "recognized", "his", "roar,", "and", "he", "rushed", "over", "and", "gnawed", "the", "rope", "with", "his", "teeth,", "setting", "the", "Lion", "free.", "The", "Mouse", "said", "'You", "laughed", "at", "the", "idea", "of", "my", "ever", "being", "able", "to", "help", "you.", "Now", "you", "know", "that", "it", "is", "possible", _
"for", "even", "a", "small", "little", "Mouse", "to", "help", "a", "great", "big", "Lion.,'", "The", "Fox", "and", "the", "Crow", _
"A", "Crow", "stole", "a", "bit", "of", "meat,", "and", "she", "perched", "in", "a", "tree,", "holding", "it", "firmly", "in", "her", "beak.", "A", "Fox", "saw", "her", "there", "and", "knew", "he", "wanted", "the", "meat", "for", "himself.", "He", "came", "up", "with", "a", "plan.", "'How", "handsome", "is", "the", "Crow,'", "he", "said.", "'Her", "shape", "is", "beautiful", "and", "her", "feathers", "bright.", "If", "only", "her", "voice", "were", "equal", "to", "her", "beauty,", "she", "would", "be", "called", "the", "Queen", "of", "Birds!'", "He", "lied", "when", "he", "said", "this,", "but", "the", "Crow", "was", "vain.", "She", "wanted", "to", "show", "the", "Fox", "that", "her", "voice", "really", "was", "as", "beautiful", "as", "he", "said.", "She", "opened", "her", "mouth", "wide", "and", "cawed", "loudly,", "dropping", "the", "meat", "to", "the", "ground.", "The", "Fox", "quickly", "picked", "it", "up", "and", "shook", "his", "head", "with", "scorn.", _
"'My", "good", "Crow,'", "he", "said,", "your", "voice", "is", "fine,", "but", "you", "don't", "have", "the", "brains", "to", "use", "it", "well.'", "The", "Two", "Dogs", _
"A", "Man", "had", "two", "dogs.", "One", "was", "a", "Hound,", "trained", "to", "assist", "him", "in", "his", "sports,", "and", "the", "other", "was", "a", "House", "Dog,", "taught", "to", "watch", "the", "house.", "When", "he", "came", "home", "after", "a", "good", "day's", "hunting,", "he", "always", "gave", "the", "House", "Dog", "a", "large", "share", "of", "his", "catch.", "The", " Hound", "felt", "very", "hurt", "at", "this.", "He", "scolded", "his", "companion,", "saying,", "'My", "work", "is", "very", "hard,", "but", "it", "is", "harder", "to", "watch", "you.", "You", "do", "not", "help", "me", "with", "the", "chase,", "but", "you", "share", "in", "my", "reward.", "It", "is", "not", "fair.'", "The", "House", "Dog", "replied,", "'Do", "not", "not", "fair.'", "The", "House", "Dog", "replied,", "'Do", "not", "blame", "me,", "my", "friend.", "If", "you", "are", "angry,", "find", "faulor", "the", "faults", "of", "its", "parents.", _
"The", "Kingdom", "of", "the", "Lion", "The", "Lion", "was", "the", "king", "of", "all", "the", "animals", "in", "the", "field", "and", "the", "forest.", "He", "was", "not", "violent", "or", "cruel,", "but", "he", "was", "as", "just", "and", "gentle", "as", "a", "king", "should", "be.", _
"During", "his", "reign,", "he", "called", "for", "a", "gathering", "of", "all", "the", "animals", "to", "draw", "up", "conditions", "for", "an", "alliance.", "In", "this", "alliance,", "the", "Wolf", "and", "the", "Lamb,", "the", "Tiger", "and", "the", "Deer,", "the", "Dog", "and", "the", "Hare,", "would", "live", "together", "in", "perfect", "peace.", "The", "Hare", "said,", "'Oh,", "how", "I", "have", "longed", "to", "see", "this", "day.", "Now", "the", "weak", "shall", "take", "their", "place", "by", "the", "side", "of", "the", "strong", "without", "fear.'", "Everyone", "agreed,", "but", "after", "the", "Hare", "said", "this,", "he", "ran", "for", "his", "life.", "After", "all,", "you", "cannot", "believe", "everything", "you", "hear.", "The", "Wolf", "and", "the", "Crane", _
"A", "Wolf", "had", "a", "bone", "stuck", "in", "his", "throat.", "He", "hired", "a", "Crane,", "for", "a", "large", "sum,", "to", "put", "her", "head", "into", "his", "mouth", "and", "draw", "out", "the", "bone.", "When", "the", "Crane", "had", "taken", "out", "the", "bone,", "he", "demanded", "the", "promised", "payment.", "The", "Wolf,", "grinning", "and", "grinding", "his", "teeth,", "said,", "'Why,", "you", "have", "already", "received", "your", "reward", "--", "you", "were", "allowed", "to", "draw", "out", "your", "head", "in", "safety", "from", "the", "mouth", "and", "jaws", "of", "a", "wolf.'", _
"In", "serving", "the", "wicked,", "expect", "no", "reward.", "Be", "thankful", "if", "you", "escape", "injury", "for", "your", "pains.", "The", "Traveler", "and", "His", "Dog", _
"A", "Man", "about", "to", "set", "out", "on", "a", "journey", "saw", "his", "Dog", "stand", "at", "the", "door", "stretching", "himself.", "He", "asked", "him", "sharply,", "'Why", "do", "you", "stand", "there?", "Everything", "is", "ready", "but", "you,", "and", "I", "will", "be", "late.", "Come", "with", "me", "this", "instant!'", "The", "Dog,", "wagging", "his", "tail,", "replied,", "'O,", "master!", "I", "am", "quite", "ready.", "In", "fact,", "it", "is", "you", "for", "whom", "I", "am", "waiting.'", _
"The", "late", "person", "often", "blames", "a", "delay", "on", "his", "more", "active", "friend.")
Label9 = fables(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find fables(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Franklin_Next()
    On Error Resume Next
    Beep
    i = i + 1
 franklin = Array("Benjamin", "Franklin", "was", "not", "only", "one", "of", "the", "Founding", "Fathers", "of", "Colonial", "America,", "but", "he", "was", "one", "of", "the", "premier", "thinkers", "and", "statesmen", "of", "that", "time.", "It", "is", "interesting,", "therefore,", "to", "take", "a", "look", "at", "some", "of", "his", "thoughts", "and", "opinions", "--", "in", "his", "own", "words:", _
"Good", "works", "and", "their", "reward.", "--", "For", "my", "own", "part,", "when", "I", "am", "employed", "in", "serving", "others,", "I", "do", "not", "look", "upon", "myself", "as", "conferring", "favors,", "but", "as", "paying", "debts.", "In", "my", "travels,", "and", "since", "my", "settlement,", "I", "have", "received", "much", "kindness", "from", "men,", "to", "whom", "I", "shall", "never", "have", "any", "opportunity", "of", "making", "the", "least", "direct", "return.", "And", "numberless", "mercies", "from", "God,", "who", "is", "infinitely", "above", "being", "benefitted", "by", "our", "services.", "Those", "kindnesses", "from", "men,", "I", "can", "therefore", "only", "return", "on", "their", "fellow", "men;", "and", "I", "can", "only", "show", "my", "gratitude", "for", "these", "mercies", "from", "God,", "by", "a", "readiness", "to", "help", "his", "other", "children", "and", "my", "brethren.", _
"For", "I", "do", "not", "think", "that", "thanks", "and", "compliments,", "though", "repeated", "weekly,", "can", "discharge", "our", "real", "obligations", "to", "each", "other,", "and", "much", "less", "those", "to", "our", "Creator.", _
"You", "will", "see", "in", "this", "my", "notion", "of", "good", "works,", "that", "I", "am", "far", "from", "expecting", "(as", "you", "suppose)", "that", "I", "shall", "ever", "merit", "Heaven", "by", "them.", "By", "Heaven", "we", "understand", "a", "state", "of", "happiness,", "infinite", "in", "degree,", "and", "eternal", "in", "duration:", "I", "can", "do", "nothing", "to", "deserve", "such", "rewards:", "He", "that", "for", "giving", "a", "draught", "of", "water", "to", "a", "thirsty", "person,", "should", "expect", "to", "be", "paid", "with", "a", "good", "plantation,", "would", "be", "modest", "in", "his", "demands,", "compared", "with", "those", "who", "think", "they", "deserve", "Heaven", "for", "the", "little", "good", "they", "do", "on", "Earth.", "Even", "the", "mixed", "imperfect", "pleasures", "we", "enjoy", "in", "this", "world", "are", "rather", "from", "God's", "goodness", "than", "our", "merit;", "how", "much", "more", "such", "happiness", "of", "Heaven.", _
"For", "my", "own", "part", "I", "have", "not", "the", "vanity", "to", "think", "I", "deserve", "it,", "the", "folly", "to", "expect", "it,", "nor", "the", "ambition", "to", "desire", "it.", "(To", "Joseph", "Huey,", "June", "6,", "1753).", _
"Inventors", "and", "inventions.", "--", "There", "are", "everywhere", "a", "number", "of", "people,", "who,", "being", "totally", "destitute", "of", "any", "inventive", "faculty", "themselves,", "do", "not", "readily", "conceive", "that", "others", "may", "possess", "it;", "they", "think", "of", "inventions", "as", "of", "miracles;", "there", "might", "be", "such", "formerly,", "but", "they", "are", "ceased.", "With", "these,", "every", "one", "who", "offers", "a", "new", "invention", "is", "deemed", "a", "pretender;", "he", "had", "it", "from", "some", "other", "country,", "or", "from", "some", "book;", "a", "man", "of", "their", "own", "acquaintance,", "one", "who", "has", "no", "more", "sense", "than", "themselves,", "could", "not", "possibly,", "in", "their", "opinion,", "have", "been", "the", "inventor", "of", "any", "thing.", "They", "are", "confirmed,", "too,", "in", "these", "sentiments,", "by", "frequent", "instances", "of", "pretensions", "to", _
"intervention,", "which", "vanity", "is", "daily", "producing.", "That", "vanity,", "too,", "though", "an", "incitement", "to", "invention,", "is,", "at", "the", "same", "time,", "the", "pest", "of", "inventors.", _
"Jealousy", "and", "envy", "deny", "the", "merit", "or", "the", "novelty", "of", "your", "invention;", "but", "vanity,", "when", "the", "novelty", "and", "merit", "are", "established,", "claims", "it", "for", "its", "own.", "The", "smaller", "your", "invention", "is,", "the", "more", "mortification", "you", "receive", "in", "having", "the", "credit", "of", "it", "disputed", "with", "you", "by", "a", "rival,", "whom", "the", "jealousy", "and", "envy", "of", "others", "are", "ready", "to", "support", "against", "you,", "at", "least", "so", "far", "as", "to", "make", "the", "point", "doubtful.", "It", "is", "not", "in", "itself", "of", "importance", "enough", "for", "a", "dispute;", "no", "one", "would", "think", "your", "proofs", "and", "reasons", "worth", "their", "attention;", "and", "yet,", "if", "you", "do", "not", "dispute", "the", "point,", "and", "demonstrate", "your", "right,", "you", "not", "only", "lose", "the", "credit", "of", "being", "in", "that", _
"instance", "ingenious,", "but", "you", "suffer", "the", "disgrace", "of", "not", "being", "ingenuous;", "not", "only", "of", "being", "a", "plagiary,", "but", "of", "being", "plagiary", "for", "trifles.", _
"Had", "the", "invention", "been", "greater,", "it", "would", "have", "disgraced", "you", "less;", "for", "men", "have", "not", "so", "contemptible", "an", "idea", "of", "him", "that", "robs", "for", "gold", "on", "the", "highway,", "as", "of", "him", "that", "can", "pick", "pockets", "for", "half-pence", "and", "farthings.", "Thus,", "through", "envy,", "jealousy,", "and", "the", "vanity", "of", "competitors", "for", "fame,", "the", "origin", "of", "many", "of", "the", "most", "extraordinary", "inventions,", "though", "produced", "within", "but", "a", "few", "centuries", "past,", "is", "involved", "in", "doubt", "and", "uncertainty.", "We", "scarce", "know", "to", "whom", "we", "are", "indebted", "for", "the", "compass,", "and", "spectacles,", "nor", "have", "even", "paper", "and", "printing,", "that", "record", "every", "thing", "else,", "been", "able", "to", "preserve", "with", "certainty", "the", "name", "and", "reputation", "of", "their", "inventors.", _
"One", "would", "not,", "therefore,", "of", "all", "faculties", "or", "qualities", "of", "the", "mind,", "wish,", "for", "a", "friend", "or", "a", "child,", "that", "he", "should", "have", "that", "of", "invention.", "(To", "John", "Lining,", "March", "18,", "1755)", _
"Error.", "--", "There", "are", "few,", "though", "convinced,", "that", "know", "how", "to", "give", "up,", "even", "an", "error,", "they", "have", "been", "once", "engaged", "in", "maintaining;", "there", "is", "therefore", "the", "more", "merit", "in", "dropping", "a", "contest", "where", "one", "thinks", "one's", "self", "right;", "'tis", "at", "least", "respectful", "to", "those", "we", "converse", "with.", "And", "indeed", "all", "our", "knowledge", "is", "so", "imperfect,", "and", "we", "are", "from", "a", "thousand", "causes", "so", "perpetually", "subject", "to", "mistake", "and", "error,", "that", "positiveness", "can", "scarce", "ever", "become", "even", "the", "most", "knowing;", "and", "modesty", "in", "advancing", "any", "opinion,", "however", "plain", "and", "true", "we", "may", "suppose", "it,", "is", "always", "decent,", "and", "generally", "more", "likely", "to", "procure", "assent.", "(To", "Mary", "Stevenson,", "[1759])", _
"Exercise.", "--", "In", "considering", "the", "different", "kinds", "of", "exercise,", "I", "have", "thought,", "that", "the", "quantum", "of", "each", "is", "to", "be", "judged", "of,", "not", "by", "time", "or", "by", "distance,", "but", "by", "the", "degree", "of", "warmth", "it", "produces", "in", "the", "body.", "Thus,", "when", "I", "observe,", "if", "I", "am", "cold", "when", "I", "get", "into", "a", "carriage", "in", "a", "morning,", "I", "may", "ride", "all", "day", "without", "being", "warmed", "by", "it;", "that,", "if", "on", "horseback", "my", "feet", "are", "cold,", "I", "may", "ride", "some", "hours", "before", "they", "become", "warm;", "but,", "if", "I", "am", "ever", "so", "cold", "on", "foot,", "I", "cannot", "walk", "an", "hour", "briskly,", "without", "glowing", "from", "head", "to", "foot", "by", "the", "quickened", "circulation;", "I", "have", "been", "ready", "to", "say,", "(using", "round", "numbers", "without", "regard", "to", "exactness,", "but", "merely", "to", "mark", "a", _
"great", "difference)", "that", "there", "is", "more", "exercise", "in", "one", "mile's", "riding", "on", "horseback,", "than", "in", "five", "in", "a", "coach;", "and", "more", "in", "one", "mile's", "walking", "on", "foot,", "than", "in", "five", "on", "horseback;", "to", "which", "I", "may", "add,", "that", "there", "is", "more", "in", "walking", "one", "mile", "up", "and", "down", "stairs,", "than", "in", "five", "on", "a", "level", "floor.")
Label9 = franklin(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find franklin(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Dna_Next()
    On Error Resume Next
    Beep
    i = i + 1
 dna = Array("At", "the", "end", "of", "the", "road", "in", "Little", "Cottonwood", "Canyon,", "near", "Salt", "Lake", "City,", "Utah,", "lies,", "Alta,", "a", "near-mythic", "location", "for", "skiing", "enthusiasts.", "In", "time,", "though,", "it", "may", "well", "achieve", "a", "similar", "status", "among", "molecular", "geneticists.", "In", "December", "1984,", "a", "conference", "co-sponsored", "by", "the", "U.S.", "Department", "of", "Energy", "was", "held", "there,", "a", _
 "conference", "that", "pondered", "a", "single", "question:", "Does", "modern", "DNA", "research", "offer", "a", "way", "of", "detecting", "tiny", "genetic", "mutations", "-- and,", "in", "particular,", "of", "observing", "any", "increase", "in", "the", "mutation", "rate", "among", "the", "survivors", "of", "the", "Hiroshima", "and", "Nagasaki", "bombings", "and", "their", "descendants?", "In", "short,", "the", "answer", "was,", "'Not", "yet;'", "however,", "in", "this", "atmosphere", _
"of", "intellectual", "fertility,", "the", "seeds", "were", "sown", "for", "a", "project", "that", "would", "make", "such", "detection", "possible", "in", "the", "future", "-- the", "Human", "Genome", "Project.", _
"In", "the", "months", "that", "followed,", "much", "deliberation", "and", "debate", "ensued.", "But", "in", "1986,", "the", "DOE", "took", "a", "bold", "and", "unilateral", "step", "by", "announcing", "its", "Human", "Genome", "Initiative,", "convinced", "that", "its", "mission", "would", "be", "well", "served", "by", "a", "comprehensive", "picture", "of", "the", "human", "genome.", _
"The", "immediate", "response", "was", "considerable", "skepticism", "-- skepticism", "about", "the", "scientific", "community's", "technological", "wherewithal", "for", "sequencing", "the", "genome", "at", "a", "reasonable", "cost,", "and", "skepticism", "about", "the", "value", "of", "the", "result,", "even", "if", "it", "could", "be", "obtained", "economically.", _
"Over", "the", "years,", "though,", "things", "have", "changed.", "Today,", "a", "worldwide", "effort", "is", "under", "way", "to", "develop", "and", "apply", "the", "technologies", "needed", "to", "completely", "map", "and", "sequence", "the", "human", "genome,", "as", "well", "as", "the", "genomes", "of", "several", "model", "organisms.", _
"Technological", "progress", "has", "been", "rapid,", "and", "it", "is", "now", "generally", "agreed", "that", "this", "international", "project", "will", "produce", "the", "complete", "sequence", "of", "the", "human", "genome", "by", "the", "year", "2005,", "if", "not", "before.", _
"And", "what", "is", "more", "important,", "the", "value", "of", "the", "project", "is", "becoming", "evident", "in", "a", "wide", "range", "of", "areas.", "Genome", "research", "is", "revolutionizing", "both", "biology", "and", "biotechnology,", "and", "it", "is", "providing", "a", "vital", "thrust", "to", "the", "increasingly", "broad", "scope", "of", "the", "biological", "sciences.", _
"The", "impact", "that", "will", "be", "felt", "in", "medicine", "and", "health", "care", "alone,", "once", "we", "identify", "and", "understand", "the", "function", "of", "all", "human", "genes,", "is", "inestimable.", "To", "that", "end,", "the", "project", "has", "already", "stimulated", "significant", "financial", "investment", "by", "large", "corporations", "and", "has", "prompted", "the", "creation", _
"of", "new", "companies", "hoping", "to", "capitalize", "on", "the", "burgeoning", "market.", _
"But", "the", "DOE's", "early,", "catalytic", "decision", "deserves", "further", "comment,", "as", "the", "implications", "are", "more", "widespread", "-- and", "more", "controversial", "-- than", "initially", "believed.", "The", "project", "could", "deliver", "and", "already", "has", "developed", "into,", "much", "more", "than", "the", "promised", "tool", "for", "assessing", ",mutation", "rates.", _
"The", "information", "generated", "would", "contribute", "not", "only", "to", "a", "new", "understanding", "of", "human", "biology,", "but", "also", "to", "a", "host", "of", "practical", "applications", "in", "the", "biotechnology", "industry", "and", "in", "the", "arenas", "of", "agriculture", "and", "environmental", "protection.", _
"A", "1987", "report", "by", "a", "DOE", "advisory", "committee", "provided", "some", "examples.")
Label9 = dna(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find dna(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub




Private Sub Cast_Next()
    On Error Resume Next
    Beep
    i = i + 1
 cast = Array("If", "you're", "like", "most", "kids,", "you've", "probably", "broken", "a", "bone.", "Or", "maybe", "you", "know", "someone", "else", "who", "has.", "That", "means", "you've", "seen", "a", "cast.", "Do", "you", "know", "how", "a", "cast", "works?", "What", "it's", "made", "of?", "Whether", "it's", "OK", "to", "draw", "on", "one?", "Keep", "typing", "to", "find", "out", "more.", _
"Think", "of", "a", "cast", "as", "a", "big", "bandage", "that", "has", "two", "layers.", "Inside", "there", "is", "a", "soft", "cotton", "layer.", "It", "rests", "next", "to", "the", "skin.", "This", "makes", "the", "cast", "more", "comfortable.", "It", "also", "protects", "the", "skin.", "Outside", "there", "is", "a", "hard", "outer", "layer.", "This", "layer", "prevents", "the", "broken", "bone", "from", "moving.", "Sometimes", "you", "can", "choose", "how", "your", "cast", "looks.", "You", "can", "pick", "cool", "colors", "or", "patterns.", _
"Casts", "come", "in", "two", "types.", "The", "first", "type", "is", "made", "of", "plaster", "of", "Paris.", "You", "may", "have", "used", "this", "material", "for", "arts", "and", "crafts.", "Plaster", "of", "Paris", "is", "a", "heavy", "white", "powder.", "When", "it's", "mixed", "with", "water,", "it", "forms", "a", "thick", "paste", "that", "gets", "hard", "fast.", _
"The", "second", "type", "is", "made", "of", "fiberglass.", "This", "is", "a", "plastic", "material.", "It", "can", "be", "shaped.", "Fiberglass", "casts", "come", "in", "lots", "of", "colors.", _
"These", "casts", "are", "lighter", "and", "stronger", "than", "plaster", "casts.", "That", "makes", "them", "easier", "to", "wear.", "It", "also", "makes", "them", "last", "longer.", _
"Your", "doctor", "will", "make", "sure", "that", "all", "the", "pieces", "of", "your", "broken", "bone", "are", "in", "their", "proper", "places.", "Then", "he'll", "put", "on", "the", "cast", "to", "prevent", "the", "pieces", "from", "moving", "while", "they", "heal.", "This", "way", "they'll", "grow", "back", "together", "into", "one", "bone.", "It", "also", "prevents", "injury", "to", "the", "area", "around", "the", "bone.", "Most", "kids", "with", "broken", "bones", "say", "it", "actually", "feels", "much", "better", "once", "the", "cast", "is", "on.", _
"How", "long", "do", "you", "have", "to", "wear", "a", "cast?", "Sometimes", "it's", "only", "for", "a", "few", "weeks.", "Sometimes", " it's", "much", "longer.", "It", "depends", "on", "the", "type", "of", "injury", "you", "have.", "It", "also", "depends", "on", "how", "badly", "the", "bone", "was", "broken.", _
"The", "cast", "may", "be", "on", "your", "arm.", "If", "so,", "you", "may", "be", "given", "a", "sling.", _
"Your", "doctor", "may", "put", "on", "the", "cast", "himself.", "Usually,", "though,", "another", "person", "who", "is", "specially", "trained", "to", "put", "on", "casts", "does", "it.", _
"You", "may", "want", "to", "keep", " your", "cast.", "Perhaps", "all", "your", "friends", "signed", "it.", _
"If", "so,", "just", "ask", "your", "doctor.", "Many", "doctors", "are", "happy", "to", "let", "people", "keep", "their", "casts.,,,")
Label9 = cast(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find cast(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub



Private Sub Photo_Next()
    On Error Resume Next
    Beep
    i = i + 1
photo = Array("As", "in", "most", "cases", "of", "great", "and", "valuable", "inventions", "in", "science", "and", "art,", "the", "English", "lay", "claim", "to", "the", "honor", "of", "having", "first", "discovered", "photogenic", "drawing.", "But", "we", "shall", "see", "in", "the", "progress", "of", "our", "discussion", "here,", "like", "many", "other", "assumptions", "of", "their", "authors,", "that", "claim", "is", "unfounded,", "and", "they", "are", "due", "no", "more", "credit", "than", "for", "the", "invention", "of", "the", "steamboat", "or", "the", "cotton", "gin.", _
"The", "English", "claim", "is", "founded", "upon", "the", "fact", "that,", "in", "1802,", "a", "Mr.", "Wedgwood", "recorded", "an", "experiment", "in", "the", "'Journal", "of", "the", "Royal", "Institution'", "as", "follows:", _
"'A", "piece", "of", "paper,", "or", "other", "convenient", "material,", "was", "placed", "upon", "a", "frame", "and", "sponged", "over", "with", "a", "solution", "of", "nitrate", "of", "silver;", "it", "was", "then", "placed", "behind", "a", "painting", "on", "glass", "and", "the", "light", "traversing", "the", "painting", "produced", "a", "kind", "of", "copy", "upon", "the", "prepared", "paper,", "those", "parts", "in", "which", "the", "rays", "were", "least", "intercepted", "being", "of", "the", "darkest", "hues.", "Here,", "however,", "terminated", "the", "experiment;", "for", "although", "both", "Mr.", "Wedgwood", "and", "Sir", "Humphry", "Davey", "experimented", "carefully,", "for", "the", "purpose", "of", "endeavoring", "to", "fix", "the", "drawings", "thus", "obtained,", "yet", "the", "object", "could", "not", "be", "accomplished,", "and", "the", "whole", "ended", "in", "failure.'", _
"This,", "by", "their", "own", "showing,", "was", "the", "earliest", "attempt", "of", "the", "English", "experimenters.", "But", "this", "much", "of", "the", "principle", "was", "known", "to", "the", "alchemists", "at", "a", "much", "earlier", "date,", "although", "the", "experiment", "was", "carried", "out", "in", "a", "practically", "different", "manner,", "as", "the", "following", "description", "amply", "proves:", _
"'Dissolve", "chalk", "in", "aquafortis", "to", "the", "consistency", "of", "milk,", "and", "add", "to", "it", "a", "strong", "solution", "of", "silver;", "keep", "this", "liquor", "in", "a", "glass", "bottle", "well", "stopped;", "then", "cutting", "out", "from", "a", "piece", "of", "paper", "the", "letters", "you", "would", "have", "appear,", "paste", "it", "on", "the", "decanter,", "and", "lay", "it", "in", "the", "sun's", "rays", "in", "such", "a", "manner", "that", "the", "rays", "may", "pass", "through", "the", "spaces", "cut", "out", "of", "the", "paper", "and", "fall", "on", "the", "surface", "of", "the", "liquor", "the", "part", "of", "the", "glass", "through", "which", "the", "rays", "pass", "will", "be", "turned", "black,", "while", "that", "under", "the", "paper", "remains", "white;", "but", "particular", "care", "must", "be", "observed", "that", "the", "bottle", "be", "not", "moved", "during", "the", "operation.'", _
"The", "alchemists", "became", "so", "intent", "on", "discovering", "the", "far-famed", "philosopher's", "stone", "that", "they", "were", "unmindful", "of", "the", "accidental", "occurrence", "of", "more", "valuable", "discoveries.", "If", "they", "had", "not,", "then", "this", "little", "experiment", "in", "chemistry", "might", "have", "induced", "them", "to", "conduct", "a", "more", "thorough", "search", "into", "the", "principle.", "In", "that", "case,", "photogenic", "art", "would", "not", "now", "be", "claimed", "as", "a", "new", "discovery.", _
"Many", "individuals", "even", "assert", "that", "the", "Jugglers", "of", "India", "possessed", "a", "great", "secret", "for", "quite", "some", "time,", "a", "secret", "by", "which", "they", "were", "enabled,", "in", "a", "brief", "space,", "to", "copy", "the", "likeness", "of", "any", "individual", "by", "the", "action", "of", "light.", "This", "fact,", "if", "it", "is", "truly", "fact", "and", "not", "a", "myth,", "may", "account", "for", "the", "celebrated", "magic", "mirrors", "said", "to", "be", "possessed", "by", "these", "jugglers,", "and", "it", "was", "the", "probable", "cause", "of", "their", "power", "over", "the", "rest", "of", "the", "people.", _
"However,", "as", "early", "as", "1556", "experimenters", "established,", "as", "fact,", "that", "a", "combination", "of", "chloride", "and", "silver,", "called", "horn", "silver,", "from", "its", "appearance,", "could", "become", "blackened", "by", "the", "sun's", "rays.", "Then,", "in", "the", "latter", "part", "of", "the", "last", "century,", "a", "Mrs.", "Fulhame", "published", "an", "experiment", "by", "she", "effected", "a", "change", "in", "color", "of", "gold", "chloride,", "also", "by", "means", "of", "light.", "She", "concluded", "that,", "in", "her", "opinion,", "words", "might", "be", "someday", "be", "written", "in", "this", "way.", "These", "incidents", "are", "considered", "as", "the", "first", "steps", "towards", "the", "discovery", "of", "the", "photogenic", "art.", _
"Mr.", "Wedgwood's", "experiments", "can", "scarcely", "be", "said", "to", "be", "any", "improvement", "on", "either", "of", "these", "instances,", "since", "he", "failed", "to", "bring", "them", "to", "practical", "usefulness;", "his", "countrymen", "will", "have", "to", "be", "satisfied", "with", "awarding", "the", "honor", "of", "its", "complete", "adaptation", "for", "practical", "purposes,", "to", "MM.", "Niepce", "and", "Daguerre", "of", "France", "and", "to", "Professors", "Draper", "and", "Morse", "of", "New", "York.", _
"MM.", "Niepce", "and", "Daguerre", "pursued", "the", "subject", "simultaneously,", "without", "either,", "however,", "being", "aware", "of", "the", "experiments", "of", "his", "colleague", "in", "science.", "When", "after", "several", "years", "a", "chance", "circumstance", "made", "them", "acquainted,", "though,", "they", "entered", "into", "co-partnership,", "and", "conjointly", "they", "brought", "the", "art", "almost", "to", "perfection.", _
"M.", "Niepce", "presented", "his", "first", "paper", "on", "the", "subject", "to", "the", "Royal", "Society", "in", "1827,", "naming", "his", "discovery", "Heliography.", "What", "led", "him", "to", "the", "study", "of", "the", "principles", "of", "the", "photogenic", "art", "is", "unclear,", "but", "it", "was", "probably", "at", "least", "in", "part", "due", "to", "the", "previously", "mentioned", "alchemists", "experiments.", "But", "M.", "Daguerre,", "who", "is", "a", "celebrated", "dioramic", "painter,", "had", "a", "more", "easily", "understood", "motivation.", "He", "wished", "to", "employ", "some", "of", "the", "singularly", "changeable", "salts", "of", "silver", "to", "produce", "a", "peculiar", "class", "of", "effects", "in", "his", "paintings.", "His", "investigation", "of", "the", "method", "resulted", "in", "the", "discovery", "of", "the", "Daguerreotype,", "or", "photogenic", "drawing", "on", "plates", "of", "copper", "coated", "with", "silver.", _
"To", "this", "gentleman--", "and", "to", "his", "liberality--", "we", "Americans", "are", "indebted", "for", "the", "free", "use", "of", "his", "invention;", "and", "the", "large", "and", "increasing", "class", "of", "Daguerrean", "artists", "of", "this", "country", "should", "hold", "him", "in", "the", "most", "profound", "respect", "for", "it.", "He", "was", "not", "willing", "to", "see", "his", "invention", "confined", "to", "a", "few", "individuals", "who", "might", "monopolize", "the", "benefits", "to", "be", "derived", "from", "its", "practice,", "and", "thereby", "shut", "out", "all", "chance", "of", "improvement.", "Like", "a", "true,", "noble-hearted", "French", "gentleman,", "he", "wanted", "his", "invention", "to", "be", "spread", "freely", "throughout", "the", "entire", "world.", "With", "these", "views", "he", "opened", "negotiations", "with", "the", "French", "government,", "negotiations", "which", "benefitted", "both", "the", "inventors.", "And", "now", "France", "has", "the", _
"'glory", "of", "endowing", "the", "whole", "world", "of", "science", "and", "art", "with", "one", "of", "the", "most", "surprising", "discoveries", "that", "honor", "the", "land.'", "Nevertheless,", "the", "invention", "has", "been", "patented", "in", "England", "and", "the", "result", "is", "what", "might", "have", "been", "expected:", "English", "pictures", "are", "far", "below", "the", "standard", "of", "excellence", "of", "those", "taken", "by", "American", "artists.", "In", "addition,", "calotype,", "the", "name", "given", "to", "one", "of", "the", "methods", "of", "photogenic", "drawing", "on", "paper", "(discovered", "and", "perfected", "by", "Mr.", "Fox", "Talbot", "of", "England)", "is", "in", "precisely", "the", "same", "predicament,", "not", "only", "in", "that", "country", "but", "in", "the", "United", "States.", "Mr.", "Talbot", "is", "the", "patentee", "in", "both", "countries.", _
"But", "while", "he", "is", "a", "man", "of", "some", "wealth,", "he", "demands", "so", "high", "a", "price", "for", "a", "single", "right", "in", "this", "country,", "that", "no", "one", "is", "foolhardy", "enough", "to", "purchase", "one.", "When", "Mr.", "Talbot's", "patent", "for", "the", "United", "States", "expires", "and", "our", "native", "scientists", "have", "the", "opportunity", "to", "proceed", "with", "improvements,", "it", "is", "fairly", "certain", "that", "the", "calotype,", "in", "their", "hands,", "will", "supercede", "the", "Daguerreotype.", _
"Let", "them,", "therefore,", "study", "the", "principles", "of", "the", "art,", "then", "experiment,", "practice,", "and", "perfect", "themselves", "in", "all", "its", "details", "so", "that", "they", "will", "be", "prepared", "when", "that", "day", "arrives.", _
"It", "is", "to", "Professor", "Samuel", "F.", "B.", "Morse,", "the", "distinguished", "inventor", "of", "the", "Magnetic", "Telegraph", "of", "New", "York,", "that", "we", "are", "indebted", "for", "the", "application", "of", "photography", "to", "portrait", "taking.", "He", "was", "in", "Paris,", "for", "the", "purpose", "of", "presenting", "to", "the", "scientific", "world", "his", "Electro-Magnetic", "Telegraph,", "at", "the", "time", "(1838)", "M.", "Daguerre", "announced", "his", "splendid", "discovery;", "its", "astounding", "results,", "as", "well", "as", "its", "bearing", "on", "the", "art", "of", "design", "claimed", "his", "attention.", "In", "a", "letter", "on", "the", "subject,", "the", "Professor", "details", "the", "following", "interesting", "facts:")
Label9 = photo(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find photo(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub




Private Sub Girl_Next()
    On Error Resume Next
    Beep
    i = i + 1
girl = Array("It", "was", "the", "last", "night", "of", "the", "year.", "It", "had", "snowed,", "and", "it", "was", "very", "cold", "outside.", "Now", "it", "was", "nearly", "dark.", "In", "the", "cold", "and", "dark", "walked", "a", "poor", "little", "girl.", "She", "had", "no", "hat.", "She", "had", "no", "shoes.", "When", "she", "left", "home,", "she", "had", "shoes.", "They", "were", "very", "large.", "They", "had", "been", "her", "mother's.", "But", "she", "had", "run", "across", "the", "street", "to", "avoid", "a", "fast", "horse.", "The", "shoes", "were", "so", "big,", "they", "fell", "off", "as", "she", "ran.", _
"Another", "child", "had", "found", "one", "shoe.", "He", "ran", "off", "with", "it.", "She", "couldn't", "find", "the", "other.", "So", "the", "little", "girl", "walked", "on", "with", "her", "tiny,", "naked", "feet.", "They", "were", "quite", "red", "and", "numb", "from", "cold.", _
"The", "girl", "had", "many", "matches.", "She", "sold", "them", "for", "money.", "She", "kept", "most", "of", "them", "in", "an", "old", "apron.", "She", "held", "a", "bundle", "of", "them", "in", "her", "hand", "so", "people", "could", "see.", "But", "nobody", "had", "bought", "any", "from", "her", "the", "whole", "day.", "No", "one", "had", "given", "her", "a", "single", "cent.", _
"And", "now", "she", "was", "hungry.", "She", "had", "no", "money", "for", "food.", "She", "couldn't", "stay", "warm.", "Her", "body", "shook", "as", "she", "walked", "along.", "The", "flakes", "of", "snow", "covered", "her", "long,", "fair", "hair.", "It", "fell", "in", "lovely", "curls", "around", "her", "neck.", "But", "she", "didn't", "feel", "pretty.", _
"Right", "now", "she", "felt", "alone.", "She", "passed", "many", "houses.", "Candles", "shone", "in", "all", "the", "windows.", "And", "the", "air", "smelled", "of", "roast", "goose.", "It", "was", "New", "Year's", "Eve.", "People", "were", "celebrating.", "They", "were", "all", "happy,", "but", "she", "was", "not.", _
"She", "found", "a", "corner", "made", "by", "two", "houses.", "She", "sat", "down", "and", "tried", "to", "keep", "out", "of", "the", "wind.", "She", "drew", "her", "feet", "up", "close", "to", "her.", "But", "she", "could", "not", "keep", "them", "warm.", "Her", "whole", "body", "grew", "colder.", "But", "she", "couldn't", "go", "home.", "She", "had", "not", "sold", "any", "matches", "today.", "She", "had", "no", "money", "to", "bring", "to", "her", "family.", "Her", "father", "would", "be", "angry.", "And", "it", "was", "cold", "at", "home,", "too.", "In", "her", "room,", "the", "wind", "whistled.", "The", "roof", "had", "large", "cracks.", "They", "were", "stopped", "with", "straw", "and", "rags.", "But", "the", "cold", "came", "in", "just", "the", "same.", _
"Her", "little", "hands", "were", "almost", "numb", "with", "cold.", "She", "had", "a", "thought.", "A", "match", "might", "bring", "her", "some", "comfort.", "If", "she", "only", "dared", "take", "one", "out", "of", "the", "bundle.", "She", "could", "draw", "it", "against", "the", "wall.", "It", "would", "light.", "She", "could", "warm", "her", "fingers", "by", "it.", "She", "took", "one", "out.", "She", "lit", "it.", "How", "it", "blazed,", "how", "it", "burned!", "It", "was", "a", "warm,", "bright", "flame.", "It", "looked", "like", "a", "candle.", "She", "held", "her", "hands", "over", "it.", "It", "felt", "wonderful.", _
"It", "seemed", "to", "the", "little", "girl", "as", "if", "she", "were", "sitting", "before", "a", "large", "fire.", "She", "stretched", "out", "her", "feet", "to", "warm", "them,", "too.", "But", "the", "small", "flame", "went", "out.", "The", "fire", "was", "gone.", "She", "had", "only", "the", "remains", "of", "the", "burnt-out", "match", "in", "her", "hand.", _
"She", "rubbed", "another", "one", "against", "the", "wall.", "It", "burned", "as", "brightly", "as", "the", "first.", "The", "light", "fell", "on", "the", "wall.", "She", "thought", "she", "could", "see", "into", "the", "room", "beyond.", "On", "the", "table", "was", "spread", "a", "snow-white", "tablecloth.", "And", "there", "was", "a", "splendid", "china", "set.", "The", "roast", "goose", "was", "hot.", "It", "was", "stuffed", "with", "apples", "and", "dried", "plums.", "Her", "mouth", "watered", "with", "hunger.", _
"She", "reached", "out", "for", "the", "goose.", "Her", "fingers", "almost", "touched", "it.", "Then,", "the", "match", "went", "out.", "Nothing", "was", "left", "but", "the", "thick,", "cold,", "damp", "wall.", _
"She", "lit", "another", "match.", "Now", "she", "was", "sitting", "under", "the", "most", "beautiful", "Christmas", "tree.", "Thousands", "of", "lights", "were", "burning", "on", "the", "green", "branches.", "Pretty", "pictures", "hung", "on", "the", "walls.", "They", "looked", "as", "lovely", "as", "the", "ones", "she", "had", "seen", "in", "the", "shop", "windows.", "The", "little", "girl", "held", "out", "her", "hands", "towards", "them.", "Just", "then,", "the", "match", "went", "out.", _
"But", "the", "lights", "of", "the", "Christmas", "tree", "rose", "higher", "and", "higher.", "She", "saw", "them", "now", "as", "stars", "in", "the", "sky.", "One", "fell", "down", "and", "formed", "a", "long", "trail", "of", "fire.", _
"'Oh,'", "said", "the", "little", "girl.", "'Someone", "has", "just", "died.'", "Her", "old", "grandmother", "had", "told", "her", "the", "story.", "When", "a", "star", "falls,", "a", "soul", "goes", "up", "to", "Heaven.", "She", "wanted", "to", "see", "more,", "so", "she", "lit", "another", "match.", "In", "the", "bright", "light", "stood", "her", "grandmother.", "She", "was", "the", "only", "person", "who", "had", "loved", "the", "girl.", "Her", "face", "was", "kind", "and", "full", "of", "love.", _
"'Grandmother!'", "cried", "the", "little", "girl.", "'Please,", "take", "me", "with", "you!'", "But", "her", "grandmother", "started", "to", "fade", "as", "the", "match", "burned", "out.", _
"'No!'", "screamed", "the", "little", "girl.", "'Don't", "go!'", "And", "she", "rubbed", "all", "of", "her", "matches", "against", "the", "wall.", "She", "wanted", "to", "keep", "her", "grandmother", "near", "her.", "And", "the", "matches", "gave", "such", "a", "bright", "light.", "It", "was", "brighter", "even", "than", "at", "noon.", "Her", "grandmother", "became", "solid", "again.", "She", "took", "the", "little", "girl", "on", "her", "arm.", "Both", "flew", "up", "into", "the", "light.", "And", "there", "was", "no", "cold.", "There", "was", "no", "hunger.", "They", "were", "both", "in", "Heaven.", "The", "little", "girl", "was", "so", "happy.")
Label9 = girl(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find girl(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub




Private Sub Eighttools_Next()
    On Error Resume Next
    Beep
    i = i + 1
eighttools = Array("Managers", "need", "a", "new", "toolkit", "to", "manage", "at", "the", "portfolio", "level.", "Some", "are", "very", "new;", "others", "may", "look", "familiar.", "But", "don't", "be", "fooled.", _
"To", "create", "new", "value", "with", "a", "brand", "portfolio", "approach", "it", "is", "necessary", "to", "use", "some", "of", "these", "familiar", "tools", "in", "unfamiliar", "ways.", "It", "is", "like", "using", "a", "hammer", "in", "space", "-", "it", "won't", "feel", "or", "look", "the", "same", "as", "when", "you", "use", "it", "on", "the", "ground.", _
"Extensions", "are", "the", "simplest,", "oldest,", "and", "most", "common", "approach", "to", "boosting", "portfolio", "returns.", "The", "key", "to", "using", "extensions", "strategically", "is", "to", "recognize", "that", "there", "are", "two", "different", "types", _
"of", "extensions", "that", "can", "be", "applied", "to", "each", "brand", "portfolio.", "For", "less", "dense", "portfolios,", "extension", "opportunities", "exist", "in", "the", "open", "spaces", "of", "the", "molecule.", "We", "call", "these", "opportunities", "interstitial", "extensions", "because", "they", "fill", "the", "gaps", "between", "the", "strategic", "brands", "in", "a", "portfolio.", "The", "second", "type", "is", "the", "boundary", "extension.", "They", "create", "new", "brands", "on", "the", "outer", "edges", "of", "an", "existing", "brand", "molecule.", _
"Positioning", "requires", "consumers", "to", "learn", "what", "a", "brand", "portfolio", "stands", "for.", "Repositioning", "requires", "they", "first", "unlearn", "what", "it", "no", "longer", "stands", "for.", "Of", "all", "the", "tools,", "none", "comes", "with", "a", "higher", "risk/reward", "ratio", "than", "repositioning.", _
"Three", "components", "play", "a", "role", "in", "a", "repositioning:", "a", "new", "communications", "campaign,", "the", "launch", "of", "new", "products", "or", "services", "that", "signal", "the", "new", "direction", "of", "the", "portfolio,", "and", "alliances", "with", "companies", "or", "brand", "portfolios", "whose", "positioning", "lies", "close", "to", "where", "you", "want", "to", "take", "your", "brand", "portfolio.", _
"Pruning", "brands", "in", "a", "portfolio", "should", "occur", "periodically,", "in", "times", "of", "strong", "and", "weak", "performance", "alike.", "Managers", "can", "also", "improve", "the", "chances", "of", "success", "by", "using", "the", "brand", "molecule", "to", "decide", "which", "brands", "can", "be", "cut", "at", "minimal", "risk.", _
"Finally,", "the", "brands", "that", "are", "cut", "should", "be", "cleanly", "severed", "from", "the", "portfolio.", _
"Co-brands,", "leveraged", "well,", "can", "help", "bring", "a", "brand", "to", "your", "target,", "consumers", "in", "ways", "not", "afforded", "by", "the", "immediate", "business", "of", "your", "current", "portfolio.", "Linking", "brand", "portfolios", "can", "create", "bridges", "to", "dynamic", "new", "growth", "areas.", "It", "can", "reduce", "the", "risk", "of", "boundary", "brand", "extensions", "and", "put", "unused", "brand", "equity", "to", "work.", "Often", "the", "risk", "is", "more", "manageable", "than", "with", "other", "tools.", _
"Amalgamation", "consists", "of", "merging", "two", "brand", "portfolios", "and", "eliminating", "one", "of", "more", "of", "the", "existing", "brand", "names", "in", "the", "process.", "It", "is", "one", "of", "the", "most", "powerful", "tools", "for", "managing", "the", "brand", "portfolio.", "Of", "all", "the", "moves", "to", "the", "brand", "portfolio,", "amalgamation", "is", "the", "most", "difficult", "to", "reverse.", "Therefore,", "it", "requires", "careful", "planning.", _
"Scaling", "is", "a", "two-part", "strategy", "that", "can", "drive", "real", "growth", "for", "a", "brand", "portfolio.", "First,", "it", "allows", "a", "brand", "portfolio", "to", "follow", "its", "natural", "market,", "then", "it", "fills", "the", "upmarket", "or", "downmarket", "void", "with", "another", "portfolio.", "To", "make", "this", "work,", "a", "new", "customer", "base", "must", "exist", "for", "the", "moved", "brand.", "Also,", "the", "upscaled", "or", "downscaled", "brand", "should", "be", "distinctive", "but", "complementary", "to", "its", "brethren.", _
"Adapted", "from", "The", "Infinite", "Asset:", "Managing", "Brands", "to", "Build", "New", "Value", "by", "Sam", "Hill", "and", "Chris", "Lederer,", "HBS", "Press.", "Copyright", "2001", "Harvard", "Business", "School", "Publishing", "Corporation.", "All", "rights", "reserved.")
Label9 = eighttools(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find eighttools(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub


Private Sub Credits_Next()
    On Error Resume Next
    Beep
    i = i + 1
credits = Array("You've", "stayed", "through", "the", "movie's", "final", "credits", "-- a", "seemingly", "endless", "scrolling", "of", "mysterious", "job", "titles", "like", "best", "boy,", "grip,", "and", "Foley", "artist", "-- and", _
"if", "you're", "like", "many", "people,", "you", "have", "no", "idea", "what", "all", "these", "individuals", "do,", "or", "whether", "they're", "just", "strange", "names", "there", "to", "make", "the", "audience", "laugh.", "Well,", "all", "the", "jobs", "are", "real,", "and", "we'll", "get", "to", "them,", "along", "with", "gaffers,", "wranglers,", "and", "second", "second", "assistant", "directors,", "but", "first", "you", "need", "to", "know", "something", "about", "how", "the", "movie", "business", "operates.", _
"Credits", "aren't", "really", "there", "for", "the", "audience,", "although", "they", "can", "be", "interesting", "to", "watch.", "They're", "really", "there", "so", "the", "studios", "and", "other", "industry", "personnel", _
"will", "know", "who", "did", "what", "on", "the", "film.", "They", "help", "with", "future", "jobs,", "with", "better", "contracts,", _
"with", "more", "deals", "and", "obviously,", "with", "getting", "more", "money", "next", "time", "around.", "Credits", "are", "where", "the", "agent", "earns", "his", "or", _
"her", "money,", "because", "it's", "the", "agent", "who", "negotiates", "not", "only", "the", "worker's", "salary,", "but", "her", "credits,", "as", "well.", _
"First,", "draw", "a", "horizontal", "line.", "That", "line", "is", "an", "accountant's", "way", "of", "describing", "the", "costs", "of", "making", "a", "movie.", "Above", _
"the", "line", "go", "the", "names", "of", "all", "the", "creative", "people", "involved.", "Who", "are", "the", "creative", "people?", "They're", "the", "ones", "with", "agents.", "They", "are", "the", "director,", "the", "producers,", "the", "writers,", "the", "stars", "and", "featured", "actors,", "and", "the", "production", "designers,", "composers,", "casting", "directors,", "and", "cinematographers.", "Below", "the", "line", "go", "all", "the", "production", "people", "-- the", "crew", "members", "-- and", "the", "out-of-pocket", "costs", "of", "making", "the", "movie,", "like", "allocations", "for", "location", "rentals,", "building", "sets,", "purchasing", "or", "renting", "equipment,", "getting", "props,", "recording", "and", "editing", "the", "sound", "track,", "buying", "film", "stock", "and", "processing", "it", "in", "the", _
"laboratory,", "generating", "the", "optical", "and", "special", "effects", "and", "the", "like.", "The", "line", "is", "not", "a", "figure", "of", "speech.", "Every", "production", "is", "described", _
"as", "consisting", "of", "'above-the-line'", "and", "'below-the-line'", "costs.", _
"The", "Directors", "Guild", "of", "America", "permits", "a", "film", "to", "list", "only", "one", "director,", "even", "when", "it's", "generally", "known", "that", "two", "or", "more", "worked", "on", "it.", _
"Except", "in", "very", "rare", "cases", "there", "is", "only", "one", "directing", "credit.", "This", "is", "very", "good", "for", "a", "director's", "ego,", "certainly", "for", "the", "one", "who", "gets", "the", "credit,", "but", "also", "for", "the", "one", "who", "doesn't", "-- particularly", "if", "he", "or", "she", "had", "started", "production", "and", "then", "been", "removed", "by", "the", "producer", "at", "the", "insistence", "of", "the", "star", "or", "the", "studio.", "The", "public", "won't", "know", "that", "the", "removed", "one", "either", "screwed", "up", "or", "incurred", "the", "wrath", "of", "those", "more", "powerful.", "In", "either", "case,", "the", "removed", "director", "will", "live", "to", "direct", "another", "day.", "More", "rarely,", "a", "director", "will", "take", "the", _
"initiative", "and", "leave", "a", "production", "because", "of", "conflicts", "that", "cannot", "be", "resolved", "again", "usually", "with", "the", "studio", "or", "the", "star.", "All", "of", "this", "will", "most", "likely", "take", "place", "either", "before", "or", "during", "the", "first", "few", "days", "of", "shooting.", _
"Now", "to", "the", "below-the-line", "credits.", "When", "the", "movie", "fades", "to", "black,", "a", "new", "set", "of", "names", "begins", "to", "scroll", "up", "the", "screen.", "Sometimes", "the", "first", "credits", "we", "see", "will", "go", "to", "the", "production", "crew,", "the", "people", "who", "worked", "on", "the", " shooting,", "and", "sometimes", "they", "will", "be", "for", "the", "cast,", "often", "in", "order", "of", "prominence", "in", "the", "film,", "though", "sometimes", "in", "order", "of", "appearance", "or", "in", "alphabetical", "order.", "The", "production", "crew", "credits", "will", "be", "where", "all", "the", "'funny'", "titles", "start", "coming", "up,", "and", "here's", "what", "they", "do:", "The", "'gaffer'", "is", "the", "chief", "electrician.", "He", "works", "for", "the", "director", "of", "photography,", "setting", "up", "all", "the", "lights", "as", "they", "are", "needed", "for", "shooting,", _
"directing", "a", "crew", "of", "other", "electricians,", "preparing", "the", "basic", "lighting", "for", "the", "next", "scenes", "to", "be", "shot,", "ordering", "all", "lighting", "equipment", "and", "supplies,", "and", "so", "forth.", "And", "guess", "who", "his", "or", "her", "top", "assistant", "is?", "Yes,", "the", "'best", "boy.'", "Sexist", "but", "accurate,", "since", "they're", "almost", "always", "men.")
Label9 = credits(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find credits(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub


Private Sub Insects_Next()
    On Error Resume Next
    Beep
    i = i + 1
insects = Array("We", "live", "in", "a", "world", "filled", "with", "insects.", "When", "you", "play", "or", "work", "outside,", "it", "is", "easy", "to", "get", "bitten", "or", "stung.", "Bees,", "fire", "ants,", "and", "ticks", "are", "some", "of", "the", "most", "common", "types", "of", "biting", "and", "stinging", "insects.", "What", "are", "they?", "How", "can", "you", "avoid", "them?", "And", "what", "should", "you", "do", "if", "you", "get", "stung?", _
"A", "'bee'", "is", "the", "term", "many", "people", "use", "to", "describe", "any", "insect", "that", "has", "wings", "and", "a", "stinger.", "But", "bees", "are", "part", "of", "a", "larger", "scientific", "group.", "This", "group", "also", "includes", "wasps,", "hornets,", "and", "yellow", "jackets.", "All", "bees", "have", "several", "wings,", "hairy", "bodies,", "and", "stingers.", "They", "include", "both", "solitary", "and", "social", "species.", "They", "also", "have", "sucking", "and", "chewing", "mouthparts", "that", "let", "them", "gather", "nectar", "and", "pollen.", _
"Wasps", "have", "slender", "bodies", "with", "a", "narrow", "abdomen.", "They", "also", "have", "two", "pairs", "of", "wings", "and", "mouths", "adapted", "for", "biting", "or", "sucking.", "They", "can", "also", "be", "solitary", "or", "social", "insects.", "One", "type", "of", "social", "wasp", "is", "the", "yellow", "jacket.", _
"Hornets", "are", "large,", "strong", "wasps.", "The", "common", "European", "species", "is", "dark", "brown", "and", "yellow.", "It", "is", "very", "aggressive,", "and", "its", "sting", "is", "very", "severe.", "They", "build", "large,", "papery", "nests.", "The", "nests", "have", "combs", "arranged", "in", "columns.", "The", "hornets", "raise", "their", "young", "in", "the", "combs.", _
"Bees", "and", "wasps", "use", "stingers", "to", "protect", "themselves.", "You", "will", "feel", "a", "sharp", "pain", "at", "the", "site", "of", "the", "sting.", "The", "area", "may", "also", "feel", "hot,", "and", "it", "may", "itch.", "That's", "because", "there's", "venom", "in", "the", "stinger.", "A", "red", "bump", "surrounded", "by", "a", "white", "area", "will", "also", "develop.", _
"When", "a", "bee", "stings,", "it", "may", "leave", "the", "stinger", "in", "your", "skin.", "Remove", "it", "carefully.", "You", "don't", "want", "to", "push", "it", "farther", "in.", "Wash", "the", "site", "with", "soap", "and", "water.", "You", "can", "also", "mix", "up", "some", "meat", "tenderizer", "and", "water.", "Soak", "a", "cotton", "ball", "in", "the", "solution.", "Then", "rub", "the", "sting", "with", "it.", "If", "you", "don't", "have", "meat", "tenderizer,", "you", "can", "hold", "an", "ice", "cube", "against", "the", "site.", _
"For", "most", "people,", "a", "bee", "sting", "is", "painful", "but", "nothing", "more.", "However,", "about", "two", "out", "of", "one", "hundred", "people", "will", "have", "an", "allergic", "reaction.", "The", "symptoms", "may", "include", "hives", "(red", "patches", "on", "the", "skin", "that", "sting", "and", "itch),", "nausea,", "dizziness,", "swelling,", "and", "difficulty", "breathing.", "If", "you", "are", "allergic,", "you", "need", "to", "get", "medical", "attention", "right", "away.", "Medication", "can", "help", "control", "the", "swelling", "and", "the", "itching.", "It", "can", "help", "your", "body", "heal", "itself.", _
"The", "best", "way", "to", "keep", "from", "getting", "stung", "is", "to", "stay", "away", "from", "bees", "in", "the", "first", "place.", "Gardens", "or", "orchards", "in", "bloom", "usually", "have", "lots", "of", "bees.", "If", "you", "have", "to", "be", "there,", "take", "care.", "Keep", "your", "clothes", "and", "shoes", "on.", "Dress", "in", "long", "pants", "and", "a", "light,", "long-sleeved", "shirt.", "Avoid", "bright-colored", "clothing", "and", "clothing", "with", "flowery", "prints.", "The", "bees", "may", "think", "they're", "real", "flowers.", "Don't", "wear", "sweet-smelling", "perfume.", _
"When", "you", "go", "on", "picnics,", "cover", "the", "food.", "Soda", "and", "cake", "attract", "bees.", "Don't", "drink", "from", "open", "cans.", "Yellow", "jackets", "like", "to", "climb", "inside", "the", "cans", "and", "will", "sting", "from", "there.", "And", "if", "you", "see", "a", "bee", "flying", "around", "you,", "don't", "swat", "at", "it", "or", "run", "around.", "This", "makes", "it", "angry,", "and", "they", "may", "sting.", "Move", "slowly", "or", "stand", "still,", "and", "the", "bee", "will", "most", "likely", "fly", "away.", _
"Fire", "ants", "are", "another", "type", "of", "common", "insect.", "They", "are", "found", "all", "throughout", "the", "Southeastern", "and", "Southwestern", "United", "States.", "They", "build", "large", "mounds", "and", "can", "inflict", "a", "painful", "sting.", "The", "most", "common", "type", "is", "the", "red", "imported", "fire", "ant.", "It", "is", "reddish-brown", "in", "color.", "It", "is", "about", "half", "the", "size", "of", "a", "pencil", "eraser.", "These", "ants", "are", "very", "hostile.", _
"Red", "imported", "fire", "ants", "live", "in", "colonies.", "First,", "they", "nest", "in", "the", "ground.", "They", "are", "usually", "found", "wherever", "there", "is", "grass.", "However,", "some", "fire", "ants", "create", "nests", "in", "the", "walls", "of", "buildings.", "As", "they", "build", "the", "colony,", "they", "create", "a", "mound", "of", "dirt", "over", "the", "nest.", "The", "mound", "can", "grow", "up", "to", "18", "inches", "high", "and", "over", "two", "feet", "wide!", _
"When", "you", "get", "bitten", "by", "a", "fire", "ant,", "you", "will", "feel", "a", "sharp", "pain", "and", "a", "burning", "at", "the", "site.", "Each", "sting", "will", "turn", "into", "an", "itchy", "blister", "over", "the", "next", "day.", _
"The", "fire", "ants'", "bite", "contains", "venom.", "This", "can", "cause", "the", "area", "to", "swell", "up", "quite", "a", "bit.", "You", "treat", "the", "bite", "the", "same", "way", "you", "treat", "a", "bee", "sting.", "Wash", "the", "area", "with", "soap", "and", "water.", "Rub", "the", "sting", "with", "a", "cotton", "ball", "soaked", "in", "meat", "tenderizer", "and", "water.", "You", "can", "also", "hold", "an", "ice", "cube", "against", "it.", "Some", "people", "have", "an", "allergic", "reaction", "to", "fire", "ants.", "The", "symptoms", "are", "just", "like", "the", "ones", "from", "a", "bee", "sting.", "If", "you", "have", "any", "of", "these", "symptoms,", "see", "a", "doctor", "as", "soon", "as", "possible.", _
"The", "best", "way", "to", "avoid", "getting", "bitten", "is", "to", "avoid", "fire", "ants", "and", "their", "mounds.", "If", "you're", "near", "a", "mound,", "keep", "your", "shoes", "on.", "And", "don't", "ever", "poke", "at", "it", "or", "try", "to", "play", "with", "it.", "The", "ants", "are", "hostile,", "and", "they", "will", "attack.", _
"Ticks", "are", "a", "third", "type", "of", "common", "insect.", "They", "are", "any", "one", "of", "a", "large", "number", "of", "parasitic", "mites.", "They", "do", "not", "have", "wings.", "They", "attach", "themselves", "to,", "and", "suck", "the", "blood", "of,", "cattle,", "dogs,", "and", "people.", "When", "filled", "with", "blood,", "they", "swell", "up", "and", "usually", "become", "red", "in", "color.", "The", "young", "are", "active", "and", "have", "six", "legs.", "Many", "carry", "diseases", "that", "can", "hurt", "animals", "and", "people.", "One", "common", "disease", "is", "Lyme", "Disease.", _
"Tick", "bites", "are", "not", "like", "bee", "stings", "or", "fire", "ant", "bites.", "They", "do", "not", "hurt.", "Sometimes", "you", "know", "you've", "been", "bitten", "because", "the", "tick", "is", "still", "attached.", "Sometimes", "all", "you", "may", "see", "is", "some", "redness", "around", "the", "area", "of", "the", "bite.", _
"You", "should", "check", "your", "clothes", "and", "body", "after", "you've", "been", "in", "tall", "grass", "or", "in", "the", "woods.", "Ticks", "like", "shrubbery.", "They", "stay", "close", "to", "the", "ground.", "They", "jump", "on", "people", "or", "animals", "that", "pass", "by.", "Make", "sure", "to", "wear", "long", "sleeves", "and", "long", "pants.", "Tuck", "your", "pants", "into", "your", "socks", "so", "they", "can't", "crawl", "underneath.", "Also", "put", "a", "tick", "collar", "on", "your", "dog", "before", "you", "take", "it", "with", "you.", "Dogs", "get", "ticks", "in", "their", "coats.", _
"You", "don't", "need", "medical", "attention", "for", "most", "tick", "bites.", "Usually", "you", "can", "simply", "pull", "the", "tick", "out", "by", "yourself.", "However,", "if", "you", "live", "in", "an", "area", "where", "ticks", "carry", "diseases,", "check", "with", "your", "doctor", "to", "see", "if", "you", "need", "medicine", "to", "treat", "the", "bites.")
Label9 = insects(i)
  With RichTextBox1
  .SetFocus
  .SelStart = 0
  .SelLength = 0
  .Find insects(i), , , rtfWholeWord
  .SelUnderline = True
  End With
Text1.SetFocus
End Sub

