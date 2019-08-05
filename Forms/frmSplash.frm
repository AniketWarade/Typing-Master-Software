VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Typing Master Software"
   ClientHeight    =   9735
   ClientLeft      =   2640
   ClientTop       =   990
   ClientWidth     =   16080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   15360
      Top             =   4320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      DragIcon        =   "frmSplash.frx":164A
      Height          =   9930
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   16080
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1200
         ScaleHeight     =   195
         ScaleWidth      =   8835
         TabIndex        =   9
         Top             =   8760
         Width           =   8895
         Begin VB.Image Image1 
            Height          =   180
            Left            =   0
            Picture         =   "frmSplash.frx":2C94
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Top             =   7560
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11160
         TabIndex        =   8
         Top             =   8400
         Width           =   4215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "4] Piyesh. K. Patil"
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
         Left            =   8400
         TabIndex        =   6
         Top             =   7440
         Width           =   3420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "3] Subodh. V. Deshmukh"
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
         Left            =   8400
         TabIndex        =   5
         Top             =   6480
         Width           =   4800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "2] Sandesh. N.Bhoir"
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
         Left            =   8400
         TabIndex        =   4
         Top             =   5520
         Width           =   3915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "1] Aniket. V. Warade"
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
         Left            =   8400
         TabIndex        =   3
         Top             =   4560
         Width           =   4005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Made By:-"
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
         Left            =   8040
         TabIndex        =   2
         Top             =   3120
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Typing Master Software"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   6720
         TabIndex        =   1
         Top             =   1080
         Width           =   9285
      End
      Begin VB.Image imgLogo 
         Height          =   7185
         Left            =   0
         Picture         =   "frmSplash.frx":2FD5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim i As Integer
Option Explicit

Private Sub Form_load()
   File1.FileName = App.Path
   X = File1.ListCount
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Load Form1
    Form1.Show
    Unload Me
End Sub

Private Sub Timer1_timer()
If (Image1.Left <= 6600) Then
    Image1.Left = Image1.Left + 1000
Else
    Image1.Left = 0
End If
If (i <= X) Then
    Label7.Caption = File1.List(i)
    i = i + 1
Else
    Load Form1
    Form1.Show
    Unload Me
End If
End Sub
