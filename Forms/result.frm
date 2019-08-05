VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "result.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   10335
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   13455
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Height          =   3015
         Left            =   3000
         TabIndex        =   2
         Top             =   6840
         Width           =   7575
         Begin VB.Label Label11 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   1080
            TabIndex        =   12
            Top             =   960
            Width           =   5415
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Comment:-"
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
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   2145
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Height          =   4935
         Left            =   3000
         TabIndex        =   1
         Top             =   1560
         Width           =   7575
         Begin VB.Label Label8 
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
            Left            =   2880
            TabIndex        =   16
            Top             =   4320
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Accuracy:-"
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
            Left            =   240
            TabIndex        =   15
            Top             =   4320
            Width           =   1650
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3120
            TabIndex        =   14
            Top             =   600
            Width           =   105
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:-"
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
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   2640
         End
         Begin VB.Label Label7 
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
            Left            =   2880
            TabIndex        =   8
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label6 
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
            Left            =   2880
            TabIndex        =   7
            Top             =   2400
            Width           =   1935
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
            Left            =   2880
            TabIndex        =   6
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Speed:-"
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
            Left            =   240
            TabIndex        =   5
            Top             =   3360
            Width           =   1770
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Speed:-"
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
            Left            =   240
            TabIndex        =   4
            Top             =   2400
            Width           =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Duration:-"
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
            Left            =   240
            TabIndex        =   3
            Top             =   1440
            Width           =   1545
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Results for the completed test."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   7785
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form8
Form8.Show
Unload Me
End Sub
