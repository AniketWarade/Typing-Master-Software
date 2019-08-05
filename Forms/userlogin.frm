VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Typing Master Software"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "userlogin.frx":0000
   LinkTopic       =   "Form19"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   11295
      Begin VB.TextBox Text1 
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
         Left            =   2880
         TabIndex        =   6
         Top             =   1320
         Width           =   5655
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "userlogin.frx":164A
         Height          =   5295
         Left            =   2880
         TabIndex        =   7
         Top             =   2040
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   9340
         _Version        =   393216
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   44
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter  Your Name :-"
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
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   4485
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
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
      Left            =   18720
      Picture         =   "userlogin.frx":165F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   7320
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"userlogin.frx":1738
      OLEDBString     =   $"userlogin.frx":17C7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "typing_user"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "User Login and Registration"
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
      Left            =   7560
      TabIndex        =   8
      Top             =   720
      Width           =   7110
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   9255
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   13455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Enter"
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
      TabIndex        =   3
      Top             =   9840
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Typing Master Software"
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
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6045
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo errmsg
For i = 0 To Adodc1.Recordset.RecordCount
If Text1.Text = DataGrid1.Text Then
GoTo Yes
Else
DataGrid1.Row = DataGrid1.Row + 1
End If
Next

Yes:
If Len(Text1.Text) > 0 Then
MsgBox "User Logged in Successfully !!!", vbInformation, "!!! User Registration !!!"
Load Form2
Form2.Show
Unload Me
Else
MsgBox "Please Enter Your Name !!!", vbInformation, "!!! Alert !!!"
End If
Exit Sub


errmsg:
If Len(Text1.Text) > 0 Then
yes_no1 = MsgBox("User " & Text1.Text & " not found, are you a new user?", vbYesNo, "!!! Alert !!!")
Adodc1.Recordset.AddNew
If yes_no1 = vbYes Then
yes_no2 = MsgBox("Do you want to create an user account?", vbYesNo, "!!! Alert !!!")
If yes_no2 = vbYes Then
DataGrid1.Text = Text1.Text
MsgBox "User Added & Logged in Successfully !!!", vbInformation, "!!! User Registration !!!"
Load Form2
Form2.Show
Unload Me
Else

End If
Else

End If
Else
MsgBox "Please Enter Your Name !!!", vbInformation, "!!! Alert !!!"
End If
DataGrid1.Row = 0
End Sub


Private Sub Command2_Click()
On Error GoTo errmsg
If Adodc1.Recordset.RecordCount > 0 Then
confirm = MsgBox("Are you sure you want to delete this user.", vbYesNo, "!!! Delete User !!!")
If confirm = vbYes Then
Adodc1.Recordset.Delete
MsgBox "User Deleted Successfully !!!", vbInformation, "!!! Alert !!!"
Else

End If
Else
MsgBox "No user available to delete !!!", vbInformation, "!!! Alert !!!"
End If
Exit Sub
errmsg:
MsgBox "Sorry user not deleted. Please try again.", vbInformation, "!!! Warning !!!"
End Sub


Private Sub Form_Load()
Adodc1.Visible = False
Text1.Text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
Form10.Label13.Caption = Text1.Text
End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1.Text = ""
DataGrid1.Row = 0
End Sub


