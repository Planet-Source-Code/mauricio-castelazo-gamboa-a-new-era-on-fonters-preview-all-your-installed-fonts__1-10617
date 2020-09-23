VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font List Viewer: By Mauricio Castelazo...Vote for this program !!!"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "fontviewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Height          =   225
      Left            =   3150
      TabIndex        =   14
      ToolTipText     =   "Check this box if you want to use a defined example text"
      Top             =   4050
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3390
      TabIndex        =   13
      Text            =   "Example Text"
      Top             =   4050
      Width           =   3945
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tip !!!"
      Height          =   315
      Left            =   1980
      TabIndex        =   7
      Top             =   4050
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "fontviewer.frx":0F7A
      Left            =   7470
      List            =   "fontviewer.frx":0F93
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4050
      Width           =   675
   End
   Begin VB.PictureBox Picture1 
      Height          =   3705
      Left            =   120
      ScaleHeight     =   3645
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   150
      Width           =   8025
      Begin VB.Frame LETRAS 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7725
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Visit my Home page ""www.cyberlatino.com.mx"""
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   900
            TabIndex        =   6
            Top             =   60
            Width           =   5955
         End
      End
      Begin VB.VScrollBar Scroll 
         Height          =   3645
         LargeChange     =   250
         Left            =   7740
         SmallChange     =   5
         TabIndex        =   1
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Fontlist"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4050
      Width           =   1875
   End
   Begin ComctlLib.ProgressBar Progress 
      Height          =   135
      Left            =   150
      TabIndex        =   10
      Top             =   4470
      Visible         =   0   'False
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   3915
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   8235
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   705
      Left            =   30
      TabIndex        =   12
      Top             =   3990
      Width           =   8235
   End
   Begin VB.Label Label3 
      Height          =   225
      Left            =   3060
      TabIndex        =   9
      Top             =   4110
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size:"
      Height          =   225
      Left            =   6720
      TabIndex        =   4
      Top             =   4140
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code will generate a fontlist with all your installed fonts
'This program doesn't use any control (like a RTF control)
'When you run the program, first, it will load all fonts and put them
'in a list, this is only because, I wanted to sort the antire font list.
'
'This code is free, and I will apreciate your comments or suggestions about it
'You can email me at: castelazo@cyberlatino.com.mx
'
'Author: Mauricio Castelazo Gamboa
'Homepage: http://www.cyberlatino.com.mx
'
'Remember, with this code, I only want to show how to make the font list, of
'course, you can add many features to the fonter, but that is not hard to do.
'Good look.
'Have fun...
'
'P.S. If you use this code in a program, I really would like to receive your
'Exe program. Also, if you give me credit would be nice too.




Dim BEFORE As Integer

Private Sub Check1_Click()
    Text1.Enabled = CBool(Check1.Value)
End Sub

Private Sub Command1_Click()
    On Error GoTo fin:
    Dim i As Integer
    Dim TOPE As Double
    
    Label3.Caption = "Creating list!!!"
    Label1(0).fontname = List1.List(0)
    Label1(0).FontSize = CInt(Combo1.Text)
    Label1(0).Left = centrar(LETRAS.Width, Label1(0).Width)
    Label1(0).ToolTipText = List1.List(0)
    If Check1.Value = 0 Then
        Label1(0).Caption = List1.List(i)
    Else
        Label1(0).Caption = Text1.Text
    End If

    TOPE = Label1(0).Height + Label1(0).Top
    Progress.Visible = True
    Progress.Max = List1.ListCount
    Progress.Value = 0
    For i = 1 To List1.ListCount - 1
        Load Label1(i)
        Label1(i).Top = TOPE
        Label1(i).AutoSize = True
        Label1(0).FontSize = Int(Combo1.Text)
        Label1(i).Font.Name = List1.List(i)
        If Check1.Value = 0 Then
            Label1(i).Caption = List1.List(i)
        Else
            Label1(i).Caption = Text1.Text
        End If
        Label1(i).ToolTipText = List1.List(i)
        Label1(i).Left = centrar(LETRAS.Width, Label1(i).Width)
        TOPE = Label1(i).Height + TOPE + 15
        Label1(i).Visible = True
        Progress.Value = i
    Next i
    Combo1.Enabled = False
    Label3.Caption = "Done!!!"
    Progress.Visible = False
    LETRAS.Height = TOPE + Label1(i - 1).Height
    Scroll.Min = 0
    Scroll.Max = (LETRAS.Height - Picture1.Height) / 10
    Command1.Enabled = False
    i = MsgBox("Now you can preview all installed fonts in a hurry!!!" & vbCrLf & "And you can double click a font to see it's character set", vbInformation, "Font List Viewer")
    Exit Sub
fin:
i = MsgBox(Err.Description, vbCritical)
End Sub


Private Function centrar(Padre As Long, Hijo As Long) As Long
    centrar = (Padre - Hijo) / 2
End Function


Private Sub Command2_Click()
    Dim i As Byte
    i = MsgBox("If you have many fonts installed on your system (more than 550)," & vbCrLf & _
             "use a font size like 14 or 16. If not, you won't be able to see all fonts.", vbInformation, "Font List Creator")
End Sub

Private Sub Form_Load()
    Dim Y As Byte
    Y = MsgBox("Thanks for downloading this code, I hope it to be useful for you." & vbCrLf & "If you like this code, please, VOTE FOR IT!!!" & vbCrLf & vbCrLf & "Else, visit my homepage www.cyberlatino.com.mx", vbInformation, "Image merger")
    LLENAR_FONTLIST List1
    Label3.Caption = "You have " & List1.ListCount & " on your system"
    Combo1.ListIndex = 2
    BEFORE = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Y As Byte
    Y = MsgBox("Remember, if you like this code, please, VOTE FOR IT!!!" & vbCrLf & "at http://www.planet-source-code.com/vb" & vbCrLf & vbCrLf & "Also, visit my homepage www.cyberlatino.com.mx", vbInformation, "Image merger")
End Sub

Private Sub Label1_Click(Index As Integer)
    Label1(BEFORE).ForeColor = vbBlack
    Label1(BEFORE).FontBold = False
    Label1(Index).ForeColor = 8914708
    Label1(Index).Font.Bold = True
    Label3.Caption = Label1(Index).Caption
    BEFORE = Index
End Sub

Private Sub Label1_DblClick(Index As Integer)
    Form2.CHARACTER Label1(Index).Caption
End Sub

Private Sub Scroll_Change()
    Dim B As Double
    B = -(Scroll.Value)
    B = B * 10
    LETRAS.Top = B
End Sub

Private Sub Scroll_Scroll()
    Dim B As Double
    B = -(Scroll.Value)
    B = B * 10
    LETRAS.Top = B
End Sub


