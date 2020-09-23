VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Viewer"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Columns         =   6
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6705
   End
End
Attribute VB_Name = "Form2"
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


Public Sub CHARACTER(fontname As String)
    Me.Icon = Form1.Icon
    List1.Clear
    Dim i As Integer
    For i = 0 To 255
        List1.AddItem Chr(i)
    Next i
    With List1
        .Font.Name = fontname
        .Font.Size = 28
        Me.Height = .Height + 450
    End With
    Me.Visible = True
End Sub

Private Sub List1_Click()
    Me.Caption = "Character Viewer   " & List1.Text & " = Alt + 0" & Asc(List1.Text)
End Sub
