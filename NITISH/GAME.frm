VERSION 5.00
Begin VB.Form GAME 
   Caption         =   "Form4"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   LinkTopic       =   "Form4"
   Picture         =   "GAME.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   18
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton button3 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      MaskColor       =   &H00800000&
      TabIndex        =   10
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button9 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      MaskColor       =   &H00800000&
      TabIndex        =   9
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button6 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button8 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      MaskColor       =   &H00800000&
      TabIndex        =   7
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button5 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button2 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      MaskColor       =   &H00800000&
      TabIndex        =   5
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button7 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00800000&
      TabIndex        =   4
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button4 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00800000&
      TabIndex        =   3
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton button1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00800000&
      TabIndex        =   2
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton new 
      Caption         =   "NEW ROUND"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2040
      TabIndex        =   1
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TIC TAC TOE"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2160
      TabIndex        =   17
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X score"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "O score"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "GAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ply1, ply2 As String
Public i, j As Integer

Private Sub button1_Click()
If Label1.Visible = True Then
button1.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button1.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button1.Enabled = False
End Sub

Private Sub button2_Click()
If Label1.Visible = True Then
button2.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button2.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button2.Enabled = False
End Sub

Private Sub button3_Click()
If Label1.Visible = True Then
button3.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button3.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button3.Enabled = False
End Sub

Private Sub button4_Click()
If Label1.Visible = True Then
button4.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label4.Visible = True Then
button4.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button4.Enabled = False
End Sub

Private Sub button5_Click()
If Label1.Visible = True Then
button5.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button5.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button5.Enabled = False
End Sub

Private Sub button6_Click()
If Label1.Visible = True Then
button6.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button6.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button6.Enabled = False
End Sub

Private Sub button7_Click()
If Label1.Visible = True Then
button7.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button7.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button7.Enabled = False
End Sub

Private Sub button8_Click()
If Label1.Visible = True Then
button8.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button8.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button8.Enabled = False
End Sub

Private Sub button9_Click()
If Label1.Visible = True Then
button9.Caption = "O"
Call txt_chk
Label1.Visible = False
Label2.Visible = True
GoTo 5000
End If
If Label2.Visible = True Then
button9.Caption = "X"
Call txt_chk
Label1.Visible = True
Label2.Visible = False
GoTo 5000
End If
5000: button9.Enabled = False
End Sub


Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub exit_Click()
Unload Me

End Sub


Private Sub Form_Load()

Label1.Caption = "player 1"
Label2.Caption = "player 2"
Form1.Picture = LoadPicture("F:\NITISH\8.jpg")
MsgBox ("Welcome to tic tac toe game by Nitish Mehta")
ply1 = InputBox("Enter Name Of Player 1", "Name Of Player 1")
ply2 = InputBox("Enter Name Of Player 2", "Name Of Player 2")
If ply1 = "" Then
ply1 = "player 1"
End If
If ply2 = "" Then
ply2 = "player 2"
End If
Label2.Visible = False

End Sub

Private Sub new_Click()
button1.Caption = ""
button1.Enabled = True
button2.Caption = ""
button2.Enabled = True
button3.Caption = ""
button3.Enabled = True
button4.Caption = ""
button4.Enabled = True
button5.Caption = ""
button5.Enabled = True
button6.Caption = ""
button6.Enabled = True
button7.Caption = ""
button7.Enabled = True
button8.Caption = ""
button8.Enabled = True
button9.Caption = ""
button9.Enabled = True
End Sub

Function txt_chk()
 i = lbl1.Caption
 j = lbl2.Caption
'player 1 win
If button1.Caption = "O" And button2.Caption = "O" And button3.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button4.Caption = "O" And button5.Caption = "O" And button6.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button7.Caption = "O" And button8.Caption = "O" And button9.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button1.Caption = "O" And button4.Caption = "O" And button7.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button2.Caption = "O" And button5.Caption = "O" And button8.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button3.Caption = "O" And button6.Caption = "O" And button9.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button1.Caption = "O" And button5.Caption = "O" And button9.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button3.Caption = "O" And button5.Caption = "O" And button7.Caption = "O" Then
MsgBox "Congratulations " & ply1 & " win"
lbl1.Caption = i + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
'player 2 win
End If
If button1.Caption = "X" And button2.Caption = "X" And button3.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button4.Caption = "X" And button5.Caption = "X" And button6.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button7.Caption = "X" And button8.Caption = "X" And button9.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button1.Caption = "X" And button4.Caption = "X" And button7.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button2.Caption = "X" And button5.Caption = "X" And button8.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button3.Caption = "X" And button6.Caption = "X" And button9.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button1.Caption = "X" And button5.Caption = "X" And button9.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
ElseIf button3.Caption = "X" And button5.Caption = "X" And button7.Caption = "X" Then
MsgBox "Congratulations " & ply2 & " win"
lbl2.Caption = j + 1
button1.Enabled = False
button2.Enabled = False
button3.Enabled = False
button4.Enabled = False
button5.Enabled = False
button6.Enabled = False
button7.Enabled = False
button8.Enabled = False
button9.Enabled = False
End If
End Function
