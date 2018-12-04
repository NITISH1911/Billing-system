VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form login 
   Caption         =   "Form4"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form4"
   Picture         =   "login.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "SHOW PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      TabIndex        =   8
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   4440
      Top             =   7200
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIGN-UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   6
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox PASS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   10080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   5040
      Width           =   4095
   End
   Begin VB.TextBox nme 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   2
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label WAIT 
      Caption         =   "PLEASE WAIT"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As New ADODB.Connection
Dim I As Integer
Dim RS As New ADODB.Recordset
Private Sub Check1_Click()
If Check1.Value = 1 Then
PASS.PasswordChar = ""
Else
PASS.PasswordChar = "*"
End If
End Sub


Private Sub Command1_Click()

While Not RS.EOF
If nme.Text = RS.Fields("name") And PASS.Text = RS.Fields("password") Then
MsgBox "LOGIN SUCCESSFULLY", vbDefaultButton3

Timer1.Enabled = True
WAIT.Visible = True
ProgressBar1.Visible = True
ProgressBar1.Enabled = True
End If
RS.MoveNext
Wend
If ProgressBar1.Enabled = False Then
MsgBox "LOGIN UNSUCCESSFUL", vbExclamation

I = I + 1
End If
If I = 3 Then
MsgBox "LIMIT OVER", vbCritical

Unload Me
End If
End Sub

Private Sub Command2_Click()
CN.Close
signup.Show
Unload Me

End Sub

Private Sub Form_Load()
I = 0
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\NITISH\LOGIN.MDB;"
CN.Open
RS.Open "signup", CN, adOpenDynamic, adLockOptimistic
Timer1.Enabled = False
Timer1.Interval = 40
ProgressBar1.Enabled = False
ProgressBar1.Visible = False
WAIT.Visible = False
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
Form1.Show
Unload Me
End If
End Sub
