VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form signup 
   Caption         =   "Form4"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form4"
   Picture         =   "signup.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dob 
      Height          =   735
      Left            =   10560
      TabIndex        =   12
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95944705
      CurrentDate     =   43151
   End
   Begin VB.TextBox EMAIL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   11
      Top             =   4920
      Width           =   4095
   End
   Begin VB.TextBox PHNO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   10
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox PASSWORD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   9
      Top             =   7320
      Width           =   4095
   End
   Begin VB.TextBox nme 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   8
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ALREADY HAVE A ACCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11040
      TabIndex        =   7
      Top             =   8880
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIGN-UP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   6
      Top             =   8880
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIGN-UP"
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
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub Command1_Click()


If nme.Text = "" Or EMAIL.Text = "" Or PHNO.Text = "" Or PASSWORD.Text = "" Or dob.Value = "" Then
MsgBox "ENTER ALL THE INFORMATION"
ElseIf Not (IsNumeric(PHNO.Text)) Then
MsgBox "Phone number is in numbers only", vbApplicationModal
Else
RS.Fields(0) = nme.Text
RS.Fields(1) = dob.Value
RS.Fields(2) = PASSWORD.Text
RS.Fields(3) = PHNO.Text
RS.Fields(4) = EMAIL.Text
RS.AddNew
MsgBox "YOU SIGNED IN"
CN.Close
login.Show
Unload Me

End If

End Sub

Private Sub Command2_Click()
CN.Close
login.Show
Unload Me
End Sub


Private Sub Form_Load()
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\NITISH\LOGIN.MDB;"
CN.Open
RS.Open "signup", CN, adOpenDynamic, adLockOptimistic
End Sub
