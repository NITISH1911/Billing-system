VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      DataField       =   "tax"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   11
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      DataField       =   "total"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   10
      Top             =   3720
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9480
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\NITISH\CUST.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\NITISH\CUST.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "customer"
      Caption         =   ""
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
   Begin VB.TextBox Text4 
      DataField       =   "phno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   6
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      DataField       =   "total_amount"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   5
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      DataField       =   "cname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Tax"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Customer  Email"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Customer    Ph.No"
      BeginProperty Font 
         Name            =   "News701 BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Picture = LoadPicture("F:\NITISH\1.jpg")
End Sub
