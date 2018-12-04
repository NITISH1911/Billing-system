VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   19305
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   62
      Top             =   8640
      Width           =   4455
   End
   Begin VB.CommandButton com 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "COST OF MEAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   61
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton reset 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   60
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton exit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   59
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton recipt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECIPT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   58
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      TabIndex        =   57
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "ADVERTISMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   56
      Top             =   7800
      Width           =   4455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000C&
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   7200
      TabIndex        =   47
      Top             =   2040
      Width           =   6015
      Begin VB.TextBox cno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   55
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox cemail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   54
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox clname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   53
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox cname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   52
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label13 
         Caption         =   "Customer Ph. NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Customer Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Customer last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000C&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   13800
      TabIndex        =   25
      Top             =   2040
      Width           =   4575
      Begin VB.CommandButton Command8 
         Caption         =   "AC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   43
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton subt 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   42
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton mul 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   41
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton div 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   40
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton add 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   39
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton equal 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   38
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   1560
         TabIndex        =   37
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   600
         TabIndex        =   36
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   2520
         TabIndex        =   35
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   1560
         TabIndex        =   34
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   600
         TabIndex        =   33
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   2520
         TabIndex        =   32
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   1560
         TabIndex        =   31
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   600
         TabIndex        =   30
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   29
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Left            =   600
         TabIndex        =   26
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.Frame Total 
      BackColor       =   &H8000000C&
      Caption         =   "Cost of Services"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7200
      TabIndex        =   3
      Top             =   6240
      Width           =   6015
      Begin VB.TextBox totxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   24
         Text            =   " "
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox ttxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   23
         Text            =   " "
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox sttxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   22
         Text            =   " "
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      Caption         =   "Item Sold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      TabIndex        =   2
      Top             =   6240
      Width           =   6015
      Begin VB.TextBox msoldtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   18
         Text            =   " "
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox bsoldtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   17
         Text            =   " "
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox dsoldtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   16
         Text            =   " "
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Meal Sold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Burger Sold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Drink Sold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Select Meal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   6015
      Begin VB.ListBox VEGCMB 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   1080
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox CHIKENCMB 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   1080
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton minus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   45
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton plus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   44
         Top             =   3240
         Width           =   495
      End
      Begin VB.CheckBox meal 
         Caption         =   "Meal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox veg 
         Caption         =   "Veg. Burger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox chicken 
         Caption         =   "Chicken Burger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox drinktxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   8
         Text            =   "0"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox drinkcmb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "Form1.frx":925F3
         Left            =   1920
         List            =   "Form1.frx":925F5
         TabIndex        =   7
         Text            =   " "
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox mealtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Text            =   " "
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox vegtxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Text            =   " "
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox chickentxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   46
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Drinks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BILLING SYSTEM"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1215
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chickenburger As Integer
Dim vegburger As Integer
Dim meal1 As Integer
Dim drink As Integer
Dim I As Integer
Dim chiken1 As Integer
Dim veg1 As Integer
Dim mel As Integer

Dim name1, last, EMAIL, no As String

Dim preval As Double
Dim curval As Double
Dim result As Double
Dim choice As String

Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub add_Click()
Text1.Text = ""
preval = curval
curval = 0
choice = "+"

End Sub

Private Sub chicken_Click()
If chicken.Value = 1 Then
chickenburger = 100
CHIKENCMB.Visible = True

Else
chickenburger = 0
CHIKENCMB.Visible = False
End If
End Sub

Private Sub CHIKENCMB_Click()
CHIKENCMB.Visible = False
End Sub

Private Sub com_Click()


A = Val(chickentxt.Text)
b = Val(vegtxt.Text)
c = Val(mealtxt.Text)
d = Val(drinktxt.Text)

dsoldtxt.Text = d
bsoldtxt.Text = A + b
msoldtxt.Text = c


If chicken.Value = 1 Then
chicken1 = 100 * A
Else
chicken1 = 0
End If

If veg.Value = 1 Then
veg1 = 80 * b
Else
veg1 = 0
End If

If meal.Value = 1 Then
mel = 200 * c
Else
mel = 0
End If

Dim drink1 As Integer
drink1 = drink * d

Dim Total As Double
Dim subtotal As Integer
Dim tax As Double

subtotal = chicken1 + veg1 + mel + drink1
tax = subtotal * (5 / 100)
Total = subtotal + tax

sttxt.Text = subtotal
ttxt.Text = tax
totxt.Text = Total
End Sub



Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text & Command1(Index).Caption
curval = Val(Text1.Text)
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
CN.Close
Unload Me
GAME.Show
End Sub

Private Sub Command8_Click()
preval = curval = 0
Text1.Text = 0
End Sub

Private Sub div_Click()
Text1.Text = ""
preval = curval
curval = 0
choice = "/"

End Sub

Private Sub equal_Click()
Select Case choice
Case "+"
result = curval + preval
Text1.Text = Str(result)

Case "-"
result = curval - preval
Text1.Text = Str(result)

Case "*"
result = curval * preval
Text1.Text = Str(result)

Case "/"
result = curval / preval
Text1.Text = Str(result)
 
End Select
 curval = result
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Form1.Picture = LoadPicture("F:\NITISH\2.jpg")
I = 0
drinkcmb.AddItem "Coke"
drinkcmb.AddItem "Sprite"
drinkcmb.AddItem "Mirinda"
drinkcmb.AddItem "7up"
drinkcmb.AddItem "Mountain Dew"
Dim j As Integer
drink = 50

CHIKENCMB.AddItem "CHICKEN MAKHNI BURGER"
CHIKENCMB.AddItem "CHICKEN CRISPY"
CHIKENCMB.AddItem "CHICKEN CHILLI"
CHIKENCMB.AddItem "CHICKEN TIKKA"
CHIKENCMB.AddItem "CHICKEN CRUNCHY"
CHIKENCMB.AddItem "CHICKEN CHEESE CRUNCHY"
CHIKENCMB.AddItem "CHICKEN SHOT"
CHIKENCMB.AddItem "CHICKEN SURPRISE"
CHIKENCMB.AddItem "CHICKEN POPCORN"
CHIKENCMB.AddItem "CHICKEN SORT"

VEGCMB.AddItem "ALOO TIKKI"
VEGCMB.AddItem "DESI STREET"
VEGCMB.AddItem "PANEER CRISPY"
VEGCMB.AddItem "VEG SURPRISE"
VEGCMB.AddItem "ALOO 'N' CHEESE"
VEGCMB.AddItem "VEGGIE CRISPY"
VEGCMB.AddItem "ALOO BECHARA"
VEGCMB.AddItem "HOT VEG LAVA"
VEGCMB.AddItem "VEG JALFREZI"
VEGCMB.AddItem "PANEER CHIILI LAVA"


name1 = cname.Text
last = clname.Text
EMAIL = cemail.Text
no = cno.Text

CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=F:\NITISH\CUST.mdb"
CN.Open
RS.Open "customer", CN, adOpenDynamic, adLockOptimistic
End Sub

Private Sub meal_Click()
If meal.Value = 1 Then
meal1 = 200
Else
meal1 = 0
End If
End Sub

Private Sub minus_Click()
I = I - 1
drinktxt.Text = I
End Sub

Private Sub mul_Click()
Text1.Text = ""
preval = curval
curval = 0
choice = "*"

End Sub

Private Sub plus_Click()
I = I + 1
drinktxt.Text = I

End Sub

Private Sub recipt_Click()
If cname.Text = "" Or cemail.Text = "" Or clname.Text = "" Or cno.Text = "" Then
MsgBox "COMPLETE CUSTOMER DETAILS", vbDefaultButton1
ElseIf Not (IsNumeric(cno.Text)) Then
MsgBox "CUSTOMER PHONE NO IS WRONG", vbApplicationModal
Else
RS.Fields(0) = cname.Text
RS.Fields(1) = clname.Text
RS.Fields(2) = cemail.Text
RS.Fields(3) = cno.Text
RS.Fields(6) = totxt.Text
RS.Fields(4) = sttxt.Text
RS.Fields(5) = ttxt.Text
RS.Fields(7) = chickentxt.Text
RS.Fields(8) = vegtxt.Text
RS.Fields(9) = mealtxt.Text
RS.Fields(10) = drinktxt.Text

RS.AddNew
DataReport1.Show
End If
End Sub

Private Sub reset_Click()
chicken.Value = 0
veg.Value = 0
meal.Value = 0

cname.Text = ""
clname.Text = ""
cemail.Text = ""
cno.Text = ""

chickentxt.Text = ""
vegtxt.Text = ""
mealtxt.Text = ""
drinkcmb.Text = ""
drinktxt.Text = ""

dsoldtxt.Text = ""
bsoldtxt.Text = ""
msoldtxt.Text = ""

sttxt.Text = ""
ttxt.Text = ""
totxt.Text = ""

End Sub

Private Sub subt_Click()
Text1.Text = ""
preval = curval
curval = 0
choice = "-"

End Sub

Private Sub veg_Click()
If veg.Value = 1 Then
vegburger = 80
VEGCMB.Visible = True

Else
vegburger = 0
VEGCMB.Visible = False
End If
End Sub

Private Sub VEGCMB_Click()
VEGCMB.Visible = False
End Sub
