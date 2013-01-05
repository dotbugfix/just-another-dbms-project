VERSION 5.00
Begin VB.Form Welcome 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1425
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   6030
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SUBMIT"
      Default         =   -1  'True
      Height          =   390
      Left            =   1680
      TabIndex        =   11
      Top             =   6525
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1425
      TabIndex        =   10
      Top             =   5640
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1425
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   4110
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SUBMIT"
      Height          =   390
      Left            =   1560
      TabIndex        =   6
      Top             =   4605
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1425
      TabIndex        =   5
      Top             =   3720
      Width           =   2325
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXISTING VISITOR"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEW VISITOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   7320
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   1200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   6045
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   5655
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   4125
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   3735
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "STAFF"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
Form2.Show
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Command3_Click()
Unload Me
form9.Show
End Sub

Private Sub Form_Load()

End Sub
