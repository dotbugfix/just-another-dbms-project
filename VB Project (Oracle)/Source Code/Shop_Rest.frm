VERSION 5.00
Begin VB.Form Shop_Rest 
   Caption         =   "Shops"
   ClientHeight    =   6420
   ClientLeft      =   4590
   ClientTop       =   2925
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9885
   Begin VB.Frame Frame2 
      Caption         =   "Snacks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   2295
      Begin VB.CheckBox Item 
         Caption         =   "Ice cream (30/-)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Item 
         Caption         =   "Chocolates (30/-)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Item 
         Caption         =   "Cotton Candy (20/-)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Item 
         Caption         =   "Popcorn (25/-)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CheckBox Item 
      Caption         =   "Puppy (150/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7320
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox Item 
      Caption         =   "Cars (300/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7320
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox Item 
      Caption         =   "Barbie doll (250/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7320
      TabIndex        =   15
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Item 
      Caption         =   "Teddy bear (200/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Item 
      Caption         =   "Yo-yo (20/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CheckBox Item 
      Caption         =   "Hats (25/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CheckBox Item 
      Caption         =   "Goggles (80/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox Item 
      Caption         =   "Balloons (5/-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   10
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2640
      Picture         =   "Shop_Rest.frx":0000
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   7
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Proceed to check out >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   6
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Bill 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Billing amount >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Toys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7200
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Visitor ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
      Begin VB.Label v_id 
         Alignment       =   2  'Center
         Caption         =   "(id)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the shopping arena!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "Shop_Rest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Bill.Text = Val(Item(0))
End Sub

Private Sub Command2_Click()
Unload Me
Map.Show
End Sub

Private Sub Command3_Click()
Map.Show
End Sub

Private Sub Form_Load()
v_id.Caption = Global_Module.visitor_id
End Sub

