VERSION 5.00
Begin VB.Form Map 
   Caption         =   "Park Map"
   ClientHeight    =   8655
   ClientLeft      =   5670
   ClientTop       =   1785
   ClientWidth     =   7875
   LinkTopic       =   "Form3"
   Picture         =   "map.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   7875
   Begin VB.CommandButton Back 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Restaurants"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Request for Assistance"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   7800
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Shops "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Water Rides"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   4
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Roller Coasters"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Thrill Rides"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gentle Rides"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transport Rides"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Park Map"
      BeginProperty Font 
         Name            =   "DecoTech"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
Unload Me
Welcome.Show
End Sub

Private Sub Command1_Click()
Unload Me
Ride_Transport.Show
End Sub

Private Sub Command2_Click()
Unload Me
Ride_Gentle.Show
End Sub

Private Sub Command3_Click()
Unload Me
Ride_Thrill.Show
End Sub

Private Sub Command4_Click()
Unload Me
Ride_Coaster.Show
End Sub

Private Sub Command5_Click()
Unload Me
Ride_Water.Show
End Sub

Private Sub Command6_Click()
Unload Me
Shop.Show
End Sub

Private Sub Command7_Click()
Unload Me
Help_request.Show
End Sub

Private Sub Command8_Click()
restaurant.Show
Unload Me
End Sub
