VERSION 5.00
Begin VB.Form Map 
   Caption         =   "Form3"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form3"
   ScaleHeight     =   7875
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Click here if you want assistants"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   6840
      Width           =   5055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Shop"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Water"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Roller Coaster"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Thrill"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gentle"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transport"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "map"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
form4.Show
End Sub

Private Sub Command2_Click()
Unload Me
form6.Show
End Sub

Private Sub Command3_Click()
Unload Me
form5.Show
End Sub

Private Sub Command4_Click()
Unload Me
form8.Show
End Sub

Private Sub Command5_Click()
Unload Me
form7.Show
End Sub

Private Sub Command7_Click()
Unload Me
form8.Show
End Sub
