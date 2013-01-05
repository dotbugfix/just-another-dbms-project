VERSION 5.00
Begin VB.Form Visitor_Existing 
   Caption         =   "Form9"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form9"
   ScaleHeight     =   3120
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Text            =   "ID"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter your Visitor ID:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "WELCOME BACK!!!"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "Visitor_Existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
form3.Show
End Sub
