VERSION 5.00
Begin VB.Form Rules_Reg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   6645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "I agree to the above rules & regulations"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   $"Dialog.frx":0000
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   6135
   End
   Begin VB.Label Label5 
      Caption         =   "Please supervise your children at all times"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   $"Dialog.frx":0096
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   6255
   End
   Begin VB.Label Label3 
      Caption         =   $"Dialog.frx":012B
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   $"Dialog.frx":01DB
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Rules And Regulations:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Rules_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
form3.Show
End Sub
