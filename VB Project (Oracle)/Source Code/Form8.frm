VERSION 5.00
Begin VB.Form staff_help 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Which kind of help do you need??"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "staff_help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
