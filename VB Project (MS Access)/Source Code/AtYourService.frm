VERSION 5.00
Begin VB.Form AtYourService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help Request"
   ClientHeight    =   3840
   ClientLeft      =   7710
   ClientTop       =   3705
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AtYourService.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your request has been submitted alongwith your contact information. One of our staff members will assist you shortly!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1575
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "AtYourService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
Welcome.Show
End Sub

