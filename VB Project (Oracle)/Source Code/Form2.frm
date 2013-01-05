VERSION 5.00
Begin VB.Form MainCounter 
   Caption         =   "Form2"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   5160
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "YES"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "YES"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Nausiea"
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Heart Problems"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Blood Pressure"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   3000
      List            =   "Form2.frx":000D
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Do you have camera?? Rupees 100/-"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Do you want locker??? Rupees:50/-"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Health Problems:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Contact No:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Your ID  is:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "NAME:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Select the age group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   " Total Entry Fee :"
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
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "MAIN COUNTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "MainCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Click()
If Combo1.List(0) = "For Kids (Age 1-5)" Then
Text1.Text = "50"


End If

If Combo1.List(1) = "Adults (Above 60)" Then
Text1.Text = "70"

End If

If Combo1.List(2) = "Others" Then
Text1.Text = "100"

End If
End Sub

Private Sub Command1_Click()
Unload Me
Dialog.Show
End Sub
