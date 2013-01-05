VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Admin_ChangeDate 
   Caption         =   "Custom Date"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ChangeDate.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   5985
   Begin VB.CommandButton SetToday 
      Caption         =   "Today's Date"
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
      Left            =   840
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton UpdateDate 
      Caption         =   "Update Date"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   68288515
      CurrentDate     =   40454
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the custom date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(For demonstration purposes)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "Admin_ChangeDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTPicker1.Value = Format(Global_Module.Today, "DD-MMM-YYYY")
With Me
.Top = 100
.Left = 100
.Height = 4365
.Width = 6100
End With
End Sub

Private Sub SetToday_Click()
DTPicker1.Value = Format(Now(), "DD-MMM-YYYY")
End Sub

Private Sub UpdateDate_Click()
Global_Module.Today = DTPicker1.Value
Welcome.date.Caption = Format(Global_Module.Today, "dd-mmm-yyyy")
MsgBox "The current date has been updated!", vbOKOnly, "Admin Landing"
Unload Me
End Sub
