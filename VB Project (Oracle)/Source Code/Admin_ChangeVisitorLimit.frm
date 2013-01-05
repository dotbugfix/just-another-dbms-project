VERSION 5.00
Begin VB.Form Admin_ChangeVisitorLimit 
   Caption         =   "Change Visitor Limit"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ChangeVisitorLimit.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   6000
   Begin VB.CommandButton UpdateLimit 
      Caption         =   "Update Visitor Limit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox VisitorLimit 
      DataField       =   "VALUE"
      DataMember      =   "GetVisitorLimit"
      DataSource      =   "OracleDB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the per-day visitor limit:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "Admin_ChangeVisitorLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
VisitorLimit.Text = Global_Module.visitor_limit
With Me
.Top = 100
.Left = 100
.Height = 4700
.Width = 6100
End With
End Sub

Private Sub UpdateLimit_Click()
OracleDB.rsGetVisitorLimit.Update "value", VisitorLimit.Text
Global_Module.visitor_limit = Val(VisitorLimit.Text)
MsgBox "The per-day visitor limit in the resort has been updated!", vbOKOnly, "Admin Landing"
Unload Me
End Sub
