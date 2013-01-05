VERSION 5.00
Begin VB.Form Admin_ChangeEntryFees 
   Caption         =   "Change Entry Fees"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ChangeEntryFees.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   5775
   Begin VB.CommandButton UpdateFees 
      Caption         =   "Update Entry Fees"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Entry_Kids 
      DataField       =   "VALUE"
      DataMember      =   "GetEntTax"
      DataSource      =   "OracleDB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Entry_Adults 
      DataField       =   "VALUE"
      DataMember      =   "GetSerTax"
      DataSource      =   "OracleDB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry fee for Kids:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry fee for Adults:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "Admin_ChangeEntryFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EntTax_Change()

End Sub

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 4200
.Width = 5900
End With

Entry_Kids.Text = Global_Module.Entry_Kids
Entry_Adults.Text = Global_Module.Entry_Adults

End Sub

Private Sub UpdateFees_Click()
OracleDB.GetEntryFee "Entry_Kids"
OracleDB.rsGetEntryFee.Update "Value", Val(Entry_Kids.Text)
OracleDB.rsGetEntryFee.Close
OracleDB.GetEntryFee "Entry_Adults"
OracleDB.rsGetEntryFee.Update "Value", Val(Entry_Adults.Text)

Global_Module.Entry_Kids = Val(Entry_Kids.Text)
Global_Module.Entry_Adults = Val(Entry_Adults.Text)

MsgBox "The entry fees have been updated!", vbOKOnly, "Admin Landing"
Unload Me
End Sub
