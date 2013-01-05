VERSION 5.00
Begin VB.Form Admin_ChangeEntTax 
   Caption         =   "Change Taxes"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ChangeEntTax.frx":0000
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Begin VB.TextBox IncomeTax 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox SerTax 
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox EntTax 
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton UpdateEntTax 
      Caption         =   "Update Taxes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Income Tax:"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax:"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entertainment Tax:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Admin_ChangeEntTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 4200
.Width = 5900
End With

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.GetEntTax
OracleDB.GetSerTax
OracleDB.GetIncomeTax

EntTax.Text = Global_Module.Ent_Tax
SerTax.Text = Global_Module.Ser_Tax
IncomeTax.Text = Global_Module.Income_Tax

MsgBox "These values of taxes are used to dynamically generate financial reports and will also affect previously present entries in the database.", vbInformation, "Demonstration"

End Sub

Private Sub UpdateEntTax_Click()
OracleDB.rsGetEntTax.Update "value", EntTax.Text
OracleDB.rsGetSerTax.Update "value", SerTax.Text
OracleDB.rsGetIncomeTax.Update "value", IncomeTax.Text
Global_Module.Ent_Tax = Val(EntTax.Text)
Global_Module.Ser_Tax = Val(SerTax.Text)
Global_Module.Income_Tax = Val(IncomeTax.Text)
MsgBox "The prescribed values of taxes have been updated!", vbOKOnly, "Admin Landing"
Unload Me
End Sub
