VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Admin_DailyRides 
   Caption         =   "Daily Report - Rides"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_DailyRides.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   5910
   Begin VB.CommandButton Proceed 
      Caption         =   "Generate Report >>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16449539
      CurrentDate     =   40453
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the date for report:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
End
Attribute VB_Name = "Admin_DailyRides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 4900
.Width = 5950
End With

DTPicker1.Value = Format(Global_Module.Today, "DD-MMM-YYYY")
End Sub

Private Sub Proceed_Click()
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.DailyRides (Format(DTPicker1.Value, "DD-MMM-YYYY"))
Report_DailyRides.Orientation = rptOrientLandscape
Report_DailyRides.Show
Unload Me
End Sub

