VERSION 5.00
Begin VB.Form Admin_AnnualFinanceLanding 
   Caption         =   "Financial Report"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_AnnualFinanceLanding.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   5940
   Begin VB.TextBox YearBox 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Text            =   "2010"
      Top             =   1680
      Width           =   615
   End
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
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select the starting financial year:"
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
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "Admin_AnnualFinanceLanding"
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

End Sub

Private Sub Proceed_Click()
'Generate StartDate & EndDate for report
StartDate = DateSerial(Val(YearBox.Text), 4, 1)
EndDate = DateSerial((Val(YearBox.Text) + 1), 3, 30)

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    

    
OracleDB.MonthlyEntryFee Format(StartDate, "dd-mmm-yyyy"), Format(EndDate, "dd-mmm-yyyy")
OracleDB.MonthlyRideRev Format(StartDate, "dd-mmm-yyyy"), Format(EndDate, "dd-mmm-yyyy")
OracleDB.MonthlyShopRevExp Format(StartDate, "dd-mmm-yyyy"), Format(EndDate, "dd-mmm-yyyy")
OracleDB.MonthlyRestRevExp Format(StartDate, "dd-mmm-yyyy"), Format(EndDate, "dd-mmm-yyyy")
With Admin_Landing
   .WindowState = vbMaximized
End With

Admin_AnnualFinanceReport.Show
Unload Me
End Sub



