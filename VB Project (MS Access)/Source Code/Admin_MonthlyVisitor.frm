VERSION 5.00
Begin VB.Form Admin_MonthlyVisitor 
   Caption         =   "Monthly Report - Visitor"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_MonthlyVisitor.frx":0000
   ScaleHeight     =   4530
   ScaleWidth      =   5955
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
      Left            =   4320
      TabIndex        =   3
      Text            =   "2010"
      Top             =   2040
      Width           =   615
   End
   Begin VB.ComboBox MonthLIst 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Admin_MonthlyVisitor.frx":4D81
      Left            =   3360
      List            =   "Admin_MonthlyVisitor.frx":4DA9
      TabIndex        =   1
      Text            =   "(month)"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Proceed 
      Caption         =   "Generate Report >>"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select a month and year for the report:"
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
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "Admin_MonthlyVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndDate As Date
Dim test As Integer
Dim StartDate As Date

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 4900
.Width = 5950
End With

MsgBox "Sample data has been entered into the database for the month of NOVEMBER 2010.", vbInformation, "Demonstration"

End Sub

Private Sub Proceed_Click()
'Generate StartDate & EndDate for report
StartDate = DateSerial(YearBox.Text, MonthLIst.ListIndex + 1, 1)
If ((MonthLIst.ListIndex + 1) Mod 2) = 1 Or (MonthLIst.ListIndex + 1) = 7 Then
    EndDate = DateSerial(YearBox.Text, MonthLIst.ListIndex + 1, 31)
Else
    EndDate = DateSerial(YearBox.Text, MonthLIst.ListIndex + 1, 30)
End If

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
OracleDB.MonthlyVisitor Format(StartDate, "dd-mmm-yyyy"), Format(EndDate, "dd-mmm-yyyy")
Report_MonthlyVisitor.Orientation = rptOrientLandscape
Report_MonthlyVisitor.Show
Unload Me
End Sub
