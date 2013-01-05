VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Staff_SalarySlip 
   Caption         =   "Salary Slip"
   ClientHeight    =   5295
   ClientLeft      =   6465
   ClientTop       =   3000
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   Picture         =   "Staff_SalarySlip.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   6225
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   600
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Proceed 
      Caption         =   "Proceed to Job >>"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5640
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label DA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label HRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HRA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Basic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Basic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(Gross)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(DA)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(HRA)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(Basic)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Role:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Salary 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Role 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label NameLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   3495
   End
End
Attribute VB_Name = "Staff_SalarySlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub Command1_Click()
 With dlgCommonDialog
        .DialogTitle = "Print"
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .ShowPrinter
    End With
End Sub

Private Sub Form_Load()
If OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Open
Else
    OracleDB.OracleProvider.Close
    OracleDB.OracleProvider.Open
End If

OracleDB.getStaffSalary Global_Module.staff_login

NameLabel.Caption = OracleDB.rsgetStaffSalary.Fields("Name_First").Value & " " & OracleDB.rsgetStaffSalary.Fields("Name_Last").Value
Role.Caption = OracleDB.rsgetStaffSalary.Fields("Role")
Salary.Caption = OracleDB.rsgetStaffSalary.Fields("Salary")

x = Val(Salary.Caption) / 1.65
Basic.Caption = x
HRA.Caption = x * 0.05
DA.Caption = x * 0.6

End Sub

Private Sub LastName_Click()

End Sub

Private Sub Proceed_Click()
OracleDB.rsgetStaffSalary.Update "Last_Claim", Format(Global_Module.Today, "dd-mmm-yyyy")

If (StrComp(OracleDB.rsgetStaffSalary.Fields("Role"), "Assistant") = 0) Then
'Check for queued up requests
            Set connVisitor = New ADODB.Connection
            Set rsVisitor = New ADODB.Recordset
            connVisitor.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
            sqlquery = "select count(*) as num from visitor where need_help=1"
            rsVisitor.Open sqlquery, connVisitor
            
            cnt = rsVisitor.Fields("num").Value
            
            rsVisitor.Close
            connVisitor.Close
            
            If cnt = 0 Then
                MsgBox "There are no visitors requesting for assistance at the moment!", vbOKOnly, "Entertainment Resort"
                Unload Me
                Welcome.Show
            Else
                Unload Me
                Staff_Assistant_landing.Show
            End If
ElseIf (StrComp(OracleDB.rsgetStaffSalary.Fields("Role"), "Mechanic") = 0) Then
        Unload Me
        Staff_Mechanic_Landing.Show
    End If
End Sub
