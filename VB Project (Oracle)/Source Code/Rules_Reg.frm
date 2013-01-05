VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Rules_Reg 
   Caption         =   "Fee Receipt"
   ClientHeight    =   9420
   ClientLeft      =   5910
   ClientTop       =   1365
   ClientWidth     =   7275
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "Rules_Reg.frx":0000
   ScaleHeight     =   628
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   Begin VB.CommandButton OKButton 
      Caption         =   "Proceed >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CheckBox Agree 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      MaskColor       =   &H000040C0&
      Picture         =   "Rules_Reg.frx":12140
      TabIndex        =   1
      Top             =   8790
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Receipt"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   480
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "I Agree to the above rules & regulations."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   8745
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   $"Rules_Reg.frx":8230CC
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   720
      TabIndex        =   20
      Top             =   7560
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Please supervise your children at all times"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   7200
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   $"Rules_Reg.frx":823163
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   720
      TabIndex        =   18
      Top             =   6240
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   $"Rules_Reg.frx":8231F8
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   720
      TabIndex        =   17
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   $"Rules_Reg.frx":8232A8
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   720
      TabIndex        =   16
      Top             =   4320
      Width           =   5895
   End
   Begin VB.Label VisitorID 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   5595
      TabIndex        =   15
      Top             =   2505
      Width           =   975
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   495
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label AdvCurr 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance/Current Booking"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label DateBox 
      BackStyle       =   0  'Transparent
      Caption         =   "dd-mm-yyyy"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Fee:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ent. Tax:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Name_Last 
      BackStyle       =   0  'Transparent
      Caption         =   " Last Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Name_First 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label EntryFee 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs (    )/-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label EntTax 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs (    )/-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "Rules_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection
Dim connLog  As ADODB.Connection
Dim rsLog  As ADODB.Recordset
Dim sqlquery As String
Dim adv_curr As Integer
Option Explicit

Private Sub Command1_Click()
    'On Error Resume Next
    'If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
'        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .Flags = .Flags + cdlPDAllPages
'        Else
'            .Flags = .Flags + cdlPDSelection
'        End If
        .ShowPrinter
'        If Err <> MSComDlg.cdlCancel Then
'            ActiveForm.rtfText.SelPrint .hDC
'        End If
    End With

End Sub

Private Sub Form_Load()
VisitorID.Caption = Global_Module.visitor_id


Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "select * from visitor where id=" & VisitorID.Caption
rsVisitor.Open sqlquery, connVisitor
    Set Name_First.DataSource = rsVisitor.DataSource
    Name_First.DataField = "Name_First"
    Set Name_Last.DataSource = rsVisitor.DataSource
    Name_Last.DataField = "Name_Last"
    Set EntryFee.DataSource = rsVisitor.DataSource
    EntryFee.DataField = "Entry_Fee"

DateBox.Caption = Format(rsVisitor.Fields("Entry_Date").Value, "dd-mmm-yyyy")
rsVisitor.Close
connVisitor.Close
EntTax.Caption = Format((EntryFee.Caption / (1 + (Global_Module.Ent_Tax / 100))) * (Global_Module.Ent_Tax / 100), "00.00")
adv_curr = DateDiff("d", Format(Global_Module.Today, "dd/mm/yyyy"), Format(DateBox.Caption, "dd/mm/yyyy"))
If (adv_curr = 0) Then
    adv_curr = DateDiff("m", Format(Global_Module.Today, "dd/mm/yyyy"), Format(DateBox.Caption, "dd/mm/yyyy"))
End If
If (adv_curr = 0) Then
    adv_curr = DateDiff("y", Format(Global_Module.Today, "dd/mm/yyyy"), Format(DateBox.Caption, "dd/mm/yyyy"))
End If
If (adv_curr = 0) Then
    AdvCurr.Caption = "Current Booking"
Else
    AdvCurr.Caption = "Advance Booking"
End If

End Sub

Private Sub OKButton_Click()
ReDim varfields(5), varvalues(5)

If (Agree.Value = 0) Then
MsgBox "Please tick on 'I Agree' before proceeding!"
Else
    Unload Me
    Set connLog = New ADODB.Connection
    Set rsLog = New ADODB.Recordset
    connLog.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
    rsLog.LockType = adLockOptimistic

    If (adv_curr = 0) Then
        Map.Show
        varfields = Array("V_ID", "ACTION_NAME", "ACTION_TYPE", "FEES_PAID", "ACTION_DATE")
        varvalues = Array(VisitorID.Caption, "Current Booking", "Entry Fee", EntryFee.Caption, Format(DateBox, "DD-MMM-YYYY"))
        rsLog.Open "select * from log_visitor", connLog
        rsLog.AddNew varfields, varvalues
        rsLog.Close
        Unload Me
        'OracleDB.logCurrentBooking Val(VisitorID.Caption), Val(EntryFee.Caption), Format(DateBox.Caption, "dd-mmm-yyyy")
    Else
        MsgBox "Your advance booking has been confirmed. Please bring this receipt at the time of arrival!", vbOKOnly, "Entertainment Resort"
        varfields = Array("V_ID", "ACTION_NAME", "ACTION_TYPE", "FEES_PAID", "ACTION_DATE")
        varvalues = Array(VisitorID.Caption, "Advance Booking", "Entry Fee", EntryFee.Caption, Format(DateBox, "DD-MMM-YYYY"))
        rsLog.Open "select * from log_visitor", connLog
        rsLog.AddNew varfields, varvalues
        rsLog.Close
        Welcome.Show
        Unload Me
    End If
    connLog.Close
End If
End Sub

