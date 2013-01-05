VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Ride_Thrill 
   Caption         =   "Thrill Rides"
   ClientHeight    =   7140
   ClientLeft      =   4770
   ClientTop       =   2460
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   Picture         =   "Ride_Thrill.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   9765
   Begin VB.Frame Frame3 
      Caption         =   "Visitor_stats"
      Height          =   1815
      Left            =   2760
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox Nausea_visitor 
         Caption         =   "Nausea"
         DataField       =   "NAUSEA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Heart_visitor 
         Caption         =   "Heart Problems"
         DataField       =   "HEART"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox bp_visitor 
         Caption         =   "Blood Pressure"
         DataField       =   "BP"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ride_stats"
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox bp_ride 
         Caption         =   "Blood Pressure"
         DataField       =   "BP"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox heart_ride 
         Caption         =   "Heart Problems"
         DataField       =   "HEART"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox nausea_ride 
         Caption         =   "Nausea"
         DataField       =   "NAUSEA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adodc_visitor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.CommandButton BackCommand 
      Caption         =   "<<Back"
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
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Proceed 
      Caption         =   "Proceed to ride >>"
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
      Left            =   6240
      TabIndex        =   0
      Top             =   6480
      Width           =   3255
   End
   Begin MSDataListLib.DataCombo RideList 
      Height          =   360
      Left            =   6960
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      Text            =   "(select)"
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "Rupee"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   7440
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "You need to pay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Fee 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   7920
      TabIndex        =   14
      Top             =   3765
      Width           =   1095
   End
   Begin VB.Label VisitorID 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   8400
      TabIndex        =   13
      Top             =   1590
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Visitor ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a ride:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Ride_Thrill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection
Dim rsRide As ADODB.Recordset
Dim connRide As ADODB.Connection
Dim rsLog As ADODB.Recordset
Dim connLog As ADODB.Connection
Dim sqlquery As String
Dim RideName As String
Dim Age As Integer

Private Sub BackCommand_Click()
Unload Me
Map.Show
End Sub
Private Sub Form_Load()
VisitorID.Caption = Global_Module.visitor_id

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.RideList_Thrill

Set RideList.RowSource = OracleDB
RideList.RowMember = "RideList_Thrill"
RideList.ListField = "NAME"

End Sub

Private Sub Proceed_Click()
If (StrComp(Fee.Caption, "") = 0) Then
    MsgBox "Please select a ride before proceeding!", vbOKOnly, "Entertainment Resort"
Else
    'Insert into log_ride table
    Set connLog = New ADODB.Connection
    Set rsLog = New ADODB.Recordset
    connLog.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
    rsLog.LockType = adLockOptimistic
    varfields = Array("V_ID", "ACTION_NAME", "ACTION_TYPE", "FEES_PAID", "ACTION_DATE")
    varvalues = Array(VisitorID.Caption, RideList.Text, "thrill", Fee.Caption, Format(Global_Module.Today, "DD-MMM-YYYY"))
    rsLog.Open "select * from log_visitor", connLog
    rsLog.AddNew varfields, varvalues
    rsLog.Close
    MsgBox "Enjoy the ride!", vbOKOnly, "Entertainment Resort"
    Unload Me
    Welcome.Show
End If
End Sub

Private Sub RideList_Change()
If Not RideList.Text = "(select)" Then
'Get visitor's age from SQL DB
Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "select * from visitor where id=" & VisitorID.Caption
rsVisitor.Open sqlquery, connVisitor

Age = rsVisitor.Fields("age").Value

'Get visitor's health problems from SQL DB
Set Heart_visitor.DataSource = rsVisitor.DataSource
Heart_visitor.DataField = "heart"
Set bp_visitor.DataSource = rsVisitor.DataSource
bp_visitor.DataField = "bp"
Set Nausea_visitor.DataSource = rsVisitor.DataSource
Nausea_visitor.DataField = "nausea"

rsVisitor.Close
connVisitor.Close

'Get Ride Fees as per age and Health Problem Stats from SQL SB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "select * from rides where type='thrill' and name='" & RideList.Text & "'"
rsRide.Open sqlquery, connRide

If Age < 15 Then
    Fee.Caption = rsRide.Fields("cost_kids").Value
Else
    Fee.Caption = rsRide.Fields("cost_adults").Value
End If

If Val(Fee.Caption) = 0 Then
    MsgBox "Sorry, visitors of your age are not allowed for this ride! Please try another ride!", vbOKOnly, "Entertainment Resort"
    RideList.Text = "(select)"
    Fee.Caption = ""
Else

Set heart_ride.DataSource = rsRide.DataSource
heart_ride.DataField = "heart"
Set bp_ride.DataSource = rsRide.DataSource
bp_ride.DataField = "bp"
Set nausea_ride.DataSource = rsRide.DataSource
nausea_ride.DataField = "nausea"

rsRide.Close
connRide.Close

If (Heart_visitor.Value = 1 And heart_ride.Value = 1) Then
    MsgBox "Sorry, you cannot enter this ride since it is not safe for patients with heart problems!", vbOKOnly, "Entertainment Resort"
    RideList.Text = "(select)"
    Fee.Caption = ""
ElseIf (bp_visitor.Value = 1 And bp_ride.Value = 1) Then
    MsgBox "Sorry, you cannot enter this ride since it is not safe for patients with blood pressure problems!", vbOKOnly, "Entertainment Resort"
    RideList.Text = "(select)"
    Fee.Caption = ""
ElseIf (Nausea_visitor.Value = 1 And nausea_ride.Value = 1) Then
    MsgBox "Sorry, you cannot enter this ride since it is not safe for patients with nausea-related problems!", vbOKOnly, "Entertainment Resort"
    RideList.Text = "(select)"
    Fee.Caption = ""
End If
End If
End If
End Sub



