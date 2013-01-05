VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Staff_Assistant_landing 
   Caption         =   "Staff Landing - Assistant"
   ClientHeight    =   4080
   ClientLeft      =   6765
   ClientTop       =   4140
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   Picture         =   "staff_assistant_landing.frx":0000
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   Begin VB.CommandButton Proceed 
      Caption         =   "Proceed to HELP!"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc adodc_visitor 
      Height          =   735
      Left            =   3840
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=./db.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=./db.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from visitor where need_help=1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Contact 
      BackStyle       =   0  'Transparent
      DataField       =   "CONTACT"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Name_Last 
      BackStyle       =   0  'Transparent
      DataField       =   "NAME_LAST"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Name_First 
      BackStyle       =   0  'Transparent
      DataField       =   "NAME_FIRST"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Help_Type 
      BackStyle       =   0  'Transparent
      DataField       =   "HELP_TYPE"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label VisitorID 
      BackStyle       =   0  'Transparent
      DataField       =   "ID"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The following visitor needs help:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Visitor ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type of help:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Staff_Assistant_landing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection
Dim rsStaff As ADODB.Recordset
Dim connStaff As ADODB.Connection
Dim cnt As Integer

Private Sub Form_Load()
'Get no. of visitors requesting for help from SQL DB
'Set connVisitor = New ADODB.Connection
'Set rsVisitor = New ADODB.Recordset
'connVisitor.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
'sqlquery = "select count(*) as num from visitor where need_help=1"
'rsVisitor.Open sqlquery, connVisitor
'
'cnt = rsVisitor.Fields("num").Value
'
'rsVisitor.Close
'connVisitor.Close
'
'If cnt = 0 Then
'    cnt = 1
'    MsgBox "There are no visitors requesting for assistance at the moment!", vbOKOnly, "Entertainment Resort"
'    'Unload Me
'    Me.Hide
'    Welcome.Show
'End If

'Populate fields with first visitor requesting for help

End Sub

Private Sub Form_Unload(Cancel As Integer)
Welcome.Show
End Sub

Private Sub Proceed_Click()
'Reset the need_help flag from visitor entry
Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "update (select * from visitor where id=" & VisitorID.Caption & ") set need_help=0,help_type=null"
connVisitor.Execute sqlquery
connVisitor.Close

'Insert entry in log_staff_assistant DB
Set connStaff = New ADODB.Connection
Set rsStaff = New ADODB.Recordset
connStaff.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
rsStaff.LockType = adLockOptimistic
rsStaff.Open "select * from log_staff_assistant", connStaff
varfields = Array("staff_login", "v_id", "help_type", "assistance_date")
varvalues = Array(Global_Module.staff_login, VisitorID.Caption, Help_Type.Caption, Format(Global_Module.Today, "DD-MMM-YYYY"))
rsStaff.AddNew varfields, varvalues
rsStaff.Close
connStaff.Close


MsgBox "The request status of the visitor has been reset!", vbOKOnly, "Entertainment Resort"
Unload Me
Welcome.Show
End Sub
