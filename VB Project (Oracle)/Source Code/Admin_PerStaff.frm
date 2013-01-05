VERSION 5.00
Begin VB.Form Admin_PerStaff 
   Caption         =   "Per Staff History"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_PerStaff.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   7695
   Begin VB.TextBox NameLast 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox NameFirst 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox StaffRole 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Admin_PerStaff.frx":5C44
      Left            =   3840
      List            =   "Admin_PerStaff.frx":5C4E
      TabIndex        =   3
      Text            =   "(select)"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox StaffList 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Admin_PerStaff.frx":5C67
      Left            =   3840
      List            =   "Admin_PerStaff.frx":5C69
      TabIndex        =   2
      Text            =   "(select)"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox TempBox 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Proceed 
      Caption         =   "Generate Report >>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the role of staff:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the staff member:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
End
Attribute VB_Name = "Admin_PerStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStaff As ADODB.Recordset
Dim connStaff As ADODB.Connection
Dim TableSelect As String
Dim sqlquery As String
Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 5700
.Width = 7750
End With

End Sub

Private Sub Proceed_Click()
If StaffList.Text = "(select)" Or StaffList.Text = "" Then
    MsgBox "Please select a staff member before generating the report!", vbOKOnly, "Admin Landing"
Else
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
sqlquery = "select * from staff where username='" & StaffList.Text & "'"
rsStaff.Open sqlquery, connStaff
NameFirst.Text = rsStaff.Fields("NAME_FIRST").Value
NameLast.Text = rsStaff.Fields("NAME_LAST").Value
rsStaff.Close
If StaffRole.Text = "Assistant" Then
    OracleDB.PerStaffAssistant (StaffList.Text)
    Report_PerStaffAssistant.Orientation = rptOrientLandscape
    Report_PerStaffAssistant.Show
Else
    OracleDB.PerStaffMechanic (StaffList.Text)
    Report_PerStaffMechanic.Orientation = rptOrientLandscape
    Report_PerStaffMechanic.Show
End If
connStaff.Close
Unload Me
End If


End Sub

Private Sub StaffList_GotFocus()
If StaffRole.Text = "(select)" Then
    MsgBox "Please select the role of staff first!", vbOKOnly, "Admin Landing"
Else

'Populate list of Staff_Usernames from SQL DB
Set connStaff = New ADODB.Connection
Set rsStaff = New ADODB.Recordset
connStaff.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "select * from staff where role='" & StaffRole.Text & "'"
rsStaff.Open sqlquery, connStaff
Set TempBox.DataSource = rsStaff.DataSource
TempBox.DataField = "username"
rsStaff.MoveFirst
StaffList.Clear
While (rsStaff.EOF = False)
    StaffList.AddItem (TempBox.Text)
    rsStaff.MoveNext
Wend
rsStaff.Close
End If

End Sub


Private Sub StaffRole_GotFocus()
StaffList.Clear
StaffList.Text = "(select)"
End Sub
