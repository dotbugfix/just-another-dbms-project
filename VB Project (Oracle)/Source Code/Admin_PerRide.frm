VERSION 5.00
Begin VB.Form Admin_PerRide 
   Caption         =   "Per Ride History"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_PerRide.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   7680
   Begin VB.ComboBox RideList 
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
      ItemData        =   "Admin_PerRide.frx":5BCA
      Left            =   3360
      List            =   "Admin_PerRide.frx":5BCC
      TabIndex        =   3
      Text            =   "(select)"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox RideType 
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
      ItemData        =   "Admin_PerRide.frx":5BCE
      Left            =   3360
      List            =   "Admin_PerRide.frx":5BE1
      TabIndex        =   2
      Text            =   "(select)"
      Top             =   1440
      Width           =   2175
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
   Begin VB.TextBox TempBox 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the ride:"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a type of ride:"
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
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "Admin_PerRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRide As ADODB.Recordset
Dim connRide As ADODB.Connection
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
If RideList.Text = "" Or RideList.Text = "(select)" Then
    MsgBox "Please select a ride before generating the report!", vbOKOnly, "Admin Landing"
Else

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
OracleDB.PerRide (RideList.Text)
Report_PerRide.Orientation = rptOrientLandscape
Report_PerRide.Show
Unload Me
End If


End Sub

Private Sub RideList_GotFocus()
If RideType.Text = "(select)" Then
    MsgBox "Please select the type of ride first!", vbOKOnly, "Admin Landing"
Else

Select Case RideType.Text
    Case "Transport"
        TableSelect = "transport"
    Case "Gentle"
        TableSelect = "gentle"
    Case "Thrill"
        TableSelect = "thrill"
    Case "Water"
        TableSelect = "water"
    Case "Roller Coaster"
        TableSelect = "coaster"
End Select


'Populate list of rides from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "select * from rides where type='" & TableSelect & "'"
rsRide.Open sqlquery, connRide
Set TempBox.DataSource = rsRide.DataSource
TempBox.DataField = "name"
rsRide.MoveFirst
RideList.Clear
While (rsRide.EOF = False)
    RideList.AddItem (TempBox.Text)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close
End If

End Sub

Private Sub RideType_GotFocus()
RideList.Clear
RideList.Text = "(select)"
End Sub
