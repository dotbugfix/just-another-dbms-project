VERSION 5.00
Begin VB.Form Admin_RideServicing 
   Caption         =   "Rides Servicing"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_RideServicing.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   5520
   Begin VB.ComboBox RideType 
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
      ItemData        =   "Admin_RideServicing.frx":67F0
      Left            =   1200
      List            =   "Admin_RideServicing.frx":6803
      TabIndex        =   3
      Text            =   "(select)"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox RideList 
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
      ItemData        =   "Admin_RideServicing.frx":6839
      Left            =   1200
      List            =   "Admin_RideServicing.frx":683B
      TabIndex        =   2
      Text            =   "(select)"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox TempBox 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Tag 
      Caption         =   "Tag Servicing >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a type of ride:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the ride:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "Admin_RideServicing"
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
.Height = 5100
.Width = 5600
End With
MsgBox "To view a list of rides which are already tagged for servicing, use the Ride Statistics view.", vbInformation, "Demonstration"



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
connRide.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "select * from rides where type='" & TableSelect & "'"
rsRide.Open sqlquery, connRide
Set TempBox.DataSource = rsRide.DataSource
TempBox.DataField = "name"
rsRide.MoveFirst
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
End Sub

Private Sub Tag_Click()
If RideType.Text = "(select)" Or RideList.Text = "(select)" Or RideList.Text = "" Then
    MsgBox "Please select a ride to tag for servicing!", vbOKOnly, "Admin Landing"
Else

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "update (select * from rides where name='" & RideList.Text & "') set need_service=1"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been tagged for servicing! Visitors will not be able to enjoy this ride until it is serviced by a mechanic!", vbOKOnly, "Admin Landing"
Unload Me
End If
End Sub
