VERSION 5.00
Begin VB.Form Admin_NewRide 
   Caption         =   "Add Ride"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "admin_newride.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   8640
   Begin VB.TextBox Cost_Kids 
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
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Make this field 0 if you want to restrict the ride to adults only"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Cost_Adults 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Op_Cost 
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
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Operating cost per visitor"
      Top             =   4560
      Width           =   855
   End
   Begin VB.CheckBox Nausea 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5520
      TabIndex        =   7
      Top             =   2760
      Width           =   200
   End
   Begin VB.CheckBox Heart 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   200
   End
   Begin VB.CheckBox BP 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   200
   End
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
      ItemData        =   "admin_newride.frx":F35B
      Left            =   2520
      List            =   "admin_newride.frx":F36E
      TabIndex        =   0
      Text            =   "(select)"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox RideName 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton AddRide 
      Caption         =   "Add Ride >>"
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
      Left            =   5760
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Pressure"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   1750
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Heart Problems"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Nausea"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   2700
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "For kids:"
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
      Left            =   720
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "For adults:"
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
      Left            =   1920
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Operating Cost:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select a type of ride:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of ride:"
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
      Left            =   720
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "Admin_NewRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlquery As String
Dim TableSelect As String
Dim rsRide As ADODB.Recordset
Dim connRide As ADODB.Connection

Private Sub AddRide_Click()
If RideType.Text = "(select)" Or RideType.Text = "" Or RideName.Text = "" Or Cost_Kids.Text = "" Or Cost_Adults.Text = "" Then
    MsgBox "Please fill in all details properly before adding a new ride to the database!", vbOKOnly, "Admin Landing"
Else

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
rsRide.LockType = adLockOptimistic
rsRide.Open "select * from rides", connRide
varfields = Array("NAME", "TYPE", "COST_KIDS", "COST_ADULTS", "OP_COST", "HEART", "BP", "NAUSEA", "NEED_SERVICE")
varvalues = Array(RideName.Text, TableSelect, Cost_Kids.Text, Cost_Adults.Text, Op_Cost.Text, Heart.Value, BP.Value, Nausea.Value, 0)
rsRide.AddNew varfields, varvalues
rsRide.Close
connRide.Close
MsgBox "The ride has been added to the database!", vbOKOnly, "Add a ride"
Unload Me
End If
End Sub

Private Sub Cost_Kids_GotFocus()
Admin_Landing.sbStatusBar.Panels(1).Text = "Make this field 0 if you want to restrict the ride to adults only."
End Sub

Private Sub Cost_Kids_LostFocus()
Admin_Landing.sbStatusBar.Panels(1).Text = "Status"
End Sub

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 5800
.Width = 8750
End With
'With Me
'    If .WindowState <> vbMaximized Then
'        .Top = 0
'        .Left = 0
'        .Height = 4290
'        .Width = 7215
'    End If
'End With
        
End Sub

Private Sub Op_Cost_GotFocus()
Admin_Landing.sbStatusBar.Panels(1).Text = "Operating cost per month"
End Sub

Private Sub Op_Cost_LostFocus()
Admin_Landing.sbStatusBar.Panels(1).Text = "Status"
End Sub

Private Sub RideType_LostFocus()
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
End Sub
