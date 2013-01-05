VERSION 5.00
Begin VB.Form Admin_RemoveRide 
   Caption         =   "Remove Ride"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7725
   Begin VB.CommandButton RemoveRide 
      Caption         =   "Remove Ride >>"
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
      Left            =   4440
      TabIndex        =   19
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ride Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   6495
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
         ItemData        =   "Admin_RemoveRide.frx":0000
         Left            =   3120
         List            =   "Admin_RemoveRide.frx":0013
         TabIndex        =   17
         Text            =   "(select)"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Search 
         Caption         =   "Search"
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
         Left            =   5280
         TabIndex        =   15
         Top             =   840
         Width           =   975
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
         Left            =   3120
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cost of Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   4215
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
            Left            =   240
            TabIndex        =   8
            ToolTipText     =   "Make this field 0 if you want to restrict the ride to adults only"
            Top             =   720
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
            Left            =   1560
            TabIndex        =   7
            Top             =   720
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
            Left            =   2880
            TabIndex        =   6
            ToolTipText     =   "Operating cost per visitor"
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "For kids:"
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
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "For adults:"
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
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Operating Cost:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Health Issues"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4440
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
         Begin VB.CheckBox BP 
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
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox Heart 
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
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox Nausea 
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
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Select a type of ride to remove:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "The details of the ride are:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Enter name of ride to search for:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Remove rides from  the resort"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Admin_RemoveRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlquery As String
Dim TableSelect As String
Dim rsRide As ADODB.Recordset
Dim connRide As ADODB.Connection
Private Sub RemoveRide_Click()
sqlquery = "delete from " & TableSelect & " where name='" & RideName.Text & "'"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the resort!", vbOKOnly, "Admin Landing"
Unload Me
End Sub

Private Sub RideType_LostFocus()
Select Case RideType.Text
    Case "Transport"
        TableSelect = "ride_transport"
    Case "Gentle"
        TableSelect = "ride_gentle"
    Case "Thrill"
        TableSelect = "ride_thrill"
    Case "Water"
        TableSelect = "ride_water"
    Case "Roller Coaster"
        TableSelect = "ride_coaster"
End Select
End Sub

Private Sub Search_Click()
If (RideName.Text = "") Then
    MsgBox "Please enter the name of the ride!", vbOKOnly, "Admin Landing"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")

'Search for ride in DB
sqlquery = "select * from " & TableSelect & " where name='" & RideName.Text & "'"
rsRide.Open sqlquery, connRide
If (rsRide.EOF = True) Then
    MsgBox "Ride does not exist! Please check the type & name!", vbOKOnly, "Admin Landing"
Else
Set Cost_Kids.DataSource = rsRide.DataSource
Cost_Kids.DataField = "cost_kids"
Set Cost_Adults.DataSource = rsRide.DataSource
Cost_Adults.DataField = "cost_adults"
Set Op_Cost.DataSource = rsRide.DataSource
Op_Cost.DataField = "op_cost"
Set Heart.DataSource = rsRide.DataSource
Heart.DataField = "heart"
Set BP.DataSource = rsRide.DataSource
BP.DataField = "bp"
Set Nausea.DataSource = rsRide.DataSource
Nausea.DataField = "nausea"
rsRide.Close
End If
End If
End Sub
