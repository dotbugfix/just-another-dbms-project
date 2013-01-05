VERSION 5.00
Begin VB.Form Admin_FireStaff 
   Caption         =   "Fire Staff"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Fire 
      Caption         =   "Fire Staff >>"
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
      Left            =   3720
      TabIndex        =   13
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Staff Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5775
      Begin VB.TextBox Role 
         DataField       =   "SALARY"
         DataSource      =   "adodcStaff"
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
         Left            =   840
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
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
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Name_First 
         DataField       =   "NAME_FIRST"
         DataSource      =   "adodcStaff"
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
         Left            =   840
         TabIndex        =   5
         Text            =   "(First Name)"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Name_Last 
         DataField       =   "NAME_LAST"
         DataSource      =   "adodcStaff"
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
         Left            =   3000
         TabIndex        =   4
         Text            =   "(Last Name)"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Username 
         DataField       =   "USERNAME"
         DataSource      =   "adodcStaff"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Salary 
         DataField       =   "SALARY"
         DataSource      =   "adodcStaff"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Role:"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "The details of the staff member are:"
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
         TabIndex        =   10
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Username to search for:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Salary:"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Fire a Staff Member"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Admin_FireStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStaff As ADODB.Recordset
Dim connStaff As ADODB.Connection

Private Sub Fire_Click()
sqlquery = "delete from staff where username='" & Username.Text & "'"
connStaff.Execute sqlquery
connStaff.Close
MsgBox "The selected staff member has been fired!", vbOKOnly, "Admin Landing"
Unload Me
End Sub

Private Sub Search_Click()
If (Username.Text = "") Then
    MsgBox "Please enter the staff username!", vbOKOnly, "Admin Landing"
Else
Set connStaff = New ADODB.Connection
Set rsStaff = New ADODB.Recordset
connStaff.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")

'Search for username in Staff DB
sqlquery = "select * from staff where username='" & Username.Text & "'"
rsStaff.Open sqlquery, connStaff
If (rsStaff.EOF = True) Then
    MsgBox "Invalid staff username!", vbOKOnly, "Admin Landing"
Else
Set Name_First.DataSource = rsStaff.DataSource
Name_First.DataField = "name_first"
Set Name_Last.DataSource = rsStaff.DataSource
Name_Last.DataField = "name_last"
Set Role.DataSource = rsStaff.DataSource
Role.DataField = "role"
Set Salary.DataSource = rsStaff.DataSource
Salary.DataField = "salary"
rsStaff.Close
End If
End If
End Sub
