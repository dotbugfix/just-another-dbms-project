VERSION 5.00
Begin VB.Form Help_Request 
   Caption         =   "Help Request"
   ClientHeight    =   6225
   ClientLeft      =   7725
   ClientTop       =   2760
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   Picture         =   "Help_Request.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   4140
   Begin VB.ComboBox HelpList 
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
      ItemData        =   "Help_Request.frx":4264
      Left            =   720
      List            =   "Help_Request.frx":4271
      TabIndex        =   1
      Text            =   "(select)"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Request 
      Caption         =   "Request >>"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label VisitorID 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   3070
      TabIndex        =   4
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Visitor ID:"
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
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Which kind of help do you need?"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "Help_request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection
Private Sub Form_Load()
VisitorID.Caption = Global_Module.visitor_id
End Sub

Private Sub Request_Click()
If HelpList.Text = "(select)" Then
    MsgBox "Please select the type of help that you need!", vbOKOnly, "Entertainment Resort"
Else
Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
rsVisitor.LockType = adLockOptimistic
sqlquery = "select count(*) as num from visitor where id=" & VisitorID.Caption & " and need_help=1"
rsVisitor.Open sqlquery, connVisitor
cnt = rsVisitor.Fields("NUM").Value
If (cnt = 1) Then
    MsgBox "You have already requested for help earlier! Your request will be attended shortly.", vbOKOnly, "Entertainment Resort"
    rsVisitor.Close
    connVisitor.Close
    Unload Me
    Welcome.Show
Else
rsVisitor.Close
varfields = Array("NEED_HELP", "HELP_TYPE")
varvalues = Array(1, HelpList.Text)
sqlquery = "select need_help,help_type from visitor where id=" & VisitorID.Caption
rsVisitor.Open sqlquery, connVisitor

rsVisitor.Update varfields, varvalues

rsVisitor.Close
connVisitor.Close

Unload Me
AtYourService.Show
End If
End If
End Sub
