VERSION 5.00
Begin VB.Form Admin_PerVisitor 
   Caption         =   "Per Visitor History"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_PerVisitor.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   7725
   Begin VB.CommandButton Login 
      Caption         =   "Search"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox VisitorID 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   735
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
   Begin VB.Label VisitorName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The details in Visitors' database are:"
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
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the Visitor ID:"
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
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "Admin_PerVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 5700
.Width = 7750
End With

End Sub

Private Sub Login_Click()
If VisitorID.Text = "" Then
    MsgBox "Please enter a VisitorID before Searching in the database!", vbOKOnly, "Admin Landing"
Else
Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
sqlquery = "select * from visitor where id=" & VisitorID.Text
rsVisitor.Open sqlquery, connVisitor
If (rsVisitor.EOF = True) Then
    MsgBox "Invalid Visitor ID!", vbOKOnly, "Admin Login"
    VisitorID.Text = ""
Else
    Name_First = rsVisitor.Fields("Name_First").Value
    Name_Last = rsVisitor.Fields("Name_Last").Value
    VisitorName.Caption = Name_First & " " & Name_Last
End If
rsVisitor.Close
connVisitor.Close
End If

End Sub

Private Sub Proceed_Click()
If VisitorID.Text = "" Then
    MsgBox "Please search for a visitor by his ID before generating the report!", vbOKOnly, "Admin Landing"
ElseIf VisitorName.Caption = "" Then
    Call Login_Click
Else
If OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Open
End If
OracleDB.PerVisitor (VisitorID.Text)
Report_PerVisitor.Orientation = rptOrientLandscape
Report_PerVisitor.Show
Unload Me
End If

End Sub

Private Sub VisitorID_Change()
VisitorName.Caption = ""
End Sub
