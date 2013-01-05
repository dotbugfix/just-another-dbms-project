VERSION 5.00
Begin VB.Form Visitor_Existing 
   Caption         =   "Visitor Login"
   ClientHeight    =   4065
   ClientLeft      =   6615
   ClientTop       =   4035
   ClientWidth     =   5640
   LinkTopic       =   "Form9"
   Picture         =   "visitor_existing.frx":0000
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   Begin VB.TextBox VisitorID 
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
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Login 
      Caption         =   "Login"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Back 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Proceed 
      Caption         =   "Proceed to Map >>"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label NameLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your Visitor ID:"
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
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your details in our database are:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   3735
   End
End
Attribute VB_Name = "Visitor_Existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVisitor As ADODB.Recordset
Dim adv_curr As Integer
Dim Name_First, Name_Last As String
Dim Visitor_Date As Date
Dim connVisitor As ADODB.Connection
Private Sub Command1_Click()
Global_Module.visitor_id = existing_vid.Text
Unload Me
Map.Show
End Sub

Private Sub Command3_Click()
'adodc_visitor.CommandType = adCmdText
'adodc_visitor.CommandText = "select * From visitor Where ID =" & existing_vid.Text
'adodc_visitor.CommandType = adCmdText
adodc_visitor.Refresh

End Sub

Private Sub Back_Click()
Unload Me
Welcome.Show
End Sub


Private Sub Label1_Click()

End Sub

Private Sub Login_Click()
If (VisitorID.Text = "") Then
    MsgBox "Please enter your visitor ID for logging in!", vbOKOnly, "Entertainment Resort"
Else
Set connVisitor = New ADODB.Connection
Set rsVisitor = New ADODB.Recordset
connVisitor.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "select * from visitor where id=" & VisitorID.Text
rsVisitor.Open sqlquery, connVisitor
If (rsVisitor.EOF = True) Then
    MsgBox "Invalid Visitor ID! If you are a new visitor, please go to the main counter!", vbOKOnly, "Entertainment Resort"
Else
    Name_First = rsVisitor.Fields("Name_First").Value
    Name_Last = rsVisitor.Fields("Name_Last").Value
    Visitor_Date = rsVisitor.Fields("entry_date").Value
    NameLabel.Caption = Name_First & " " & Name_Last
End If
rsVisitor.Close
connVisitor.Close
End If
End Sub

Private Sub Proceed_Click()
If (Name_First = "") Then
    MsgBox "Please login to the database before proceeding!", vbOKOnly, "Entertainment Resort"
Else
adv_curr = DateDiff("d", Global_Module.Today, Visitor_Date)
If (adv_curr = 0) Then
    Global_Module.visitor_id = VisitorID.Text
    Unload Me
    Map.Show
ElseIf (adv_curr < 0) Then
    MsgBox "Sorry, your booking date has already passed! Please check-in again at the main counter!", vbOKOnly, "Entertainment Resort"
    Unload Me
    Welcome.Show
Else
    MsgBox "Sorry, your advance booking date has not yet arrived! Please visit on the date of your booking!", vbOKOnly, "Entertainment Resort"
    Unload Me
    Welcome.Show
End If
End If
End Sub

Private Sub visitorid_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If Not KeyAscii = 8 Then    'Allow Backspace
    KeyAscii = 0
    End If
End If
End Sub

