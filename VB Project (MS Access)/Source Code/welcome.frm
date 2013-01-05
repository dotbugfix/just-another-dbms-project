VERSION 5.00
Begin VB.Form Welcome 
   Caption         =   "Welcome"
   ClientHeight    =   7995
   ClientLeft      =   3435
   ClientTop       =   2220
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   Begin VB.OptionButton OptionStaff 
      Caption         =   "Staff"
      Height          =   200
      Left            =   9480
      TabIndex        =   12
      Top             =   1440
      Width           =   200
   End
   Begin VB.OptionButton OptionAdmin 
      Caption         =   "Admin"
      Height          =   200
      Left            =   8280
      TabIndex        =   11
      Top             =   1440
      Value           =   -1  'True
      Width           =   200
   End
   Begin VB.TextBox Password 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   9480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1965
   End
   Begin VB.CommandButton Login 
      Caption         =   "Login >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10440
      TabIndex        =   2
      Top             =   1320
      Width           =   1140
   End
   Begin VB.TextBox UserName 
      Height          =   345
      Left            =   9480
      TabIndex        =   0
      Top             =   360
      Width           =   1965
   End
   Begin VB.CommandButton ExistingVisitor 
      Caption         =   "Existing Visitor Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton NewVisitor 
      Caption         =   "New Visitor Check-in "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Admin 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
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
      Left            =   8520
      TabIndex        =   3
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   8280
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   8280
      TabIndex        =   9
      Top             =   380
      Width           =   1095
   End
   Begin VB.Label date 
      BackStyle       =   0  'Transparent
      Caption         =   "(date)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStaff As ADODB.Recordset
Dim connStaff As ADODB.Connection
Dim rsVisitor As ADODB.Recordset
Dim connVisitor As ADODB.Connection
Dim Role As String
Dim PasswordIn As String
Dim cnt As Integer
Private Sub cmdOK_Click()

End Sub
Private Sub txtUserName_Change()

End Sub

Private Sub AdminSubmit_Click()
Unload Me
Admin_Landing.Show
End Sub

Private Sub ExistingVisitor_Click()
Unload Me
Visitor_Existing.Show
End Sub

Private Sub Form_Load()
date.Caption = Format(Global_Module.Today, "DD-MMM-YYYY")
End Sub

Private Sub NewVisitor_Click()
    Unload Me
    MainCounter.Show
End Sub

Private Sub Login_Click()
If OptionStaff.Value = True Then
    
    If (Username.Text = "") Then
        MsgBox "Please enter the staff username!", vbOKOnly, "Staff Login"
    ElseIf (Password.Text = "") Then
        MsgBox "Please enter the staff password!", vbOKOnly, "Staff Login"
    Else
    Set connStaff = New ADODB.Connection
    Set rsStaff = New ADODB.Recordset
    connStaff.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
    
    'Search for username in Staff DB
    sqlquery = "select count (*) as num from staff where username='" & Username.Text & "'"
    rsStaff.Open sqlquery, connStaff
    cnt = rsStaff.Fields("NUM").Value
    
    If cnt = 0 Then
        MsgBox "Invalid staff username! Please try again!", vbOKOnly, "Staff Login"
        
    Else
    rsStaff.Close
    sqlquery = "select * from staff where username='" & Username.Text & "'"
    rsStaff.Open sqlquery, connStaff
    Role = rsStaff.Fields("role").Value
    
    PasswordIn = rsStaff.Fields("password").Value
    
    If (StrComp(Password.Text, PasswordIn) = 0) Then
        Global_Module.staff_login = Username.Text
        'Show Salary Slip
        If DatePart("d", Global_Module.Today) = 1 Then
            If DateDiff("m", Format(Global_Module.Today, "dd-mmm-yyyy"), Format(rsStaff.Fields("last_claim"), "dd-mmm-yyyy")) < 0 Then
                Unload Me
                Staff_SalarySlip.Show
                GoTo finished
            End If
        End If
        'Normal Login
        If (StrComp(Role, "Assistant") = 0) Then
        'Check for queued up requests
            Set connVisitor = New ADODB.Connection
            Set rsVisitor = New ADODB.Recordset
            connVisitor.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
            sqlquery = "select count(*) as num from visitor where need_help=1"
            rsVisitor.Open sqlquery, connVisitor
            
            cnt = rsVisitor.Fields("num").Value
            
            rsVisitor.Close
            connVisitor.Close
            
            If cnt = 0 Then
                MsgBox "There are no visitors requesting for assistance at the moment!", vbOKOnly, "Entertainment Resort"
            Else
            Unload Me
            Staff_Assistant_landing.Show
            End If
            'Unload Staff_Assistant_landing
        ElseIf (StrComp(Role, "Mechanic") = 0) Then
            Unload Me
            Staff_Mechanic_Landing.Show
        End If
        
    Else
        MsgBox "Invalid password! Please try again!", vbOKOnly, "Staff Login"
    End If
    End If
finished:
    Username.Text = ""
    Password.Text = ""
    rsStaff.Close
    connStaff.Close
    End If
    
Else        'Admin Login

    If (Username.Text = "") Then
        MsgBox "Please enter the admin username!", vbOKOnly, "Admin Login"
    ElseIf (Password.Text = "") Then
        MsgBox "Please enter the admin password!", vbOKOnly, "Admin Login"
    Else
    Set connStaff = New ADODB.Connection
    Set rsStaff = New ADODB.Recordset
    connStaff.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\db.mdb;Mode=ReadWrite;Persist Security Info=False")
    
    'Search for username in Staff DB
    sqlquery = "select count (*) as num from administrator where username='" & Username.Text & "'"
    rsStaff.Open sqlquery, connStaff
    cnt = rsStaff.Fields("NUM").Value
    
    If cnt = 0 Then
        MsgBox "Invalid Admin username! Please try again!", vbOKOnly, "Staff Login"
    Else
    rsStaff.Close
    sqlquery = "select * from administrator where username='" & Username.Text & "'"
    rsStaff.Open sqlquery, connStaff
        
    PasswordIn = rsStaff.Fields("password").Value
    
    If (StrComp(Password.Text, PasswordIn) = 0) Then
        Unload Me
        Admin_Landing.Show
    Else
        MsgBox "Invalid password! Please try again!", vbOKOnly, "Admin Login"
    End If
    End If
    Username.Text = ""
    Password.Text = ""
    rsStaff.Close
    connStaff.Close
    End If
End If

End Sub
