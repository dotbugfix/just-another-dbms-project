VERSION 5.00
Begin VB.Form Admin_ChangeAdminPass 
   Caption         =   "c(To create a new admin, enter a new username && leave the next field blank)"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ChangeAdminPass.frx":0000
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   Begin VB.CommandButton Appoint 
      Caption         =   "Change Password >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox ExistingPass 
      DataField       =   "PASSWORD"
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   2085
   End
   Begin VB.TextBox ConfirmPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4200
      Width           =   2085
   End
   Begin VB.TextBox Password 
      DataField       =   "PASSWORD"
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   2085
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(Leave blank to remove admin)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(To create a new admin, enter a new username && leave the next field blank)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   960
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New  Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Admin_ChangeAdminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isnew As Integer
Dim sqlquery As String
Dim connAdmin As ADODB.Connection
Dim rsAdmin As ADODB.Recordset

Private Sub Appoint_Click()
Set rsAdmin = New ADODB.Recordset
Set connAdmin = New ADODB.Connection

'Check for existing admin
connAdmin.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
sqlquery = "select count(*) as num from administrator where username='" & Username.Text & "'"
rsAdmin.Open sqlquery, connAdmin

isnew = rsAdmin.Fields("NUM").Value

rsAdmin.Close

If isnew = 1 Then
    If Username.Text = "" Or ExistingPass.Text = "" Then
        MsgBox "Please fill up all required fields before continuing!", vbOKOnly, "Admin Landing"
    ElseIf (StrComp(ConfirmPassword.Text, Password.Text) <> 0) Then
        MsgBox "Please confirm the new password correctly!", vbOKOnly, "Admin Landing"
    ElseIf Password.Text = "" Then
    'Remove admin
        rsAdmin.LockType = adLockOptimistic
        sqlquery = "select * from administrator where username='" & Username.Text & "'"
        rsAdmin.Open sqlquery, connAdmin
        rsAdmin.Delete
        rsAdmin.Close
        MsgBox "The administrator has been removed from the database!", vbOKOnly, "Admin Landing"
        connAdmin.Close
        Unload Me
    Else
    'Update password
        rsAdmin.LockType = adLockOptimistic
        sqlquery = "select * from administrator where username='" & Username.Text & "'"
        rsAdmin.Open sqlquery, connAdmin
        If (StrComp(ExistingPass.Text, rsAdmin.Fields("password").Value) <> 0) Then
            MsgBox "The existing password is incorrect!", vbOKOnly, "Admin Landing"
        Else
        varfields = Array("USERNAME", "PASSWORD")
        varvalues = Array(Username.Text, Password.Text)
        rsAdmin.Update varfields, varvalues
        rsAdmin.Close
        MsgBox "The password has been updated in the database!", vbOKOnly, "Admin Landing"
        connAdmin.Close
        Unload Me
        End If
    End If
Else
'Add new admin
    If Username.Text = "" Or Password.Text = "" Or ConfirmPassword.Text = "" Then
        MsgBox "Please fill up all required fields before continuing!", vbOKOnly, "Admin Landing"
    ElseIf (StrComp(ConfirmPassword.Text, Password.Text) <> 0) Then
        MsgBox "Please confirm the new password correctly!", vbOKOnly, "Admin Landing"
    Else
        rsAdmin.LockType = adLockOptimistic
        rsAdmin.Open "select * from administrator", connAdmin
        varfields = Array("USERNAME", "PASSWORD")
        varvalues = Array(Username.Text, Password.Text)
        rsAdmin.AddNew varfields, varvalues
        rsAdmin.Close
        MsgBox "The new administrator has been added to the database!", vbOKOnly, "Admin Landing"
        connAdmin.Close
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 6165
.Width = 5820
End With

End Sub

