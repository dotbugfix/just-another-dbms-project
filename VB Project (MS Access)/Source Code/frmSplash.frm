VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   9390
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   9390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   9390
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   5160
      Top             =   2160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6240
      Top             =   720
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()

End Sub

Private Sub lblProductName_Click()

End Sub

Private Sub Form_Load()
Global_Module.Today = Now()

OracleDB.GetEntTax
Global_Module.Ent_Tax = OracleDB.rsGetEntTax.Fields("value").Value

OracleDB.GetSerTax
Global_Module.Ser_Tax = OracleDB.rsGetSerTax.Fields("value").Value

OracleDB.GetIncomeTax
Global_Module.Income_Tax = OracleDB.rsGetIncomeTax.Fields("value").Value

OracleDB.GetVisitorLimit
Global_Module.visitor_limit = OracleDB.rsGetVisitorLimit.Fields("value").Value

OracleDB.GetEntryFee "Entry_Kids"
Global_Module.Entry_Kids = OracleDB.rsGetEntryFee.Fields("value").Value
OracleDB.rsGetEntryFee.Close

OracleDB.GetEntryFee "Entry_Adults"
Global_Module.Entry_Adults = OracleDB.rsGetEntryFee.Fields("value").Value







InitCommonControlsVB

End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

'Private Sub Timer1_Timer()
'Timer1.Interval = 3000
'Timer1.Enabled = True
'Unload Me
'Welcome.Show
'End Sub

Private Sub Title_Click()

End Sub


'// open file (quotes are used so that the actual value that is passed is "C:\test.doc"
'Private Sub cmdOpen_Click()
'End Sub
'
''// open url
'Private Sub cmdOpen_Click()
'    ShellExecute 0, vbNullString, "http://www.vbweb.co.uk/", vbNullString, vbNullString, vbNormalFocus
'End Sub
'
''// open email address
'Private Sub cmdOpen_Click()
'    ShellExecute 0, vbNullString, "mailto:support@vbweb.co.uk", vbNullString, vbNullString, vbNormalFocus
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim str As String



End Sub

Private Sub Timer2_Timer()
Timer2.Interval = 15
Timer2.Enabled = True
If ProgressBar1.Value < 100 Then
    ProgressBar1.Value = ProgressBar1.Value + 1
Else
    Unload Me
    Welcome.Show
End If

End Sub
