VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MainCounter 
   Caption         =   "Main Counter"
   ClientHeight    =   6315
   ClientLeft      =   4410
   ClientTop       =   2580
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   Picture         =   "maincounter.frx":0000
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel Check-in"
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
      Left            =   7800
      TabIndex        =   32
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Calculate_Fee 
      Caption         =   "Calculate Entry Fee >>"
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
      Left            =   6120
      TabIndex        =   10
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Age 
      DataField       =   "AGE"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Name_First 
      DataField       =   "NAME_FIRST"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Text            =   "(First Name)"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Name_Last 
      DataField       =   "NAME_LAST"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   1
      Text            =   "(Last Name)"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox ContactNo 
      DataField       =   "CONTACT"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Ent_Tax 
      DataField       =   "ENT_TAX"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Text            =   "(enttax)"
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check3 
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   3720
      Width           =   195
   End
   Begin VB.CheckBox Check2 
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   4080
      Width           =   195
   End
   Begin VB.CheckBox Check1 
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   4440
      Width           =   195
   End
   Begin VB.CheckBox Locker 
      DataField       =   "LOCKER"
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   5160
      Width           =   195
   End
   Begin VB.CheckBox Camera 
      DataField       =   "CAMERA"
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Top             =   5160
      Width           =   195
   End
   Begin VB.TextBox DateBox 
      DataField       =   "ENTRY_DATE"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Bindings        =   "maincounter.frx":A92C
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16580611
      CurrentDate     =   40453
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >>"
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
      Left            =   8640
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adodc_visitor 
      Height          =   615
      Left            =   6120
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "VISITOR"
      Caption         =   "adodc_visitor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Camera   (Rs.100/- extra)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   31
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Locker Facility (Rs.100/- extra)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   30
      Top             =   5040
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
      Left            =   1320
      TabIndex        =   29
      Top             =   4440
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
      Left            =   1320
      TabIndex        =   28
      Top             =   4080
      Width           =   1575
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
      Left            =   1320
      TabIndex        =   27
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label vid1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Visitor ID is:"
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
      Left            =   7200
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label vid2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "You may use this ID to login again to the resort!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label VisitorID 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      DataField       =   "ID"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   855
      Left            =   8970
      TabIndex        =   24
      Top             =   1350
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "The entry fee will be based on the choices that you opt for on this form."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   23
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your age:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Entry Fee is: (inclusive of all taxes)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   21
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label EntryFee 
      BackStyle       =   0  'Transparent
      Caption         =   "(fee)"
      DataField       =   "ENTRY_FEE"
      DataSource      =   "adodc_visitor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   19
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Label locker1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your locker number is same as your VisitorID!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   960
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
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
      Left            =   960
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date:"
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
      Left            =   960
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Depending on your health problems, you may be advised not to board certain rides as a safety precaution."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   11
      Top             =   3675
      Width           =   2655
   End
End
Attribute VB_Name = "MainCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim entry_date As Date
Dim v_id As Integer
Dim cnt As Integer
Dim entry_fee As Integer

Private Sub Age_keypress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
End If
End Sub

Private Sub Calculate_Fee_Click()
If Age.Text = "" Or Name_First.Text = "" Or Name_Last.Text = "" Then
    MsgBox "Please fill all fields before continuing!", vbOKOnly, "Main Counter"
ElseIf Len(ContactNo.Text) < 10 Then
    MsgBox "Please check your contact number!", vbOKOnly, "Main Counter"
ElseIf DateDiff("d", Global_Module.Today, DTPicker1.Value) < 0 Then
    MsgBox "Please make a booking for today or in the future!", vbOKOnly, "Main Counter"
Else
entry_fee = 0
If Val(Age.Text) < 15 Then
    entry_fee = entry_fee + Global_Module.Entry_Kids
Else
    entry_fee = entry_fee + Global_Module.Entry_Adults
End If

If Locker.Value = 1 Then
    entry_fee = entry_fee + 100
    locker1.Visible = True
End If

If Camera.Value = 1 Then
    entry_fee = entry_fee + 100
End If

EntryFee.Caption = entry_fee

'Auto-Generate Visitor ID

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
OracleDB.CountID
'Set TempBox.DataSource = OracleDB.rsCountID.DataSource
'TempBox.DataField = "num_of_visitor"


v_id = OracleDB.rsCountID.Fields("num_of_visitor").Value + 1
VisitorID.Caption = v_id
VisitorID.Visible = True
vid1.Visible = True
vid2.Visible = True
Global_Module.visitor_id = Val(VisitorID.Caption)

OracleDB.rsCountID.Close
OracleDB.OracleProvider.Close
'Calculate Ent. Tax
Ent_Tax.Text = Format((EntryFee.Caption / (1 + (Global_Module.Ent_Tax / 100))) * (Global_Module.Ent_Tax / 100), "00.00")

End If

End Sub

Private Sub Camera_Click()
EntryFee.Caption = ""
End Sub

Private Sub Cancel_Click()
Unload Me
Welcome.Show
End Sub

Private Sub Command1_Click()
If EntryFee.Caption = "" Then
    MsgBox "Please calculate your entry fee first before proceeding!", vbOKOnly, "Main Counter"
Else
'Check if Visitor Limit has been reached for seleted date

If OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Open
End If
OracleDB.getVisitors_on_Date (Format(DTPicker1.Value, "DD-MMM-YYYY"))
'Set TempBox3.DataSource = OracleDB.rsgetVisitors_on_Date.DataSource
'TempBox3.DataField = "num"

cnt = OracleDB.rsgetVisitors_on_Date.Fields("num").Value
OracleDB.rsgetVisitors_on_Date.Close
OracleDB.OracleProvider.Close

If (cnt < Global_Module.visitor_limit) Then
    DateBox.Text = Format(DTPicker1.Value, "dd-mmm-yyyy")
    adodc_visitor.Recordset.Save
    Unload Me
    Load Rules_Reg
    Rules_Reg.Show
Else
    MsgBox "Sorry, we're fully booked for the date you wish to visit our resort! Please visit some other time!", vbOKOnly, "Entertainment Resort"
End If
End If
End Sub

Private Sub ContactNo_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
End If

End Sub



Private Sub Form_Load()
adodc_visitor.Recordset.AddNew
DTPicker1.Value = Format(Global_Module.Today, "DD-MMM-YYYY")
End Sub

Private Sub Frame3_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Locker_Click()
EntryFee.Caption = ""
End Sub

Private Sub name_first_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
    If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
    End If
End If
End Sub



Private Sub name_last_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
    If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
    End If
End If
End Sub

