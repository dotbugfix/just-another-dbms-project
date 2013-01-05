VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form restaurant 
   Caption         =   "restaurant"
   ClientHeight    =   6570
   ClientLeft      =   4590
   ClientTop       =   2640
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   Picture         =   "restaurant.frx":0000
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   675
   Begin VB.CommandButton RemoveFromCart 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9000
      TabIndex        =   14
      Top             =   3120
      Width           =   405
   End
   Begin VB.CommandButton Checkout 
      Caption         =   "Proceed to check out >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7680
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Back 
      Caption         =   "<<Back"
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
      Left            =   480
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton AddtoCart 
      Caption         =   "Add to Order"
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
      Left            =   720
      TabIndex        =   2
      Top             =   4680
      Width           =   2175
   End
   Begin VB.ListBox Cart 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   7320
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin MSDataListLib.DataList ItemCost 
      Bindings        =   "restaurant.frx":17EFD
      Height          =   1770
      Left            =   2280
      TabIndex        =   0
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   3122
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList ItemList 
      Bindings        =   "restaurant.frx":17F14
      DataSource      =   "OracleDB"
      Height          =   1770
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3122
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo ItemType 
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      _Version        =   393216
      ListField       =   ""
      Text            =   "(select)"
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Billing Amount:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   5820
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "Rupee"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5835
      Width           =   255
   End
   Begin VB.Label Bill 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   5820
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Current Order:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Visitor ID:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label v_id 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(id)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7770
      TabIndex        =   8
      Top             =   3870
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a type of item that you wish to order:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "restaurant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Dim slquery As String
Dim item_sp, item_cp As Integer
Dim total_sp, total_cp As Integer

Private Sub AddtoCart_Click()
If ItemList.Text = "" Then
    MsgBox "Please select an item to add to your order!", vbOKOnly, "Entertainment Resort"
Else
Cart.AddItem ItemList.Text
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
sqlquery = "select item_sp,item_cp from restaurant where item_name='" & ItemList.Text & "'"
rs.Open sqlquery, conn
item_sp = rs.Fields("item_sp").Value
item_cp = rs.Fields("item_cp").Value
total_sp = total_sp + item_sp
total_cp = total_cp + item_cp
rs.Close
conn.Close
Bill.Caption = total_sp 'update bill amount
End If
End Sub

Private Sub Back_Click()
Map.Show
Unload Me
End Sub

Private Sub Checkout_Click()
If total_sp = 0 Then
    MsgBox "Please order some food before proceeding!", vbOKOnly, "Entertainment Resort"
Else

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
sqlquery = "select * from log_visitor"
rs.Open sqlquery, conn, adOpenDynamic, adLockOptimistic
    varfields = Array("v_id", "action_name", "action_type", "fees_paid", "cp", "action_date")
    varvalues = Array(Val(v_id.Caption), "Dining", "Restaurant", Val(Bill.Caption), total_cp, Format(Global_Module.Today, "dd-mmm-yyyy"))
rs.AddNew varfields, varvalues
rs.Close
conn.Close
MsgBox "Enjoy the food!", vbOKOnly, "Restautants"
Unload Me
Welcome.Show
End If
End Sub

Private Sub Form_Load()
v_id.Caption = Global_Module.visitor_id
total_sp = 0
total_cp = 0

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.getRestItemList

Set ItemType.RowSource = OracleDB
ItemType.RowMember = "getRestItemList"
ItemType.ListField = "ITEM_TYPE"

End Sub

Private Sub ItemType_Change()
OracleDB.OracleProvider.Close
OracleDB.OracleProvider.Open
OracleDB.PopulateRestItems ItemType.Text
ItemList.RowMember = "PopulateRestItems"
ItemList.ListField = "ITEM_NAME"
ItemList.Refresh
ItemList.ReFill

ItemCost.RowMember = "PopulateRestItems"
ItemCost.ListField = "ITEM_SP"
ItemCost.Refresh
ItemCost.ReFill
End Sub

Private Sub ItemType_GotFocus()
ItemList.RowMember = ""
ItemList.ListField = ""
ItemList.Refresh

ItemCost.RowMember = ""
ItemCost.ListField = ""
ItemCost.Refresh


End Sub

Private Sub RemoveFromCart_Click()
sel = Cart.ListIndex
If sel >= 0 Then
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
    sqlquery = "select item_sp,item_cp from restaurant where item_name='" & ItemList.Text & "'"
    rs.Open sqlquery, conn
    item_sp = rs.Fields("item_sp").Value
    item_cp = rs.Fields("item_cp").Value
    total_sp = total_sp - item_sp
    total_cp = total_cp - item_cp
    rs.Close
    conn.Close
    Cart.RemoveItem sel
    Bill.Caption = total_sp 'update bill amount
Else
    MsgBox "Please select an item in your order to cancel!", vbOKOnly, "Entertainment Resort"
End If

End Sub
