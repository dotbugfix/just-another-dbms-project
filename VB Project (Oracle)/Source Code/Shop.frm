VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Shop 
   Caption         =   "Shops"
   ClientHeight    =   6600
   ClientLeft      =   5910
   ClientTop       =   2880
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   Picture         =   "Shop.frx":0000
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
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
      Left            =   6600
      TabIndex        =   12
      Top             =   2640
      Width           =   405
   End
   Begin MSDataListLib.DataList ItemCost 
      Bindings        =   "Shop.frx":27F51
      Height          =   1980
      Left            =   4080
      TabIndex        =   7
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   3493
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
   End
   Begin VB.ListBox Cart 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4800
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton AddtoCart 
      Caption         =   "Add to cart"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin MSDataListLib.DataList ItemList 
      Bindings        =   "Shop.frx":27F68
      DataSource      =   "OracleDB"
      Height          =   1980
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3493
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo ItemType 
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   "ITEM_TYPE"
      Text            =   "(select)"
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
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
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
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
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bill Amount:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   5790
      Width           =   2175
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
      Left            =   3720
      TabIndex        =   13
      Top             =   5775
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
      Left            =   4020
      TabIndex        =   11
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cart:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   4875
      Width           =   1455
   End
   Begin VB.Label v_id 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(id)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   4875
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a type of item:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2220
      Width           =   2175
   End
End
Attribute VB_Name = "Shop"
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
    MsgBox "Please select an item to add to the cart!", vbOKOnly, "Entertainment Resort"
Else
Cart.AddItem ItemList.Text
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open ("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott")
sqlquery = "select item_sp,item_cp from shop where item_name='" & ItemList.Text & "'"
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



Private Sub Checkout_Click()
If total_sp = 0 Then
    MsgBox "Please buy some items before proceeding!", vbOKOnly, "Entertainment Resort"
Else
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.Open "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
sqlquery = "select * from log_visitor"
rs.Open sqlquery, conn, adOpenDynamic, adLockOptimistic
    varfields = Array("v_id", "action_name", "action_type", "fees_paid", "cp", "action_date")
    varvalues = Array(Val(v_id.Caption), "Shopping", "Shop", Val(Bill.Caption), total_cp, Format(Global_Module.Today, "dd-mmm-yyyy"))
rs.AddNew varfields, varvalues
rs.Close
conn.Close
MsgBox "Thankyou for shopping!", vbOKOnly, "Shops"
Unload Me
Welcome.Show
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me
Map.Show
End Sub

Private Sub Command3_Click()
Map.Show
End Sub

Private Sub DataList1_Click()

End Sub

Private Sub DataList1_GotFocus()

End Sub

Private Sub Form_Load()
v_id.Caption = Global_Module.visitor_id
total_sp = 0
total_cp = 0

If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.getShopItemList

Set ItemType.RowSource = OracleDB
ItemType.RowMember = "getShopItemList"
ItemType.ListField = "ITEM_TYPE"

End Sub

Private Sub ItemType_Change()
OracleDB.OracleProvider.Close
OracleDB.OracleProvider.Open
OracleDB.PopulateShopItems ItemType.Text
ItemList.RowMember = "PopulateShopItems"
ItemList.ListField = "ITEM_NAME"
ItemList.Refresh
ItemList.ReFill

ItemCost.RowMember = "PopulateShopItems"
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
    sqlquery = "select item_sp,item_cp from shop where item_name='" & Cart.Text & "'"
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
    MsgBox "Please select an item in the cart to remove!", vbOKOnly, "Entertainment Resort"
End If
End Sub
