VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Admin_ShopAddItem 
   Caption         =   "Shops & Restaurants - Add Item"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_ShopAddItem.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   6015
   Begin VB.TextBox ItemName 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox CP 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox SP 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton AddItem 
      Caption         =   "Add Item >>"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ComboBox ShopType 
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
      ItemData        =   "Admin_ShopAddItem.frx":627E
      Left            =   3240
      List            =   "Admin_ShopAddItem.frx":6288
      TabIndex        =   1
      Text            =   "(select)"
      Top             =   1120
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo ItemType 
      Bindings        =   "Admin_ShopAddItem.frx":629E
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "ITEM_TYPE"
      Text            =   "(select)"
      Object.DataMember      =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2805
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Price:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3420
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3420
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the type of shop:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1170
      Width           =   2655
   End
End
Attribute VB_Name = "Admin_ShopAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddItem_Click()
If ShopType.Text = "(select)" Or ItemType.Text = "(select)" Or ItemName.Text = "" Or CP.Text = "" Or SP.Text = "" Then
    MsgBox "Please fill in all details about the new item before adding it to the database!", vbOKOnly, "Entertainment Resort"
Else
    
    varfields = Array("item_type", "item_name", "item_cp", "item_sp")
    varvalues = Array(ItemType.Text, ItemName.Text, CP.Text, SP.Text)
If ShopType.Text = "Shop" Then
    If OracleDB.rsShop.State = closed Then
        OracleDB.Shop
    End If
    OracleDB.rsShop.AddNew varfields, varvalues
ElseIf ShopType.Text = "Restaurant" Then
    If OracleDB.rsRestaurant.State = closed Then
        OracleDB.restaurant
    End If
    OracleDB.rsRestaurant.AddNew varfields, varvalues
End If
MsgBox "The new item has been added!", vbOKOnly, "Admin Landing"
Unload Me
End If
End Sub

Private Sub CP_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
End If

End Sub

Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 5310
.Width = 6135
End With
End Sub

Private Sub ItemType_GotFocus()
If ShopType.Text = "Shop" Then
    If Not OracleDB.OracleProvider.State = closed Then
        OracleDB.OracleProvider.Close
    End If
    OracleDB.OracleProvider.Open
    OracleDB.getShopItemList
    ItemType.RowMember = "getShopItemList"
    MsgBox "The types of items already present in the database are shown here. You may manually edit the contents of this combo box to create a new item type.", vbInformation, "Demonstration"
ElseIf ShopType.Text = "Restaurant" Then
    If Not OracleDB.OracleProvider.State = closed Then
        OracleDB.OracleProvider.Close
    End If
    OracleDB.OracleProvider.Open
    OracleDB.getRestItemList
    ItemType.RowMember = "getRestItemList"
    MsgBox "The types of items already present in the database are shown here. You may manually edit the contents of this combo box to create a new item type.", vbInformation, "Demonstration"
Else
    MsgBox "Please select the type of shop first!", vbOKOnly, "Admin Landing"
End If

ItemType.Refresh
ItemType.ReFill
End Sub

Private Sub ShopType_GotFocus()
ItemType.RowMember = ""
ItemType.Text = ""
End Sub

Private Sub SP_KeyPress(KeyAscii As Integer)

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    
        If Not KeyAscii = 8 Then    'Allow Backspace
            KeyAscii = 0
        End If
End If

End Sub
