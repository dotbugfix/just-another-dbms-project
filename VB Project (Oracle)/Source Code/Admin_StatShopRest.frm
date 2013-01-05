VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Admin_StatShopRest 
   Caption         =   "Shop & Restaurant Statistics"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   5460
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Shops"
      TabPicture(0)   =   "Admin_StatShopRest.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(2)=   "ItemType"
      Tab(0).Control(3)=   "Label3"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Restaurants"
      TabPicture(1)   =   "Admin_StatShopRest.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MenuType"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Aggregates"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   4695
         Begin VB.Label MenuCount 
            DataField       =   "Count_Menu"
            DataMember      =   "Count_Menu"
            DataSource      =   "OracleDB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3720
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Total number of Items:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Aggregates"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   3
         Top             =   3360
         Width           =   4695
         Begin VB.Label Label16 
            Caption         =   "Total number of Items:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label ItemCount 
            DataSource      =   "OracleDB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3720
            TabIndex        =   4
            Top             =   360
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Admin_StatShopRest.frx":0038
         Height          =   2055
         Left            =   -74760
         TabIndex        =   2
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ITEM_NAME"
            Caption         =   "Item Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ITEM_CP"
            Caption         =   "Cost Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "ITEM_SP"
            Caption         =   "Selling Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1950.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo ItemType 
         Height          =   360
         Left            =   -72600
         TabIndex        =   9
         Top             =   480
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Admin_StatShopRest.frx":004F
         Height          =   2055
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ITEM_NAME"
            Caption         =   "Item Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ITEM_CP"
            Caption         =   "Cost Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "ITEM_SP"
            Caption         =   "Selling Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1950.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo MenuType 
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
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
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   13
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label3 
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
         Left            =   -74760
         TabIndex        =   10
         Top             =   540
         Width           =   2175
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Shops && Restaurants"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Admin_StatShopRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 6250
.Width = 5580
End With

OracleDB.getShopItemList

Set ItemType.RowSource = OracleDB
ItemType.RowMember = "getShopItemList"
ItemType.ListField = "ITEM_TYPE"

OracleDB.getRestItemList

Set MenuType.RowSource = OracleDB
MenuType.RowMember = "getRestItemList"
MenuType.ListField = "ITEM_TYPE"

'DataGrid1.Refresh
'DataGrid2.Refresh
End Sub

Private Sub ItemType_Change()
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.ShopCount ItemType.Text
DataGrid1.DataMember = "StatShop"
DataGrid1.Refresh

'ItemCount.DataMember = ""
'ItemCount.DataField = ""

ItemCount.Caption = OracleDB.rsShopCount.Fields("CountItems").Value

End Sub

Private Sub MenuType_Click(Area As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.RestCount MenuType.Text
DataGrid2.DataMember = "StatRest"
DataGrid2.Refresh

'ItemCount.DataMember = ""
'ItemCount.DataField = ""

MenuCount.Caption = OracleDB.rsRestCount.Fields("CountMenu").Value

End Sub
