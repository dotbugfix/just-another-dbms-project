VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Admin_StatRide 
   Caption         =   "Ride Statistics"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   11145
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Thrill Rides"
      TabPicture(0)   =   "Admin_StatRide.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Transport Rides"
      TabPicture(1)   =   "Admin_StatRide.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Gentle Rides"
      TabPicture(2)   =   "Admin_StatRide.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(1)=   "DataGrid3"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Water Rides"
      TabPicture(3)   =   "Admin_StatRide.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label22"
      Tab(3).Control(1)=   "DataGrid4"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Roller Coasters"
      TabPicture(4)   =   "Admin_StatRide.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(1)=   "DataGrid5"
      Tab(4).Control(2)=   "Frame5"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame5 
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   23
         Top             =   2640
         Width           =   6135
         Begin VB.Label Label25 
            DataField       =   "Number_of_Rides"
            DataMember      =   "CoasterRideStat"
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
            Left            =   5160
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Total number of Roller Coasters:"
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
            TabIndex        =   26
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label23 
            Caption         =   "Number of rides needing service:"
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
            TabIndex        =   25
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label6 
            DataField       =   "Needs_Service"
            DataMember      =   "CoasterRideStat"
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
            Left            =   5160
            TabIndex        =   24
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   18
         Top             =   2640
         Width           =   6135
         Begin VB.Label Label21 
            DataField       =   "Number_of_Rides"
            DataMember      =   "WaterRideStat"
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
            Left            =   4680
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "Total number of Water Rides:"
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
            TabIndex        =   21
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label19 
            Caption         =   "Number of rides needing service:"
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
            TabIndex        =   20
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label5 
            DataField       =   "Needs_Service"
            DataMember      =   "WaterRideStat"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   840
            Width           =   855
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   13
         Top             =   2640
         Width           =   6135
         Begin VB.Label Label17 
            DataField       =   "Number_of_Rides"
            DataMember      =   "GentleRideStat"
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
            Left            =   4800
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Total number of Gentle Rides:"
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
            TabIndex        =   16
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label15 
            Caption         =   "Number of rides needing service:"
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
            TabIndex        =   15
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label4 
            DataField       =   "Needs_Service"
            DataMember      =   "GentleRideStat"
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
            Left            =   5160
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   8
         Top             =   2640
         Width           =   6135
         Begin VB.Label Label13 
            DataField       =   "Number_of_Rides"
            DataMember      =   "TransportRideStat"
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
            Left            =   5160
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Total number of Transport Rides:"
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
            TabIndex        =   11
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label11 
            Caption         =   "Number of rides needing service:"
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
            TabIndex        =   10
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label3 
            DataField       =   "Needs_Service"
            DataMember      =   "TransportRideStat"
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
            Left            =   5160
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
      End
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
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   2640
         Width           =   6135
         Begin VB.Label Label9 
            DataField       =   "Needs_Service"
            DataMember      =   "ThrillRideStat"
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
            Left            =   5160
            TabIndex        =   7
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Number of rides needing service:"
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
            TabIndex        =   6
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label Label2 
            Caption         =   "Total number of Thrill Rides:"
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
            TabIndex        =   4
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label1 
            DataField       =   "Num_of_rides"
            DataMember      =   "ThrillRideStat"
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
            Left            =   4560
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Admin_StatRide.frx":008C
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         DataMember      =   "Ride_Thrill"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NAME"
            Caption         =   "Name"
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
            DataField       =   "COST_KIDS"
            Caption         =   "Kids Fee"
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
            DataField       =   "COST_ADULTS"
            Caption         =   "Adults Fee"
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
         BeginProperty Column03 
            DataField       =   "OP_COST"
            Caption         =   "Operating Cost"
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
         BeginProperty Column04 
            DataField       =   "BP"
            Caption         =   "BP Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "HEART"
            Caption         =   "Heart Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NAUSEA"
            Caption         =   "Nausea Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "NEED_SERVICE"
            Caption         =   "Servicing tag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Admin_StatRide.frx":00A3
         Height          =   2055
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         DataMember      =   "Ride_Transport"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NAME"
            Caption         =   "Name"
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
            DataField       =   "COST_KIDS"
            Caption         =   "Kids Fee"
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
            DataField       =   "COST_ADULTS"
            Caption         =   "Adults Fee"
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
         BeginProperty Column03 
            DataField       =   "OP_COST"
            Caption         =   "Operating Cost"
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
         BeginProperty Column04 
            DataField       =   "BP"
            Caption         =   "BP Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "HEART"
            Caption         =   "Heart Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0.000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NAUSEA"
            Caption         =   "Nausea Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0.000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "NEED_SERVICE"
            Caption         =   "Servicing tag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "Admin_StatRide.frx":00BA
         Height          =   2055
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         DataMember      =   "Ride_Gentle"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NAME"
            Caption         =   "Name"
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
            DataField       =   "COST_KIDS"
            Caption         =   "Kids Fee"
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
            DataField       =   "COST_ADULTS"
            Caption         =   "Adults Fee"
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
         BeginProperty Column03 
            DataField       =   "OP_COST"
            Caption         =   "Operating Cost"
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
         BeginProperty Column04 
            DataField       =   "BP"
            Caption         =   "BP Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0.000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "HEART"
            Caption         =   "Heart Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NAUSEA"
            Caption         =   "Nausea Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "NEED_SERVICE"
            Caption         =   "Servicing tag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Admin_StatRide.frx":00D1
         Height          =   2055
         Left            =   -74760
         TabIndex        =   33
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         DataMember      =   "Ride_Water"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NAME"
            Caption         =   "Name"
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
            DataField       =   "COST_KIDS"
            Caption         =   "Kids Fee"
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
            DataField       =   "COST_ADULTS"
            Caption         =   "Adults Fee"
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
         BeginProperty Column03 
            DataField       =   "OP_COST"
            Caption         =   "Operating Cost"
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
         BeginProperty Column04 
            DataField       =   "BP"
            Caption         =   "BP Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "HEART"
            Caption         =   "Heart Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NAUSEA"
            Caption         =   "Nausea Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "NEED_SERVICE"
            Caption         =   "Servicing tag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "Admin_StatRide.frx":00E8
         Height          =   2055
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         DataMember      =   "Ride_Coaster"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NAME"
            Caption         =   "Name"
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
            DataField       =   "COST_KIDS"
            Caption         =   "Kids Fee"
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
            DataField       =   "COST_ADULTS"
            Caption         =   "Adults Fee"
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
         BeginProperty Column03 
            DataField       =   "OP_COST"
            Caption         =   "Operating Cost"
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
         BeginProperty Column04 
            DataField       =   "BP"
            Caption         =   "BP Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "HEART"
            Caption         =   "Heart Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NAUSEA"
            Caption         =   "Nausea Prob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "NEED_SERVICE"
            Caption         =   "Servicing tag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "You may change the parameters of a perticular ride here. You may also remove a ride from the resort."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68520
         TabIndex        =   36
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "You may change the parameters of a perticular ride here. You may also remove a ride from the resort."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68520
         TabIndex        =   34
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "You may change the parameters of a perticular ride here. You may also remove a ride from the resort."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68520
         TabIndex        =   32
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "You may change the parameters of a perticular ride here. You may also remove a ride from the resort."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68520
         TabIndex        =   30
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "You may change the parameters of a perticular ride here. You may also remove a ride from the resort."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6480
         TabIndex        =   28
         Top             =   3120
         Width           =   3855
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Ride Statistics"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Admin_StatRide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me
.Top = 100
.Left = 100
.Height = 5800
.Width = 11200
End With

DataGrid1.Refresh
DataGrid2.Refresh
DataGrid3.Refresh
DataGrid4.Refresh
DataGrid5.Refresh
End Sub

