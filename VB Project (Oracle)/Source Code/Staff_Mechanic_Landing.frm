VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Staff_Mechanic_Landing 
   Caption         =   "Staff Landing - Mechanic"
   ClientHeight    =   4155
   ClientLeft      =   5685
   ClientTop       =   4200
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   8055
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
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
      TabPicture(0)   =   "Staff_Mechanic_Landing.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ThrillService"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Transport Rides"
      TabPicture(1)   =   "Staff_Mechanic_Landing.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TransportService"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Picture2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Gentle Rides"
      TabPicture(2)   =   "Staff_Mechanic_Landing.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GentleService"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Picture3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Water Rides"
      TabPicture(3)   =   "Staff_Mechanic_Landing.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WaterService"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Picture4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Roller Coasters"
      TabPicture(4)   =   "Staff_Mechanic_Landing.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "RollerService"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Picture5"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.PictureBox Picture5 
         Height          =   2535
         Left            =   -71280
         Picture         =   "Staff_Mechanic_Landing.frx":008C
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   25
         Top             =   480
         Width           =   3735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Service Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox CoasterRideList 
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
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "The following rides are pending for servicing:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton RollerService 
         Caption         =   "Service Ride >>"
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
         Left            =   -73920
         TabIndex        =   21
         Top             =   2160
         Width           =   2535
      End
      Begin VB.PictureBox Picture4 
         Height          =   2535
         Left            =   -71280
         Picture         =   "Staff_Mechanic_Landing.frx":9154
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   20
         Top             =   480
         Width           =   3735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Service Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox WaterRideList 
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
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label5 
            Caption         =   "The following rides are pending for servicing:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton WaterService 
         Caption         =   "Service Ride >>"
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
         Left            =   -73920
         TabIndex        =   16
         Top             =   2160
         Width           =   2535
      End
      Begin VB.PictureBox Picture3 
         Height          =   2535
         Left            =   -71280
         Picture         =   "Staff_Mechanic_Landing.frx":CFAF
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Service Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox GentleRideList 
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
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "The following rides are pending for servicing:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton GentleService 
         Caption         =   "Service Ride >>"
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
         Left            =   -73920
         TabIndex        =   11
         Top             =   2160
         Width           =   2535
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Left            =   -71280
         Picture         =   "Staff_Mechanic_Landing.frx":13043
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   10
         Top             =   480
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Service Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox TransportRideList 
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
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label3 
            Caption         =   "The following rides are pending for servicing:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton TransportService 
         Caption         =   "Service Ride >>"
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
         Left            =   -73920
         TabIndex        =   6
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton ThrillService 
         Caption         =   "Service Ride >>"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Service Ride"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3495
         Begin VB.ComboBox ThrillRideList 
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
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "The following rides are pending for servicing:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   2535
         Left            =   3720
         Picture         =   "Staff_Mechanic_Landing.frx":157C8
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   245
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Staff Landing (Ride Servicing)"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Staff_Mechanic_Landing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRide As ADODB.Recordset
Dim connRide As ADODB.Connection
Dim cnt As Integer
Dim sqlquery As String

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Welcome.Show
End Sub

Private Sub GentleRideList_GotFocus()
GentleRideList.Clear
'Get no. of rides pending for servicing from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select count(*) as num from rides where type='gentle' and need_service='1'", connRide

cnt = rsRide.Fields("NUM").Value
rsRide.Close
connRide.Close

If cnt = 0 Then
    MsgBox "There are no rides of this type pending for service!", vbOKOnly, "Entertainment Resort"
Else
'Populate list of rides from SQL DB

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select * from rides where type='gentle' and need_service='1'", connRide

rsRide.MoveFirst
While (rsRide.EOF = False)
    GentleRideList.AddItem (rsRide.Fields("NAME").Value)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close

End If
End Sub

Private Sub GentleService_Click()
If GentleRideList.Text = "" Then
    MsgBox "Please select a ride for servicing!", vbOKOnly, "Entertainment Resort"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "update (select need_service from rides where type='gentle' and name='" & GentleRideList.Text & "') set need_service=0"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the service queue! Visitors can now enjoy this ride again.", vbOKOnly, "Entertainment Resort"
'Insert into log
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "insert into log_staff_mechanic values('" & Global_Module.staff_login & "','" & GentleRideList.Text & "','gentle','" & Format(Global_Module.Today, "dd-mmm-yyyy") & "')"
connRide.Execute sqlquery
connRide.Close

GentleRideList.Clear
End If
End Sub

Private Sub CoasterRideList_GotFocus()
CoasterRideList.Clear
'Get no. of rides pending for servicing from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select count(*) as num from rides where type='coaster' and need_service='1'", connRide

cnt = rsRide.Fields("NUM").Value
rsRide.Close
connRide.Close

If cnt = 0 Then
    MsgBox "There are no rides of this type pending for service!", vbOKOnly, "Entertainment Resort"
Else
'Populate list of rides from SQL DB

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select * from rides where type='coaster' and need_service='1'", connRide

rsRide.MoveFirst
While (rsRide.EOF = False)
    CoasterRideList.AddItem (rsRide.Fields("NAME").Value)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close

End If
End Sub

Private Sub RollerService_Click()
If CoasterRideList.Text = "" Then
    MsgBox "Please select a ride for servicing!", vbOKOnly, "Entertainment Resort"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "update (select need_service from rides where type='coaster' and name='" & CoasterRideList.Text & "') set need_service=0"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the service queue! Visitors can now enjoy this ride again.", vbOKOnly, "Entertainment Resort"
'Insert into log
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "insert into log_staff_mechanic values('" & Global_Module.staff_login & "','" & CoasterRideList.Text & "','coaster','" & Format(Global_Module.Today, "dd-mmm-yyyy") & "')"
connRide.Execute sqlquery
connRide.Close

CoasterRideList.Clear

End If
End Sub

Private Sub ThrillRideList_GotFocus()
ThrillRideList.Clear
'Get no. of rides pending for servicing from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select count(*) as num from rides where type='thrill' and need_service='1'", connRide

cnt = rsRide.Fields("NUM").Value
rsRide.Close
connRide.Close

If cnt = 0 Then
    MsgBox "There are no rides of this type pending for service!", vbOKOnly, "Entertainment Resort"
Else
'Populate list of rides from SQL DB

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select * from rides where type='thrill' and need_service='1'", connRide

rsRide.MoveFirst
While (rsRide.EOF = False)
    ThrillRideList.AddItem (rsRide.Fields("NAME").Value)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close
End If

End Sub

Private Sub ThrillService_Click()
If ThrillRideList.Text = "" Then
    MsgBox "Please select a ride for servicing!", vbOKOnly, "Entertainment Resort"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "update (select need_service from rides where type='thrill' and name='" & ThrillRideList.Text & "') set need_service=0"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the service queue! Visitors can now enjoy this ride again.", vbOKOnly, "Entertainment Resort"
'Insert into log
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "insert into log_staff_mechanic values('" & Global_Module.staff_login & "','" & ThrillRideList.Text & "','thrill','" & Format(Global_Module.Today, "dd-mmm-yyyy") & "')"
connRide.Execute sqlquery
connRide.Close
ThrillRideList.Clear
End If
End Sub

Private Sub TransportRideList_GotFocus()
TransportRideList.Clear
'Get no. of rides pending for servicing from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select count(*) as num from rides where type='transport' and need_service='1'", connRide

cnt = rsRide.Fields("NUM").Value
rsRide.Close
connRide.Close

If cnt = 0 Then
    MsgBox "There are no rides of this type pending for service!", vbOKOnly, "Entertainment Resort"
Else
'Populate list of rides from SQL DB

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select * from rides where type='transport' and need_service='1'", connRide

rsRide.MoveFirst
While (rsRide.EOF = False)
    TransportRideList.AddItem (rsRide.Fields("NAME").Value)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close
End If

End Sub

Private Sub TransportService_Click()
If TransportRideList.Text = "" Then
    MsgBox "Please select a ride for servicing!", vbOKOnly, "Entertainment Resort"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "update (select need_service from rides where type='transport' and name='" & TransportRideList.Text & "') set need_service=0"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the service queue! Visitors can now enjoy this ride again.", vbOKOnly, "Entertainment Resort"
'Insert into log
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "insert into log_staff_mechanic values('" & Global_Module.staff_login & "','" & TransportRideList.Text & "','transport','" & Format(Global_Module.Today, "dd-mmm-yyyy") & "')"
connRide.Execute sqlquery
connRide.Close

TransportRideList.Clear
End If
End Sub

Private Sub WaterRideList_GotFocus()
WaterRideList.Clear
'Get no. of rides pending for servicing from SQL DB
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select count(*) as num from rides where type='water' and need_service='1'", connRide

cnt = rsRide.Fields("NUM").Value
rsRide.Close
connRide.Close

If cnt = 0 Then
    MsgBox "There are no rides of this type pending for service!", vbOKOnly, "Entertainment Resort"
Else
'Populate list of rides from SQL DB

Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
rsRide.Open "select * from rides where type='water' and need_service='1'", connRide

rsRide.MoveFirst
While (rsRide.EOF = False)
    WaterRideList.AddItem (rsRide.Fields("NAME").Value)
    rsRide.MoveNext
Wend
rsRide.Close
connRide.Close

End If

End Sub

Private Sub WaterService_Click()
If WaterRideList.Text = "" Then
    MsgBox "Please select a ride for servicing!", vbOKOnly, "Entertainment Resort"
Else
Set connRide = New ADODB.Connection
Set rsRide = New ADODB.Recordset
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "update (select need_service from rides where type='water' and name='" & WaterRideList.Text & "') set need_service=0"
connRide.Execute sqlquery
connRide.Close
MsgBox "The selected ride has been removed from the service queue! Visitors can now enjoy this ride again.", vbOKOnly, "Entertainment Resort"
'Insert into log
connRide.Open ("Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True")
sqlquery = "insert into log_staff_mechanic values('" & Global_Module.staff_login & "','" & WaterRideList.Text & "','water','" & Format(Global_Module.Today, "dd-mmm-yyyy") & "')"
connRide.Execute sqlquery
connRide.Close

WaterRideList.Clear
End If

End Sub
