VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Admin_Landing 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administrator Landing"
   ClientHeight    =   7635
   ClientLeft      =   2835
   ClientTop       =   2280
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Picture         =   "admin_Landing.frx":0000
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   10440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "1/13/2011"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:38 PM"
         EndProperty
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
   End
   Begin VB.Menu Menu_db 
      Caption         =   "&Database"
      Index           =   1
      Begin VB.Menu menu_Ride 
         Caption         =   "Ride"
         Index           =   1
         Begin VB.Menu new_Ride 
            Caption         =   "Add Ride"
            Index           =   1
            Shortcut        =   ^R
         End
         Begin VB.Menu remove_ride 
            Caption         =   "Remove Ride"
            Index           =   2
         End
         Begin VB.Menu service_ride 
            Caption         =   "Servicing..."
            Index           =   3
         End
      End
      Begin VB.Menu menu_Staff 
         Caption         =   "Staff"
         Index           =   2
         Begin VB.Menu staff_Appoint 
            Caption         =   "Appoint Staff"
            Index           =   1
         End
         Begin VB.Menu staff_Fire 
            Caption         =   "Fire Staff"
            Index           =   2
         End
      End
      Begin VB.Menu menu_Shop 
         Caption         =   "Shops && Restaurants"
         Index           =   3
         Begin VB.Menu shop_additem 
            Caption         =   "Add Item"
            Index           =   1
         End
         Begin VB.Menu shop_remove 
            Caption         =   "Remove Item"
            Index           =   2
         End
      End
      Begin VB.Menu db_adminPass 
         Caption         =   "Change Admin Password"
         Index           =   4
      End
      Begin VB.Menu menu_print 
         Caption         =   "Print..."
         Index           =   5
      End
   End
   Begin VB.Menu menu_Statistics 
      Caption         =   "&Statistics"
      Index           =   2
      Begin VB.Menu stat_visitor 
         Caption         =   "Visitors"
         Index           =   1
      End
      Begin VB.Menu statRide 
         Caption         =   "Rides"
         Index           =   2
      End
      Begin VB.Menu stat_shop_res 
         Caption         =   "Shops && Restaurants"
         Index           =   3
      End
      Begin VB.Menu stat_Staff 
         Caption         =   "Staff"
         Index           =   4
      End
   End
   Begin VB.Menu menu_History 
      Caption         =   "&History"
      Index           =   3
      Begin VB.Menu stat_PerRide 
         Caption         =   "Per Ride History"
         Index           =   1
      End
      Begin VB.Menu stat_PerVisitor 
         Caption         =   "Per Visitor History"
         Index           =   2
      End
      Begin VB.Menu stat_PerStaff 
         Caption         =   "Per Staff History"
         Index           =   3
      End
      Begin VB.Menu history_resortLog 
         Caption         =   "Resort Log"
         Index           =   4
      End
      Begin VB.Menu history_ServicingLog 
         Caption         =   "Servicing Log"
         Index           =   5
      End
   End
   Begin VB.Menu menu_reports 
      Caption         =   "R&eports"
      Index           =   4
      Begin VB.Menu rep_monthly 
         Caption         =   "Monthly Reports"
         Index           =   1
         Begin VB.Menu rep_mon_visitor 
            Caption         =   "Visitors"
            Index           =   1
         End
         Begin VB.Menu rep_mon_rides 
            Caption         =   "Rides"
            Index           =   2
         End
      End
      Begin VB.Menu rep_daily 
         Caption         =   "Daily Reports"
         Index           =   2
         Begin VB.Menu rep_day_visitors 
            Caption         =   "Visitors"
            Index           =   1
         End
         Begin VB.Menu rep_day_rides 
            Caption         =   "Rides"
            Index           =   2
         End
      End
   End
   Begin VB.Menu finance 
      Caption         =   "&Finance"
      Index           =   5
      Begin VB.Menu Finance_Monthly 
         Caption         =   "Monthly Report"
         Index           =   1
      End
      Begin VB.Menu finance_annual 
         Caption         =   "Annual Report"
         Index           =   2
      End
      Begin VB.Menu finance_EntryFee 
         Caption         =   "Change Entry Fees"
         Index           =   3
      End
      Begin VB.Menu menu_ChangeEntTax 
         Caption         =   "Change Taxes"
         Index           =   4
      End
   End
   Begin VB.Menu menu_Date 
      Caption         =   "&Realtime Settings"
      Index           =   6
      Begin VB.Menu menu_ChangeDate 
         Caption         =   "Change Current Date"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu menu_VisitorLimit 
         Caption         =   "Change Visitor Limit"
         Index           =   2
      End
   End
   Begin VB.Menu menu_logout 
      Caption         =   "&Logout"
      Index           =   7
   End
End
Attribute VB_Name = "Admin_Landing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub db_adminPass_Click(Index As Integer)
Admin_ChangeAdminPass.Show
End Sub

Private Sub finance_annual_Click(Index As Integer)
Admin_AnnualFinanceLanding.Show
End Sub

Private Sub finance_EntryFee_Click(Index As Integer)
Admin_ChangeEntryFees.Show
End Sub

Private Sub Finance_Monthly_Click(Index As Integer)
Admin_MonthlyFinanceLanding.Show
End Sub

Private Sub history_resortLog_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
With Me
   .WindowState = vbMaximized
End With
    
OracleDB.ResortLog

Report_ResortLog.Orientation = rptOrientLandscape
Report_ResortLog.Show
End Sub

Private Sub history_ServicingLog_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
OracleDB.ServicingLog

Report_ServicingLog.Orientation = rptOrientLandscape
Report_ServicingLog.Show
End Sub

Private Sub MDIForm_Load()
'    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
'    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    'LoadNewDoc
End Sub


Private Sub LoadNewDoc()
    'Static lDocumentCount As Long
    'Dim frmD As DataReport1
    'lDocumentCount = lDocumentCount + 1
    'Set frmD = New DataReport1
    'frmD.Caption = "Document " & lDocumentCount
    'frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Welcome.Show
End Sub

Private Sub menu_ChangeDate_Click(Index As Integer)
MsgBox "Use this function to test various validation criteria of the resort such as advance booking, staff salary slips etc. Refer the documentation for further details.", vbInformation, "Demonstration"
Admin_ChangeDate.Show
End Sub

Private Sub menu_ChangeEntTax_Click(Index As Integer)
Admin_ChangeEntTax.Show
End Sub

Private Sub menu_Logout_Click(Index As Integer)
Unload Me
Welcome.Show
End Sub

Private Sub menu_print_Click(Index As Integer)

MsgBox "This function prints the active MDI Child that is open in this form.", vbInformation, "Demonstration"
 
 On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
'        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .Flags = .Flags + cdlPDAllPages
'        Else
'            .Flags = .Flags + cdlPDSelection
'        End If
        .ShowPrinter
'        If Err <> MSComDlg.cdlCancel Then
'            ActiveForm.rtfText.SelPrint .hDC
'        End If
    End With
End Sub

Private Sub menu_VisitorLimit_Click(Index As Integer)
Admin_ChangeVisitorLimit.Show
End Sub

Private Sub new_Ride_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_NewRide.Show
End Sub

Private Sub remove_ride_Click(Index As Integer)
MsgBox "Rides can be removed from the resort by using the Ride Statistics table view", vbOKOnly, "Admin Landing"
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.ThrillRideStat
OracleDB.CoasterRideStat
OracleDB.WaterRideStat
OracleDB.TransportRideStat
OracleDB.GentleRideStat

Admin_StatRide.Show
End Sub

Private Sub rep_day_rides_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
Admin_DailyRides.Show
End Sub

Private Sub rep_day_visitors_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_DailyVisitor.Show
End Sub

Private Sub rep_mon_rides_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_MonthlyRides.Show
End Sub

Private Sub rep_mon_visitor_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_MonthlyVisitor.Show
End Sub

Private Sub service_ride_Click(Index As Integer)
Admin_RideServicing.Show
End Sub

Private Sub shop_additem_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_ShopAddItem.Show
End Sub

Private Sub shop_remove_Click(Index As Integer)
MsgBox "Items can be removed from the resort by using the Shops & Restaurants Statistics table view", vbOKOnly, "Admin Landing"
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_StatShopRest.Show

End Sub

Private Sub staff_Appoint_Click(Index As Integer)
Admin_StaffAppoint.Show
End Sub

Private Sub staff_Fire_Click(Index As Integer)
MsgBox "Staff can be removed from their jobs in the resort by using the Staff Statistics table view", vbOKOnly, "Admin Landing"
Admin_StatStaff.Show
End Sub

Private Sub stat_PerRide_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
Admin_PerRide.Show
End Sub

Private Sub stat_PerStaff_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
Admin_PerStaff.Show
End Sub

Private Sub stat_PerVisitor_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

Admin_PerVisitor.Show
End Sub

Private Sub stat_shop_res_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
'OracleDB.CountItems
'OracleDB.Count_Menu
    
Admin_StatShopRest.Show
End Sub

Private Sub stat_Staff_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
    
OracleDB.StaffStats
    
Admin_StatStaff.Show
End Sub

Private Sub stat_visitor_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open
OracleDB.CountID

With Me
   .WindowState = vbMaximized
End With

Admin_StatVisitor.Show
End Sub

Private Sub statRide_Click(Index As Integer)
If Not OracleDB.OracleProvider.State = closed Then
    OracleDB.OracleProvider.Close
End If
    OracleDB.OracleProvider.Open

OracleDB.ThrillRideStat
OracleDB.CoasterRideStat
OracleDB.WaterRideStat
OracleDB.TransportRideStat
OracleDB.GentleRideStat

Admin_StatRide.Show
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub
Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
   

End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub
