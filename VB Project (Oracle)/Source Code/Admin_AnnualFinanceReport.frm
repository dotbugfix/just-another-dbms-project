VERSION 5.00
Begin VB.Form Admin_AnnualFinanceReport 
   Caption         =   "Annual Finance Report"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Admin_AnnualFinanceReport.frx":0000
   ScaleHeight     =   11085
   ScaleWidth      =   11880
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "(All amounts in Rs.)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4200
      TabIndex        =   58
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Perticular"
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
      Left            =   960
      TabIndex        =   57
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ride"
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
      Left            =   3480
      TabIndex        =   56
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
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
      Left            =   4920
      TabIndex        =   55
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Restaurant"
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
      Left            =   6480
      TabIndex        =   54
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   9840
      TabIndex        =   53
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   52
      Top             =   1680
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11070
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   51
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Fee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   50
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   49
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   48
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   47
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Ent. Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   46
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   45
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      Left            =   8520
      TabIndex        =   44
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   43
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Profit (Before Tax)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   42
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   41
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Income Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   40
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   39
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Profit (After Tax)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1125
      TabIndex        =   38
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   960
      Y2              =   7900
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   960
      Y2              =   7900
   End
   Begin VB.Line Line4 
      X1              =   6120
      X2              =   6120
      Y1              =   960
      Y2              =   7900
   End
   Begin VB.Line Line5 
      X1              =   8400
      X2              =   8400
      Y1              =   960
      Y2              =   7900
   End
   Begin VB.Line Line6 
      X1              =   9600
      X2              =   9600
      Y1              =   960
      Y2              =   7900
   End
   Begin VB.Line Line7 
      X1              =   480
      X2              =   11070
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line8 
      X1              =   480
      X2              =   11070
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line9 
      X1              =   480
      X2              =   11070
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line10 
      X1              =   480
      X2              =   11070
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label RevenueRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "AMOUNT"
      DataMember      =   "MonthlyRideRev"
      DataSource      =   "OracleDB"
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
      Left            =   3120
      TabIndex        =   37
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label RevenueShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "REVENUE"
      DataMember      =   "MonthlyShopRevExp"
      DataSource      =   "OracleDB"
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
      Left            =   4680
      TabIndex        =   36
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label RevenueRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "REVENUE"
      DataMember      =   "MonthlyRestRevExp"
      DataSource      =   "OracleDB"
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
      Left            =   6600
      TabIndex        =   35
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label RevenueEntry 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "ENTRY_FEE"
      DataMember      =   "MonthlyEntryFee"
      DataSource      =   "OracleDB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label RevenueTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   33
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label ExpenseRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "AMOUNT"
      DataMember      =   "AnnualRideExp"
      DataSource      =   "OracleDB"
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
      Left            =   3120
      TabIndex        =   32
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label ExpenseShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "CP"
      DataMember      =   "MonthlyShopRevExp"
      DataSource      =   "OracleDB"
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
      Left            =   4680
      TabIndex        =   31
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label ExpenseRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "CP"
      DataMember      =   "MonthlyRestRevExp"
      DataSource      =   "OracleDB"
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
      Left            =   6600
      TabIndex        =   30
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label ExpenseSalary 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      DataField       =   "AMOUNT"
      DataMember      =   "AnnualSalary"
      DataSource      =   "OracleDB"
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
      Left            =   8280
      TabIndex        =   29
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label ExpenseEntTax 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   8280
      TabIndex        =   28
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label ExpenseTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label ProfitBeforeTax 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   26
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label IncomeTax 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   25
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label ProfitAfterTax 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   24
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label ExpenseSerTaxRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label ExpenseSerTaxShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label ExpenseSerTaxRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label ProfitBeforeTaxRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   3120
      TabIndex        =   20
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label ProfitBeforeTaxShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   4680
      TabIndex        =   19
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label ProfitBeforeTaxRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   6600
      TabIndex        =   18
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label PercentProfit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
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
      Left            =   9600
      TabIndex        =   17
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contribution to Profit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   16
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "7."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label PercentProfitRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   3120
      TabIndex        =   14
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label PercentProfitShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label PercentProfitRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   6600
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Line Line11 
      X1              =   480
      X2              =   11070
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label ProfitabilityRest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   6600
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label ProfitabilityShop 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label ProfitabilityRide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "6."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Profitability"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Profitability 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(amt)"
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
      Left            =   9600
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Line Line12 
      X1              =   480
      X2              =   11100
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Month 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year:"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Year 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label NoteSerTax 
      BackStyle       =   0  'Transparent
      Caption         =   "(note sertax)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   8880
      Width           =   7575
   End
   Begin VB.Label NoteEntTax 
      BackStyle       =   0  'Transparent
      Caption         =   "(note enttax)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   8520
      Width           =   7935
   End
   Begin VB.Label NoteIncomeTax 
      BackStyle       =   0  'Transparent
      Caption         =   "(note incometax)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   9240
      Width           =   7815
   End
End
Attribute VB_Name = "Admin_AnnualFinanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me
.Top = 100
.Left = 3000
.Height = 10380
.Width = 11655
End With

Year.Caption = Admin_AnnualFinanceLanding.YearBox.Text & " - " & Val(Admin_AnnualFinanceLanding.YearBox.Text) + 1

RevenueTotal.Caption = Format((Val(RevenueRide.Caption) + Val(RevenueShop.Caption) + Val(RevenueRest.Caption) + Val(RevenueEntry.Caption)), "0.00")

ExpenseEntTax.Caption = Format((Val(RevenueEntry.Caption) / (1 + (Global_Module.Ent_Tax / 100))) * (Global_Module.Ent_Tax / 100), "00.00")
ExpenseSerTaxRide.Caption = Format((Val(RevenueRide.Caption) / (1 + (Global_Module.Ser_Tax / 100))) * (Global_Module.Ser_Tax / 100), "00.00")
ExpenseSerTaxShop.Caption = Format((Val(RevenueShop.Caption) / (1 + (Global_Module.Ser_Tax / 100))) * (Global_Module.Ser_Tax / 100), "00.00")
ExpenseSerTaxRest.Caption = Format((Val(RevenueRest.Caption) / (1 + (Global_Module.Ser_Tax / 100))) * (Global_Module.Ser_Tax / 100), "00.00")
ExpenseTotal.Caption = Format((Val(ExpenseRide.Caption) + Val(ExpenseShop.Caption) + Val(ExpenseRest.Caption) + Val(ExpenseSalary.Caption) + Val(ExpenseEntTax.Caption) + Val(ExpenseSerTaxRide.Caption) + Val(ExpenseSerTaxShop.Caption) + Val(ExpenseSerTaxRest.Caption)), "0.00")

ProfitBeforeTax.Caption = Format((Val(RevenueTotal.Caption) - Val(ExpenseTotal.Caption)), "0.00")
ProfitBeforeTaxRide.Caption = Format((Val(RevenueRide.Caption) - Val(ExpenseRide.Caption) - Val(ExpenseSerTaxRide.Caption)), "0.00")
ProfitBeforeTaxShop.Caption = Format((Val(RevenueShop.Caption) - Val(ExpenseShop.Caption) - Val(ExpenseSerTaxShop.Caption)), "0.00")
ProfitBeforeTaxRest.Caption = Format((Val(RevenueRest.Caption) - Val(ExpenseRest.Caption) - Val(ExpenseSerTaxRest.Caption)), "0.00")
ProfitWithoutOtherExp = Val(ProfitBeforeTaxRide.Caption) + Val(ProfitBeforeTaxShop.Caption) + Val(ProfitBeforeTaxRest.Caption)

IncomeTax.Caption = Format(((Global_Module.Income_Tax / 100) * Abs(Val(ProfitBeforeTax.Caption))), "0.00")
IncomeTaxWithoutOtherExp = Format(((Global_Module.Income_Tax / 100) * ProfitWithoutOtherExp), "0.00")

ProfitAfterTax.Caption = Format((Val(ProfitBeforeTax.Caption) - Val(IncomeTax.Caption)), "0.00")

Profitability.Caption = Format(((Val(ProfitAfterTax.Caption) / Val(RevenueTotal.Caption)) * 100), "0.00") & "%"
ProfitabilityRide.Caption = Format(((Val(ProfitBeforeTaxRide.Caption) / Val(RevenueRide.Caption)) * 100), "0.00") & "%"
ProfitabilityShop.Caption = Format(((Val(ProfitBeforeTaxShop.Caption) / Val(RevenueShop.Caption)) * 100), "0.00") & "%"
ProfitabilityRest.Caption = Format(((Val(ProfitBeforeTaxRest.Caption) / Val(RevenueRest.Caption)) * 100), "0.00") & "%"

PercentProfitRide.Caption = Format(((Val(ProfitBeforeTaxRide.Caption) / ProfitWithoutOtherExp) * 100), "0.00") & "%"
PercentProfitShop.Caption = Format(((Val(ProfitBeforeTaxShop.Caption) / ProfitWithoutOtherExp) * 100), "0.00") & "%"
PercentProfitRest.Caption = Format(((Val(ProfitBeforeTaxRest.Caption) / ProfitWithoutOtherExp) * 100), "0.00") & "%"

NoteEntTax.Caption = "Entertainment Tax @ " & Global_Module.Ent_Tax & "% paid directly by visitors at the time of entry."
NoteSerTax.Caption = "Service Tax @ " & Global_Module.Ser_Tax & "% paid directly by visitors for all purchases."
NoteIncomeTax.Caption = "Income Tax @ " & Global_Module.Income_Tax & "% paid by the resort on total profit earned."
End Sub

Private Sub Form_Unload(Cancel As Integer)
With Admin_Landing
    .WindowState = vbNormal
    .Top = 1935
    .Left = 3390
End With

End Sub

