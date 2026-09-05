VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form rptSalesTrend 
   Caption         =   "Sales Trend"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGraph 
      Caption         =   "Graph"
      Height          =   3570
      Left            =   840
      TabIndex        =   42
      Top             =   645
      Visible         =   0   'False
      Width           =   1290
Begin AceSpread fpsAxis_0
         Height          =   1365
         Left            =   7425
         TabIndex        =   43
         Top             =   6480
         Width           =   4275
         _Version        =   458752
         _ExtentX        =   7541
         _ExtentY        =   2408
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   1
      End
      Begin VB.Frame fraDevision 
         Caption         =   "Grid Division"
         Height          =   585
         Left            =   90
         TabIndex        =   51
         Top             =   6465
         Width           =   2580
         Begin VB.TextBox txtRowDiv 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   720
            TabIndex        =   53
            Text            =   "14"
            Top             =   210
            Width           =   450
         End
         Begin VB.TextBox txtColDiv 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2070
            TabIndex        =   52
            Text            =   "20"
            Top             =   210
            Width           =   450
         End
         Begin VB.Label lblRowDiv 
            Caption         =   "Row (Y)"
            Height          =   195
            Left            =   60
            TabIndex        =   55
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblColDiv 
            Caption         =   "Column (X)"
            Height          =   195
            Left            =   1230
            TabIndex        =   54
            Top             =   255
            Width           =   855
         End
      End
      Begin VB.Frame fraAxis 
         Caption         =   "Y - Axis"
         Height          =   570
         Left            =   75
         TabIndex        =   44
         Top             =   210
         Width           =   7275
         Begin VB.TextBox txtStart 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   210
            Width           =   1500
         End
         Begin VB.TextBox txtEnd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   210
            Width           =   1500
         End
         Begin VB.TextBox txtTarget 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5715
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   210
            Width           =   1500
         End
         Begin VB.Label lblStart 
            Caption         =   "Start Value"
            Height          =   195
            Left            =   75
            TabIndex        =   50
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lblEnd 
            Caption         =   "End Value"
            Height          =   195
            Left            =   2460
            TabIndex        =   49
            Top             =   255
            Width           =   750
         End
         Begin VB.Label lblTarget 
            Caption         =   "Targe Value"
            Height          =   195
            Left            =   4785
            TabIndex        =   48
            Top             =   270
            Width           =   930
         End
      End
      Begin TabDlg.SSTab sstGraph 
         Height          =   5640
         Left            =   90
         TabIndex        =   63
         Top             =   810
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   9948
         _Version        =   393216
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Sales Trend"
         TabPicture(0)   =   "rptSalesTrend.frx":0425
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pctGraph(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "rptSalesTrend.frx":0441
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "rptSalesTrend.frx":045D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.CommandButton Command1 
            Caption         =   "&Print"
            Height          =   315
            Left            =   -65670
            TabIndex        =   72
            Top             =   6300
            Width           =   960
         End
         Begin VB.Frame Criteria 
            Caption         =   "Criteria"
            Height          =   1830
            Left            =   -71400
            TabIndex        =   70
            Top             =   1890
            Width           =   4200
Begin AceSpread fpsQty
               Height          =   1545
               Left            =   75
               TabIndex        =   71
               ToolTipText     =   "Enter Acceptance Creteria Details"
               Top             =   225
               Width           =   4050
               _Version        =   458752
               _ExtentX        =   7144
               _ExtentY        =   2725
               _StockProps     =   64
               AutoClipboard   =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   5
               MaxRows         =   1
               ScrollBars      =   2
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Down Time Reason"
            Height          =   1830
            Left            =   -72135
            TabIndex        =   69
            Top             =   3705
            Width           =   5835
         End
         Begin VB.PictureBox pctGraph 
            Height          =   5190
            Index           =   0
            Left            =   75
            ScaleHeight     =   5130
            ScaleWidth      =   11400
            TabIndex        =   68
            Top             =   375
            Width           =   11460
            Begin MSChart20Lib.MSChart mscGraph 
               Height          =   5250
               Index           =   0
               Left            =   240
               OleObjectBlob   =   "rptSalesTrend.frx":0A21
               TabIndex        =   76
               Top             =   -180
               Width           =   11310
            End
         End
         Begin VB.PictureBox pctGraph2 
            Height          =   4695
            Left            =   -74940
            ScaleHeight     =   4635
            ScaleWidth      =   11445
            TabIndex        =   66
            Top             =   360
            Width           =   11505
            Begin MSChart20Lib.MSChart mscGraph2 
               Height          =   5025
               Left            =   -165
               OleObjectBlob   =   "rptSalesTrend.frx":4444
               TabIndex        =   67
               Top             =   -180
               Width           =   11865
            End
         End
         Begin VB.Frame fraGPaynter 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4635
            Left            =   -74880
            TabIndex        =   64
            Top             =   345
            Width           =   11370
Begin AceSpread fpsGPaynter
               Height          =   4380
               Left            =   90
               TabIndex        =   65
               ToolTipText     =   "Enter the Item Details"
               Top             =   180
               Width           =   10830
               _Version        =   458752
               _ExtentX        =   19103
               _ExtentY        =   7726
               _StockProps     =   64
               AllowMultiBlocks=   -1  'True
               ColsFrozen      =   2
               DisplayColHeaders=   0   'False
               DisplayRowHeaders=   0   'False
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   14
               MaxRows         =   15
               OperationMode   =   1
               SelectBlockOptions=   8
            End
            Begin VB.Shape shpGPaynter 
               BackColor       =   &H00E9C5AF&
               BackStyle       =   1  'Opaque
               FillColor       =   &H000000FF&
               Height          =   4380
               Left            =   10935
               Top             =   180
               Width           =   330
            End
         End
         Begin VB.OLE OLEChart 
            AutoActivate    =   3  'Automatic
            Class           =   "Excel.Chart.8"
            Height          =   1530
            Left            =   -70335
            OleObjectBlob   =   "rptSalesTrend.frx":A367
            TabIndex        =   73
            Top             =   2745
            Width           =   2280
         End
      End
      Begin VB.Frame fraRGGC 
         Height          =   585
         Left            =   90
         TabIndex        =   58
         Top             =   6990
         Width           =   5250
         Begin VB.CommandButton cmdGraphClose 
            Caption         =   "&Close"
            Height          =   390
            Left            =   3915
            TabIndex        =   59
            Top             =   150
            Width           =   1290
         End
         Begin VB.CommandButton cmdGraphPreview 
            Caption         =   "Graph Pre&view "
            Height          =   390
            Left            =   2625
            TabIndex        =   60
            Top             =   150
            Width           =   1300
         End
         Begin VB.CommandButton cmdGraphPrint 
            Caption         =   "&Graph Print"
            Height          =   390
            Left            =   1335
            TabIndex        =   61
            Top             =   150
            Width           =   1300
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   390
            Left            =   45
            TabIndex        =   62
            Top             =   150
            Width           =   1300
         End
      End
      Begin VB.ListBox lstYAxis 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         ItemData        =   "rptSalesTrend.frx":1057F
         Left            =   2745
         List            =   "rptSalesTrend.frx":10581
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   6645
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.ListBox lstYAxis 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         ItemData        =   "rptSalesTrend.frx":10583
         Left            =   2700
         List            =   "rptSalesTrend.frx":10585
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   6570
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame fraTrend 
      Height          =   4440
      Index           =   0
      Left            =   -60
      TabIndex        =   12
      Top             =   525
      Visible         =   0   'False
      Width           =   840
Begin AceSpread fpsTrend_0
         Height          =   5280
         Left            =   90
         TabIndex        =   13
         ToolTipText     =   "Enter the Item Details"
         Top             =   180
         Width           =   9960
         _Version        =   458752
         _ExtentX        =   17568
         _ExtentY        =   9313
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   255
         MaxCols         =   13
         MaxRows         =   19
         SelectBlockOptions=   0
      End
      Begin VB.Frame fraPC 
         Height          =   630
         Index           =   0
         Left            =   6030
         TabIndex        =   14
         Top             =   5460
         Width           =   4005
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   400
            Index           =   0
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Click to Exit from the Current Screen"
            Top             =   165
            Width           =   1300
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   400
            Index           =   0
            Left            =   1350
            TabIndex        =   15
            ToolTipText     =   "Click to Clear the Data"
            Top             =   165
            Width           =   1300
         End
         Begin VB.CommandButton cmdShowGraph 
            Caption         =   "Show &Graph"
            Height          =   405
            Index           =   0
            Left            =   60
            TabIndex        =   25
            Top             =   165
            Width           =   1300
         End
      End
      Begin VB.Frame fraOrient 
         Caption         =   "Print Orientation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   4575
         TabIndex        =   17
         Top             =   5460
         Width           =   1395
         Begin VB.ComboBox cmbOrient 
            Height          =   330
            Index           =   0
            ItemData        =   "rptSalesTrend.frx":113FA
            Left            =   75
            List            =   "rptSalesTrend.frx":11404
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   210
            Width           =   1245
         End
      End
      Begin VB.Frame fraOrient 
         Caption         =   "If need,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   1
         Left            =   2895
         TabIndex        =   74
         Top             =   5460
         Width           =   1620
         Begin VB.CheckBox chkPrintedDate 
            Caption         =   "&Printed Date"
            Height          =   420
            Left            =   375
            TabIndex        =   75
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkFullView 
         Caption         =   "Full View"
         Height          =   420
         Left            =   2910
         TabIndex        =   24
         Top             =   5655
         Width           =   1095
      End
      Begin VB.Label lblTrend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Record(s) :  21 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   5445
         Width           =   1710
      End
   End
   Begin VB.Frame fraGroup 
      Caption         =   "[ Sales Trend Monthwise Analysis ]"
      Height          =   2505
      Left            =   4320
      TabIndex        =   26
      Top             =   480
      Width           =   4125
      Begin VB.CommandButton CmdStep 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   3540
         TabIndex        =   79
         Top             =   1905
         Width           =   345
      End
      Begin VB.Frame FraStep 
         Caption         =   "Month"
         Height          =   900
         Index           =   0
         Left            =   210
         TabIndex        =   29
         Top             =   225
         Width           =   3285
         Begin VB.CheckBox chkMonth 
            Caption         =   "12 Months"
            Height          =   210
            Left            =   1125
            TabIndex        =   80
            Top             =   195
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin MSComCtl2.DTPicker dtpTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   2010
            TabIndex        =   30
            Top             =   495
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMM-yyyy"
            Format          =   103677955
            CurrentDate     =   37273
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   510
            TabIndex        =   31
            Top             =   495
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMM-yyyy"
            Format          =   103677955
            CurrentDate     =   37273
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   210
            Left            =   105
            TabIndex        =   33
            Top             =   540
            Width           =   360
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   210
            Left            =   1800
            TabIndex        =   32
            Top             =   540
            Width           =   180
         End
      End
      Begin VB.CommandButton CmdStep 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   3525
         TabIndex        =   28
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton CmdStep 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3525
         TabIndex        =   27
         Top             =   1350
         Width           =   345
      End
      Begin VB.Frame FraStep 
         Caption         =   "Chart Type"
         Enabled         =   0   'False
         Height          =   660
         Index           =   1
         Left            =   225
         TabIndex        =   77
         Top             =   1080
         Width           =   3285
         Begin VB.ComboBox cmbChartType 
            Height          =   330
            ItemData        =   "rptSalesTrend.frx":1141D
            Left            =   915
            List            =   "rptSalesTrend.frx":11445
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   225
            Width           =   1515
         End
      End
      Begin VB.Frame FraStep 
         Caption         =   "Customer"
         Enabled         =   0   'False
         Height          =   570
         Index           =   2
         Left            =   210
         TabIndex        =   34
         Top             =   1725
         Width           =   3300
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   195
            Left            =   465
            TabIndex        =   36
            Top             =   255
            Value           =   -1  'True
            Width           =   570
         End
         Begin VB.OptionButton optSelective 
            Caption         =   "Selective"
            Height          =   255
            Left            =   2040
            TabIndex        =   35
            Top             =   225
            Width           =   960
         End
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "Select Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   4290
      TabIndex        =   0
      Top             =   4635
      Visible         =   0   'False
      Width           =   4695
Begin AceSpread fpsList
         Height          =   1950
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   4560
         _Version        =   458752
         _ExtentX        =   8043
         _ExtentY        =   3440
         _StockProps     =   64
         DAutoSizeCols   =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         OperationMode   =   4
         RowHeaderDisplay=   0
         ScrollBars      =   2
      End
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   60
         TabIndex        =   2
         Top             =   2070
         Width           =   4575
         Begin VB.Frame fraClearSelection 
            Height          =   465
            Left            =   3165
            TabIndex        =   3
            Top             =   0
            Width           =   1410
            Begin VB.CommandButton cmdClearSelection 
               Caption         =   "Clear S&election"
               Height          =   270
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Click to OK"
               Top             =   150
               Width           =   1290
            End
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total && Selected Record(s) :  0  of  21 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   60
            TabIndex        =   5
            Top             =   165
            Width           =   3075
         End
      End
      Begin VB.Frame Frame6 
         Height          =   555
         Left            =   60
         TabIndex        =   6
         Top             =   2430
         Width           =   4575
         Begin VB.TextBox txtFind 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1245
            TabIndex        =   7
            ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
            Top             =   180
            Width           =   3255
         End
         Begin VB.Label lblFind 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Customer Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   8
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.Frame fraFind 
         Height          =   420
         Left            =   60
         TabIndex        =   9
         Top             =   2880
         Width           =   4575
         Begin VB.CheckBox chkPMAllowed 
            Caption         =   "Partial Matches Allowed"
            Height          =   210
            Left            =   2505
            TabIndex        =   11
            Top             =   165
            Value           =   1  'Checked
            Width           =   2040
         End
         Begin VB.CheckBox chkSRow 
            Caption         =   "Search Row"
            Height          =   210
            Left            =   60
            TabIndex        =   10
            Top             =   165
            Value           =   1  'Checked
            Width           =   1260
         End
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4920
      TabIndex        =   37
      Top             =   2985
      Width           =   2925
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptSalesTrend.frx":117F6
         Left            =   810
         List            =   "rptSalesTrend.frx":117FD
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   225
         Width           =   1995
      End
      Begin VB.ComboBox cmbOrderby 
         Height          =   330
         ItemData        =   "rptSalesTrend.frx":11817
         Left            =   810
         List            =   "rptSalesTrend.frx":1181E
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   615
         Width           =   1995
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   210
         Left            =   405
         TabIndex        =   41
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblOrderby 
         AutoSize        =   -1  'True
         Caption         =   "Order by"
         Height          =   210
         Left            =   135
         TabIndex        =   40
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame Frame4 
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4380
      TabIndex        =   20
      Top             =   4005
      Width           =   3990
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   2625
         TabIndex        =   21
         Top             =   165
         Width           =   1300
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   1335
         TabIndex        =   22
         Top             =   165
         Width           =   1300
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Enabled         =   0   'False
         Height          =   400
         Left            =   45
         TabIndex        =   23
         Top             =   165
         Width           =   1300
      End
   End
End
Attribute VB_Name = "rptSalesTrend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim d3 As Byte

'Section 1 : ( Main Control(s) ) :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurrDate As Date
  
  Try
  CurrDate = Format(ServerDate, "MMM-yyyy")
  
  chkMonth.Value = 1
  DTPFrom.Value = DateAdd("m", -11, Format(CurrDate, "MMM-yyyy"))
  DTPTo.Value = Format(CurrDate, "MMM-yyyy")

  d3 = 0
  If mDocSeq(1) = 1 Then
    rptSalesTrend.Caption = "Sales Trend Monthwise Analysis"
    fraGroup.Caption = "[ Sales Trend Monthwise Analysis ]"
  ElseIf mDocSeq(1) = 2 Then
    rptSalesTrend.Caption = "Sales Return Monthwise Analysis"
    fraGroup.Caption = "[ Sales Return Monthwise Analysis ]"
  ElseIf mDocSeq(1) = 3 Then
    rptSalesTrend.Caption = "New Products Monthwise Analysis"
    fraGroup.Caption = "[ New Products Monthwise Analysis ]"
  Else
    'Skip
  End If
  
  cmbType.ListIndex = 0
  cmbOrderby.ListIndex = 0
  cmbChartType.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub chkMonth_Click()
  Dim CurrDate As Date
  
' On Error Resume Next
  
  If chkMonth.Value = 1 Then
    CurrDate = Format(ServerDate, "MMM-yyyy")
    
    DTPFrom.Value = DateAdd("m", -11, Format(CurrDate, "MMM-yyyy"))
    DTPTo.Value = Format(CurrDate, "MMM-yyyy")
  End If
End Sub

Private Sub chkFullView_Click()

  Try

  If chkFullView = 0 Then
    Call Spread_HideCol(fpsTrend(d3), True, 4, 7)
  Else
    Call Spread_HideCol(fpsTrend(d3), False, 4, 7)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub chkPrintedDate_Click()
' On Error Resume Next
  
  If chkPrintedDate.Value = 1 Then
    fpsTrend(d3).SetText 2, 3, "Printed Date : " & Format(Trim$(ServerDate), "dd/mm/yyyy")
    Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignRight, 2, 3, 2, 3)
  Else
    fpsTrend(d3).SetText 2, 3, ""
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  
  cmbType.ListIndex = 0
  cmbOrderby.ListIndex = 0
  fraList.Visible = False
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
 
  Try
  Me.MousePointer = vbHourglass
  
'  CmdStep(1).Enabled = False

  If mDocSeq(1) = 1 Or mDocSeq(1) = 2 Then
  
    Sql = "Select Distinct a.CustomerCode [Customer Code],  b. PartyName As [Customer Name]"
    Sql = Sql & " From SalesInvoice a"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
    Sql = Sql & " On a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid"
    Sql = Sql & " and a.Compid=b.Compid"
    
    If mDocSeq(1) = 1 Then
 Sql = Sql & " Where Convert(VarChar(7),a.InvoiceDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
End If
    If mDocSeq(1) = 2 Then
 Sql = Sql & " Where Convert(VarChar(7),a.OrderDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
End If
    
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"   'and a.Finyear='" & gFinyear & "'"
          
  ElseIf mDocSeq(1) = 3 Then
    GoTo ExitHere
  Else
    GoTo ExitHere
  End If
 
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  fpsList.MaxRows = fpsList.DataRowCnt
'
'  If fpsList.DataRowCnt <= 0 Then
'    sShowMessage "Record(s) Not Found"
'    OptAll.Value = True
'    GoTo ExitHere
'  Else
'  'For Search
'    Call Set_SearchTextOption(1)
'    fraList.Visible = True
'    Call fpsList_Click(1, 0)
'  End If

 ' fpsList.ClearSelection
' cmdClearSelection.Enabled = False
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    OptAll.Value = True
    GoTo ExitHere
  Else
    Call Set_SearchTextOption(1)
    fpsList.Enabled = True
 fpsList.Visible = True
    fpsList.ClearSelection
    fraList.Visible = True
    Call fpsList_Click(1, 2)
    Call ADORS_Spread_DataFill(rs, fpsList, True)
    
    fpsList.MaxRows = fpsList.DataRowCnt
    Call Spread_Set_Focus(fpsList, 1, 2)
     
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpFrom_Change()
' On Error Resume Next
  
  If chkMonth.Value = 1 Then
    DTPTo.Value = DateAdd("m", 11, Format(DTPFrom.Value, "MMM-yyyy"))
  Else
    If Format(DTPFrom.Value, "yyyy-MM") <= Format(DTPTo.Value, "yyyy-MM") Then
      'Skip
    Else
      sShowMessage "'From Month' is less than or equal to 'To Month'"
      DTPFrom.Value = DTPTo.Value
    End If
  End If
End Sub

Private Sub dtpTo_Change()
' On Error Resume Next
  
  If chkMonth.Value = 1 Then
    DTPFrom.Value = DateAdd("m", -11, Format(DTPTo.Value, "MMM-yyyy"))
  Else
    If Format(DTPTo.Value, "yyyy-MM") >= Format(DTPFrom.Value, "yyyy-MM") Then
      'Skip
    Else
      sShowMessage "'To Month' is greater than or equal to 'From Month'"
      DTPTo.Value = DTPFrom.Value
    End If
  End If
End Sub

Private Sub fpsList_Click(sender as object, e as ClickEventArgs)
  Dim i As Integer
  Dim retVal As Variant
  
' On Error Resume Next
'  If optItemAll.Value <> True Then
'    optItemAll.Value = True
'    fraList.Left = 3630
'  End If
  
  If Row = 0 And Col > 0 Then
    fpsList.Tag = Col
      fpsList.Row = 0
fpsList.Row2 = 0
      
      fpsList.BlockMode = True
      For i = 1 To fpsList.MaxCols
        fpsList.Col = i
fpsList.Col2 = i
fpsList.GetText i, 0, retVal
        fpsList.CellType = CellTypeButton
        fpsList.TypeButtonText = Trim$(retVal)
        fpsList.TypeButtonTextColor = RGB(0, 0, 0)
      Next i
      
      fpsList.Col = Col
fpsList.Col2 = Col
      fpsList.CellType = CellTypeButton
      If mDocSeq(4) = 2 Then
        lblFind.Caption = IIf(Col = 1, "Doc. Type", IIf(Col = 2, "Doc. No.", IIf(Col = 3, "Doc. Date", "Revision No.")))
      ElseIf mDocSeq(4) = 3 Then
        lblFind.Caption = IIf(Col = 1, "Doc. Type", IIf(Col = 2, "Doc. No.", IIf(Col = 3, "Doc. Date", IIf(Col = 4, "User ID", "User Name"))))
      Else
        lblFind.Caption = IIf(Col = 1, "Customer Code", "Customer Name")
      End If
      fpsList.TypeButtonText = Trim$(lblFind.Caption)
      fpsList.TypeButtonTextColor = RGB(0, 128, 0)
      fpsList.BlockMode = False
      
  End If
  If fpsList.SelectionCount >= 1 Then
    cmdClearSelection.Enabled = True
  Else  '0
    cmdClearSelection.Enabled = False
  End If
  lblList.Caption = "Total && Selected Record(s) :  " & fpsList.SelectionCount & "   of   " & fpsList.DataRowCnt
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound >= Index Then
    If OptSelective.Value = True Then
      If fpsList.SelectionCount >= 1 Then
        CmdStep(Index).Enabled = False
      Else
        sShowMessage "Plaese select one Customer from Customer List"
        If fpsList.Visible = True Then
 fpsList.SetFocus
End If
        CmdStep(Index).Enabled = True
      End If
    End If
'    If OptPrdrandom = True Then
'      If fpsList2.SelectionCount >= 1 Then
'        fpsList2.Enabled = False
'        CmdStep(Index).Enabled = False
'      End If
'    End If
  End If
End Sub

Private Sub cmdView_Click()
  Dim i As Integer
  Dim j As Integer

  Try
  Me.MousePointer = vbHourglass
'Call Display_Frame
  
    fpsTrend(d3).MaxRows = 8
    fpsTrend(d3).ClearRange 1, 1, fpsTrend(d3).MaxCols, 6, True
    Call Spread_HideCol(fpsTrend(d3), True, 4, 7)
    Call Spread_Clear_Row(fpsTrend(d3), 1, fpsTrend(d3).MaxCols, 1, 6)
    fpsAxis(0).MaxRows = 1
    Call Spread_Clear_Row(fpsAxis(0), 1, fpsAxis(0).MaxCols, 1, fpsAxis(0).MaxRows)
    
    If sPrintSalesTrend = False Then
      cmdprint(0).Tag = 0
      GoTo ExitHere
    Else
      If mDocSeq(1) = 2 Then
 Call Spread_HideCol(fpsTrend(d3), False, 7, 7)
End If
      cmdprint(0).Tag = 1
      cmbOrient(0).ListIndex = 1
      cmdShowGraph(0).Tag = 1
      Call cmdShowGraph_Click(0)
      cmdShowGraph(0).Tag = 0
      cmdview.Tag = 1
Call Print_PageWise
      Call Display_Frame
    End If
    cmdprint(0).Tag = 0
    chkPrintedDate.Value = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClose_Click(ByVal Index As Integer)
  
  Try
  Me.MousePointer = vbHourglass

    fraTrend(d3).Visible = False
    fraTrend(d3).Top = 25
    fraTrend(d3).Left = 25
    fraTrend(d3).Height = 50
    fraTrend(d3).Width = 50
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdShowGraph_Click(ByVal Index As Integer)
  
  Try
  Me.MousePointer = vbHourglass
      
  If Val(cmdShowGraph(0).Tag) <> 1 Then
      fraGraph.Top = 0
      fraGraph.Left = 55
      fraGraph.Height = 7940
      fraGraph.Width = 11800
      fraGraph.Visible = True  End If
    
  sstGraph.Tag = 1
  Call sstGraph_Click(0)
  
  If Val(cmdShowGraph(0).Tag) = 1 Then
    'Sales Trend Graph
      Clipboard.Clear
      mscGraph(0).Tag = mscGraph(0).Height
      pctGraph(0).Tag = mscGraph(0).Width
      If mDocSeq(1) = 1 Then
        mscGraph(0).Height = 5500 '2250
        mscGraph(0).Width = 8500  '4485
      ElseIf mDocSeq(1) = 2 Then
        mscGraph(0).Height = 5500 '2250
        mscGraph(0).Width = 8500  '4485
      ElseIf mDocSeq(1) = 3 Then
        mscGraph(0).Height = 5500 '2250
        mscGraph(0).Width = 8500  '4485
      Else
          'Skip
      End If
      mscGraph(0).EditCopy
      mscGraph(0).Height = Val(mscGraph(0).Tag)
      mscGraph(0).Width = Val(pctGraph(0).Tag)
      fpsTrend(d3).Row = Val(fpsTrend(d3).Tag)
fpsTrend(d3).Row2 = Val(fpsTrend(d3).Tag)
fpsTrend(d3).Col = 11
fpsTrend(d3).Col2 = 11
fpsTrend(d3).CellType = CellTypePicture
'      .TypePictStretch = True
      fpsTrend(d3).TypePictMaintainScale = True
      fpsTrend(d3).TypePictPicture = Clipboard.GetData(vbCFDIB)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdprint_Click(ByVal Index As Integer)

  Try
  Me.MousePointer = vbHourglass

  cmdview.Tag = 0
Call Print_PageWise
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdRefresh_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
    Call sstGraph_Click(sstGraph.Tab)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdGraphPrint_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  GphPrint = 1
  Call cmdGraphPreview_Click
  GphPrint = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdGraphPreview_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  Set TmpChart = IIf(Trim$(cmdGraphPreview.Tag) = "0", mscGraph(0), mscGraph2)
  Set TmpSpread = IIf(Trim$(cmdGraphPreview.Tag) = "0", fpsAxis(0), fpsAxis(1))
  
'  TmpChart.Top = 1695
'  TmpChart.Left = 1650
  TmpChart.Height = 5790
  TmpChart.Width = 14865
    
'Graph From mscGraph To Clipboard
  Clipboard.Clear
    IIf(Trim$(cmdGraphPreview.Tag).EditCopy
  
  TmpChart.Top = -210
  TmpChart.Left = -180
  TmpChart.Height = 5025
  TmpChart.Width = 11865

  FrmPreGraph.Show vbModal
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdGraphClose_Click()

  Try
  Me.MousePointer = vbHourglass
  
    fraGraph.Visible = False
    fraGraph.Top = 25
    fraGraph.Left = 25
    fraGraph.Height = 50
    fraGraph.Width = 50
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClear_Click()
  Dim CurrDate As Date
  
' On Error Resume Next
  CurrDate = Format(ServerDate, "MMM-yyyy")
  
  Call StepClear(Me)
  chkMonth.Value = 1
  DTPFrom.Value = DateAdd("m", -11, Format(CurrDate, "MMM-yyyy"))
  DTPTo.Value = Format(CurrDate, "MMM-yyyy")
  OptAll.Value = True
  CmdStep(0).Enabled = True
  cmbChartType.ListIndex = 0
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

'Section : Function/Procedure/Subroutine :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Function sPrintSalesTrend() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As String
  Dim i As Integer
  Dim Ln As Integer
  Dim mRev(2) As String
  
  Try
  sPrintSalesTrend = False

  If mDocSeq(1) = 1 Then 'Sales Trend
  
    Sql = "Select x.InvoiceType, x.InvMonth, Sum(x.InvoiceQty) As InvQty,"
    Sql = Sql & " Sum(x.SalesGross) As SalesGross, Sum(x.SalesNetExcise) As SalesNetExcise"
    Sql = Sql & " From ("
    Sql = Sql & " Select a.InvoiceType, a.InvoiceDate,"
    Sql = Sql & " Convert(VarChar(7),a.InvoiceDate,126) As InvMonth,"
    Sql = Sql & " Cast(a.InvoiceQty As Decimal(25,0)) As InvoiceQty,"
    Sql = Sql & " Cast(a.Rate As Decimal(25,0)) As InvoiceRate,"
    Sql = Sql & " Cast(a.InvoiceQty*a.Rate As Decimal(15,0)) As SalesNetExcise,"
    Sql = Sql & " Cast(a.TotalAmount As Decimal(25,0)) As SalesGross,"
    Sql = Sql & " a.CustomerCode"
    Sql = Sql & " From SalesInvoice a"
    Sql = Sql & " Where Convert(VarChar(7),a.InvoiceDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
    
    If OptSelective Then
      If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Vendor"
GoTo ExitHere
End If
      Sql = Sql & " and a.CustomerCode In("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 1
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
    End If
    
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"   'and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " ) x"
    Sql = Sql & " Group by x.InvoiceType, x.InvMonth"
    
    If cmbType.ListIndex = 0 Then
      Sql = Sql & " Order by x.InvMonth"
    Else
      'Skip
    End If
    
  ElseIf mDocSeq(1) = 2 Then  'Sales Retuen
    
    Sql = "Select y.OrdMonth, Sum(y.GdsVal) As GdsVal, Sum(y.RetVal) As RetVal,"
    Sql = Sql & " Sum(y.SalesGross) As SalesGross,"
    Sql = Sql & " Cast(IsNull((Sum(IsNull(y.RetVal,'0'))/Sum(y.GdsVal))*100,'0') As Decimal(25,2)) As RetPerc"
    Sql = Sql & " From ("
    Sql = Sql & " Select x.OrderType, x.OrderNo, x.CustomerCode, x.ItemCode, x.OrdMonth,"
    Sql = Sql & " (IsNull(x.InvQty,'0')*IsNull(x.Rate,'0')) As GdsVal, (IsNull(b.RetQty,'0')*IsNull(x.Rate,'0')) As RetVal," 'IsNull(b.RetQty,'0') As RetQty,
    Sql = Sql & " IsNull(((IsNull(x.InvQty,'0')*IsNull(x.Rate,'0'))-(IsNull(b.RetQty,'0')*IsNull(x.Rate,'0'))),'0') As SalesGross, x.Rate"
    Sql = Sql & " From("
    Sql = Sql & " Select a.OrderType, a.OrderNo, a.CustomerCode, a.ItemCode, "
    Sql = Sql & " Cast(a.Rate As Decimal(25,2)) As Rate,"
    Sql = Sql & " Convert(VarChar(7),a.OrderDate,126) As OrdMonth,"
    Sql = Sql & " Sum(Cast(a.InvoiceQty As Decimal(25,0))) As InvQty"
    Sql = Sql & " From SalesInvoice a"
    Sql = Sql & " Where Convert(VarChar(7),a.OrderDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
    
    If OptSelective Then
      If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Vendor"
GoTo ExitHere
End If
      Sql = Sql & " and a.CustomerCode In("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 1
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
    End If
    
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"   'and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " and a.Type In('OE', 'Labour','VAT','EOU')"
    Sql = Sql & " Group by Convert(VarChar(7),a.OrderDate,126), a.OrderType, a.OrderNo,"
    Sql = Sql & " a.CustomerCode, a.ItemCode, a.Rate"
    Sql = Sql & " ) x"
    Sql = Sql & " Full Outer Join"
    Sql = Sql & " (Select b.RefType, b.RefNo, b.VendorCode, b.ItemCode,"
    Sql = Sql & " Convert(VarChar(7),b.RefDate,126) As RefMonth,"
    Sql = Sql & " Sum(Cast(b.AcceptedQty As Decimal(25,0))) As AccQty,"
    Sql = Sql & " Sum(Cast(b.RejectionQty As Decimal(25,0))) As RetQty"
    Sql = Sql & " From ReceiptEntry b"
    Sql = Sql & " Where b.InwardType='R'"
    Sql = Sql & " and Convert(VarChar(7),b.RefDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
    
    If OptSelective Then
      If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Vendor"
GoTo ExitHere
End If
      Sql = Sql & " and b.VendorCode In("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 1
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
    End If
    
    Sql = Sql & " and b.Branchid='" & gBranchID & "' and b.Compid='" & gCompID & "'"   'and b.Finyear='" & gFinyear & "'"
    Sql = Sql & " Group by Convert(VarChar(7),b.RefDate,126), b.RefType, b.RefNo,"
    Sql = Sql & " b.VendorCode, b.ItemCode"
    Sql = Sql & " ) b"
    Sql = Sql & " On x.OrderType=b.RefType and x.OrderNo=b.RefNo and x.OrdMonth=b.RefMonth"
    Sql = Sql & " and x.ItemCode=b.ItemCode and x.CustomerCode=b.VendorCode"
    Sql = Sql & " ) y"
    Sql = Sql & " Group by y.OrdMonth"
    
    If cmbType.ListIndex = 0 Then
      Sql = Sql & " Order by y.OrdMonth"
    Else
      'Skip
    End If
    
  ElseIf mDocSeq(1) = 3 Then  'New Products
  
    Sql = "Select x.RevMonth, Sum(x.SalesValue) As SalesValue,"
    Sql = Sql & " Sum(X.CntIC) As NoProducts"
    Sql = Sql & " From ("
    Sql = Sql & " Select a.itemcode As ItemCode, IsNull(a.itemdescription,'') As ItemDesc,"
    Sql = Sql & " Cast(a.Rate  As Decimal(25,0)) As Rate, Count(a.itemcode) As CntIC,"
    Sql = Sql & " Cast((Count(a.itemcode)*Rate) As Decimal(25,0)) As SalesValue,"
    Sql = Sql & " a.RevisionNo As RevNo, IsNull(a.RevisionDate,'') As RevDate,"
    Sql = Sql & " Convert(VarChar(7), A.RevisionDate, 126) As RevMonth"
    Sql = Sql & " From INV_MATERIAL_MASTER a"
    Sql = Sql & " Where a.RevisionNo='0' and a.RevisionNo<>''"
    Sql = Sql & " and Convert(VarChar(7),a.RevisionDate,126) Between '" & Format(DTPFrom, "yyyy-MM") & "' and '" & Format(DTPTo, "yyyy-MM") & "'"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"   'and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " Group by a.ItemCode, a.itemdescription, a.Rate, a.RevisionNo, a.RevisionDate"
    Sql = Sql & " ) x"
    Sql = Sql & " Group by x.RevMonth"
    
    If cmbType.ListIndex = 0 Then
      Sql = Sql & " Order by x.RevMonth"
    Else
      'Skip
    End If
    
  End If
      
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If rs.RecordCount <= 0 Then
    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  Ln = 7
lstYAxis(0).Clear
  fpsTrend(d3).MaxRows = 8 + (IIf(rs.RecordCount <= 40, 40, rs.RecordCount) * 1)
  fpsAxis(0).MaxRows = rs.RecordCount * 2
  Call Spread_PageHeader(fpsTrend(d3))
  
    
    For i = 1 To rs.RecordCount
      Ln = Ln + 1
      fpsTrend(d3).SetText 2, Ln, (Ln - 7)
fpsTrend(d3).RowHeight(Ln) = 9.84 * 1
      Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignCenter, 2, Val(Ln), 2, Val(Ln))
      
      If mDocSeq(1) = 1 Then
        fpsTrend(d3).SetText 3, Ln, Format(rs("InvMonth"), "MMM-yyyy")  'X Axis
        fpsTrend(d3).SetText 4, Ln, Trim$(rs("InvQty"))         'Monthwise Sum of Invoice Qty
        fpsTrend(d3).SetText 8, Ln, rs("SalesNetExcise")        'Monthwise Goods Value - Y Axis For Set Max. VAlue
        fpsTrend(d3).SetText 9, Ln, Trim$(rs("SalesGross"))     'Monthwise Invoice Value
        lstYAxis(0).AddItem Space(20 - Len(Trim$(Int(Val(rs("SalesNetExcise")) / 1)))) & Val(rs("SalesNetExcise"))
        fpsAxis(0).SetText 1, i, Format(rs("InvMonth"), "MMM-yyyy")
        fpsAxis(0).SetText 2, i, Val(DigitsForAxis(Trim$(Val(rs("SalesGross"))), 2))
        fpsAxis(0).SetText 3, i, Val(DigitsForAxis(Trim$(Val(rs("SalesNetExcise"))), 2))
        
      ElseIf mDocSeq(1) = 2 Then
        fpsTrend(d3).SetText 3, Ln, Format(rs("OrdMonth"), "MMM-yyyy")  'X Axis
        fpsTrend(d3).SetText 4, Ln, Trim$(rs("GdsVal"))         'Monthwise Sum of Invoice Qty
        fpsTrend(d3).SetText 7, Ln, rs("RetVal")
        fpsTrend(d3).SetText 8, Ln, rs("SalesGross")            'Monthwise Goods Value - Y Axis For Set Max. VAlue
        fpsTrend(d3).SetText 9, Ln, Trim$(rs("RetPerc")) & "%"  'Monthwise Invoice Value
        lstYAxis(0).AddItem Space(20 - Len(Trim$(Int(Val(rs("SalesGross")) / 1)))) & Val(rs("SalesGross"))
        lstYAxis(1).AddItem Space(20 - Len(Trim$(Int(Val(rs("RetVal")) / 1)))) & Val(rs("RetVal"))
        fpsAxis(0).SetText 1, i, Format(rs("OrdMonth"), "MMM-yyyy")
        fpsAxis(0).SetText 2, i, Val(DigitsForAxis(Trim$(Val(rs("SalesGross"))), 2))
        fpsAxis(0).SetText 3, i, Val(DigitsForAxis(Trim$(Val(rs("RetVal"))), 2))
        fpsAxis(0).SetText 4, i, rs("RetPerc") & "%"
        
      ElseIf mDocSeq(1) = 3 Then
        fpsTrend(d3).SetText 3, Ln, Format(rs("RevMonth"), "MMM-yyyy")  'X Axis
        fpsTrend(d3).SetText 4, Ln, Trim$(rs("NoProducts"))   'Monthwise Sum of Invoice Qty
        fpsTrend(d3).SetText 8, Ln, rs("NoProducts")          'Monthwise Goods Value - Y Axis For Set Max. VAlue
        fpsTrend(d3).SetText 9, Ln, Trim$(rs("SalesValue"))   'Monthwise Invoice Value
        lstYAxis(0).AddItem Space(20 - Len(Trim$(Int(Val(rs("SalesValue")) / 1)))) & Val(rs("SalesValue"))
        fpsAxis(0).SetText 1, i, Format(rs("RevMonth"), "MMM-yyyy")
        fpsAxis(0).SetText 2, i, Val(DigitsForAxis(Trim$(Val(rs("SalesValue"))), 2))
        fpsAxis(0).SetText 3, i, Val(DigitsForAxis(Trim$(Val(rs("NoProducts"))), 2))
      End If
      
      rs.MoveNext
    Next i
    fpsTrend(d3).AddCellSpan 11, 8, 2, 1
    fpsTrend(d3).AddCellSpan 10, 9, 1, 22
    fpsTrend(d3).AddCellSpan 11, 9, 1, 22
    fpsTrend(d3).AddCellSpan 13, 9, 1, 22
    
    For i = 10 To 30
      If i = 10 Then
        fpsTrend(d3).SetCellBorder 12, Val(i - 1), 12, Val(i - 1), CellBorderIndexLeft, 0, CellBorderStyleBlank
      End If
      fpsTrend(d3).SetCellBorder 12, Val(i), 12, Val(i), CellBorderIndexTop, 0, CellBorderStyleBlank
      fpsTrend(d3).SetCellBorder 12, Val(i), 12, Val(i), CellBorderIndexLeft, 0, CellBorderStyleBlank
      If i = 26 Then
        fpsTrend(d3).SetCellBorder 12, Val(i + 1), 12, Val(i + 1), CellBorderIndexLeft, 0, CellBorderStyleBlank
        fpsTrend(d3).SetCellBorder 12, Val(i + 1), 12, Val(i + 1), CellBorderIndexRight, 0, CellBorderStyleBlank
      End If
    Next i
    
    fpsTrend(d3).SetCellBorder 11, 9, 12, 30, CellBorderIndexOutline, vbBlue, CellBorderStyleSolid
    fpsTrend(d3).Tag = 9  'For Picture (Chart) insertion
    
    If mDocSeq(1) = 1 Then
      fpsTrend(d3).AddCellSpan 12, (11 + 7), 1, 2
fpsTrend(d3).SetText 12, (11 + 7), "x Sales Net" & vbCrLf & "    Of Excise"

      fpsTrend(d3).Row = (11 + 7)
fpsTrend(d3).Row2 = (11 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(128, 0, 0)
      fpsTrend(d3).AddCellSpan 12, (13 + 7), 1, 2
fpsTrend(d3).SetText 12, (13 + 7), "+ Sales Gross"

      fpsTrend(d3).Row = (13 + 7)
fpsTrend(d3).Row2 = (13 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(0, 0, 128)
      
      fpsTrend(d3).SetCellBorder 12, (11 + 7), 12, (16 + 7), CellBorderIndexOutline, 0, CellBorderStyleSolid
      fpsTrend(d3).SetCellBorder 12, (17 + 7), 12, (17 + 7), CellBorderIndexTop, 0, CellBorderStyleSolid
      
    ElseIf mDocSeq(1) = 2 Then
      fpsTrend(d3).CellType = CellTypeEdit
fpsTrend(d3).TypeEditMultiLine = True

      fpsTrend(d3).AddCellSpan 12, (11 + 7), 1, 2
fpsTrend(d3).SetText 12, (11 + 7), "+ Sales Return "

      fpsTrend(d3).Row = (11 + 7)
fpsTrend(d3).Row2 = (11 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(128, 0, 0)
      fpsTrend(d3).AddCellSpan 12, (13 + 7), 1, 2
fpsTrend(d3).SetText 12, (13 + 7), "x Sales Gross"

      fpsTrend(d3).Row = (13 + 7)
fpsTrend(d3).Row2 = (13 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(0, 0, 128)
'      .AddCellSpan 12, (15 + 7), 1, 2: .SetText 12, (15 + 7), "o %(Sales Return/" & vbCrLf & "Sales Gross)"

      fpsTrend(d3).Row = (15 + 7)
fpsTrend(d3).Row2 = (15 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(128, 0, 0)
      
      fpsTrend(d3).SetCellBorder 12, (11 + 7), 12, (16 + 7), CellBorderIndexOutline, 0, CellBorderStyleSolid
      fpsTrend(d3).SetCellBorder 12, (17 + 7), 12, (17 + 7), CellBorderIndexTop, 0, CellBorderStyleSolid
      
    ElseIf mDocSeq(1) = 3 Then
      fpsTrend(d3).CellType = CellTypeEdit
fpsTrend(d3).TypeEditMultiLine = True

      fpsTrend(d3).AddCellSpan 12, (11 + 7), 1, 2
fpsTrend(d3).SetText 12, (11 + 7), "O-No. of " & vbCrLf & "Products "

      fpsTrend(d3).Row = (11 + 7)
fpsTrend(d3).Row2 = (11 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(128, 0, 0)
      fpsTrend(d3).AddCellSpan 12, 13, 1, 2
fpsTrend(d3).SetText 12, 13, "O-Sales Value"

      fpsTrend(d3).Row = (13 + 7)
fpsTrend(d3).Row2 = (13 + 7)
fpsTrend(d3).Col = 12
fpsTrend(d3).Col2 = 12
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
      fpsTrend(d3).ForeColor = RGB(128, 0, 0)
      
      fpsTrend(d3).SetCellBorder 12, (11 + 7), 12, (16 + 7), CellBorderIndexOutline, 0, CellBorderStyleSolid
      fpsTrend(d3).SetCellBorder 12, (17 + 7), 12, (17 + 7), CellBorderIndexTop, 0, CellBorderStyleSolid
    End If
    
    If fpsTrend(d3).DataRowCnt <= 30 Then
      Sql = Trim$(24 + 7)
    Else  'above Row 27
      Sql = Trim$(fpsTrend(d3).DataRowCnt + 1)
    End If
'Call Spread_Set_Focus(fpsTrend(d3), (Val(Sql) + 5), 11)
    If fpsTrend(d3).DataRowCnt <= 30 Then
 fpsTrend(d3).AddCellSpan 2, Val(Sql), 8, 1
End If
    fpsTrend(d3).AddCellSpan 11, 31, 2, 1
    If fpsTrend(d3).DataRowCnt >= 31 Then
 fpsTrend(d3).AddCellSpan 2, Val(Sql), 13, 1
End If
    fpsTrend(d3).AddCellSpan 2, (Val(Sql) + 1), 13, 3
    fpsTrend(d3).AddCellSpan 2, (Val(Sql) + 4), 13, 1
    fpsTrend(d3).SetText 2, (Val(Sql) + 4), Space(40) & "Prepared by" & Space(170) & "Approved by"
    fpsTrend(d3).Row = (Val(Sql) + 4)
fpsTrend(d3).Row2 = (Val(Sql) + 4)
fpsTrend(d3).Col = 2
fpsTrend(d3).Col2 = fpsTrend(d3).MaxCols
fpsTrend(d3).FontSize = 8
fpsTrend(d3).FontBold = True
    fpsTrend(d3).SetCellBorder 2, (Val(Sql) + 4), 13, (Val(Sql) + 4), CellBorderIndexOutline, vbBlue, CellBorderStyleSolid
    fpsTrend(d3).AddCellSpan 2, (Val(Sql) + 5), 13, 3
    fpsTrend(d3).SetText 1, (Val(Sql) + 8), "0"
    fpsTrend(d3).SetCellBorder 2, (Val(Sql) + 1), 13, (Val(Sql) + 7), CellBorderIndexOutline, vbBlue, CellBorderStyleSolid
    fpsTrend(d3).SetCellBorder 2, (Val(Sql) + 1), 13, (Val(Sql) + 7), CellBorderIndexLeft, vbBlue, CellBorderStyleSolid
    fpsTrend(d3).AddCellSpan 2, (Val(Sql) + 8), 13, 1
    fpsTrend(d3).MaxRows = fpsTrend(d3).DataRowCnt
    fpsAxis(0).MaxRows = fpsAxis(0).DataRowCnt
    lblTrend(d3).Caption = "Total Record(s) : " & Ln
    
  
'End or High Value For Y Axis
  If lstYAxis(0).ListCount >= 1 Then
    Sql = ""
sRow = ""
txtEnd.Tag = ""
    sRow = Len(Trim$(Int(Val(lstYAxis(0).List(lstYAxis(0).ListCount - 1)) / 1)))
    txtEnd.Tag = Int(Val(lstYAxis(0).List(lstYAxis(0).ListCount - 1)) / 1)
    
'    If mDocSeq(1) = 2 Then
'      txtStart.Tag = Int(Val(lstYAxis(0).List(lstYAxis(0).ListCount - 1)) / 1)
'    Else
'      'Skip
'    End If

    Sql = ""
    For i = 1 To Val(sRow) - 1
      If i = 1 Then
        Sql = "1"
      Else
        Sql = Sql & "0"
      End If
    Next i
        
    txtEnd.Tag = Val(DigitsForAxis(Trim$(Val(txtEnd.Tag)) + Val(Sql), 1))
  End If

  rs.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function DigitsForAxis(ByVal mDig As String,ByVal  mOpt As Byte) As String
  
  Try
    
  If Len(Trim$((Int(Val(mDig)) / 1))) = 9 Then
    If mOpt = 1 Then
 mDig = Left(Trim$(mDig), 4) & "." & Mid(Trim$(mDig), 5, 3)
End If
    If mOpt = 2 Then
 mDig = Left(Trim$(mDig), 4) & "." & Mid(Trim$(mDig), 5, 5)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 8 Then
    If mOpt = 1 Then
 mDig = Left(Trim$(mDig), 3) & "." & Mid(Trim$(mDig), 4, 3)
End If
    If mOpt = 2 Then
 mDig = Left(Trim$(mDig), 3) & "." & Mid(Trim$(mDig), 4, 5)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 7 Then
    If mOpt = 1 Then
 mDig = Left(Trim$(mDig), 2) & "." & Mid(Trim$(mDig), 3, 3)
End If
    If mOpt = 2 Then
 mDig = Left(Trim$(mDig), 2) & "." & Mid(Trim$(mDig), 3, 5)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 6 Then
    If mOpt = 1 Then
 mDig = Left(Trim$(mDig), 1) & "." & Mid(Trim$(mDig), 2, 3)
End If
    If mOpt = 2 Then
 mDig = Left(Trim$(mDig), 1) & "." & Mid(Trim$(mDig), 2, 5)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 5 Then
    If mOpt = 1 Then
 mDig = "0." & Mid(Trim$(mDig), 1, 3)
End If
    If mOpt = 2 Then
 mDig = "0." & Mid(Trim$(mDig), 1, 5)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 4 Then
    If mOpt = 1 Then
 mDig = "0.0" & Mid(Trim$(mDig), 1, 3)
End If
    If mOpt = 2 Then
 mDig = "0.0" & Mid(Trim$(mDig), 1, 4)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 3 Then
    If mOpt = 1 Then
 mDig = "0.00" & Mid(Trim$(mDig), 1, 3)
End If
    If mOpt = 2 Then
 mDig = "0.00" & Mid(Trim$(mDig), 1, 3)
End If
  
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 2 Then
    If mOpt = 1 Then
 mDig = "0.000" & Mid(Trim$(mDig), 1, 2)
End If
    If mOpt = 2 Then
 mDig = "0.000" & Mid(Trim$(mDig), 1, 2)
End If
    
  ElseIf Len(Trim$((Int(Val(mDig)) / 1))) = 1 Then
    If mOpt = 1 Then
 mDig = "0.0000" & Mid(Trim$(mDig), 1, 1)
End If
    If mOpt = 2 Then
 mDig = "0.0000" & Mid(Trim$(mDig), 1, 1)
End If
    
  Else
    mDig = Trim$(mDig)
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Spread_PageHeader(ByRef fps As AceSpread)
  Dim CurrDate As String
  Dim i As Integer
  Dim sRow As String
  
  Try
  
  CurrDate = Format(ServerDate, "dd/mm/yyyy")
        
    fps.SetText 2, 1, Trim$(gCompName)
    fps.SetText 2, 7, "S. NO."
    fps.SetText 2, 10, "SALES TREND CHART"
    Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignCenter, 2, 7, 2, 7)
    If mDocSeq(1) = 1 Then
    
      fps.SetText 2, 2, "SALES TREND [ MONTHWISE ]" & vbCrLf & Format(DTPFrom, "MMM-yyyy") & " - " & Format(DTPTo, "MMM-yyyy")
      If OptSelective Then
        If fpsList.SelectionCount = 1 Then
          sRow = 0
          For i = 1 To fpsList.SelectionCount
            sRow = fpsList.GetMultiSelItem(sRow)
            fpsList.Row = Val(sRow)
fpsList.Col = 2
            If Trim$(fpsList.Text) <> "" Then
              sRow = Trim$(fpsList.Text)
              Exit For
            End If
          Next i
          fps.SetText 2, 5, Space(5) & Trim$(sRow)
        End If
      End If
'      .Row = 7
'      .CellType = CellTypeEdit: .TypeEditMultiLine = True
      fps.SetText 3, 7, "MONTH(S)"
      fps.SetText 8, 7, "SALES NET OF EXCISE"
      fps.SetText 9, 7, "SALES GROSS"
      fpsAxis(0).SetText 1, 0, "X - Axis                        (Month(s))"
      fpsAxis(0).SetText 2, 0, "Y - Axis                        (Sales Gross)"
      
    ElseIf mDocSeq(1) = 2 Then
    
      fps.SetText 2, 2, "SALES RETURN [ MONTHWISE ]" & vbCrLf & Format(DTPFrom, "MMM-yyyy") & " - " & Format(DTPTo, "MMM-yyyy")
      If OptSelective Then
        If fpsList.SelectionCount = 1 Then
          sRow = 0
          For i = 1 To fpsList.SelectionCount
            sRow = fpsList.GetMultiSelItem(sRow)
            fpsList.Row = Val(sRow)
fpsList.Col = 2
            If Trim$(fpsList.Text) <> "" Then
              sRow = Trim$(fpsList.Text)
              Exit For
            End If
          Next i
          fps.SetText 2, 5, Space(5) & Trim$(sRow)
        End If
      End If
      fps.SetText 2, 7, "MONTH(S)"
      fps.SetText 7, 7, "SALES RETURN"
      fps.SetText 8, 7, "SALES GROSS"
      fps.SetText 9, 7, "PERC.(%) OF SALES RETURN"
      fpsAxis(0).SetText 1, 0, "X - Axis                        (Month(s))"
      fpsAxis(0).SetText 2, 0, "Y - Axis                        (Sales Gross)"
      
    ElseIf mDocSeq(1) = 3 Then
    
      fps.SetText 2, 2, "NEW PRODUCTS [ MONTHWISE ]" & vbCrLf & Format(DTPFrom, "MMM-yyyy") & " - " & Format(DTPTo, "MMM-yyyy")
      fps.SetText 2, 7, "MONTH(S)"
      fps.SetText 8, 7, "NO. OF PRODUCTS"
      fps.SetText 9, 7, "SALES VALUE"
      fpsAxis(0).SetText 1, 0, "X - Axis                        (Month(s))"
      fpsAxis(0).SetText 2, 0, "Y - Axis                        (Sales Value)"
      
    Else
      'Skip
    End If
        
    Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignCenter, 2, 1, 2, 2)
    Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignCenter, 2, 2, 2, 2)
    fps.Row = 1
fps.Row2 = 1
fps.Col = 2
fps.Col2 = fps.MaxCols
fps.FontSize = 12
fps.FontBold = True

    fps.Row = 2
fps.Row2 = 2
fps.Col = 2
fps.Col2 = fps.MaxCols
fps.FontSize = 12
fps.FontBold = True
    
    If chkPrintedDate.Value = 1 Then
      fps.SetText 2, 3, "Printed Date : " & Format(Trim$(CurrDate), "dd/mm/yyyy")
      Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignRight, 2, 3, 2, 3)
    End If
    Call Spread_CellHorAlign_Set(fpsTrend(d3), TypeHAlignRight, 2, 4, 2, 4)

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Print_PageWise()
  Dim mPageNo As Integer
  Dim rCntSet As Integer
  Dim rHPrev As Integer
  Dim rHNext As Integer
  Dim Ln As Long
  Dim LnCnt As Long
  Dim i As Integer
  
  Try
    
    If fpsTrend(d3).DataRowCnt >= 6 Then
    
      Ln = 0
mPageNo = 0
rCntSet = 0
LnCnt = 6
      For i = 6 To fpsTrend(d3).DataRowCnt
        Ln = Ln + 1
        If Ln = 35 Then
          LnCnt = LnCnt + Ln
          rCntSet = rCntSet + 1
          mPageNo = mPageNo + 1
'          .Row = Ln - 49: .Row2 = Ln - 49: .Col = 1: .Col2 = 1
'          .BackColor = RGB(255, 128, 128)
'          .Row = Ln: .Row2 = Ln
'          .BackColor = RGB(185, 0, 0)
          fpsTrend(d3).SetText 2, 4, " Page(s) : " & mPageNo
          If rCntSet = 1 Then
            Call Spread_HideRow(fpsTrend(d3), True, LnCnt, fpsTrend(d3).MaxRows)
            If Val(cmdview.Tag) <> 1 Then Call PrintSection
          Else
            Call Spread_HideRow(fpsTrend(d3), True, 6, (LnCnt - Ln))
            Call Spread_HideRow(fpsTrend(d3), False, (LnCnt - Ln), (LnCnt - 1))
            Call Spread_HideRow(fpsTrend(d3), True, LnCnt, fpsTrend(d3).MaxRows)
            If Val(cmdview.Tag) <> 1 Then Call PrintSection
          End If
          Ln = 0
        End If
        
        If i = fpsTrend(d3).DataRowCnt And Ln <> 0 Then
          LnCnt = LnCnt + Ln
          rCntSet = rCntSet + 1
          mPageNo = mPageNo + 1
          fpsTrend(d3).SetText 2, 4, "   Page(s) : " & mPageNo
          If rCntSet = 0 Then
            If Val(cmdview.Tag) <> 1 Then Call PrintSection
          ElseIf rCntSet = 1 Then
            Call Spread_HideRow(fpsTrend(d3), True, LnCnt, fpsTrend(d3).MaxRows)
            If Val(cmdview.Tag) <> 1 Then Call PrintSection
          Else
            Call Spread_HideRow(fpsTrend(d3), True, 6, (LnCnt - Ln))
            Call Spread_HideRow(fpsTrend(d3), False, (LnCnt - Ln), (LnCnt - 1))
            If Val(cmdview.Tag) <> 1 Then Call PrintSection
          End If
        End If
      Next i
      Call Spread_HideRow(fpsTrend(d3), False, 1, fpsTrend(d3).MaxRows)
      fpsTrend(d3).SetText 2, 4, "   Page(s) : " & " 1  To  " & mPageNo

      If Val(cmdview.Tag) <> 1 Then sShowMessage "Print Page(s) :  1  To  " & mPageNo & vbCrLf & "Print Over"
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Display_Frame()
  
  Try
    
    fraTrend(d3).Top = 0
    fraTrend(d3).Left = 60
    fraTrend(d3).Height = 7770
    fraTrend(d3).Width = 11805
    fraTrend(d3).Visible = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sstGraph_Click(ByVal PreviousTab As Integer)
  Dim retVal As Variant

  Try
  Me.MousePointer = vbHourglass

  If Val(sstGraph.Tag) = 1 Then
    If sstGraph.Tab <> 0 Then
      sstGraph.Tab = 0
cmdGraphPreview.Tag = "0"
      GoTo ExitHere
    End If
  'Sales Trend Graph
    fpsAxis(0).Visible = True   ': fpsAxis(1).Visible = False

    sstGraph.Tag = 0
cmdGraphPreview.Tag = "0"
    txtStart.Text = 0#     'Low Value
    txtEnd.Text = Val(txtEnd.Tag)   'High Value
'    txtTarget.Text = txtEnd.Text
    lblStart.Tag = Val(txtStart.Text)
lblEnd.Tag = Val(txtEnd.Tag)
    Call PrepareChart
    If Val(cmdShowGraph(0).Tag) = 1 Then
 sstGraph.Tab = 1
End If
    GoTo ExitHere
  End If
  
  Select Case sstGraph.Tab
    Case 0:   'Sales Trend Graph
      fpsAxis(0).Visible = True
      txtStart.Text = 0   'Low Value
      txtEnd.Text = Val(txtEnd.Tag)   'High Value
'      txtTarget.Text = txtEnd.Text
      cmdGraphPreview.Tag = "0"
    Case 1: '?

    Case 2: '?
    
    Case Else
      cmdGraphPreview.Tag = ""
  End Select
  
  Call PrepareChart
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Sub PrintSection()
  
  Try
  
    fpsTrend(d3).PrintUseDataMax = True
    fpsTrend(d3).PrintPageOrder = PageOrderDownThenOver
    If mDocSeq(4) = 3 Or mDocSeq(4) = 5 Then
      fpsTrend(d3).PrintOrientation = PrintOrientationPortrait
    Else
      fpsTrend(d3).PrintOrientation = PrintOrientationLandscape
    End If
    fpsTrend(d3).PrintMarginTop = 300
    fpsTrend(d3).PrintCenterOnPageH = True
    fpsTrend(d3).PrintFirstPageNumber = 1
    fpsTrend(d3).PrintPageStart = 1
    fpsTrend(d3).BorderStyle = BorderStyleFixedSingle
'    .PrintBorder = True
    fpsTrend(d3).Action = ActionPrint
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub PrepareChart()

  Try
  
  Select Case sstGraph.Tab
    Case 0, 1, 2:   'Chart
        'Vertical Axis Divided as per Given Range Details
        With IIf(sstGraph.Tab.Plot.Axis(VtChAxisIdY).ValueScale
          IIf(sstGraph.Tab.Auto = False
          IIf(sstGraph.Tab.Minimum = Val(txtStart) 'Low Value
          IIf(sstGraph.Tab.Maximum = Val(txtEnd)   'High Value
          
'          .MajorDivision = Val(txtRowDiv)
'          .MinorDivision = Val(txtColDiv)
          .Plot.Axis(VtChAxisIdY).ValueScale.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
          .Plot.Axis(VtChAxisIdY).ValueScale.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = Val(txtStart) 'Low Value
          .Plot.Axis(VtChAxisIdY).ValueScale.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Val(txtEnd)   'High Value
          
'          .MajorDivision = Val(txtRowDiv)
'          .MinorDivision = Val(txtColDiv)        
'        With .Plot.Axis(VtChAxisIdY2).ValueScale
'          .Auto = False
'          .Minimum = Val(txtstart) 'Low Value
'          .Maximum = Val(txtEnd)   'High Value
'
''          .MajorDivision = Val(txtRowDiv)
''          .MinorDivision = Val(txtColDiv)
'        End With
        
      'For ChartType
        Call ChartType_2D3D
        
        If mDocSeq(1) = 1 Then
          If OptSelective Then
            .Title.Text = "SALES TREND [CUSTOMERWISE] MONTHWISE ANALYSIS"
          Else
            .Title.Text = "SALES TREND MONTHWISE ANALYSIS"
          End If
          .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "MONTH(S)"
          .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "SALES VALUE (In Lakh(s))"
          .FootnoteText = gCompName
        ElseIf mDocSeq(1) = 2 Then
          .Title.Text = "SALES RETURN MONTHWISE ANALYSIS"
          .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "MONTH(S)"
          .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "SALES VALUE (In Lakh(s))"
          .FootnoteText = gCompName
        ElseIf mDocSeq(1) = 2 Then
          .Title.Text = "NEW PRODUCTS MONTHWISE ANALYSIS"
          .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "MONTH(S)"
          .Plot.Axis(VtChAxisIdY).AxisTitle.Text = "SALES VALUE (In Lakh(s))"
          .FootnoteText = gCompName
        Else
          'Skip
        End If
             
        'Horizontal Axis Divided as per Given Time Range
        
    '      .Auto = False
    '      .Maximum = Val(dtpFrom)
    '      .Minimum = Val(dtpTo)
          .Plot.Axis(VtChAxisIdX).ValueScale.Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = 20
          
      End With
    Case Else
      'Skip
  End Select
  
  Call PlotValues
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub ChartType_2D3D()

  Try
  
    Select Case cmbChartType.ListIndex
      Case 0:   '2D Chart
        mscGraph(0).ChartType = VtChChartType2dBar
      Case 1:
        mscGraph(0).ChartType = VtChChartType2dLine
      Case 2:
        mscGraph(0).ChartType = VtChChartType2dCombination
      Case 3:
        mscGraph(0).ChartType = VtChChartType2dXY
      Case 4:
        mscGraph(0).ChartType = VtChChartType2dPie
      Case 5:
        mscGraph(0).ChartType = VtChChartType3dArea
      Case 6:
        mscGraph(0).ChartType = VtChChartType2dStep
      Case 7:   '3D Chart
        mscGraph(0).ChartType = VtChChartType3dBar
      Case 8:
        mscGraph(0).ChartType = VtChChartType3dLine
      Case 9:
        mscGraph(0).ChartType = VtChChartType3dCombination
      Case 10:
        mscGraph(0).ChartType = VtChChartType3dArea
      Case 11:
        mscGraph(0).ChartType = VtChChartType3dStep
      Case Else
        sShowMessage "'Chart Type' Not Found"
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub PlotValues()
  Dim rCnt As Long
  Dim addRCnt As Long
  Dim arrData()
  Dim retVal As Variant
  Dim tmp As Variant
  Dim i As Integer
  
  Try
  
  
  Select Case sstGraph.Tab
    Case 0, 1, 2: 'Sales Trend
        rCnt = fpsAxis(0).DataRowCnt
        ReDim arrData(rCnt, 1 To 15)
        
      'Time (Spread - Col=1)  X-Axis
        For i = 0 To rCnt - 1
          fpsAxis(0).GetText 1, i + 1, retVal
          arrData(i, 1) = " " & Trim$(retVal)
        Next i
        
        mscGraph2.Plot.Axis(VtChAxisIdX).CategoryScale.LabelTick = True
        mscGraph2.ChartData = arrData
                      
      'Actual Reading (Spread - Col=2)  Y-Axis
        For i = 0 To rCnt - 1
          fpsAxis(0).GetText 2, i + 1, retVal
          arrData(i, 2) = Round(Val(retVal), 4)
        Next i
        
'        If mDocSeq(1) = 2 Then
'        'Actual Reading (Spread - Col=3)  Y2-Axis
'          For i = 0 To rCnt - 1
'            .GetText 4, i + 1, retVal
'            arrData(i, 3) = Val(Left(Trim$(retVal), Len(Trim$(retVal)) - 1))
'          Next i
'        End If
            
        For i = 0 To rCnt - 1
          fpsAxis(0).GetText 3, i + 1, retVal
          arrData(i, 3) = Round(Val(retVal), 4)
        Next i

'      'Target Value
'        For i = 0 To rCnt - 1
'          arrData(i, 10) = Val(txtTarget)
'        Next i
        
        mscGraph(0).ChartData = arrData      
    Case 3: '?
        rCnt = fpsAxis(1).DataRowCnt
        ReDim arrData(rCnt, 1 To 15)
        
      'Time (Spread - Col=1)  X-Axis
        For i = 0 To rCnt - 1
          fpsAxis(1).GetText 1, i + 1, retVal
          arrData(i, 1) = " " & Trim$(retVal)
        Next i
        
'        mscGraph2.Plot.Axis(VtChAxisIdX).CategoryScale.LabelTick = True
        mscGraph2.ChartData = arrData
                      
      'Actual Reading (Spread - Col=2)  Y-Axis
        For i = 0 To rCnt - 1
          fpsAxis(1).GetText 2, i + 1, retVal
          arrData(i, 2) = Round(Val(retVal), 4)
        Next i

'      'Target Value
'        For i = 0 To rCnt - 1
'          arrData(i, 10) = Val(txtTarget)
'        Next i
        
        mscGraph2.ChartData = arrData
    Case Else
      'Skip
  End Select
Catch ex As Exception
Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Set_SearchTextOption(ByVal mCl As Integer)
  Dim retVal As Variant
  Dim i As Integer
  
  Try
  
    fpsList.Row = 0
fpsList.Row2 = 0
    
    fpsList.BlockMode = True
    For i = 1 To fpsList.MaxCols
      fpsList.Col = i
fpsList.Col2 = i
fpsList.GetText i, 0, retVal
      fpsList.CellType = CellTypeButton
      fpsList.TypeButtonText = Trim$(retVal)
      fpsList.TypeButtonTextColor = RGB(0, 0, 0)
    Next i
    
    fpsList.Col = 1
fpsList.Col2 = 1
fpsList.GetText 1, 0, retVal
    fpsList.CellType = CellTypeButton
    fpsList.TypeButtonText = Trim$(retVal)
    fpsList.TypeButtonTextColor = RGB(0, 128, 0)
    fpsList.BlockMode = False
  
    fpsList.Tag = 1
    
  If mCl = 1 Then
    txtFind.Text = ""
txtFind.Tag = 0
lblFind.Tag = 0

    chkSRow.Value = 1
chkPMAllowed.Value = 1
fpsList.Tag = 2
  ElseIf mCl = 2 Then
    txtFind.Text = ""
txtFind.Tag = 0
lblFind.Tag = 0

    chkSRow.Value = 1
chkPMAllowed.Value = 1
fpsList.Tag = 2

'    txtFind2.Text = "": txtFind2.Tag = 0: lblFind2.Tag = 0

'    chkSRow2.Value = 1: chkPMAllowed2.Value = 1: fpsList2.Tag = 2
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Design_fpsAxis_0(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsAxis(0)

End Sub

Private Sub Design_fpsQty(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsQty

End Sub

Private Sub Design_fpsGPaynter(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsGPaynter

End Sub

Private Sub Design_fpsTrend_0(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsTrend(0)

End Sub

Private Sub Design_fpsList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsList


fpsList.Click +=  new EventHandler(fpsList_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_fpsAxis_0("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesTrend.frm", fpsAxis(0), "0")

	Call Design_fpsQty("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesTrend.frm", fpsQty, "")

	Call Design_fpsGPaynter("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesTrend.frm", fpsGPaynter, "")

	Call Design_fpsTrend_0("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesTrend.frm", fpsTrend(0), "0")

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesTrend.frm", fpsList, "")

End Sub
