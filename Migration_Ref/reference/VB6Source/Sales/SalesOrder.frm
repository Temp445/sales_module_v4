VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSalesOrder 
   Caption         =   "Sales Order"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   118
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5325
      TabIndex        =   117
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4035
      TabIndex        =   116
      ToolTipText     =   "Click to Save the Data"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   8040
      TabIndex        =   124
      Top             =   5790
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   119
         Top             =   0
         Width           =   900
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1890
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   122
         Top             =   1845
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   120
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   0
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   1485
         Left            =   45
         TabIndex        =   121
         Top             =   345
         Visible         =   0   'False
         Width           =   3555
         _Version        =   458752
         _ExtentX        =   6271
         _ExtentY        =   2619
         _StockProps     =   64
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
         ScrollBars      =   2
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   990
         TabIndex        =   127
         Top             =   1890
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   126
         Top             =   15
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin TabDlg.SSTab tabSales 
      Height          =   3855
      Left            =   270
      TabIndex        =   115
      Top             =   1905
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Terms Details"
      TabPicture(0)   =   "SalesOrder.frx":0314
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTerms"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Order List"
      TabPicture(1)   =   "SalesOrder.frx":0330
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSchedule"
      Tab(1).Control(1)=   "spOrder"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Packing Details"
      TabPicture(2)   =   "SalesOrder.frx":034C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBuyer"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Packing List"
      TabPicture(3)   =   "SalesOrder.frx":0368
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraCarriage"
      Tab(3).Control(1)=   "txtReference"
      Tab(3).Control(2)=   "sprGoods"
      Tab(3).Control(3)=   "lblRef"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "General Details / Departments"
      TabPicture(4)   =   "SalesOrder.frx":0384
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtGeneralDetails"
      Tab(4).Control(1)=   "spDept"
      Tab(4).Control(2)=   "Label25"
      Tab(4).Control(3)=   "Label26"
      Tab(4).ControlCount=   4
      Begin VB.Frame fraSchedule 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1575
         Left            =   -74955
         TabIndex        =   132
         Top             =   2190
         Width           =   7950
         Begin VB.TextBox txtItemCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   53
            Top             =   105
            Width           =   2025
         End
Begin AceSpread sprSchedule
            Height          =   1035
            Left            =   45
            TabIndex        =   54
            Top             =   480
            Width           =   7860
            _Version        =   458752
            _ExtentX        =   13864
            _ExtentY        =   1826
            _StockProps     =   64
            ButtonDrawMode  =   4
            DAutoCellTypes  =   0   'False
            DAutoHeadings   =   0   'False
            DAutoSave       =   0   'False
            DAutoSizeCols   =   0
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   1
            ScrollBars      =   2
         End
         Begin VB.Label lblItemCode 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Item Code"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   330
            TabIndex        =   52
            Top             =   157
            Width           =   705
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraBuyer 
         ForeColor       =   &H00FF0000&
         Height          =   1710
         Left            =   -74910
         TabIndex        =   131
         Top             =   375
         Width           =   11220
         Begin VB.TextBox txtBState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   30
            TabIndex        =   72
            ToolTipText     =   "Sales Quote Date"
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox txtBCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   74
            ToolTipText     =   "Sales Quote Date"
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox txtBPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6555
            MaxLength       =   20
            TabIndex        =   68
            Top             =   810
            Width           =   1860
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8430
            MaxLength       =   50
            TabIndex        =   70
            ToolTipText     =   "Sales Quote Date"
            Top             =   810
            Width           =   2640
         End
         Begin VB.TextBox txtBAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            TabIndex        =   62
            Top             =   810
            Width           =   3660
         End
         Begin VB.TextBox txtBPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5655
            TabIndex        =   66
            ToolTipText     =   "Sales Quote Date"
            Top             =   810
            Width           =   885
         End
         Begin VB.TextBox txtBCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            TabIndex        =   64
            Top             =   810
            Width           =   1890
         End
         Begin VB.TextBox txtBAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   58
            Top             =   315
            Width           =   3660
         End
         Begin VB.TextBox txtBAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7425
            TabIndex        =   60
            Top             =   315
            Width           =   3660
         End
         Begin VB.TextBox txtBuyerName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   50
            TabIndex        =   56
            Top             =   315
            Width           =   3660
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   71
            Top             =   1140
            Width           =   375
         End
         Begin VB.Label lblcountry 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2880
            TabIndex        =   73
            Top             =   1125
            Width           =   570
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6555
            TabIndex        =   67
            Top             =   630
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   8430
            TabIndex        =   69
            Top             =   630
            Width           =   360
         End
         Begin VB.Label lblBPinCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5655
            TabIndex        =   65
            Top             =   630
            Width           =   585
         End
         Begin VB.Label lblBCity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   63
            Top             =   630
            Width           =   270
         End
         Begin VB.Label lblBAdd3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   61
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblBAdd1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   57
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblBAdd2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7425
            TabIndex        =   59
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblBuyerName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Buyer Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   55
            Top             =   135
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00FF0000&
         Height          =   1770
         Left            =   -74910
         TabIndex        =   130
         Top             =   1980
         Width           =   11220
         Begin VB.TextBox txtConsignName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   50
            TabIndex        =   76
            Top             =   345
            Width           =   3660
         End
         Begin VB.TextBox txtAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7425
            TabIndex        =   80
            Top             =   345
            Width           =   3660
         End
         Begin VB.TextBox txtAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   78
            Top             =   345
            Width           =   3660
         End
         Begin VB.TextBox txtCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            TabIndex        =   84
            Top             =   855
            Width           =   1890
         End
         Begin VB.TextBox txtPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5655
            TabIndex        =   86
            ToolTipText     =   "Sales Quote Date"
            Top             =   855
            Width           =   885
         End
         Begin VB.TextBox txtAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            TabIndex        =   82
            Top             =   855
            Width           =   3660
         End
         Begin VB.TextBox txtBEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8430
            MaxLength       =   50
            TabIndex        =   90
            ToolTipText     =   "Sales Quote Date"
            Top             =   855
            Width           =   2640
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6555
            MaxLength       =   20
            TabIndex        =   88
            Top             =   855
            Width           =   1860
         End
         Begin VB.TextBox txtCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   94
            ToolTipText     =   "Sales Quote Date"
            Top             =   1365
            Width           =   2790
         End
         Begin VB.TextBox txtState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   30
            TabIndex        =   92
            ToolTipText     =   "Sales Quote Date"
            Top             =   1365
            Width           =   2790
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7410
            TabIndex        =   79
            Top             =   150
            Width           =   720
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   77
            Top             =   165
            Width           =   720
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consignee Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   75
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   81
            Top             =   675
            Width           =   720
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   83
            Top             =   675
            Width           =   270
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5655
            TabIndex        =   85
            Top             =   675
            Width           =   585
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   8430
            TabIndex        =   89
            Top             =   675
            Width           =   360
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6555
            TabIndex        =   87
            Top             =   675
            Width           =   450
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2880
            TabIndex        =   93
            Top             =   1170
            Width           =   570
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   91
            Top             =   1185
            Width           =   375
         End
      End
      Begin VB.Frame fraCarriage 
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   -74910
         TabIndex        =   110
         Top             =   2430
         Width           =   7725
         Begin VB.TextBox txtPortName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   107
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.TextBox txtDesName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   109
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.TextBox txtLoadName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   105
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtRecName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   101
            ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
            Top             =   180
            Width           =   2430
         End
         Begin VB.TextBox txtCarriageNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   103
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtCarriage 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   99
            ToolTipText     =   "Sales Quotation Type"
            Top             =   180
            Width           =   2430
         End
         Begin VB.Label lblDesName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final Destination"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4020
            TabIndex        =   108
            Top             =   922
            Width           =   1170
         End
         Begin VB.Label lblLoadName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port of Loading"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4080
            TabIndex        =   104
            Top             =   577
            Width           =   1110
         End
         Begin VB.Label lblPortName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port of Discharge"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   106
            Top             =   922
            Width           =   1260
         End
         Begin VB.Label lblCarriage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Carriage by"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   98
            Top             =   232
            Width           =   1125
         End
         Begin VB.Label lblReceivingName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place Of Receipt"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3990
            TabIndex        =   100
            Top             =   232
            Width           =   1215
         End
         Begin VB.Label lblCarriageNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vessel / Filght No"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   102
            Top             =   577
            Width           =   1260
         End
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73590
         TabIndex        =   96
         ToolTipText     =   "Sales Quotation Type"
         Top             =   585
         Width           =   2430
      End
      Begin VB.TextBox txtGeneralDetails 
         Appearance      =   0  'Flat
         Height          =   3135
         Left            =   -74820
         MaxLength       =   3000
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   555
         Width           =   6900
      End
      Begin VB.Frame fraTerms 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   90
         TabIndex        =   128
         Top             =   330
         Width           =   11220
         Begin VB.TextBox txtRevisionLevel 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7950
            MaxLength       =   10
            TabIndex        =   50
            ToolTipText     =   "Enter the Remarks"
            Top             =   1305
            Width           =   1176
         End
         Begin VB.ComboBox cmbScheduleType 
            Height          =   330
            ItemData        =   "SalesOrder.frx":0763
            Left            =   6540
            List            =   "SalesOrder.frx":076D
            Style           =   2  'Dropdown List
            TabIndex        =   48
            ToolTipText     =   "Select the Sales Order Type"
            Top             =   1305
            Width           =   1410
         End
         Begin VB.TextBox txtDelIns 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   36
            ToolTipText     =   "Select the Delivery Instruction. Press F2 or Double Click for Help"
            Top             =   300
            Width           =   3675
         End
         Begin VB.TextBox txtFreight 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   40
            ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtInsurance 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   34
            ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
            Top             =   300
            Width           =   3675
         End
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   44
            ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
            Top             =   1305
            Width           =   3675
         End
         Begin VB.TextBox txtMode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   42
            ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtForm 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   38
            ToolTipText     =   "Select the Form. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtPack 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   32
            ToolTipText     =   "Select the Packing & Forwarding. Press F2 or Double Click for Help"
            Top             =   300
            Width           =   3675
         End
         Begin VB.TextBox txtFormula 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   46
            ToolTipText     =   "Select the Formula Name. Press F2 / DblClick Key for Help."
            Top             =   1305
            Width           =   2760
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revision Level"
            Height          =   195
            Left            =   7950
            TabIndex        =   49
            Top             =   1125
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule Type"
            Height          =   210
            Left            =   6540
            TabIndex        =   47
            Top             =   1125
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   7440
            TabIndex        =   41
            Top             =   630
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3765
            TabIndex        =   33
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Form"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   75
            TabIndex        =   37
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packing && Forwarding"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   75
            TabIndex        =   31
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Freight"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3765
            TabIndex        =   39
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terms of Payment"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   75
            TabIndex        =   43
            Top             =   1125
            Width           =   1305
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Instruction"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   7455
            TabIndex        =   35
            Top             =   120
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula Name"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3765
            TabIndex        =   45
            Top             =   1125
            Width           =   1020
         End
      End
Begin AceSpread spDept
         Height          =   3135
         Left            =   -67890
         TabIndex        =   114
         Top             =   555
         Width           =   4080
         _Version        =   458752
         _ExtentX        =   7197
         _ExtentY        =   5530
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
         MaxCols         =   3
         MaxRows         =   100
         ScrollBars      =   2
      End
Begin AceSpread sprGoods
         Height          =   1455
         Left            =   -74895
         TabIndex        =   97
         Top             =   975
         Width           =   11175
         _Version        =   458752
         _ExtentX        =   19711
         _ExtentY        =   2566
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
         MaxCols         =   8
         MaxRows         =   1
      End
Begin AceSpread spOrder
         Height          =   1740
         Left            =   -74910
         TabIndex        =   51
         ToolTipText     =   "Enter the Sales Item Details"
         Top             =   435
         Width           =   11220
         _Version        =   458752
         _ExtentX        =   19791
         _ExtentY        =   3069
         _StockProps     =   64
         Enabled         =   0   'False
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
         MaxCols         =   12
         MaxRows         =   1
         ProcessTab      =   -1  'True
         SelectBlockOptions=   0
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Reference"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   95
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departments"
         Height          =   210
         Left            =   -67890
         TabIndex        =   113
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "General Details"
         Height          =   210
         Left            =   -74820
         TabIndex        =   111
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraQuotation 
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   7125
      TabIndex        =   123
      Top             =   -15
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtQuoType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Sales Quotation Type"
         Top             =   330
         Width           =   765
      End
      Begin VB.TextBox txtQuoDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Sales Quote Date"
         Top             =   330
         Width           =   1140
      End
      Begin VB.TextBox txtQuoNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quot.Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1875
         TabIndex        =   15
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.No."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   870
         TabIndex        =   13
         Top             =   150
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Width           =   705
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Order Source"
      Height          =   720
      Left            =   3435
      TabIndex        =   7
      ToolTipText     =   "Select the Sales Order Source"
      Top             =   -15
      Width           =   3675
      Begin VB.OptionButton optWithOutItem 
         Caption         =   "Without Item Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1995
         TabIndex        =   10
         ToolTipText     =   "Select if the Order is through Quotation"
         Top             =   300
         Width           =   1620
      End
      Begin VB.OptionButton optquotation 
         Caption         =   "Quotation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   9
         ToolTipText     =   "Select if the Order is through Quotation"
         Top             =   300
         Width           =   990
      End
      Begin VB.OptionButton optdirect 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select if it is a Direct Order"
         Top             =   300
         Value           =   -1  'True
         Width           =   720
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   720
      Left            =   270
      TabIndex        =   6
      Top             =   -15
      Width           =   3150
      Begin MSComCtl2.DTPicker dporddate 
         Height          =   315
         Left            =   1890
         TabIndex        =   5
         ToolTipText     =   "Sales Order Date"
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59047939
         CurrentDate     =   37250
      End
      Begin VB.TextBox txtordnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Sales Order No"
         Top             =   315
         Width           =   930
      End
      Begin VB.ComboBox cmbordtype 
         Height          =   330
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select the Sales Order Type"
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1890
         TabIndex        =   4
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   960
         TabIndex        =   3
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Type"
         Height          =   210
         Left            =   75
         TabIndex        =   0
         Top             =   135
         Width           =   825
      End
   End
   Begin VB.Frame fraCustomer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   1230
      Left            =   270
      TabIndex        =   129
      Top             =   630
      Width           =   10680
      Begin VB.TextBox txtPostingAc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   825
         Width           =   4620
      End
      Begin VB.TextBox txtrefno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   20
         ToolTipText     =   "Enter Ref. No"
         Top             =   315
         Width           =   1890
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   4620
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   28
         ToolTipText     =   "Enter the Remarks"
         Top             =   825
         Width           =   4605
      End
      Begin MSComCtl2.DTPicker dprefdate 
         Height          =   315
         Left            =   6645
         TabIndex        =   22
         ToolTipText     =   "Enter the Ref. Date"
         Top             =   315
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59047939
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   8085
         TabIndex        =   24
         ToolTipText     =   "Enter the Valid From Date"
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59047939
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   9360
         TabIndex        =   26
         ToolTipText     =   "Enter the Valid Till Date"
         Top             =   315
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59047939
         CurrentDate     =   37250
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4725
         TabIndex        =   29
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   210
         Left            =   105
         TabIndex        =   27
         Top             =   645
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         Height          =   210
         Left            =   9360
         TabIndex        =   25
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid From"
         Height          =   210
         Left            =   8085
         TabIndex        =   23
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer PO Date"
         Height          =   210
         Left            =   6645
         TabIndex        =   21
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer PO No"
         Height          =   210
         Left            =   4740
         TabIndex        =   19
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   17
         Top             =   135
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FrmLName As String
Dim SKey As String

Public mCallMenu As eSalesCallType

'Column Nos for Order Spread
Private Enum eItem
  mItem = 1
  mDesc = 2
  mPartNo = 3
  mRevNo = 4
  mUOMCode = 5
  mUOM = 6
  mSchType = 7
  mQty = 8
  mRate = 9
  mAssRate = 10
  mGoodsValue = 11
  mAssGoodsValue = 12
End Enum

'column nos for container details
Private Enum eGoods
    mContainerNo = 1
    mPackageKind = 2
    mDescription = 3
    mQuantity = 4
    mNetWeight = 5
    mLength = 6
    mWidth = 7
    mHeight = 8
End Enum

'Schedule Spread Column Numbers
Private Enum eSch
  mSchDate = 1
  mSchQty = 2
  mLocation = 3
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Sub SetSheet(ByVal fRow As Long)
' On Error Resume Next
  
    If sprSchedule.SheetCount < fRow Then
      sprSchedule.SheetCount = fRow
      Call CopySpreadStructure(sprSchedule, 1, , fRow)
      sprSchedule.sheet = fRow
      sprSchedule.MaxRows = 1
    End If
    sprSchedule.sheet = fRow
    sprSchedule.ActiveSheet = fRow
End Sub

Private Sub LoadDept()
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  Set snap = New AceADODB.Recordset
  
  Sql = "select Code, Description from TermsMaster where TypeFlg = 'Department'"
  Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  Sql = Sql + " order by Description"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spDept.MaxRows = snap.RecordCount
    For i = 1 To snap.RecordCount
      spDept.SetText 1, i, snap("Code")
      spDept.SetText 2, i, snap("Description")
      snap.MoveNext
    Next i
  End If
  snap.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OrderCalulation()
  Dim tot As Double
  Dim atot As Double
  Dim i As Long
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass

  If spOrder.DataRowCnt = 0 Then
 GoTo ExitHere
End If
  
  tot = 0
atot = 0
  For i = 1 To spOrder.DataRowCnt
    tot = tot + Val(SpVal(spOrder, i, 9))
    atot = atot + Val(SpVal(spOrder, i, 10))
  Next i
  
  If txtFormula <> "" Then
'    Calculation Tot, atot, Val(txtFormula.Tag), SKey  'To calculate the formula
'
'    Sql = " select a.Code,(select b.description from com_formula_parameter b "
'    Sql = Sql + " where b.code=a.Code and b.Branchid=a.Branchid and b.Compid=a.Compid) Description,"
'    Sql = Sql + " a.amt,a.Sequence,a.Printing from com_formula_tmp a where a.FormulaCode=" & Val(txtFormula.Tag) & ""
'    Sql = Sql + " and a.printing=1 and SKey = '" + SKey + "' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
'    Sql = Sql + " order by a.sequence"
'    Set snap = New ADODB.Recordset
'    If RSOpen(snap, Sql) = False Then GoTo ExitHere
'    If Not snap.EOF Then
'      snap.MoveLast: snap.MoveFirst
'      spCalc.MaxRows = snap.RecordCount
'      For i = 1 To snap.RecordCount
'        spCalc.SetText 1, i, snap("Code")
'        spCalc.SetText 2, i, snap("Description")
'        spCalc.SetText 3, i, snap("Amt")
'        spCalc.SetText 4, i, snap("Sequence")
'        spCalc.SetText 5, i, snap("Printing")
'        snap.MoveNext
'      Next i
'    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Function SaveSalOrderScheduleData(ByVal sRow As Long) As Boolean
  Dim cmd As AceADODB.Command
  Dim retval As Variant
  Dim rCnt As Long    'Row Count
  Dim ItemCode As String
  
  Try
  SaveSalOrderScheduleData = False
  
  Set cmd = New AceADODB.Command
  cmd.ActiveConnection = db
  cmd.CommandText = "SAL_Ins_Order_Schedule"
  cmd.CommandType = adCmdStoredProc
    sprSchedule.sheet = sRow
    spOrder.GetText eItem.mItem, sRow, retval
    ItemCode = Trim$(retval)
    For rCnt = 1 To sprSchedule.DataRowCnt
      sprSchedule.GetText eSch.mSchDate, rCnt, retval
      If Trim$(retval) <> "" Then
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum.Text)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate.Value, "yyyy-MM-dd"))
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(ItemCode))
        spOrder.GetText eItem.mPartNo, sRow, retval
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retval))
        sprSchedule.GetText eSch.mSchDate, rCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, , Format(retval, "yyyy-MM-dd"))
        sprSchedule.GetText eSch.mSchQty, rCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(retval))
        sprSchedule.GetText eSch.mLocation, rCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(retval))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
    Next rCnt
    
  SaveSalOrderScheduleData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
End Try
End Function

Private Function ValidSchedule(ByVal SchRow As Long) As Boolean
  Dim retval As Variant
  Dim CurSheet As Long
  Dim rCnt As Long
  Dim PrevDate As Long
  Dim CurDate As Long
  Dim fItemCode As String

  Try
  ValidSchedule = False
  
'    SchRow = Spread_Get_SheetNoByName(sprSchedule, fSheetName)
    CurSheet = SchRow
    spOrder.GetText eItem.mItem, CurSheet, retval
    fItemCode = Trim$(retval)
    retval = CurSheet
    If retval > 0 Then
      sprSchedule.sheet = Val(retval)
      If sprSchedule.DataRowCnt <= 0 Then
        sShowMessage "No Schedule found for the Item  " & fItemCode
        txtitemCode.Text = fItemCode
        sprSchedule.ActiveSheet = CurSheet
        sprSchedule.SetActiveCell eSch.mSchDate, Val(retval)
'        .SetFocus
        GoTo ExitHere
      Else
        For rCnt = 1 To sprSchedule.DataRowCnt
          sprSchedule.GetText eSch.mSchDate, rCnt, retval
          If Trim$(retval) = "" Then
            sShowMessage "Schedule Date not given"
            txtitemCode.Text = fItemCode
            sprSchedule.ActiveSheet = CurSheet
            sprSchedule.SetActiveCell eSch.mSchDate, rCnt
            sprSchedule.SetFocus
            GoTo ExitHere
          Else
            If IsDate(retval) = False Then
              sShowMessage "Invalid Schedule Date given"
              txtitemCode.Text = fItemCode
              sprSchedule.ActiveSheet = CurSheet
              sprSchedule.SetActiveCell eSch.mSchDate, rCnt
              sprSchedule.SetFocus
              GoTo ExitHere
            End If
          End If
          If Format(retval, "yyyyMMdd") < Format(dpvalidfrom, "yyyyMMdd") Then
            sShowMessage "Schedule Date Should be Greater or Equal to the Order ValidFrom Date"
            txtitemCode.Text = fItemCode
            sprSchedule.ActiveSheet = CurSheet
            sprSchedule.SetActiveCell eSch.mSchDate, rCnt
            sprSchedule.SetFocus
            GoTo ExitHere
          End If
          
          If Format(retval, "yyyyMMdd") > Format(dpvalidto, "yyyyMMdd") Then
            sShowMessage "Schedule Date Should be Lesser or Equal to the Order ValidTo Date"
            txtitemCode.Text = fItemCode
            sprSchedule.ActiveSheet = CurSheet
            sprSchedule.SetActiveCell eSch.mSchDate, rCnt
            sprSchedule.SetFocus
            GoTo ExitHere
          End If
          
          If rCnt = 1 Then
            PrevDate = Format(dporddate, "yyyyMMdd")
            CurDate = Format(retval, "yyyyMMdd")
          Else
            PrevDate = CurDate
            CurDate = Format(retval, "yyyyMMdd")
          End If
          If PrevDate >= CurDate Then
            If (rCnt = 1 And PrevDate > CurDate) Then
              sShowMessage "Schedule Date Should be Greater than or Equal to Current Date."
              txtitemCode.Text = fItemCode
              sprSchedule.ActiveSheet = CurSheet
              sprSchedule.SetActiveCell eSch.mSchDate, rCnt
              sprSchedule.SetFocus
              GoTo ExitHere
            ElseIf rCnt > 1 Then
              sShowMessage "Schedule Date Should be Greater than the Previous Date."
              txtitemCode.Text = fItemCode
              sprSchedule.ActiveSheet = CurSheet
              sprSchedule.SetActiveCell eSch.mSchDate, rCnt
              sprSchedule.SetFocus
              GoTo ExitHere
            End If
          End If
          sprSchedule.GetText eSch.mSchQty, rCnt, retval
          If Val(retval) = 0 Then
            sShowMessage "No Schedule Qty is given for the Item " & fItemCode & " in the row " & rCnt
            txtitemCode.Text = fItemCode
            sprSchedule.ActiveSheet = CurSheet
            sprSchedule.SetActiveCell eSch.mSchQty, rCnt
            sprSchedule.SetFocus
            GoTo ExitHere
          End If
        Next rCnt
        
        'sprItem.GetText eItem.mRecdQty, Val(fSheetName), retVal
        retval = Val(SpVal(spOrder, Val(SchRow), eItem.mQty))
        If Val(retval) <> Spread_Col_Total(sprSchedule, eSch.mSchQty) Then
          spOrder.GetText eItem.mItem, Val(SchRow), retval
          fItemCode = retval
          sShowMessage "Schedule Qty is not matched with the Order Qty for the Item " & fItemCode
          txtitemCode.Text = fItemCode
          sprSchedule.ActiveSheet = CurSheet
          sprSchedule.SetActiveCell eSch.mSchQty, 1
          sprSchedule.SetFocus
          GoTo ExitHere
        End If
      End If
    End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmbordtype_Click()
  Try
  Me.MousePointer = vbHourglass

  Call SetDocNo
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetDocNo()
  Try

  txtordnum = GetDocNo("SAL_ORDER_HDR", "OrderType", "OrderNo", Trim$(cmbordtype), "O")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmbScheduleType_Click()
  Dim i As Long

  Try
  Me.MousePointer = vbHourglass

  For i = 1 To spOrder.MaxRows
    If cmbScheduleType = "Open" Then
      Call Spread_Lock(spOrder, True, eItem.mQty, i)
      spOrder.SetText eItem.mQty, i, ""
      fraSchedule.Visible = False
    Else ' Specific
      Call Spread_Lock(spOrder, False, eItem.mQty, i)
'      spOrder.SetText eItem.mQty, i, ""
      fraSchedule.Visible = True
    End If
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  Form_Clear Me
  CurrDate Me
  
  Select Case mCallMenu
    Case mSalesOrder
      Call LoadDocType(mSalOrder)
    Case mExportOrder
      Call LoadDocType(mSalOrderExport)
  End Select
  
  fraOrder.Enabled = True
  fraSource.Enabled = True
  fraQuotation.Enabled = True
  fraTerms.Enabled = False
  spOrder.Enabled = False
  cmdSave.Enabled = False
  fraSchedule.Enabled = False
  spOrder.MaxRows = 1
  
  sprSchedule.SheetCount = spOrder.MaxRows
  Call Spread_Clear(sprSchedule)
  sprSchedule.MaxRows = 1
  
  cmbScheduleType.Enabled = True
  cmbScheduleType.ListIndex = 0
  optdirect.Value = True
  fraQuotation.Visible = False
  txtcustcd.Enabled = True
  txtcustcd.SetFocus
  CurrDate Me
'  Loadordtype
  LoadDept
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim ordnorec As AceADODB.Recordset
  Dim i As Long
  Dim tmp As Variant
  Dim rCnt As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  Set cmd = New AceADODB.Command
  Set ordnorec = New AceADODB.Recordset
  
  CurDate = ServerDate
  
  If Chkbefsave = False Then
    GoTo ExitHere
  End If
  
  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetDocNo
  
  If optdirect Then
    tmp = "D"
  ElseIf optquotation Then
    tmp = "Q"
  ElseIf optWithOutItem Then
    tmp = "W"
  End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_INS_ORDER_HDR"
  cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
  cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
  cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("OrderSource", adChar, adParamInput, 1, Trim$(tmp))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("ReferanceNo", adVarChar, adParamInput, 20, Trim$(txtRefNo))
  cmd.Parameters.Append cmd.CreateParameter("ReferanceDate", adDBTimeStamp, adParamInput, , Format(dprefdate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(dpvalidfrom, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(dpvalidto, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("OrderingType", adVarChar, adParamInput, 10, Trim$(cmbScheduleType.Text))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtRemarks))
  cmd.Parameters.Append cmd.CreateParameter("FormCode", adInteger, adParamInput, , Val(txtForm.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PackingCode", adInteger, adParamInput, , Val(txtPack.Tag))
  cmd.Parameters.Append cmd.CreateParameter("DeliveryCode", adInteger, adParamInput, , Val(txtDelIns.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PaymentCode", adInteger, adParamInput, , Val(txtTerms.Tag))
  cmd.Parameters.Append cmd.CreateParameter("InsuranceCode", adInteger, adParamInput, , Val(txtInsurance.Tag))
  cmd.Parameters.Append cmd.CreateParameter("ModeCode", adInteger, adParamInput, , Val(txtMode.Tag))
  cmd.Parameters.Append cmd.CreateParameter("FreightCode", adInteger, adParamInput, , Val(txtFreight.Tag))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , Trim$(txtGrandTotal.Text))
  cmd.Parameters.Append cmd.CreateParameter("PostingAc", adInteger, adParamInput, , Val(txtPostingAc.Tag))
  cmd.Parameters.Append cmd.CreateParameter("RevisionLevel", adVarChar, adParamInput, 10, Trim$(txtRevisionLevel))
  cmd.Parameters.Append cmd.CreateParameter("GeneralDetails", adVarChar, adParamInput, 3000, Trim$(txtGeneralDetails))
  cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "N")
  cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute

    For i = 1 To spOrder.DataRowCnt
      commClear cmd
      cmd.CommandText = "SAL_INS_ORDER_DTL"
      cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
      cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
      cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
      spOrder.GetText mItem, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      spOrder.GetText eItem.mDesc, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemDescription", adVarChar, adParamInput, 50, Trim$(tmp))
      spOrder.GetText mPartNo, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
'      .GetText mRevNo, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adVarChar, adParamInput, 5, "0")  'Trim$(tmp))
      spOrder.GetText mUOMCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("UOM", adInteger, adParamInput, , Val(tmp))
      spOrder.GetText mQty, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
      spOrder.GetText mRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      spOrder.GetText mAssRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "N")
      cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
      
      If Trim$(cmbScheduleType.Text) = "Specific" Then
        If SaveSalOrderScheduleData(i) = False Then 'Sales Order Schedule Save
          GoTo ExitHere
        End If
      End If
      
      If optquotation Then 'Updating Quotation table
        commClear cmd
        cmd.CommandText = "SAL_UP_QUOTATION_DTL"
        cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuoType))
        cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
        cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(txtQuoDate, "yyyy-MM-dd"))
        spOrder.GetText 1, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adChar, adParamInput, 20, Trim$(tmp))
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
    Next i  
    For i = 1 To spDept.DataRowCnt
      spDept.GetText 3, i, tmp
      If Val(tmp) = 1 Then
        commClear cmd
        cmd.CommandText = "SAL_Ins_Order_Department"
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
        spDept.GetText 1, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("DepartmentCode", adInteger, adParamInput, , Val(tmp))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
    Next i  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If Trim$(cmbScheduleType.Text) = "Specific" Then
    If SaveCalcDetails = False Then
      GoTo ExitHere
    End If
  End If
  
  If mCallMenu = mExportOrder Then
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
    cmd.Parameters.Append cmd.CreateParameter("CarriageName", adVarChar, adParamInput, 50, Trim$(txtCarriage))
    cmd.Parameters.Append cmd.CreateParameter("CarriageNo", adVarChar, adParamInput, 20, Trim$(txtCarriageNo))
    cmd.Parameters.Append cmd.CreateParameter("DischargePort", adVarChar, adParamInput, 50, Trim$(txtPortName))
    cmd.Parameters.Append cmd.CreateParameter("ReceivingPlace", adVarChar, adParamInput, 50, Trim$(txtRecName))
    cmd.Parameters.Append cmd.CreateParameter("LoadingPort", adVarChar, adParamInput, 50, Trim$(txtLoadName))
    cmd.Parameters.Append cmd.CreateParameter("DestinationPlace", adVarChar, adParamInput, 50, Trim$(txtDesName))
    cmd.Parameters.Append cmd.CreateParameter("Reference", adVarChar, adParamInput, 20, Trim$(txtReference))
    cmd.Parameters.Append cmd.CreateParameter("Buyer", adVarChar, adParamInput, 50, Trim$(txtBuyerName))
    cmd.Parameters.Append cmd.CreateParameter("Consignee", adVarChar, adParamInput, 50, Trim$(txtConsignName))
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "R") ' for Sales order Record
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
          
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST_ADDRESS"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
    cmd.Parameters.Append cmd.CreateParameter("Type", adVarChar, adParamInput, 50, "ConSignee")
    cmd.Parameters.Append cmd.CreateParameter("Name", adVarChar, adParamInput, 50, Trim$(txtConsignName))
    cmd.Parameters.Append cmd.CreateParameter("Address1", adVarChar, adParamInput, 50, Trim$(txtAdd1))
    cmd.Parameters.Append cmd.CreateParameter("Address2", adVarChar, adParamInput, 50, Trim$(txtAdd2))
    cmd.Parameters.Append cmd.CreateParameter("Address3", adVarChar, adParamInput, 50, Trim$(txtAdd3))
    cmd.Parameters.Append cmd.CreateParameter("City", adVarChar, adParamInput, 20, Trim$(txtCity))
    cmd.Parameters.Append cmd.CreateParameter("PinCode", adVarChar, adParamInput, 10, Trim$(txtPinCode))
    cmd.Parameters.Append cmd.CreateParameter("Phone", adVarChar, adParamInput, 20, Trim$(txtPhone))
    cmd.Parameters.Append cmd.CreateParameter("Email", adVarChar, adParamInput, 50, Trim$(txtEmail))
    cmd.Parameters.Append cmd.CreateParameter("State", adVarChar, adParamInput, 30, Trim$(txtState))
    cmd.Parameters.Append cmd.CreateParameter("Country", adVarChar, adParamInput, 30, Trim$(txtCountry))
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "R")
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
      
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST_ADDRESS"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
    cmd.Parameters.Append cmd.CreateParameter("Type", adVarChar, adParamInput, 50, "Buyer")
    cmd.Parameters.Append cmd.CreateParameter("Name", adVarChar, adParamInput, 50, Trim$(txtBuyerName))
    cmd.Parameters.Append cmd.CreateParameter("Address1", adVarChar, adParamInput, 50, Trim$(txtBAdd1))
    cmd.Parameters.Append cmd.CreateParameter("Address2", adVarChar, adParamInput, 50, Trim$(txtBAdd2))
    cmd.Parameters.Append cmd.CreateParameter("Address3", adVarChar, adParamInput, 50, Trim$(txtBAdd3))
    cmd.Parameters.Append cmd.CreateParameter("City", adVarChar, adParamInput, 20, Trim$(txtBCity))
    cmd.Parameters.Append cmd.CreateParameter("PinCode", adVarChar, adParamInput, 10, Trim$(txtBPinCode))
    cmd.Parameters.Append cmd.CreateParameter("Phone", adVarChar, adParamInput, 20, Trim$(txtBPhone))
    cmd.Parameters.Append cmd.CreateParameter("Email", adVarChar, adParamInput, 50, Trim$(txtBEmail))
    cmd.Parameters.Append cmd.CreateParameter("State", adVarChar, adParamInput, 30, Trim$(txtBState))
    cmd.Parameters.Append cmd.CreateParameter("Country", adVarChar, adParamInput, 30, Trim$(txtBCountry))
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "R")
    Call AddCommonCmdParameters(cmd)
    cmd.Execute

      For rCnt = 1 To sprGoods.DataRowCnt
        commClear cmd
        cmd.CommandText = "SAL_INS_PACKINGLIST_GOODS"
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
        sprGoods.GetText eGoods.mContainerNo, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("ContainerNo", adVarChar, adParamInput, 20, Trim$(tmp))
        sprGoods.GetText eGoods.mPackageKind, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("PackageKind", adVarChar, adParamInput, 50, Trim$(tmp))
        sprGoods.GetText eGoods.mDescription, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("GoodsDescription", adVarChar, adParamInput, 50, Trim$(tmp))
        sprGoods.GetText eGoods.mQuantity, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Quantity", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eGoods.mNetWeight, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("NetWeight", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
        sprGoods.GetText eGoods.mLength, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Length", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eGoods.mWidth, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Width", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eGoods.mHeight, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Height", adInteger, adParamInput, , Val(tmp))
        cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "R")
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      Next rCnt
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Order Number :" & Trim$(cmbordtype) & Trim$(txtordnum) & "  Is Generated"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set ordnorec = Nothing
  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Function SaveCalcDetails() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    
    CurDate = ServerDate
    
    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(cmbordtype) & "' DocType," & Val(txtordnum) & " DocNo,"
    Sql = Sql & "'" & Format(dporddate, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,'',"
    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(cmbordtype) & "', " & Val(txtordnum) & ", '" & Format(dporddate, "yyyy-MM-dd") & "',"
    Sql = Sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, a.IsValueEditable, "
    Sql = Sql & " a.ValueUsed, a.IsValueEditable, '', a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  SaveCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub dporddate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dprefdate_Change()
' On Error Resume Next
  dpvalidfrom = dprefdate
End Sub

Private Sub dprefdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dpvalidfrom_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dpvalidto_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  
  Me.MousePointer = vbHourglass
  
  If TypeOf Me.ActiveControl Is fpSpread Then
    GoTo ExitHere
  ElseIf Me.ActiveControl.Name = "txtGeneralDetails" Then
    GoTo ExitHere
  End If
  If KeyAscii = 13 Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try

  CurrDate Me
'  Loadordtype
  LoadDept
  cmbScheduleType.ListIndex = 0
  SKey = Format(Now, "yyyyMMddHHmmss") & gUserID
  If mCallMenu = mExportOrder Then
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = True
    tabSales.TabVisible(2) = True
    tabSales.TabVisible(3) = True
  Else
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = True
    tabSales.TabVisible(2) = False
    tabSales.TabVisible(3) = False
  End If
  lblItemCode.Caption = "Item Code"
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

Private Sub optdirect_Click()
  Try
  
  fraQuotation.Visible = False
  txtcustcd.Enabled = True
  Call Spread_Lock(spOrder, True, eItem.mDesc, 1, mPartNo, -1)
  Call Spread_Lock(spOrder, True, eItem.mRate, 1, mAssRate, -1)
  lblItemCode.Caption = "Item Code"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optquotation_Click()
  Try
  
  fraQuotation.Visible = True
  txtcustcd.Enabled = False
  Call Spread_Lock(spOrder, True, eItem.mDesc, 1, mPartNo, -1)
  Call Spread_Lock(spOrder, True, eItem.mRate, 1, mAssRate, -1)
  lblItemCode.Caption = "Item Code"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optWithOutItem_Click()
  Try
  
  Call Spread_Lock(spOrder, False, eItem.mDesc, 1, mPartNo, -1)
  Call Spread_Lock(spOrder, False, eItem.mRate, 1, mAssRate, -1)
  lblItemCode.Caption = "Part No"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spOrder_Click(sender as object, e as ClickEventArgs)
  Dim retval As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If optWithOutItem.Value = True Then
    spOrder.GetText eItem.mPartNo, Row, retval
    If Trim$(retval) = "" Then
 GoTo ExitHere
End If
    txtitemCode.Text = Trim$(retval)
    retval = Spread_Get_SheetNoByName(sprSchedule, CStr(retval)) ' Retrieve Which Sheet no
    sprSchedule.ActiveSheet = Trim$(retval)
  Else
    spOrder.GetText eItem.mPartNo, Row, retval
    If Trim$(retval) = "" Then
 GoTo ExitHere
End If
    retval = Spread_Get_SheetNoByName(sprSchedule, CStr(retval)) ' Retrieve Which Sheet no
    sprSchedule.ActiveSheet = Trim$(retval)
    spOrder.GetText eItem.mItem, Row, retval
    txtitemCode.Text = Trim$(retval)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spOrder_DblClick(sender as object, e as DblClickEventArgs)
  Try
  Me.MousePointer = vbHourglass

  Sprhlp
  Avoid_Duplicate spOrder, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spOrder_EditMode(sender as object, e as EditModeEventArgs)
  Dim retval As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mPartNo
          spOrder.GetText eItem.mPartNo, Row, retval
          If Trim$(retval) <> "" Then
            Call SetSheet(Row)
            sprSchedule.SheetName = Trim$(retval)
            txtitemCode = Trim$(retval)
          End If
        Case eItem.mQty, eItem.mRate, eItem.mAssRate
          spOrder.GetText Col, Row, retval
          If Val(retval) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
'            Call StartCalcFormula(IIf(Trim$(txtGrandTotal.Text) = "", True, False))
          End If
          
'          Select Case col
'            Case eItem.mQty
'              If Val(retVal) > 0 Then
'                If optquotation = True Then
'                  .SetActiveCell eItem.mItem, Row + 1
'                Else
'                  .SetActiveCell eItem.mRate, Row
'                End If
'              End If
'            Case eItem.mRate
'              If Val(retVal) > 0 Then
'                .SetActiveCell eItem.mAssRate, Row
'              End If
'            Case eItem.mAssRate
'              If Val(retVal) > 0 Then
'                .SetActiveCell eItem.mItem, Row + 1
'              End If
'          End Select
      End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spOrder_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim tmp As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If KeyCode = vbKeyF2 Then
    If spOrder.ActiveCol = 1 Then Sprhlp '''Procedure for Help
    If spOrder.ActiveCol = 5 Then spOrder_DblClick spOrder.ActiveCol, spOrder.ActiveRow
  ElseIf spOrder.ActiveCol = 11 And KeyCode = 13 And optWithOutItem Then
    spOrder.MaxRows = spOrder.DataRowCnt + 1
  ElseIf KeyCode = 46 And spOrder.ActiveCol = 1 Then
    spOrder.GetText eItem.mPartNo, spOrder.ActiveRow, tmp
    tmp = Spread_Get_SheetNoByName(sprSchedule, CStr(tmp)) ' Retrieve Which Sheet no
    Call sprSchedule.DeleteSheets(Val(tmp), 1)
    
    Spread_Clear_Row spOrder, 1, spOrder.MaxCols, spOrder.ActiveRow, spOrder.ActiveRow
    spOrder.DeleteRows spOrder.ActiveRow, 1
    spOrder.MaxRows = spOrder.DataRowCnt + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spOrder_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retval As Variant
  Dim sFirstRow As Long

  Try
  Me.MousePointer = vbHourglass

  If Row <> NewRow Then
    SetSheet (NewRow)
    spOrder.GetText eItem.mItem, NewRow, retval
    txtitemCode.Text = Trim$(retval)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprGoods_EditMode(sender as object, e as EditModeEventArgs)
  Try
    If Mode = 0 Then
      sprGoods.MaxRows = sprGoods.DataRowCnt + 1
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprSchedule_EditMode(sender as object, e as EditModeEventArgs)
' On Error Resume Next
  If cmbScheduleType.ListIndex = 1 Then 'Specific
      If Mode = 0 And Col = eSch.mSchDate Then
        Call Spread_Clear_Duplicate(sprSchedule, Row, eSch.mSchDate)
        sprSchedule.MaxRows = IIf(sprSchedule.DataRowCnt = 0, 1, sprSchedule.DataRowCnt + 1)
      End If
  End If
End Sub

Private Sub sprSchedule_KeyUp(sender as object, e as KeyUpEventArgs)
  
  Try
  Me.MousePointer = vbHourglass
  
  If cmbScheduleType.ListIndex = 1 Then 'Specific
      If KeyCode = vbKeyDelete And sprSchedule.EditMode = False Then
        sprSchedule.DeleteRows sprSchedule.ActiveRow, 1
        sprSchedule.MaxRows = sprSchedule.DataRowCnt + 1
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustcd_Change()
  Dim snap As New AceADODB.Recordset
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If Trim(txtcustcd) = "" Then
 GoTo ExitHere
End If
  Sql = " select a.FormCode,(select b.Description from TERMSMASTER b where "
  Sql = Sql + " b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid "
  Sql = Sql + " and b.Compid=a.Compid) FormDesc, a.PackingCode,(select b1.Description "
  Sql = Sql + " from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding' "
  Sql = Sql + " and b1.Code=a.PackingCode and b1.Branchid=a.Branchid "
  Sql = Sql + " and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,(select b2.Description "
  Sql = Sql + " from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction' "
  Sql = Sql + " and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid "
  Sql = Sql + " and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,(select b3.Description "
  Sql = Sql + " from TERMSMASTER b3 where b3.TypeFlg='Payment Terms'"
  Sql = Sql + " and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid "
  Sql = Sql + " and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,(select "
  Sql = Sql + " b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance'"
  Sql = Sql + " and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid "
  Sql = Sql + " and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,(select "
  Sql = Sql + " b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid "
  Sql = Sql + " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,(select b6.Description "
  Sql = Sql + " from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode "
  Sql = Sql + " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc,a.CustomerFormulaCode,"
  Sql = Sql + " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where "
  Sql = Sql + " b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.CustomerFormulaCode and b7.Branchid=a.Branchid and b7.Compid=a.Compid) FrmLName"
  Sql = Sql + " from FAS_PARTY_MASTER a where a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.PartyCode = " & Val(txtcustcd.Tag) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  FrmLName = ""
  If Not snap.EOF Then
    If Not IsNull(snap("CustomerFormulaCode")) = True Then txtFormula.Tag = snap("CustomerFormulaCode")
    If optquotation Then
      If Not IsNull(snap("FrmLName")) = True Then FrmLName = snap("FrmLName")
    Else
      If Not IsNull(snap("FrmLName")) = True Then txtFormula.Text = snap("FrmLName")
    End If
    
    If Not IsNull(snap("DeliveryCode")) = True Then
 txtDelIns.Tag = snap("DeliveryCode")
End If
    If Not IsNull(snap("DeliveryDesc")) = True Then
 txtDelIns.Text = snap("DeliveryDesc")
End If
    
    If Not IsNull(snap("FormCode")) = True Then
 txtForm.Tag = snap("FormCode")
End If
    If Not IsNull(snap("FormDesc")) = True Then
 txtForm.Text = snap("FormDesc")
End If
    
    If Not IsNull(snap("PackingCode")) = True Then
 txtPack.Tag = snap("PackingCode")
End If
    If Not IsNull(snap("PackingDesc")) = True Then
 txtPack.Text = snap("PackingDesc")
End If
    
    If Not IsNull(snap("InsuranceCode")) = True Then
 txtInsurance.Tag = snap("InsuranceCode")
End If
    If Not IsNull(snap("InsuranceDesc")) = True Then
 txtInsurance.Text = snap("InsuranceDesc")
End If

    If Not IsNull(snap("FreightCode")) = True Then
 txtFreight.Tag = snap("FreightCode")
End If
    If Not IsNull(snap("FreightDesc")) = True Then
 txtFreight.Text = snap("FreightDesc")
End If

    If Not IsNull(snap("PaymentCode")) = True Then
 txtTerms.Tag = snap("PaymentCode")
End If
    If Not IsNull(snap("PaymentDesc")) = True Then
 txtTerms.Text = snap("PaymentDesc")
End If

    If Not IsNull(snap("ModeCode")) = True Then
 txtMode.Tag = snap("ModeCode")
End If
    If Not IsNull(snap("ModeDesc")) = True Then
 txtMode.Text = snap("ModeDesc")
End If
  End If
  snap.Close
  Set snap = Nothing
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustcd_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "select distinct (select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
  Sql = Sql + " and b.Compid=a.Compid) PartyName, a.CustomerCode from SAL_Customer_ItemMaster a "
  Sql = Sql + " where a.Branchid = '" & gBranchID & "' and a.compid='" & gCompID & "'"
  Sql = Sql & " order by a.CustomerCode"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer") = True Then
    txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
    fraOrder.Enabled = False
    fraSource.Enabled = False
    fraQuotation.Enabled = False
    fraTerms.Enabled = True
    spOrder.Enabled = True
    cmdSave.Enabled = True
    fraSchedule.Enabled = True
    txtRefNo.SetFocus
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustcd_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)

  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
    txtcustcd_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

'Private Sub Loadordtype()
'  Dim ordtyperec As ADODB.Recordset
'  Dim Sql As String
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Set ordtyperec = New ADODB.Recordset
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='O'"
'  Sql = Sql + " and ModuleName='SAL' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
'  Sql = Sql + " order by DocumentType desc"
'  If RSOpen(ordtyperec, Sql) = False Then GoTo ExitHere
'  If Not ordtyperec.EOF Then
'    cmbordtype.Clear
'    Do Until ordtyperec.EOF
'      cmbordtype.AddItem ordtyperec("DocumentType")
'      ordtyperec.MoveNext
'    Loop
'    cmbordtype.ListIndex = 0
'    txtordnum = GetNo("SAL_ORDER_HDR", "OrderType", "OrderNo", Trim$(cmbordtype), "O", "SAL")
'  End If
'  ordtyperec.Close
'ExitHere:
'  On Error GoTo 0
'  Set ordtyperec = Nothing
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

Private Sub txtDelIns_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Delivery Instruction' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Delivery Instruction", "Description") = True Then
    txtDelIns.Text = getvalue(1)
    txtDelIns.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtForm_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Form'"
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Form's", "Description") = True Then
    txtForm.Text = getvalue(1)
    txtForm.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_Change()
   
  Try
  Me.MousePointer = vbHourglass
  
  If spOrder.DataRowCnt <= 0 Then
    GoTo ExitHere                 'No need to calculate formula
  End If
  
  If Len(Trim$(txtFormula.Tag)) > 0 Then
    If StartCalcFormula = False Then
      GoTo ExitHere
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select FrmLName, FrmID from COM_FORMULAMAST_HDR"
  Sql = Sql & " where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula's", "Formula Name") = True Then
    txtFormula.Tag = getvalue(2)
    txtFormula.Text = getvalue(1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFormula_DblClick
End If
End Sub

Private Sub txtFreight_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Freight'"
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of FreightCode", "Description") = True Then
    txtFreight.Text = getvalue(1)
    txtFreight.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtPack_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtPack_DblClick
End If
End Sub

Private Sub txtInsurance_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtInsurance_DblClick
End If
End Sub

Private Sub txtDelIns_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtDelIns_DblClick
End If
End Sub

Private Sub txtForm_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtForm_DblClick
End If
End Sub

Private Sub txtFreight_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFreight_DblClick
End If
End Sub

Private Sub txtMode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtMode_DblClick
End If
End Sub

Private Sub txtPostingAc_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
      
  Sql = "Select LedgerName, LedgerCode from FAS_LedgerMaster Where Left(LrRoute, 2) = '0E'" 'left(LrRoute,4) not in ('0206', '0303') "
  Sql = Sql & " and isActive = '1' and CompCode = '" & gCompID & gBranchID & "'"
  Sql = Sql & " Order by LedgerName"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Accounts", "Account Name") = True Then
    txtPostingAc.Tag = getvalue(2)    'AcCode
    txtPostingAc.Text = getvalue(1)    'AcName
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtPostingAc_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtPostingAc_DblClick
End If
End Sub

Private Sub txtTerms_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtTerms_DblClick
End If
End Sub

Private Sub txtInsurance_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description,Code from TERMSMASTER Where Typeflg = 'Insurance' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Insurance", "Description") = True Then
    txtInsurance.Text = getvalue(1)
    txtInsurance.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtMode_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Mode' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtMode.Text = getvalue(1)
    txtMode.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtPack_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Packing & Forwarding' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtPack.Text = getvalue(1)
    txtPack.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtTerms_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Payment Terms' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Payment Terms", "Description") = True Then
    txtTerms.Text = getvalue(1)
    txtTerms.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtquono_DblClick()
  Dim quorec As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long

  Try
  Me.MousePointer = vbHourglass
  
'  Set quorec = New ADODB.Recordset
  Set snap1 = New AceADODB.Recordset
  
  Sql = "select distinct a.QuotationType,a.QuotationNo,a.QuotationDate,a.Revisionno,(select b.PartyCode "
  Sql = Sql + " from FAS_PARTY_MASTER b where b.PartyName = a.CustomerName "
  Sql = Sql + " and b.Branchid ='" & gBranchID & "' and b.Compid = '" & gCompID & "') PartyCode,"
  Sql = Sql + " CustomerName,ValidTo,a.FormulaCode, (select top 1 b7.FrmLName from "
  Sql = Sql + " COM_FormulaMast_Hdr b7 where b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.FormulaCode and b7.Branchid=a.Branchid "
  Sql = Sql + " and b7.Compid=a.Compid) FrmLName from SalesQuotation a"
  Sql = Sql + " where OrderNo = 0 and a.compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "111101000", "List of Quotation Nos.", "Quotation Type", "Quotation No", "Quotation Date", "Revisionno", "Customer") = True Then
    If getvalue(4) = 0 Then
      If MsgBox("Please create the Entry in the Customer Master for the Customer " & getvalue(5) & vbLf & "Do you Create Now?", vbYesNo, App.Title) = vbYes Then
        frmPartyMaster.Show
        frmPartyMaster.tbPartyName = getvalue(5)
      End If
      GoTo ExitHere
    End If
    txtQuoType = getvalue(1)
    txtQuoNo = getvalue(2)
    txtQuoDate = getvalue(3)
    dpvalidfrom.MinDate = txtQuoDate
    dpvalidto.MinDate = txtQuoDate
    txtcustcd.Tag = getvalue(5)
    txtcustcd = getvalue(6)
    txtQuoDate.Tag = getvalue(7)
    txtFormula.Tag = getvalue(8)
    txtFormula.Text = getvalue(9)
  
    Sql = "select (select top 1 b.ItemCode from INV_MATERIAL_MASTER b where "
    Sql = Sql + " b.ItemDescription=a.ItemDescription and b.Branchid=a.Branchid "
    Sql = Sql + " and b.Compid=a.Compid and b.status = 1) ItemCode,a.ItemDescription,(select top 1 c.PartNo "
    Sql = Sql + " from SAL_Customer_ItemMaster c where c.Itemcode=a.ItemCode "
    Sql = Sql + " and c.Branchid=a.Branchid and c.Compid=a.Compid and c.CustomerCode=a.CustomerCode) PartNo,"
    Sql = Sql + " d.UOMDesc,d.UOMDecimal,d.UOM,a.Qty,a.Rate,a.AssessableRate"
    Sql = Sql + " from SalesQuotation a left outer join MATERIAL_MASTER d "
    Sql = Sql + " on a.ItemCode=d.ItemCode and a.Branchid=d.Branchid "
    Sql = Sql + " and a.Compid=d.Compid where a.QuotationType='" & Trim$(txtQuoType) & "'"
    Sql = Sql + " and QuotationNo=" & Val(txtQuoNo) & " and QuotationDate="
    Sql = Sql + " convert(datetime,'" + Format(txtQuoDate, "yyyy-MM-dd") + "',111) and d.Status = 1"
    Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Set quorec = New AceADODB.Recordset
    If RSOpen(quorec, Sql) = False Then GoTo ExitHere
    If Not quorec.EOF Then
      quorec.MoveLast
quorec.MoveFirst
      spOrder.MaxRows = quorec.RecordCount
      For i = 1 To quorec.RecordCount
        If IsNull(quorec("ItemCode")) Then
          sShowMessage "Please create the Entry in the Item Master & Customer Item Master for the Item " & quorec("ItemDescription")
          Call CmdClear_Click
          GoTo ExitHere
        End If
        txtitemCode.Text = quorec("ItemCode")
        spOrder.SetText eItem.mItem, i, quorec("ItemCode")
        spOrder.SetText eItem.mDesc, i, quorec("ItemDescription")
        spOrder.SetText eItem.mPartNo, i, quorec("PartNo")
        spOrder.SetText eItem.mUOMCode, i, quorec("UOM")
        spOrder.SetText eItem.mUOM, i, quorec("UOMDesc")
        cmbScheduleType = "Specific"
'        cmbScheduleType.Enabled = False
        If Trim$(cmbScheduleType) = "Yearly" Then
          Call Spread_Lock(spOrder, True, eItem.mQty, i)
        Else
          Call Spread_Lock(spOrder, False, eItem.mQty, i)
        End If
'        Spread_Decimal_Set spOrder, 6, 6, i, i, quorec("UOMDecimal")
        Call Spread_Set_Decimal(spOrder, CheckNull(quorec("UOMDecimal"), 0), eItem.mQty, i)
'        Call Spread_Lock(spOrder, False, eItem.mQty, i, 8)
        spOrder.SetText eItem.mQty, i, quorec("Qty")
        spOrder.SetText eItem.mRate, i, quorec("Rate")
        spOrder.SetText eItem.mAssRate, i, quorec("AssessableRate")
        spOrder.SetText eItem.mGoodsValue, i, Format(quorec("Qty") * quorec("Rate"), "0.00")
        spOrder.SetText eItem.mAssGoodsValue, i, Format(quorec("Qty") * quorec("AssessableRate"), "0.00")
        Call SetSprGoodsValue(i)
        quorec.MoveNext
      Next i
    End If
    
    Call SetTotalGoodsValue
    Call CalcButtonStatus(Me, True)
    
'    OrderCalulation  'to calculate the Order Value
    fraOrder.Enabled = False
    fraSource.Enabled = False
    fraQuotation.Enabled = False
    fraTerms.Enabled = True
    spOrder.Enabled = True
    cmdSave.Enabled = True
    fraSchedule.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set quorec = Nothing
  Set snap1 = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtQuoNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
    txtquono_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtrefno_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = Asc(UCase(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtremarks_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = Asc(UCase(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim CurrDate As Date
  Dim ret As Integer
  Dim tmp As Variant
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  
  Try
  
  If Trim$(cmbordtype) = "" Then
    sShowMessage "Select the Order Type"
    cmbordtype.Enabled = True
    cmbordtype.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  If Trim$(txtcustcd) = "" Then
    sShowMessage "Select the Customer"
    txtcustcd.Enabled = True
    txtcustcd.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  If Trim$(txtRefNo) = "" Then
    sShowMessage "Enter Reference No"
    txtRefNo.Enabled = True
    txtRefNo.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  If dprefdate > Format(CurrDate, "dd/MM/yyyy") Then
    sShowMessage "Reference Date should be less the or equal to Current Date"
    dprefdate.Enabled = True
    dprefdate.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
'  If dpvalidfrom < CurrDate Then
'    sShowMessage "Valid From Date should be greater than or equal to Current Date"
'    dpvalidfrom.Enabled = True
'    dpvalidfrom.SetFocus
'    Chkbefsave = False
'    GoTo ExitHere
'  End If

  If dpvalidto < dpvalidfrom Then
    sShowMessage "Valid Till Date should be greater than or equal to From Date"
    dpvalidto.Enabled = True
    dpvalidto.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  If Not Trim$(txtQuoDate.Tag) = "" Then
    If dporddate > CDate(txtQuoDate.Tag) Then
      ret = MsgBox("Order Date is greater than Valid till Date of Quotation" & vbCrLf & " Do you want to continue creating the Order ?", vbYesNo + vbQuestion, App.Title)

      If Val(ret) <> 6 Then txtQuoNo = ""
txtQuoNo.SetFocus
GoTo ExitHere
    End If
  End If
  
  If Trim$(txtPostingAc) = "" Then
    sShowMessage "Select the Posting A/c"
    txtPostingAc.Enabled = True
    txtPostingAc.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  If Trim$(txtFormula) = "" Then
    sShowMessage "Please Select The Formula"
    txtFormula.Enabled = True
    txtFormula.SetFocus
    Chkbefsave = False
    GoTo ExitHere
  End If
  
'  If spCalc.DataRowCnt = 0 Then
'    sShowMessage "Please check the Calculation"
'    txtFormula.SetFocus
'    Chkbefsave = False
'    GoTo ExitHere
'  End If

  If spOrder.DataRowCnt = 0 Then
    sShowMessage "Select atleast one item"
    spOrder.Enabled = True
    spOrder.SetFocus
    Spread_Set_Focus spOrder, 1, 1
    Chkbefsave = False
    GoTo ExitHere
  End If

    For i = 1 To spOrder.DataRowCnt
      If optWithOutItem Then
        spOrder.GetText eItem.mDesc, i, tmp
        If Trim$(tmp) = "" Then
          sShowMessage "Enter the Item Description"
          spOrder.SetFocus
          Spread_Set_Focus spOrder, i, eItem.mDesc
          Chkbefsave = False
          GoTo ExitHere
        End If
        
        spOrder.GetText eItem.mPartNo, i, tmp
        If Trim$(tmp) = "" Then
          sShowMessage "Enter the Item Part No"
          spOrder.SetFocus
          Spread_Set_Focus spOrder, i, eItem.mPartNo
          Chkbefsave = False
          GoTo ExitHere
        End If
      End If
      
      If cmbScheduleType = "Specific" Then
        spOrder.GetText eItem.mQty, i, tmp
        If Val(tmp) = 0 Then
          sShowMessage "Enter the Quantity"
          Spread_Set_Focus spOrder, i, eItem.mQty
          Chkbefsave = False
          GoTo ExitHere
        End If
        
        spOrder.GetText eItem.mItem, i, tmp
        If ValidSchedule(Val(i)) = False Then
          GoTo ExitHere
        End If
      End If
      
      spOrder.GetText eItem.mRate, i, tmp
      If Val(tmp) = 0 Then
        sShowMessage "Enter the Rate"
        Spread_Set_Focus spOrder, i, eItem.mRate
        Chkbefsave = False
        GoTo ExitHere
      End If
      
      spOrder.GetText eItem.mAssRate, i, tmp
      If Val(tmp) = 0 Then
        sShowMessage "Enter the Ass. Rate"
        Spread_Set_Focus spOrder, i, eItem.mAssRate
        Chkbefsave = False
        GoTo ExitHere
      End If
    Next  
  ret = 0
    For i = 1 To spDept.DataRowCnt
      spDept.GetText 3, i, tmp
      If Val(tmp) = 1 Then
        ret = 1
      End If
    Next i
  
  If ret = 0 Then
    sShowMessage "Select atleast one Department"
    spDept.SetFocus
    tabSales.Tab = 4
    Spread_Set_Focus spDept, i, 3
    Chkbefsave = False
    GoTo ExitHere
  End If
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Sprhlp()
  Dim Sql As String
  Dim i As Long
  Dim sRow As Long
  
  Try
  
  If spOrder.ActiveCol = eItem.mItem And optdirect Then
    Sql = "select distinct a.ItemCode,c.ItemDescription,a.PartNo,c.UOM,c.UOMDesc,c.UOMDecimal,a.Rate,a.AssessableRate,a.FormulaCode,"
    Sql = Sql + "(select top 1 b.FrmLName from COM_FormulaMast_Hdr b where b.ModuleCode = " & gModuleCode & " and b.FrmID=a.FormulaCode"
    Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid) FromulaName from SAL_Customer_ItemMaster a "
    Sql = Sql + " left outer join MATERIAL_MASTER c on a.ItemCode = c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
    Sql = Sql + " where CustomerCode = '" & Trim$(txtcustcd.Tag) & "' and a.rate is not null and a.Branchid = '" & gBranchID & "'"
    Sql = Sql + " and a.compid='" & gCompID & "' and c.Status = 1"
    If FetchValue_New(True, db.ConnectionString, Sql, "11101010000", "List of Items for the Customer", "Item Code", "Item Description", "Part No", "UOM", "Rate") = True Then
      sRow = 0
      For i = 1 To UBound(getvalue, 1)
        If getvalue(i, 1) <> "" And (getvalue(1, 9) = getvalue(i, 9)) Then
          sRow = spOrder.DataRowCnt + 1
          spOrder.SetText eItem.mItem, sRow, getvalue(i, 1) 'ItemCode
          spOrder.SetText eItem.mDesc, sRow, getvalue(i, 2) 'ItemDesc
          spOrder.SetText eItem.mPartNo, sRow, getvalue(i, 3) 'Part No
          spOrder.SetText eItem.mUOMCode, sRow, getvalue(i, 4) 'UOMCode
          spOrder.SetText eItem.mUOM, sRow, getvalue(i, 5) 'UOM
          If Trim$(cmbScheduleType) = "Yearly" Then
            Call Spread_Lock(spOrder, True, eItem.mQty, sRow)
          Else
            Call Spread_Lock(spOrder, False, eItem.mQty, sRow)
          End If
'          Spread_Decimal_Set spOrder, eItem.mQty, eItem.mQty, SRow, SRow, Val(getvalue(i, 6)) 'UOM Decimal
          Call Spread_Set_Decimal(spOrder, Val(getvalue(i, 6)), eItem.mQty, sRow)  'UOM Decimal
          spOrder.SetText eItem.mQty, sRow, 0 ' 0 Qty
          Call Spread_Lock(spOrder, True, eItem.mRate, sRow, eItem.mAssRate)
          spOrder.SetText eItem.mRate, sRow, Val(getvalue(i, 7)) 'Rate
          spOrder.SetText eItem.mAssRate, sRow, Val(getvalue(i, 8)) 'Asseassable Rate
          spOrder.MaxRows = spOrder.DataRowCnt + 1
          
          If cmbScheduleType.ListIndex = 1 Then ' If specific only Schedule Come
            txtitemCode.Text = getvalue(i, 1) 'ItemCode
            Call SetSheet(sRow)
            sprSchedule.SheetName = Trim$(getvalue(i, 3))
          End If
        End If
      Next i
      If Trim(SpVal(spOrder, 1, 1)) <> "" Then
 Spread_Set_Focus spOrder, 1, 5
End If
    End If
  ElseIf spOrder.ActiveCol = eItem.mUOM And optWithOutItem Then
    Sql = "select SDesc, Code, NoofDecimals from INV_UOM_Master where Branchid = '" & gBranchID & "'"
    Sql = Sql + " and Compid = '" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "100", "List of UOM's", "UOM") = True Then
      spOrder.SetText mUOM, spOrder.ActiveRow, getvalue(1)
      spOrder.SetText mUOMCode, spOrder.ActiveRow, getvalue(2)
'      Spread_Decimal_Set spOrder, mQty, mQty, spOrder.ActiveRow, spOrder.ActiveRow, Val(getvalue(3))
      Call Spread_Set_Decimal(spOrder, Val(getvalue(3)), mQty, spOrder.ActiveRow)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retval As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(spOrder.DataRowCnt - 1)
      ReDim SCValList(spOrder.DataRowCnt * 2 - 1)
      
      For rCnt = 1 To spOrder.DataRowCnt
        'Item List
        spOrder.GetText eItem.mItem, rCnt, retval
        Item = Trim$(retval)
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        spOrder.GetText eItem.mQty, rCnt, retval
        Qty = Val(retval)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        spOrder.GetText eItem.mRate, rCnt, retval
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retval)
        
        spOrder.GetText eItem.mAssRate, rCnt, retval
        SCValList((rCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retval)
      Next rCnt
  Else
    ReDim ItemList(0)
    ItemList(0) = "0" & Chr(145) & "0"
    
    ReDim SCValList(1)
    SCValList(0) = eParamID.mGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Text)
    SCValList(1) = eParamID.mAGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Tag)
  End If

  If CalculateFormula(Val(txtcustcd.Tag), Val(txtFormula.Tag), txtUserKey, SCValList, ItemList, txtGrandTotal, sprCalc, eCalc.mDesc, eCalc.mValue, , "Sales Order", False, ShowPromptScreen) = False Then

    GoTo ExitHere
  End If
  
  Call CalcButtonStatus(Me, False)

  StartCalcFormula = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CalcButtonStatus(ByVal FrmObj As Form,ByVal  Enable As Boolean) As Boolean
  Try
  CalcButtonStatus = False
  
    If Enable Then
      FrmObj.cmdCalc.Enabled = True
      FrmObj.sprCalc.Visible = False
      FrmObj.txtGrandTotal.Visible = False
      FrmObj.lblGrandTotal.Visible = False
    Else
      FrmObj.cmdCalc.Enabled = False
      FrmObj.sprCalc.Visible = True
      FrmObj.txtGrandTotal.Visible = True
      FrmObj.lblGrandTotal.Visible = True
    End If
  
  CalcButtonStatus = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdCalc_Click()

  Try
  Me.MousePointer = vbHourglass

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  Call CalcButtonStatus(Me, False)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetSprGoodsValue(ByVal fRow As Long)
  Dim retval(1 To 3) As Variant
  Dim tmpTot As Double
    
  Try
  
    spOrder.GetText eItem.mQty, fRow, retval(1)
    retval(2) = 1
    
    spOrder.GetText eItem.mRate, fRow, retval(3)
    tmpTot = (Val(retval(1)) / Val(retval(2)) * Val(retval(3)))
    spOrder.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    spOrder.GetText eItem.mAssRate, fRow, retval(3)
    tmpTot = (Val(retval(1)) / Val(retval(2)) * Val(retval(3)))
    spOrder.SetText eItem.mAssGoodsValue, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(spOrder, eItem.mGoodsValue), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(spOrder, eItem.mAssGoodsValue), "0.00")
End Sub

Private Function SprTotCalc(ByRef fSpread As AceSpread,ByVal  fCol As Long) As Double
  Dim rCnt As Long
  Dim retval As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, rCnt, retval
      Total = Total + Val(retval)
    Next rCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub Design_sprSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprSchedule


sprSchedule.EditMode +=  new EventHandler(sprSchedule_EditMode)
sprSchedule.KeyUp +=  new EventHandler(sprSchedule_KeyUp)

End Sub

Private Sub Design_spDept(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDept

End Sub

Private Sub Design_sprGoods(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprGoods


sprGoods.EditMode +=  new EventHandler(sprGoods_EditMode)

End Sub

Private Sub Design_spOrder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spOrder


spOrder.Click +=  new EventHandler(spOrder_Click)
spOrder.DblClick +=  new EventHandler(spOrder_DblClick)
spOrder.EditMode +=  new EventHandler(spOrder_EditMode)
spOrder.KeyDown +=  new EventHandler(spOrder_KeyDown)
spOrder.LeaveCell +=  new EventHandler(spOrder_LeaveCell)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrder.frm", sprCalc, "")

	Call Design_sprSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrder.frm", sprSchedule, "")

	Call Design_spDept("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrder.frm", spDept, "")

	Call Design_sprGoods("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrder.frm", sprGoods, "")

	Call Design_spOrder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrder.frm", spOrder, "")

End Sub
