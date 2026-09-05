VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSalesInvoice 
   Caption         =   "Sales Invoice - Create"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
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
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraJORefDc 
      Height          =   1890
      Left            =   135
      TabIndex        =   129
      Top             =   7860
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtitemCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   330
         Width           =   1845
      End
      Begin VB.TextBox txtsentqty 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   131
         ToolTipText     =   "Enter the Vehicle No"
         Top             =   855
         Width           =   960
      End
      Begin VB.TextBox txtrecdqty 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   130
         ToolTipText     =   "Enter the PLA Sl No"
         Top             =   1395
         Width           =   960
      End
Begin AceSpread sprJODC
         Height          =   1665
         Left            =   45
         TabIndex        =   133
         Top             =   165
         Width           =   8385
         _Version        =   458752
         _ExtentX        =   14790
         _ExtentY        =   2937
         _StockProps     =   64
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         ScrollBars      =   2
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "ItemCode"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   8535
         TabIndex        =   136
         Top             =   150
         Width           =   660
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Sent Qty"
         Height          =   210
         Left            =   8520
         TabIndex        =   135
         Top             =   675
         Width           =   630
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Recd Qty"
         Height          =   210
         Left            =   8520
         TabIndex        =   134
         Top             =   1215
         Width           =   675
      End
   End
Begin AceSpread spDCDetails
      Height          =   1185
      Left            =   135
      TabIndex        =   65
      Top             =   9810
      Visible         =   0   'False
      Width           =   5625
      _Version        =   458752
      _ExtentX        =   9922
      _ExtentY        =   2090
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
      MaxCols         =   5
      ScrollBars      =   2
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source of Invoice"
      Height          =   705
      Left            =   6345
      TabIndex        =   58
      ToolTipText     =   "Select the Invoice Source"
      Top             =   75
      Width           =   2265
      Begin VB.OptionButton optdc 
         Caption         =   "D&C"
         Height          =   210
         Left            =   870
         TabIndex        =   4
         Top             =   300
         Width           =   525
      End
      Begin VB.OptionButton optorder 
         Caption         =   "&Order"
         Height          =   210
         Left            =   1470
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optdirect 
         Caption         =   "&Direct"
         Height          =   210
         Left            =   75
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   720
      End
   End
   Begin VB.Frame fraInvoice 
      Height          =   705
      Left            =   3405
      TabIndex        =   59
      Top             =   75
      Width           =   2910
      Begin VB.ComboBox cmbinvtype 
         Height          =   330
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Invoice Type"
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox txtinvnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Invoice No"
         Top             =   315
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dpinvdate 
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         ToolTipText     =   "Invoice Date"
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
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
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   137428993
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inv. No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1140
         TabIndex        =   64
         Top             =   135
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Inv. Type"
         Height          =   210
         Left            =   90
         TabIndex        =   61
         Top             =   135
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Inv. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1650
         TabIndex        =   60
         Top             =   135
         Width           =   630
      End
   End
   Begin VB.Frame fraCustomer 
      Height          =   885
      Left            =   195
      TabIndex        =   62
      Top             =   735
      Width           =   11235
      Begin VB.TextBox TxtVoucher 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   390
         Width           =   2790
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "SalesInvoice.frx":0883
         Left            =   8475
         List            =   "SalesInvoice.frx":08A8
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select the Type of Invoice"
         Top             =   390
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   330
         Left            =   9990
         TabIndex        =   8
         ToolTipText     =   "Click to OK"
         Top             =   390
         Width           =   420
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   390
         Width           =   4245
      End
      Begin MSComCtl2.DTPicker DtpVrDate 
         Height          =   330
         Left            =   7170
         TabIndex        =   125
         Top             =   390
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   137428993
         CurrentDate     =   39483
      End
      Begin VB.TextBox TxtVrRoute 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   127
         ToolTipText     =   "Invoice No"
         Top             =   375
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblVoucherDate 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Date"
         Height          =   210
         Left            =   7140
         TabIndex        =   126
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4335
         TabIndex        =   124
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblcustcode 
         AutoSize        =   -1  'True
         Caption         =   "* Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   63
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Type of Invoice"
         Height          =   210
         Left            =   8460
         TabIndex        =   66
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label lblVrRoute 
         AutoSize        =   -1  'True
         Caption         =   "Route"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7875
         TabIndex        =   128
         Top             =   225
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin TabDlg.SSTab tabSales 
      Height          =   5670
      Left            =   180
      TabIndex        =   67
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10001
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Order List"
      TabPicture(0)   =   "SalesInvoice.frx":0918
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "spItem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCalc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Packing Details"
      TabPicture(1)   =   "SalesInvoice.frx":0934
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCarriage"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Packing List"
      TabPicture(2)   =   "SalesInvoice.frx":0950
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblRef"
      Tab(2).Control(1)=   "sprGoods"
      Tab(2).Control(2)=   "txtReference"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame2 
         ForeColor       =   &H00FF0000&
         Height          =   3750
         Left            =   -70110
         TabIndex        =   110
         Top             =   375
         Width           =   5055
         Begin VB.TextBox txtState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   45
            ToolTipText     =   "Sales Quote Date"
            Top             =   2979
            Width           =   2790
         End
         Begin VB.TextBox txtCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   46
            ToolTipText     =   "Sales Quote Date"
            Top             =   3330
            Width           =   2790
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   43
            ToolTipText     =   "Sales Quote Date"
            Top             =   2283
            Width           =   1890
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   44
            ToolTipText     =   "Sales Quote Date"
            Top             =   2631
            Width           =   3660
         End
         Begin VB.TextBox txtAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            TabIndex        =   40
            ToolTipText     =   "Sales Quote Date"
            Top             =   1239
            Width           =   3660
         End
         Begin VB.TextBox txtPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            TabIndex        =   42
            ToolTipText     =   "Sales Quote Date"
            Top             =   1935
            Width           =   630
         End
         Begin VB.TextBox txtCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   41
            ToolTipText     =   "Sales Quote Date"
            Top             =   1587
            Width           =   1875
         End
         Begin VB.TextBox txtAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   38
            ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
            Top             =   543
            Width           =   3660
         End
         Begin VB.TextBox txtAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   39
            ToolTipText     =   "Sales Quote Date"
            Top             =   891
            Width           =   3660
         End
         Begin VB.TextBox txtConsignName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   37
            ToolTipText     =   "Enter the Buyer Name"
            Top             =   195
            Width           =   3660
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   915
            TabIndex        =   120
            Top             =   3030
            Width           =   375
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   720
            TabIndex        =   119
            Top             =   3375
            Width           =   570
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   840
            TabIndex        =   118
            Top             =   2340
            Width           =   450
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   930
            TabIndex        =   117
            Top             =   2685
            Width           =   360
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   705
            TabIndex        =   116
            Top             =   1980
            Width           =   585
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1020
            TabIndex        =   115
            Top             =   1635
            Width           =   270
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   570
            TabIndex        =   114
            Top             =   1290
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Consignee Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   113
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   570
            TabIndex        =   112
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   570
            TabIndex        =   111
            Top             =   945
            Width           =   720
         End
      End
      Begin VB.Frame fraCalc 
         BorderStyle     =   0  'None
         Height          =   2940
         Left            =   7890
         TabIndex        =   106
         Top             =   2400
         Width           =   3660
         Begin VB.TextBox txtGoodsValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1875
            Locked          =   -1  'True
            TabIndex        =   24
            ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
            Top             =   15
            Width           =   1470
         End
         Begin VB.TextBox txtGrandTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   26
            Top             =   2445
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.CommandButton cmdCalc 
            Caption         =   "C&alculate"
            Enabled         =   0   'False
            Height          =   330
            Left            =   465
            TabIndex        =   23
            Top             =   0
            Width           =   900
         End
Begin AceSpread sprCalc
            Height          =   2085
            Left            =   45
            TabIndex        =   25
            Top             =   345
            Visible         =   0   'False
            Width           =   3570
            _Version        =   458752
            _ExtentX        =   6297
            _ExtentY        =   3678
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
         Begin VB.TextBox txtUserKey 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   600
            TabIndex        =   107
            Top             =   15
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblUserKey 
            AutoSize        =   -1  'True
            Caption         =   "User Key"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   109
            Top             =   30
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblGrandTotal 
            AutoSize        =   -1  'True
            Caption         =   "Grand Total"
            Height          =   195
            Left            =   1020
            TabIndex        =   108
            Top             =   2505
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73605
         TabIndex        =   53
         ToolTipText     =   "Sales Quotation Type"
         Top             =   435
         Width           =   2430
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00FF0000&
         Height          =   3750
         Left            =   -74850
         TabIndex        =   68
         Top             =   375
         Width           =   4725
         Begin VB.TextBox txtBuyerName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   27
            ToolTipText     =   "Enter the Buyer Name"
            Top             =   195
            Width           =   3660
         End
         Begin VB.TextBox txtBAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   29
            ToolTipText     =   "Sales Quote Date"
            Top             =   891
            Width           =   3660
         End
         Begin VB.TextBox txtBAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   28
            ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
            Top             =   543
            Width           =   3660
         End
         Begin VB.TextBox txtBCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   20
            TabIndex        =   31
            ToolTipText     =   "Sales Quote Date"
            Top             =   1587
            Width           =   1875
         End
         Begin VB.TextBox txtBPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            TabIndex        =   32
            ToolTipText     =   "Sales Quote Date"
            Top             =   1935
            Width           =   630
         End
         Begin VB.TextBox txtBAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            TabIndex        =   30
            ToolTipText     =   "Sales Quote Date"
            Top             =   1239
            Width           =   3660
         End
         Begin VB.TextBox txtBEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   50
            TabIndex        =   34
            ToolTipText     =   "Sales Quote Date"
            Top             =   2631
            Width           =   3660
         End
         Begin VB.TextBox txtBPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   20
            TabIndex        =   33
            ToolTipText     =   "Sales Quote Date"
            Top             =   2283
            Width           =   1890
         End
         Begin VB.TextBox txtBCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   30
            TabIndex        =   36
            ToolTipText     =   "Sales Quote Date"
            Top             =   3330
            Width           =   2790
         End
         Begin VB.TextBox txtBState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   30
            TabIndex        =   35
            ToolTipText     =   "Sales Quote Date"
            Top             =   2979
            Width           =   2790
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   78
            Top             =   943
            Width           =   720
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   77
            Top             =   595
            Width           =   720
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Buyer Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   76
            Top             =   247
            Width           =   885
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   75
            Top             =   1291
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   690
            TabIndex        =   74
            Top             =   1639
            Width           =   270
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   375
            TabIndex        =   73
            Top             =   1987
            Width           =   585
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   72
            Top             =   2683
            Width           =   360
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   510
            TabIndex        =   71
            Top             =   2335
            Width           =   450
         End
         Begin VB.Label lblcountry 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   390
            TabIndex        =   70
            Top             =   3382
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   585
            TabIndex        =   69
            Top             =   3031
            Width           =   375
         End
      End
Begin AceSpread sprGoods
         Height          =   2670
         Left            =   -74880
         TabIndex        =   54
         Top             =   810
         Width           =   11415
         _Version        =   458752
         _ExtentX        =   20135
         _ExtentY        =   4710
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
         ScrollBars      =   2
      End
Begin AceSpread spOrder
         Height          =   1740
         Left            =   -74925
         TabIndex        =   98
         ToolTipText     =   "Enter the Sales Item Details"
         Top             =   2175
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
         MaxCols         =   10
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SelectBlockOptions=   0
      End
Begin AceSpread spItem
         Height          =   1905
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   "Enter the Invoice Details"
         Top             =   435
         Width           =   11355
         _Version        =   458752
         _ExtentX        =   20029
         _ExtentY        =   3360
         _StockProps     =   64
         Enabled         =   0   'False
         ColsFrozen      =   2
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
         MaxCols         =   19
      End
      Begin VB.Frame fraDetails 
         Enabled         =   0   'False
         Height          =   3240
         Left            =   135
         TabIndex        =   79
         Top             =   2280
         Width           =   7740
         Begin VB.TextBox txtLCBalAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   5970
            MaxLength       =   20
            TabIndex        =   144
            ToolTipText     =   "Enter the No of Packets"
            Top             =   1380
            Width           =   1710
         End
         Begin VB.TextBox txtLCAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   5970
            MaxLength       =   20
            TabIndex        =   142
            ToolTipText     =   "Enter the No of Packets"
            Top             =   990
            Width           =   1710
         End
         Begin VB.TextBox txtLCNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5970
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   138
            ToolTipText     =   "Enter the PLA Sl No"
            Top             =   210
            Width           =   1710
         End
         Begin VB.TextBox txtFormula 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   22
            ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help."
            Top             =   2835
            Width           =   2760
         End
         Begin VB.TextBox txtMode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
            Top             =   1413
            Width           =   2040
         End
         Begin VB.TextBox txtAI1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   10
            ToolTipText     =   "Enter the Additional Information1"
            Top             =   195
            Width           =   3330
         End
         Begin VB.TextBox txtAI2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            MaxLength       =   50
            TabIndex        =   11
            ToolTipText     =   "Enter the Additional Information1"
            Top             =   541
            Width           =   3330
         End
         Begin VB.TextBox txtNarration 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   1230
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "Enter the Narration"
            Top             =   887
            Width           =   3330
         End
         Begin VB.TextBox txtVehicleNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            MaxLength       =   15
            TabIndex        =   14
            ToolTipText     =   "Enter the Vehicle No"
            Top             =   1759
            Width           =   1425
         End
         Begin VB.TextBox txtSlNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3450
            MaxLength       =   10
            TabIndex        =   15
            ToolTipText     =   "Enter the PLA Sl No"
            Top             =   1785
            Width           =   960
         End
         Begin VB.TextBox txtPackets 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4980
            MaxLength       =   20
            TabIndex        =   21
            ToolTipText     =   "Enter the No of Packets"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.TextBox txtInsPolicyNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   17
            ToolTipText     =   "Insurance Policy No"
            Top             =   2115
            Width           =   1860
         End
         Begin VB.TextBox txtInsFor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   19
            ToolTipText     =   "Enter the Insured Amount"
            Top             =   2460
            Width           =   1125
         End
         Begin VB.TextBox txtInsSlNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2910
            MaxLength       =   10
            TabIndex        =   20
            ToolTipText     =   "Enter the Insurance Sl No"
            Top             =   2490
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker dtpPLADate 
            Height          =   330
            Left            =   4440
            TabIndex        =   16
            ToolTipText     =   "Enter the PLA Date"
            Top             =   1785
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   145555457
            CurrentDate     =   38238
         End
         Begin MSComCtl2.DTPicker dtpInsDate 
            Height          =   315
            Left            =   3105
            TabIndex        =   18
            ToolTipText     =   "Enter the Policy Date"
            Top             =   2130
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   145555457
            CurrentDate     =   38238
         End
         Begin MSComCtl2.DTPicker dtpLCDate 
            Height          =   330
            Left            =   5970
            TabIndex        =   139
            ToolTipText     =   "Enter the PLA Date"
            Top             =   600
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   145555457
            CurrentDate     =   41759.5153240741
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Balance LC Amt."
            Height          =   210
            Left            =   4710
            TabIndex        =   145
            Top             =   1425
            Width           =   1200
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "LC Amt."
            Height          =   210
            Left            =   5340
            TabIndex        =   143
            Top             =   1065
            Width           =   570
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "LC Date"
            Height          =   210
            Left            =   5340
            TabIndex        =   141
            Top             =   645
            Width           =   570
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "LC No"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   5475
            TabIndex        =   140
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblFormulaName 
            AutoSize        =   -1  'True
            Caption         =   "* Formula Name"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   60
            TabIndex        =   90
            Top             =   2850
            Width           =   1125
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Mode"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   765
            TabIndex        =   89
            Top             =   1470
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Information1"
            Height          =   210
            Left            =   315
            TabIndex        =   88
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Information2"
            Height          =   210
            Left            =   270
            TabIndex        =   87
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Narration"
            Height          =   210
            Left            =   495
            TabIndex        =   86
            Top             =   885
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Vehicle No"
            Height          =   210
            Left            =   375
            TabIndex        =   85
            Top             =   1815
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "PLA Sl No"
            Height          =   210
            Left            =   2730
            TabIndex        =   84
            Top             =   1830
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No of Pkts"
            Height          =   210
            Left            =   4200
            TabIndex        =   83
            Top             =   2550
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Policy No"
            Height          =   210
            Left            =   495
            TabIndex        =   82
            Top             =   2175
            Width           =   660
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Insured For"
            Height          =   210
            Left            =   330
            TabIndex        =   81
            Top             =   2520
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Sl. No."
            Height          =   210
            Left            =   2385
            TabIndex        =   80
            Top             =   2535
            Width           =   465
         End
      End
      Begin VB.Frame fraCarriage 
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   -74850
         TabIndex        =   91
         Top             =   4065
         Width           =   7725
         Begin VB.TextBox txtPortName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   49
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.TextBox txtDesName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   52
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.TextBox txtLoadName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   51
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtRecName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   50
            ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
            Top             =   180
            Width           =   2430
         End
         Begin VB.TextBox txtCarriageNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   48
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtCarriage 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   47
            ToolTipText     =   "Sales Quotation Type"
            Top             =   180
            Width           =   2430
         End
         Begin VB.Label lblDesName 
            AutoSize        =   -1  'True
            Caption         =   "Final Destination"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4020
            TabIndex        =   97
            Top             =   922
            Width           =   1170
         End
         Begin VB.Label lblLoadName 
            AutoSize        =   -1  'True
            Caption         =   "Port of Loading"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4080
            TabIndex        =   96
            Top             =   577
            Width           =   1110
         End
         Begin VB.Label lblPortName 
            AutoSize        =   -1  'True
            Caption         =   "Port of Discharge"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   95
            Top             =   922
            Width           =   1260
         End
         Begin VB.Label lblCarriage 
            AutoSize        =   -1  'True
            Caption         =   "Pre Carriage by"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   94
            Top             =   232
            Width           =   1125
         End
         Begin VB.Label lblReceivingName 
            AutoSize        =   -1  'True
            Caption         =   "Place Of Receipt"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3990
            TabIndex        =   93
            Top             =   232
            Width           =   1215
         End
         Begin VB.Label lblCarriageNo 
            AutoSize        =   -1  'True
            Caption         =   "Vessel / Filght No"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   92
            Top             =   577
            Width           =   1260
         End
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         Caption         =   "Other Reference"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   99
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker dtpRefDate 
      Height          =   315
      Left            =   6255
      TabIndex        =   102
      ToolTipText     =   "Enter the Quotation Date"
      Top             =   6765
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   146538499
      CurrentDate     =   37250
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   7455
      Width           =   1300
   End
   Begin VB.TextBox txtRefType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   4485
      Locked          =   -1  'True
      TabIndex        =   101
      ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
      Top             =   6765
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtRefNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5385
      Locked          =   -1  'True
      TabIndex        =   100
      ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
      Top             =   6765
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5370
      TabIndex        =   56
      ToolTipText     =   "Click to Clear the Data"
      Top             =   7455
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4080
      TabIndex        =   55
      ToolTipText     =   "Click to Save the Data"
      Top             =   7455
      Width           =   1300
   End
Begin AceSpread SprRecdItem
      Height          =   1185
      Left            =   10665
      TabIndex        =   137
      Top             =   7965
      Visible         =   0   'False
      Width           =   3705
      _Version        =   458752
      _ExtentX        =   6535
      _ExtentY        =   2090
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
      MaxCols         =   2
      ScrollBars      =   2
   End
   Begin VB.Label Label29 
      Height          =   495
      Left            =   5340
      TabIndex        =   122
      Top             =   4050
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   495
      Left            =   5175
      TabIndex        =   121
      Top             =   3735
      Width           =   1215
   End
   Begin VB.Label lblRefNo 
      AutoSize        =   -1  'True
      Caption         =   "Ord. No"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5655
      TabIndex        =   105
      Top             =   6570
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblRefType 
      AutoSize        =   -1  'True
      Caption         =   "Ord. Type"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4485
      TabIndex        =   104
      Top             =   6570
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblOrdDate 
      AutoSize        =   -1  'True
      Caption         =   "Ord. Date"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   6255
      TabIndex        =   103
      Top             =   6570
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmSalesInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
 
Public isRowRetrived As Boolean
Public mCallMenu As eSalesCallType

Private SKey As String
Private RefUserKey As String

Private Const MaxItemCount As Integer = 8

Private Enum eItem
  mItemCode = 1
  mItemDesc = 2
  mWhsCode = 3
  mWhsName = 4
  mRackCode = 5
  mUOM = 6
  mStockQty = 7
  mInvQty = 8
  mRate = 9
  mARate = 10
  mGV = 11
  mAGV = 12
  mNoofPkts = 13
  mOType = 14
  mONo = 15
  mODate = 16
  mBalQty = 17
  mSchType = 18
  mSetNo = 19
End Enum

Private Enum eItemGoods
  mContainerNo = 1
  mPackageKind = 2
  mDescription = 3
  mQuantity = 4
  mNetWeight = 5
  mLength = 6
  mWidth = 7
  mHeight = 8
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Enum eDc
  mDCNo = 1
  mDCDate = 2
  mrecdItem = 3
  mNorms = 4
  mDCQty = 5
  mBalQty = 6
  mSelectQty = 7
End Enum

Private Sub DeleteTmpFormulaCalcRef()
  Dim Sql As String
  
  Try
  
  Sql = " Delete Com_TmpFormulaCalc_Ref Where UserKey = '" & Trim$(RefUserKey) & "' and ModuleCode = " & gModuleCode & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'" 'and Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function InsertTmpFormulaCalcRef() As Boolean
  Dim Sql As String
  
  Dim rCnt As Integer
  Dim retVal As Variant
  
  Dim OrdType As String
  Dim ORDNo As Integer
  Dim OrdDate As Date
  
  Dim CurDate As Date
  
  Try
  InsertTmpFormulaCalcRef = False
  
  RefUserKey = CreateUserKey
  CurDate = ServerDate
  
    For rCnt = 1 To spItem.DataRowCnt
      Sql = " Insert into Com_TmpFormulaCalc_Ref  "
      Sql = Sql & " Select '" & Trim$(RefUserKey) & "', "
      Sql = Sql & " " & gModuleCode & ","
      spItem.GetText eItem.mOType, rCnt, retVal
      OrdType = retVal
      Sql = Sql & " '" & Trim$(OrdType) & "', "
      spItem.GetText eItem.mONo, rCnt, retVal
      ORDNo = retVal
      Sql = Sql & " " & Val(ORDNo) & ", "
      spItem.GetText eItem.mODate, rCnt, retVal
'      OrdDate = Format(retVal)
      Sql = Sql & " '" & Format(OrdDate, "yyyy-MM-dd") & "',"
      Sql = Sql & " '1', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
      Sql = Sql & " Where Not Exists (Select * From Com_TmpFormulaCalc_Ref "
      Sql = Sql & " Where UserKey = '" & Trim$(RefUserKey) & "' and ModuleCode = " & gModuleCode & ""
      Sql = Sql & " and DocType = '" & Trim$(OrdType) & "' and DocNo = " & Val(ORDNo) & " and DocDate = '" & Format(OrdDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "')" 'and Finyear = '" & gFinyear & "' )"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    Next rCnt

  'Insert for ValRefType
  Sql = " Insert into Com_TmpFormulaCalc_Ref Values "
  Sql = Sql & " ('" & Trim$(RefUserKey) & "', " & gModuleCode & ", '" & Trim$(cmbinvtype.Text) & "', " & Val(txtinvnum.Text) & ", '" & Format(dpinvdate, "yyyy-MM-dd") & "', "
  Sql = Sql & " '0', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  InsertTmpFormulaCalcRef = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function LoadPackingDetails() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim Cnt As Integer
  Dim Order(2) As String
  
  Try

  LoadPackingDetails = False
  
'  Sql = " Select Type, Name, Address1, Address2, Address3, City, PinCode, State, Country, Phone, Email "
'  Sql = Sql & " From SAL_PACKINGLIST_ADDRESS"
'  Sql = Sql & " Where OrderType = 'INV' and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "' "
'  Sql = Sql & " and Finyear = '" & gFinyear & "' and OrderNo = (Select Max(OrderNo) as OrderNo"
'  Sql = Sql & " From SAL_PACKINGLIST_ADDRESS Where OrderType = 'INV' and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "') "

  Sql = " Select b.Type, b.Name, b.Address1, b.Address2, b.Address3, b.City, b.PinCode, b.State, b.Country, B.Phone, B.Email,"
  Sql = Sql & " b.OrderNo, b.OrderDate, b.OrderType"
  Sql = Sql & " From SAL_PACKINGLIST_ADDRESS b"
  Sql = Sql & " Left Outer Join SAL_Invoice_Hdr a"
  Sql = Sql & " On b.OrderDate = a.InvoiceDate and b.OrderNo = a.InvoiceNo and b.OrderType = a.InvoiceType"
  Sql = Sql & " and b.BranchID = a.BranchID and b.CompID = a.CompID"
  Sql = Sql & " Where a.CustomerCode = " & Trim$(txtcustcd.Tag) & " and b.BranchId = '" & gBranchID & "' and b.CompId = '" & gCompID & "'" 'and a.Finyear = '" & gFinyear & "'"
  Sql = Sql & " and b.OrderNo = (Select Max(OrderNo) as OrderNo"
  Sql = Sql & " From SAL_PACKINGLIST_ADDRESS Where OrderType = (Select DocumentType From Com_DocumentType_Master"
  Sql = Sql & " Where ModuleName = 'SAL' and PertainingTo = 'I' and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "')"
  Sql = Sql & " and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "')"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    Order(0) = rs("OrderNo")
    Order(1) = rs("OrderDate")
    Order(2) = rs("OrderType")
    For Cnt = 1 To rs.RecordCount
      If rs("type") = "Buyer" Then
        txtBuyerName.Text = Trim$(rs("Name"))
        txtBAdd1.Text = CheckNull(rs("Address1"))
        txtBAdd2.Text = CheckNull(rs("Address2"))
        txtBAdd3.Text = CheckNull(rs("Address3"))
        txtBCity.Text = CheckNull(rs("City"))
        txtBPhone.Text = CheckNull(rs("Phone"))
        txtBEmail.Text = CheckNull(rs("Email"))
        txtBCountry.Text = CheckNull(rs("Country"))
      ElseIf rs("type") = "Consignee" Then
        txtConsignName.Text = Trim$(rs("Name"))
        txtAdd1.Text = CheckNull(rs("Address1"))
        txtAdd2.Text = CheckNull(rs("Address2"))
        txtAdd3.Text = CheckNull(rs("Address3"))
        txtCity.Text = CheckNull(rs("City"))
        txtPhone.Text = CheckNull(rs("Phone"))
        txtEmail.Text = CheckNull(rs("Email"))
        txtCountry.Text = CheckNull(rs("Country"))
      End If
      rs.MoveNext
    Next Cnt
  
'    Sql = " Select CarriageNo, CarriageName, DischargePort, ReceivingPlace, LoadingPort, DestinationPlace From SAL_PACKINGLIST"
'    Sql = Sql & " Where OrderType = 'INV' and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "' "
'    Sql = Sql & " and Finyear = '" & gFinyear & "' and OrderNo = (Select Max(OrderNo) as OrderNo "
'    Sql = Sql & " From SAL_PACKINGLIST Where  OrderType = 'INV' and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "') "

    Sql = " Select CarriageNo, CarriageName, DischargePort, ReceivingPlace, LoadingPort, DestinationPlace "
    Sql = Sql & " From SAL_PACKINGLIST"
    Sql = Sql & " Where OrderNo = " & Order(0) & " and OrderDate = '" & Format(Order(1), "yyyy-MM-dd") & "' and OrderType = '" & Trim$(Order(2)) & "' "
    Sql = Sql & " and BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "'" 'and Finyear = '" & gFinyear & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
      txtCarriage.Text = Trim$(rs("CarriageName"))
      txtCarriageNo.Text = CheckNull(rs("CarriageNo"))
      txtPortName.Text = CheckNull(rs("DischargePort"))
      txtRecName.Text = CheckNull(rs("ReceivingPlace"))
      txtLoadName.Text = CheckNull(rs("LoadingPort"))
      txtDesName.Text = CheckNull(rs("DestinationPlace"))
'      rs.MoveNext
    End If
  End If
  
  LoadPackingDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub PackingList(ByVal OrderType As String,ByVal  OrderNo As Integer,ByVal  OrderDate As Date)
  Dim Sql As String
  Dim CurDate As Date
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  isRowRetrived = False
  
  Set rs = New AceADODB.Recordset
  Sql = "select CarriageName, CarriageNo, DischargePort, ReceivingPlace, LoadingPort,"
  Sql = Sql & " DestinationPlace, Reference, Buyer, Consignee from Sal_PackingList"
  Sql = Sql & " where OrderType='" & OrderType & "' and OrderNo=" & OrderNo & ""
  Sql = Sql & " and OrderDate='" & Format(OrderDate, "yyyy-MM-dd") & "' and  OrderFlag='R'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
'  Set rs = New ADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If

  If Not rs.EOF Then
    txtCarriage = rs("CarriageName")
    txtCarriageNo = rs("CarriageNo")
    txtPortName = rs("DischargePort")
    txtRecName = rs("ReceivingPlace")
    txtLoadName = rs("LoadingPort")
    txtDesName = rs("DestinationPlace")
    txtReference = rs("Reference")
    txtBuyerName = rs("Buyer")
    txtConsignName = rs("Consignee")
    isRowRetrived = True
  End If

  Sql = "select ContainerNo, PackageKind, GoodsDescription,"
  Sql = Sql & "Quantity, NetWeight, Length, Width, Height from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrderType & "' and OrderNo=" & OrderNo & ""
  Sql = Sql & " and OrderDate='" & Format(OrderDate, "yyyy-MM-dd") & "' and OrderFlag='R'" 'For Invoice Record
  Sql = Sql & " and Branchid='" & gBranchID & "' and  Compid='" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
        
  If Not rs.EOF Then
    isRowRetrived = True
    rs.MoveFirst
      sprGoods.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprGoods.SetText eItemGoods.mContainerNo, rCnt, rs("ContainerNo")
        sprGoods.SetText eItemGoods.mPackageKind, rCnt, rs("PackageKind")
        sprGoods.SetText eItemGoods.mDescription, rCnt, rs("GoodsDescription")
        sprGoods.SetText eItemGoods.mQuantity, rCnt, rs("Quantity")
        sprGoods.SetText eItemGoods.mNetWeight, rCnt, rs("NetWeight")
        sprGoods.SetText eItemGoods.mLength, rCnt, rs("Length")
        sprGoods.SetText eItemGoods.mWidth, rCnt, rs("Width")
        sprGoods.SetText eItemGoods.mHeight, rCnt, rs("Height")
      rs.MoveNext
      Next rCnt
  End If

  Sql = "select Type ,Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email from Sal_PackingList_Address"
  Sql = Sql & " where OrderType='" & OrderType & "' and OrderNo=" & OrderNo & ""
  Sql = Sql & " and OrderDate='" & Format(OrderDate, "yyyy-MM-dd") & "' and OrderFlag='R'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Do While Not rs.EOF
    If rs("Type") = "ConSignee" Then
      txtConsignName = rs("Name")
      txtAdd1 = rs("Address1")
      txtAdd2 = rs("Address2")
      txtAdd3 = rs("Address3")
      txtCity = rs("City")
      txtPinCode = rs("PinCode")
    ElseIf rs("Type") = "Buyer" Then
      txtBuyerName = rs("Name")
      txtBAdd1 = rs("Address1")
      txtBAdd2 = rs("Address2")
      txtBAdd3 = rs("Address3")
      txtBCity = rs("City")
      txtBPinCode = rs("PinCode")
    End If
    rs.MoveNext
  Loop
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
 End Sub
 
Private Sub Orders()
  Dim Sql As String
  Dim CurDate As Date
  Dim rs As AceADODB.Recordset
  Dim i As Integer
  Dim j As Long
  Dim sRow As Long
  
  Dim OrdType As String
  Dim ORDNo As Integer
  Dim OrdDate As Date
  
  Try
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
  If mCallMenu = mSalesOrder Then
    Sql = "Execute SalesOrderList '" + Format(CurDate, "yyyy-MM-dd") + "', " & Val(txtcustcd.Tag) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "'"
  Else
    Sql = "Execute SalesExportOrderList '" + Format(CurDate, "yyyy-MM-dd") + "'," & Val(txtcustcd.Tag) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "'"
  End If
  
  If FetchValue_New(True, db.ConnectionString, Sql, "111000101111000110", "List of Item's", "Part No", "Item Code", "Item Description", "Rate", "Order Type", "Order No", "Order Date", "Balance Qty", "Reference No", "SetNo") = True Then
    spItem.MaxRows = UBound(getvalue, 1)
    
    If ValidateItemCount(MaxItemCount, spItem.MaxRows) = True Then
      For i = 1 To UBound(getvalue, 1)
  '      If Trim(getvalue(i, 9)) = Trim(getvalue(1, 9)) And Val(getvalue(i, 10)) = Val(getvalue(1, 10)) And Trim(getvalue(i, 11)) = Trim(getvalue(1, 11)) Then
          If i = 1 Then
            txtFormula.Tag = Val(getvalue(i, 14)) 'Formula Code
            txtFormula.Text = Trim$(getvalue(i, 15)) 'Formula Name
            txtAI1.Text = Trim$(getvalue(i, 18)) 'Payment Terms
          End If
          
          If Val(getvalue(i, 14)) = Val(txtFormula.Tag) Then    'Proceed if it's same formula
            sRow = spItem.DataRowCnt + 1
            spItem.SetText eItem.mItemCode, sRow, Trim(getvalue(i, 2)) 'Item Code
            spItem.SetText eItem.mItemDesc, sRow, getvalue(i, 3) 'Item Desc
            spItem.SetText eItem.mUOM, sRow, getvalue(i, 4) ' UomDesc
            Call Spread_Set_Decimal(spItem, Val(getvalue(i, 5)), eItem.mStockQty, sRow, eItem.mInvQty)
            Call Spread_Set_Decimal(spItem, Val(getvalue(i, 5)), eItem.mBalQty, sRow)
            spItem.SetText eItem.mRate, sRow, getvalue(i, 7)  'Rate
            spItem.SetText eItem.mARate, sRow, getvalue(i, 8) 'Assessable Rate
            
            spItem.SetText eItem.mOType, sRow, getvalue(i, 9) 'Order Type
            OrdType = getvalue(i, 9)
            txtRefType = OrdType
            
            spItem.SetText eItem.mONo, sRow, getvalue(i, 10) 'Order No
            ORDNo = getvalue(i, 10)
            txtRefNo = ORDNo
            
            spItem.SetText eItem.mODate, sRow, getvalue(i, 11) 'Order Date
            OrdDate = getvalue(i, 11)
            dtpRefDate = OrdDate
            
            spItem.SetText eItem.mBalQty, sRow, getvalue(i, 12) 'Bal Qty
            spItem.SetText eItem.mSchType, sRow, getvalue(i, 13)  ' Schedule Type
            spItem.SetText eItem.mSetNo, sRow, getvalue(i, 17)
            
            Call Spread_Lock(spItem, True, eItem.mInvQty, sRow)
          End If
  '      End If
        
        If Trim$(getvalue(i, 9)) = "JO" Then
          fraJORefDc.Visible = True
  
          Call SetSheet(sprJODC, sRow)
          Call SetSheet(SprRecdItem, sRow)
                  
          Sql = " Select Distinct FromCode from ItemGroupKey_Dtl"
          Sql = Sql & " where ToCode='" & Trim(getvalue(i, 2)) & "' and ITemGroupKey=" & Val(getvalue(i, 17))
          Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
          Sql = Sql & " order by FromCode"
          If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
          
          SprRecdItem.sheet = sRow
          SprRecdItem.MaxRows = rs.RecordCount
          For j = 1 To rs.RecordCount
            SprRecdItem.SetText 1, j, rs("FromCode")
            rs.MoveNext
          Next j
        End If
      Next i
    End If
    
    If mCallMenu = mExportOrder Then
      Call PackingList(OrdType, ORDNo, OrdDate)
    End If
    
    TermsDetails   'To display the Terms & Formula
    fraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
    cmdCalc.Enabled = True
    
    txtFormula.Enabled = False
    
    cmdSave.Enabled = True
    spItem.SetFocus
    Spread_Set_Focus spItem, 1, 4
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetSheet(ByRef Spr As AceSpread,ByVal  fRow As Long)
' On Error Resume Next
    If Spr.SheetCount < fRow Then
      Spr.SheetCount = fRow
      Call CopySpreadStructure(Spr, 1, , fRow)
    End If
    Spr.sheet = fRow
    Spr.ActiveSheet = fRow
End Sub

Private Sub TermsDetails()
  Dim snap As New AceADODB.Recordset
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Sql = " select a.CustomerFormulaCode,(select distinct b.FrmLName from COM_FormulaMast_Hdr b where"
  Sql = Sql + " b.ModuleCode = " & gModuleCode & " and b.FrmID = a.CustomerFormulaCode and b.Branchid = a.Branchid "
  Sql = Sql + " and b.Compid = a.Compid) FrmLName, a.ModeCode,"
  Sql = Sql + " (select b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid) ModeDesc"
  Sql = Sql + " from FAS_PARTY_MASTER a where a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.PartyCode = " & Val(txtcustcd.Tag) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
'    If Not IsNull(snap("CustomerFormulaCode")) = True Then: txtFormula.Tag = snap("CustomerFormulaCode")

'    If Not IsNull(snap("FrmLName")) = True Then: txtFormula = snap("FrmLName")
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

Private Sub Direct()
  Dim Sql As String
  Dim i As Long
  
  Dim CurDate As Date
  
  Try

  CurDate = ServerDate
  
  Sql = " Select a.ItemCode, c.ItemDescription, a.PartNo, c.UOMDecimal, a.Rate,"
  Sql = Sql & " a.AssessableRate, c.UomDesc from SAL_Customer_ItemMaster a "
  Sql = Sql & " Left Outer Join MATERIAL_MASTER c "
  Sql = Sql & " On a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=C.Compid "
  Sql = Sql & " Where a.CustomerCode=" & Val(txtcustcd.Tag) & " and c.Status = 1"
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.ValidTo >= '" & Format(CurDate, "yyyy-MM-dd") & "'"

  Sql = Sql & " Union "
  Sql = Sql & " (Select a.ItemCode,c.ItemDescription,a.ItemCode,c.UOMDecimal,a.Rate, "
  Sql = Sql & " a.AssessableRate, c.UomDesc from PUR_Vendor_ItemMaster a "
  Sql = Sql & " Left Outer Join MATERIAL_MASTER c "
  Sql = Sql & " On a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=C.Compid "
  Sql = Sql & " Where a.VendorCode=" & Val(txtcustcd.Tag) & " and c.Status = 1"
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.ValidTo >= '" & Format(CurDate, "yyyy-MM-dd") & "')"

  If Trim$(cmbType.Text) = "Supplementary" Then
    ' For Supplementary Invoice
    Sql = Sql & " Union "
    Sql = Sql & " (Select c.ItemCode,c.ItemDescription,c.ItemCode,c.UOMDecimal,c.Rate, "
    Sql = Sql & " c.Rate as AssessableRate, c.UomDesc from MATERIAL_MASTER c "
    Sql = Sql & " Where c.Status = 1"
    Sql = Sql & " and c.Branchid='" & gBranchID & "' and c.Compid='" & gCompID & "' and c.ItemNature='M' "
    Sql = Sql & " and c.ItemCode='SUBINV0001'"
    Sql = Sql & " )"
  End If

  If FetchValue_New(True, db.ConnectionString, Sql, "1110000", "List of Item's", "Item Code", "Item Description", "Part No") = True Then
    spItem.MaxRows = UBound(getvalue, 1)
    
    If ValidateItemCount(MaxItemCount, spItem.MaxRows) = True Then
      For i = 1 To UBound(getvalue, 1)
        spItem.SetText eItem.mItemCode, i, Trim(getvalue(i, 1)) 'Item Code
        spItem.SetText eItem.mItemDesc, i, getvalue(i, 2) 'Item Desc
  '      Spread_Decimal_Set spItem, eItem.mStockQty, eItem.mInvQty, i, i, Val(getvalue(i, 4))
        Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), eItem.mStockQty, i, eItem.mInvQty)
  '      Spread_Decimal_Set spItem, eItem.mBalQty, eItem.mBalQty, i, i, Val(getvalue(i, 4))
        Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), eItem.mBalQty, i)
        spItem.SetText eItem.mRate, i, getvalue(i, 5)  'Rate
        spItem.SetText eItem.mARate, i, getvalue(i, 6) 'Assessable Rate
        spItem.SetText eItem.mUOM, i, getvalue(i, 7)
        
        If cmbType = "Supplementary" Then
          Call Spread_Lock(spItem, False, eItem.mInvQty, i, eItem.mARate)
        Else
          Call Spread_Lock(spItem, True, eItem.mInvQty, i, eItem.mARate)
        End If
      Next i
    End If
    
    TermsDetails   'To display the Terms & Formula
    fraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
    cmdSave.Enabled = True
    cmdCalc.Enabled = True
    spItem.SetFocus
    Spread_Set_Focus spItem, 1, 4
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function ValidateData() As Boolean
  Dim i As Long
  Dim tmp As Variant
  
  Try
  ValidateData = False
  
  
  For i = 1 To spItem.DataRowCnt
    If optdc = False Then
      If cmbType <> "Supplementary" Then
        spItem.GetText eItem.mWhsName, i, tmp
        If Trim$(tmp) = "" Then
          sShowMessage "Please select the Warehouse"
          spItem.SetFocus
          Spread_Set_Focus spItem, i, eItem.mWhsName
          GoTo ExitHere
        End If
      End If
      
      spItem.GetText eItem.mInvQty, i, tmp
      If Val(tmp) = 0 Then
        sShowMessage "Please enter the Invoice Qty"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eItem.mInvQty
        GoTo ExitHere
      End If
      
      spItem.GetText eItem.mNoofPkts, i, tmp
      If Val(tmp) = 0 Then
        sShowMessage "Please Enter the No Of Packets"
        Call Spread_Set_Focus(spItem, i, eItem.mNoofPkts)
        GoTo ExitHere
      End If
      
      
      spItem.GetText eItem.mOType, i, tmp
      If Trim$(tmp) = "JO" And optorder Then
        spItem.GetText eItem.mInvQty, i, tmp
        SprRecdItem.sheet = i
        If Val(tmp) <> Val(Spread_Col_Total(SprRecdItem, 2) / SprRecdItem.DataRowCnt) Then
          sShowMessage "Sum of Reference DC Qty Equval to Invoice Qty"
          Call Spread_Set_Focus(spItem, i, eItem.mInvQty)
          GoTo ExitHere
        End If
      End If
    End If
  Next i

  If Trim$(txtFormula) = "" Then
    sShowMessage "Please select the Formula"
    txtFormula.SetFocus
    GoTo ExitHere
  End If
  
'  If (Val(txtcustcd.Tag) = 11) And cmbType = "OE" Then
'    If Trim$(txtLCNo.Text) = "" Then
'      sShowMessage "Please select the LC No"
'      txtLCNo.SetFocus
'      GoTo ExitHere
'  End If
    
  If Trim$(txtLCNo.Text) <> "" Then
    If Val(txtGrandTotal.Text) > Val(txtLCBalAmt.Text) Then
      sShowMessage "Invoice Amount should be less than or equal to LC Balance Amount"
      txtLCNo.SetFocus
      GoTo ExitHere
    End If
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateItemCount(ByVal fItemCount As Integer,ByVal  fMaxRows As Integer) As Boolean
  Dim IsProceedToFillItem As Boolean
  
  Try
  
  ValidateItemCount = False
  
  If fMaxRows > fItemCount Then
    sShowMessage "Maximum" & " " & Val(fItemCount) & " Items are allowed to create Invoice"
    IsProceedToFillItem = False
  Else
    IsProceedToFillItem = True
  End If
          
  ValidateItemCount = IsProceedToFillItem
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function WareHouseBalQty(ByVal Row As Long,ByVal  ItemCode As String) As String
  Dim retVal1 As Variant
  Dim retVal As Variant
  Dim ResCode As String
  Dim sRow As Long
  Dim wRow As Long
  Dim sQty As Double
  Dim WareCode As String
  Dim alrCode(50) As String
  Dim sflag As Boolean
  Dim i As Integer
  Dim j As Integer
  
  Try
  j = 1
sflag = True
  sRow = spItem.SearchCol(eItem.mItemCode, 0, spItem.MaxRows, Trim$(ItemCode), SearchFlagsValue)
  sQty = 0
WareCode = ""
ResCode = ""
  Do While sRow <> -1
   sQty = 0
    spItem.GetText eItem.mWhsCode, sRow, retVal
    spItem.GetText eItem.mItemCode, sRow, retVal1
    sflag = True
    For i = 1 To j
      If alrCode(i) = Trim$(retVal) Then
       sflag = False
       Exit For
      End If
    Next i
    If sflag = True Then
      wRow = spItem.SearchCol(eItem.mWhsCode, 0, spItem.MaxRows, Trim$(retVal), SearchFlagsValue)
      Do While wRow <> -1
        spItem.GetText eItem.mItemCode, wRow, retVal
        If Trim$(retVal1) = Trim$(retVal) Then
          spItem.GetText eItem.mInvQty, wRow, retVal
          sQty = sQty + Val(retVal)
          spItem.GetText eItem.mWhsCode, wRow, retVal
          alrCode(j) = Trim$(retVal)
          j = j + 1
        End If
        spItem.GetText eItem.mWhsCode, wRow, retVal
        wRow = spItem.SearchCol(eItem.mWhsCode, wRow, spItem.MaxRows, Trim$(retVal), SearchFlagsValue)
      Loop
      spItem.GetText eItem.mWhsCode, sRow, retVal
      WareCode = WareCode + "'" + Trim$(retVal) + "'" + " " + "Then" + " " + CStr(Val(sQty)) + " when "
    End If
    sRow = spItem.SearchCol(eItem.mItemCode, sRow, spItem.MaxRows, Trim$(ItemCode), SearchFlagsValue)
  Loop
  
  If WareCode <> "" Then
    WareCode = "case b.WareHouseCode when " & "" & Left(WareCode, Len(Trim$(WareCode)) - 4) & "Else" & " 0" & " End"
  End If
  WareHouseBalQty = IIf(WareCode = "", 0, WareCode)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Private Sub LoadInvtype()
'  Dim InvTypeRec As ADODB.Recordset
'  Dim Sql As String
'
'  On Error GoTo ErrHandler
'
'  Set InvTypeRec = New ADODB.Recordset
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='I'"
'  Sql = Sql + " and ModuleName='SAL' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
'  If RSOpen(InvTypeRec, Sql) = False Then GoTo ExitHere
'  If Not InvTypeRec.EOF Then
'    Do Until InvTypeRec.EOF
'      cmbinvtype.AddItem InvTypeRec("DocumentType")
'      InvTypeRec.MoveNext
'    Loop
'    cmbinvtype.ListIndex = 0
'    txtinvnum = GetNo("SAL_INVOICE_HDR", "InvoiceType", "InvoiceNo", Trim$(cmbinvtype), "I", "SAL")
'  End If
'  InvTypeRec.Close
'ExitHere:
'  On Error GoTo 0
'  Set InvTypeRec = Nothing
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

Private Sub cmbinvtype_Click()
  
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

  txtinvnum = GetDocNo("SAL_INVOICE_HDR", "InvoiceType", "InvoiceNo", Trim$(cmbinvtype), "I")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()

  Try
  Me.MousePointer = vbHourglass

  Call Form_Clear(Me)
  
  Select Case mCallMenu
    Case mSalesOrder
      Call LoadDocType(mSalInv)
      Call Spread_HideCol(spItem, False, 4, 7)
    Case mExportOrder
      Call LoadDocType(mSalInvExport)
  End Select
   
  gModuleCode = 4
  fraInvoice.Enabled = True
  fraSource.Enabled = True
  fraCustomer.Enabled = True
  spItem.Enabled = False
  fraDetails.Enabled = False
  cmdSave.Enabled = False
  cmdCalc.Enabled = False
  cmbType.ListIndex = 0
  lblFormulaName.Enabled = False
  txtFormula.Enabled = False

  sprJODC.SheetCount = 1
  SprRecdItem.SheetCount = 1
  Call Spread_Clear(sprJODC)
  Call Spread_Clear(SprRecdItem)
  sprJODC.MaxRows = 1
  SprRecdItem.MaxRows = 1
  fraJORefDc.Visible = False
  
'  LoadInvtype
  cmbinvtype.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

     Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  gModuleCode = 4
  Unload Me
End Sub

 Private Sub CmdOk_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If Trim$(cmbinvtype) = "" Then
    sShowMessage "Please Select the Invoice Type"
    cmbinvtype.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtcustcd) = "" Then
    sShowMessage "Please enter the Customer"
    txtcustcd.SetFocus
    GoTo ExitHere
  End If
  
  If TxtVoucher.Text = "" Then
    sShowMessage "Pls Select the VoucherName."
    TxtVoucher.SetFocus
    GoTo ExitHere
  End If
    
  CurDate = ServerDate
  If Format(DtpVrDate, "yyyyMMdd") > Format(CurDate, "yyyyMMdd") Then
    sShowMessage "VoucherDate should Less than (or) equal to VendorBilldate"
    DtpVrDate.SetFocus
    GoTo ExitHere
  End If
  
  SKey = Format(Now, "yyyyMMddHHmmss") & gUserID
  
  If optdirect Then
 Direct 'list the Direct
End If
  If optdc Then
 DC 'list the dc
End If
  If optorder Then
 Orders 'list the Orders
End If
  
  If cmbType = "Rejection" Then
    gModuleCode = 9
  Else
    gModuleCode = 4
  End If
  
  If gReportID = "PMI" Then
    If LoadPackingDetails = False Then
      GoTo ExitHere
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim tmp As Variant
  Dim sRate As Double
  Dim Stime As Date
  Dim i As Long
  Dim rCnt As Long
  Dim CurDate As Date

  Try
  
  Set snap = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
   
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  CurDate = ServerDate
  
  sRate = 0

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetDocNo
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc

  commClear cmd
  cmd.CommandText = "SAL_INS_INVOICE_HDR"
  cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(cmbinvtype))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceTime", adDBTimeStamp, adParamInput, , Format(CurDate, "HH:mm"))
  If optdirect Then
    cmd.Parameters.Append cmd.CreateParameter("InvoicSource", adChar, adParamInput, 1, "D")
  ElseIf optorder Then
    cmd.Parameters.Append cmd.CreateParameter("InvoicSource", adChar, adParamInput, 1, "O")
  ElseIf optdc Then
    cmd.Parameters.Append cmd.CreateParameter("InvoicSource", adChar, adParamInput, 1, "C")
  End If
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("AInformation1", adVarChar, adParamInput, 50, Trim$(txtAI1.Text))
  cmd.Parameters.Append cmd.CreateParameter("AInformation2", adVarChar, adParamInput, 50, Trim$(txtAI2.Text))
  cmd.Parameters.Append cmd.CreateParameter("Narration", adVarChar, adParamInput, 150, Trim$(txtNarration.Text))
  cmd.Parameters.Append cmd.CreateParameter("ModeCode", adInteger, adParamInput, , Val(txtMode.Tag))
  cmd.Parameters.Append cmd.CreateParameter("VehicleNo", adVarChar, adParamInput, 15, Trim$(txtVehicleNo.Text))
  cmd.Parameters.Append cmd.CreateParameter("PLASlNo", adVarChar, adParamInput, 10, Trim$(txtSlNo.Text))
  If IsNull(dtpPLADate) Then
    cmd.Parameters.Append cmd.CreateParameter("PLADate", adDBTimeStamp, adParamInput, , Null)
  Else
    cmd.Parameters.Append cmd.CreateParameter("PLADate", adDBTimeStamp, adParamInput, , Format(dtpPLADate, "yyyy-MM-dd"))
  End If
  cmd.Parameters.Append cmd.CreateParameter("NoofPackets", adVarChar, adParamInput, 20, Trim$(txtPackets.Text))
  cmd.Parameters.Append cmd.CreateParameter("InsurancePolicyNo", adVarChar, adParamInput, 20, Trim$(txtInsPolicyNo.Text))
  If IsNull(dtpInsDate) Then
    cmd.Parameters.Append cmd.CreateParameter("InsurancePolicyDate", adDBTimeStamp, adParamInput, , Null)
  Else
    cmd.Parameters.Append cmd.CreateParameter("InsurancePolicyDate", adDBTimeStamp, adParamInput, , Format(dtpInsDate, "yyyy-MM-dd"))
  End If
  cmd.Parameters.Append cmd.CreateParameter("InsuranceFor", adDouble, adParamInput, , Val(txtInsFor))
  cmd.Parameters.Append cmd.CreateParameter("InsuranceSlNo", adVarChar, adParamInput, 15, Trim$(txtInsSlNo.Text))
  cmd.Parameters.Append cmd.CreateParameter("Type", adChar, adParamInput, 20, Trim$(cmbType))
  cmd.Parameters.Append cmd.CreateParameter("RemovelTime", adDBTimeStamp, adParamInput, , Format(DateAdd("n", 30, CurDate), "HH:mm"))
  cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , Val(txtGrandTotal.Text))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute
  
    For i = 1 To spItem.DataRowCnt
      spItem.GetText eItem.mRate, i, tmp
sRate = Val(tmp)
      
      commClear cmd
      cmd.CommandText = "SAL_INS_INVOICE_DTL"
      cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(cmbinvtype))
      cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
      cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
      spItem.GetText eItem.mSetNo, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      spItem.GetText eItem.mItemCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      spItem.GetText eItem.mInvQty, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Qty", adDouble, adParamInput, , Val(tmp))
      
'      .GetText eItem.mRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(sRate, "0.00"))
      
      spItem.GetText eItem.mARate, i, tmp
      tmp = IIf(Val(tmp) = 0, sRate, Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      
      spItem.GetText eItem.mNoofPkts, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("NoofPkts", adInteger, adParamInput, , Val(tmp))
      
      spItem.GetText eItem.mOType, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(tmp))
      
      spItem.GetText eItem.mONo, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(tmp))
      spItem.GetText eItem.mODate, i, tmp
      If Trim$(tmp) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(Trim$(tmp), "yyyy-MM-dd"))
      Else
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Null)
      End If
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
      
      If optdc = False Then
        If cmbType <> "Supplementary" Then
          Stime = Format(ServerDate, "hh:mm:ss")
          
          commClear cmd
          cmd.CommandText = "SAL_INV_INS_MATERIALTRAN_DTL"
          cmd.Parameters.Append cmd.CreateParameter("DOCType", adVarChar, adParamInput, 3, Trim$(cmbinvtype))
          cmd.Parameters.Append cmd.CreateParameter("DOCNo", adInteger, adParamInput, , Val(txtinvnum))
          cmd.Parameters.Append cmd.CreateParameter("DOCDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-mm-dd"))
          cmd.Parameters.Append cmd.CreateParameter("TransactionDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
          cmd.Parameters.Append cmd.CreateParameter("TransactionTime", adDBTimeStamp, adParamInput, , Stime)
          spItem.GetText eItem.mItemCode, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("TransactionItem", adVarChar, adParamInput, 20, Trim$(tmp))
          spItem.GetText eItem.mInvQty, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("TransactionQty", adDouble, adParamInput, , Val(tmp))
          cmd.Parameters.Append cmd.CreateParameter("Norms", adDouble, adParamInput, , 1)
          spItem.GetText eItem.mWhsCode, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("FromWarehouseCode", adVarChar, adParamInput, 6, Trim$(tmp))
          spItem.GetText eItem.mRackCode, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("FromRackBinCode", adVarChar, adParamInput, 6, Trim$(tmp))
          cmd.Parameters.Append cmd.CreateParameter("ToWarehouseCode", adVarChar, adParamInput, 6, Trim(txtcustcd.Tag))
          cmd.Parameters.Append cmd.CreateParameter("ToRackBinCode", adVarChar, adParamInput, 6, "")
          spItem.GetText eItem.mSetNo, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
          cmd.Parameters.Append cmd.CreateParameter("FromItem", adVarChar, adParamInput, 20, "")
          spItem.GetText eItem.mItemCode, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ToItem", adVarChar, adParamInput, 20, "")
          cmd.Parameters.Append cmd.CreateParameter("Justification", adVarChar, adParamInput, 4, "")
          Call AddCommonCmdParameters(cmd)
          cmd.Execute
        End If
      End If

      'Jobwork Reference DC
      spItem.GetText eItem.mOType, i, tmp
      If Trim$(tmp) = "JO" And optorder Then '        sprJODC.sheet = i
        For rCnt = 1 To sprJODC.DataRowCnt
          Sql = " Execute SAL_INS_Invoice_JobworkDC"

          sprJODC.GetText eDc.mDCNo, rCnt, tmp
          Sql = Sql & " @DcNo = '" & Trim$(tmp) & "',"

          sprJODC.GetText eDc.mDCDate, rCnt, tmp
          Sql = Sql & " @DcDate = '" & Format(Trim$(tmp), "yyyy-MM-dd") & "',"

          spItem.GetText eItem.mSetNo, i, tmp
          Sql = Sql & " @ItemGroupKey = " & Val(tmp) & ","

          spItem.GetText eItem.mItemCode, i, tmp
          Sql = Sql & " @ToItemCode = '" & Trim$(tmp) & "',"

          sprJODC.GetText eDc.mrecdItem, rCnt, tmp
          Sql = Sql & " @ItemCode = '" & Trim$(tmp) & "',"

          sprJODC.GetText eDc.mSelectQty, rCnt, tmp
          Sql = Sql & " @Qty = " & Val(tmp) & ","

          Sql = Sql & " @INVType = '" & Trim$(cmbinvtype) & "',"
          Sql = Sql & " @INVNo = " & Val(Trim$(txtinvnum.Text)) & ","
          Sql = Sql & " @INVDate = '" & Format(dpinvdate, "yyyy-MM-dd") & "',"

          Sql = Sql & ReturnCommonNamedParametersSqlForSave(CurDate, , , , , , False)
          If DBExecuteSql(db, Sql) = False Then GoTo ExitHere

        Next rCnt
      End If
    Next i  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If SaveCalcDetails = False Then
    GoTo ExitHere
  End If
  
  'To FAS Posting
  If FASAccountPosting(CurDate) = False Then
    GoTo ExitHere
  End If
  
  For i = 1 To spDCDetails.DataRowCnt
    commClear cmd
    cmd.CommandText = "SAL_UP_DC_DTL"
    spDCDetails.GetText 1, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("DCType", adVarChar, adParamInput, 3, Trim$(tmp))
    spDCDetails.GetText 2, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("DCNo", adInteger, adParamInput, , Val(tmp))
    spDCDetails.GetText 3, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("DCDate", adDBTimeStamp, adParamInput, , Format(Trim$(tmp), "yyyy-MM-dd"))
    spDCDetails.GetText 4, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
    cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(cmbinvtype))
    cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
    cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
  Next i
  
  If mCallMenu = mExportOrder Then
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adChar, adParamInput, 3, Trim$(cmbinvtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
    cmd.Parameters.Append cmd.CreateParameter("CarriageName", adVarChar, adParamInput, 50, Trim$(txtCarriage))
    cmd.Parameters.Append cmd.CreateParameter("CarriageNo", adVarChar, adParamInput, 20, Trim$(txtCarriageNo))
    cmd.Parameters.Append cmd.CreateParameter("DischargePort", adVarChar, adParamInput, 50, Trim$(txtPortName))
    cmd.Parameters.Append cmd.CreateParameter("ReceivingPlace", adVarChar, adParamInput, 50, Trim$(txtRecName))
    cmd.Parameters.Append cmd.CreateParameter("LoadingPort", adVarChar, adParamInput, 50, Trim$(txtLoadName))
    cmd.Parameters.Append cmd.CreateParameter("DestinationPlace", adVarChar, adParamInput, 50, Trim$(txtDesName))
    cmd.Parameters.Append cmd.CreateParameter("Reference", adVarChar, adParamInput, 20, Trim$(txtReference))
    cmd.Parameters.Append cmd.CreateParameter("Buyer", adVarChar, adParamInput, 50, Trim$(txtBuyerName))
    cmd.Parameters.Append cmd.CreateParameter("Consignee", adVarChar, adParamInput, 50, Trim$(txtConsignName))
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "I")
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
                
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST_ADDRESS"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adChar, adParamInput, 3, Trim$(cmbinvtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
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
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "I")
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
        
    commClear cmd
    cmd.CommandText = "SAL_INS_PACKINGLIST_ADDRESS"
    cmd.Parameters.Append cmd.CreateParameter("OrderType", adChar, adParamInput, 3, Trim$(cmbinvtype))
    cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
    cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
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
    cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "I")
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
          
      For rCnt = 1 To sprGoods.DataRowCnt
        commClear cmd
        cmd.CommandText = "SAL_INS_PACKINGLIST_GOODS"
        
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbinvtype))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dpinvdate, "yyyy-MM-dd"))
       sprGoods.GetText eItemGoods.mContainerNo, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("ContainerNo", adVarChar, adParamInput, 20, Trim$(tmp))
        sprGoods.GetText eItemGoods.mPackageKind, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("PackageKind", adVarChar, adParamInput, 50, Trim$(tmp))
        sprGoods.GetText eItemGoods.mDescription, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("GoodsDescription", adVarChar, adParamInput, 50, Trim$(tmp))
        sprGoods.GetText eItemGoods.mQuantity, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Quantity", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eItemGoods.mNetWeight, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("NetWeight", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
        sprGoods.GetText eItemGoods.mLength, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Length", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eItemGoods.mWidth, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Width", adInteger, adParamInput, , Val(tmp))
        sprGoods.GetText eItemGoods.mHeight, rCnt, tmp
        cmd.Parameters.Append cmd.CreateParameter("Height", adInteger, adParamInput, , Val(tmp))
        cmd.Parameters.Append cmd.CreateParameter("OrderFlag", adChar, adParamInput, 1, "I")
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      Next rCnt
  End If
  
  ' Update Itemdescription, Qty and Uom as Narration in FAS_VoucherTranMast Table
  If UpdateNarrationToFASVrMast = False Then
    GoTo ExitHere
  End If
  
  If Trim$(txtLCNo.Text) <> "" Then    'Val(txtcustcd.Tag) = 11
    Sql = " Execute SAL_Invoice_LCRef_Ins"
    Sql = Sql & " @InvType = '" & Trim$(cmbinvtype) & "',"
    Sql = Sql & " @InvNo = " & Val(Trim$(txtinvnum.Text)) & ","
    Sql = Sql & " @InvDate = '" & Format(dpinvdate, "yyyy-MM-dd") & "',"
    Sql = Sql & " @CustomerCode = " & Val(Trim$(txtcustcd.Tag)) & ","
    Sql = Sql & " @LCNo = '" & Trim$(txtLCNo.Text) & "',"
    Sql = Sql & " @LCDate = '" & Format(dtpLCDate.Value, "yyyy-MM-dd") & "',"
    Sql = Sql & " @InvAmount = " & Val(txtGrandTotal.Text) & ","
    Sql = Sql & ReturnCommonNamedParametersSqlForSave(CurDate, , , , , , False)
    If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  End If
 
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)
  GoTo ExitHere
Finally

  Set cmd = Nothing
  Set snap = Nothing
End Try
End Sub

Private Function UpdateNarrationToFASVrMast() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Integer
  Dim Narr As String
  Dim retVal As Variant
  Dim ItemCode As String
  Dim Qty As Long
    
  Try
  UpdateNarrationToFASVrMast = False
  
  Sql = " Select DocType, DocNo, DocDate, VrRoute From FAS_VoucherTranRef"
  Sql = Sql & " Where VrRoute = '0K' and CompCode = '" & gFasCompCode & "' and RefType = '" & (cmbinvtype.Text) & "'"
  Sql = Sql & " and RefNo = " & (txtinvnum.Text) & " and RefDate = '" & Format(dpinvdate, "yyyy-MM-dd") & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
      Narr = ""
      For rCnt = 1 To spItem.DataRowCnt
        spItem.GetText eItem.mItemDesc, rCnt, retVal
        ItemCode = Trim$(retVal)
        spItem.GetText eItem.mInvQty, rCnt, retVal
        Qty = Val(retVal)
        spItem.GetText eItem.mUOM, rCnt, retVal 'Uom
        Narr = Narr & IIf(Trim$(Narr) = "", "", "," & vbCrLf) & Trim$(ItemCode) & " " & Val(Qty) & " " & Trim$(retVal)
      Next rCnt
    
    If Narr <> "" Then
      Narr = Narr & "."
      
      Sql = " Update FAS_VoucherTranMast Set Narration = '" & MaskSingleQuotes(Trim$(Narr)) & "'"
      Sql = Sql & " Where VrRoute = '" & rs("VrRoute") & "' and DocType = '" & Trim$(rs("DocType")) & "' and DocNo = " & rs("DocNo") & " and DocDate = '" & Format(rs("DocDate"), "yyyy-MM-dd") & "' and CompCode = '" & gFasCompCode & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    End If
  End If
  
  UpdateNarrationToFASVrMast = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function SaveCalcDetails() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    CurDate = ServerDate
    
''''''''''''    sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(cmbinvtype) & "' DocType," & Val(txtinvnum) & " DocNo,"
''''''''''''    sql = sql & "'" & Format(dpinvdate, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
''''''''''''    sql = sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,"
''''''''''''    sql = sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
''''''''''''    sql = sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
''''''''''''    sql = sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
''''''''''''    sql = sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
''''''''''''    sql = sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
''''''''''''    db.Execute sql
''''''''''''
''''''''''''    sql = "Insert into COM_FormulaCalc_Dtl"
''''''''''''    sql = sql & " Select '" & Trim$(cmbinvtype) & "', " & Val(txtinvnum) & ", '" & Format(dpinvdate, "yyyy-MM-dd") & "',"
''''''''''''    sql = sql & " " & gModuleCode & " ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, "
''''''''''''    sql = sql & " a.IsValueEditable, a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
''''''''''''    sql = sql & " from COM_tmpFormulaCalc_Dtl a"
''''''''''''    sql = sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
''''''''''''    sql = sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
''''''''''''    db.Execute sql

    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(cmbinvtype) & "' DocType," & Val(txtinvnum) & " DocNo,"
    Sql = Sql & "'" & Format(dpinvdate, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,'',"
    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear,'" & Format(CurDate, "yyyy-MM-dd HH:MM:SS  ") & "'"
    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(cmbinvtype) & "', " & Val(txtinvnum) & ", '" & Format(dpinvdate, "yyyy-MM-dd") & "',"
    Sql = Sql & " " & gModuleCode & " ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq,a.KeyCol, a.IsTypePercentage, a.IsValueEditable, a.ValueUsed, a.IsValueEditable,'' ,a.UserID, a.BranchID, a.CompID, a.Finyear,'" & Format(CurDate, "yyyy-MM-dd HH:MM:SS  ") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub dpinvdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
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
  
'  LoadInvtype
   
  Call LoadInitialData
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub LoadInitialData()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try

  If mCallMenu = mExportOrder Then
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = True
    tabSales.TabVisible(2) = True
  Else
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = False
    tabSales.TabVisible(2) = False
  End If
  
  DtpVrDate = Format(ServerDate, "dd-MM-yyyy")
  DtpVrDate.MaxDate = DtpVrDate
   
  cmbType.ListIndex = 0
  
  Sql = " Select top 1 VoucherCode, VoucherName, (VrRoute+VrWidth)as Path From FAS_VoucherMaster"
  Sql = Sql & " Where (VrRoute+VrWidth) Like '0k%' and IsActive = 1 and CompCode = '" & gFasCompCode & "'"
  Sql = Sql & " Order by VrRoute+VrWidth"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    TxtVoucher.Text = rs("VoucherName")
    TxtVoucher.Tag = rs("VoucherCode")
    TxtVrRoute.Text = rs("Path")
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    gModuleCode = 4
  End If
End Sub

Private Sub optdc_Click()
' On Error Resume Next
  
  If optdc.Value = True Then
    lblFormulaName.Enabled = False
    txtFormula.Enabled = False
  End If
End Sub

Private Sub optdirect_Click()
' On Error Resume Next
  
  If optdirect.Value = True Then
    lblFormulaName.Enabled = True
    txtFormula.Enabled = True
  End If
End Sub

Private Sub optorder_Click()
' On Error Resume Next
  
  If optorder.Value = True Then
    lblFormulaName.Enabled = False
    txtFormula.Enabled = False
  End If
End Sub

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  Dim retVal As Variant
  Dim fCurrDate As Date
 
  Try
  Me.MousePointer = vbHourglass
  
  'Get ServerDate
  fCurrDate = ServerDate
  
'  If Row <> 1 Then
'    spItem.GetText eItem.mInvQty, IIf(Row > 1, Row - 1, Row), retVal
'  Else
    retVal = 1
'  End If
  spItem.GetText eItem.mStockQty, Row, tmp
   
  If Col = eItem.mWhsName And Row > 0 And optdc = False And cmbType <> "Supplementary" And Val(retVal) > 0 And Val(tmp) = 0 Then
    spItem.GetText 1, Row, tmp
    
    spItem.GetText eItem.mItemCode, Row, retVal
    retVal = WareHouseBalQty(Row, CStr(retVal))
    
    Sql = " select b.WarehouseName,a.WarehouseCode,a.RackBinCode,(a.CBQty-" & Trim$(retVal) & ") CBQty"
    Sql = Sql & " From dbo.fn_INV_AsOnDate_Stock_V2('" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "',  '" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a,"
    'Sql = Sql + " From INV_STOCK_MASTER a, "
    Sql = Sql + " INV_WAREHOUSE_MASTER b "
    Sql = Sql + " Where b.WarehouseStatus = 1 "
    Sql = Sql + " and a.ItemCode = '" + Trim$(tmp) + "' and (a.CBQty-" & Trim$(retVal) & ") >0"
    Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Sql = Sql + " and a.WarehouseCode=b.WarehouseCode and a.Branchid=b.Branchid "
    Sql = Sql + " and a.Compid=b.Compid"
    If FetchValue_New(False, db.ConnectionString, Sql, "1011", "List of Stock Details", "Warehouse Name", "Rack & Bin", "Stock Qty") = True Then
      spItem.SetText eItem.mWhsCode, Row, getvalue(2)
      spItem.SetText eItem.mWhsName, Row, getvalue(1)
      spItem.SetText eItem.mRackCode, Row, getvalue(3)
      spItem.SetText eItem.mStockQty, Row, Format(getvalue(4), "0.000")
      spItem.GetText eItem.mSchType, Row, tmp
      Call Spread_Lock(spItem, False, eItem.mInvQty, Row)
      If Trim$(tmp) <> "Open" And optorder Then
        spItem.GetText eItem.mBalQty, Row, tmp
        Call Spread_Set_Max(spItem, IIf(Val(getvalue(4)) > Val(tmp), Val(tmp), Val(getvalue(4))), eItem.mInvQty, Row)
      Else
        Call Spread_Set_Max(spItem, Format(getvalue(4), "0.000"), eItem.mInvQty, Row)
      End If
      Spread_Set_Focus spItem, Row, eItem.mInvQty
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

'  If Mode = 1 Then: GoTo ExitHere
'
'  spItem.GetText eItem.mInvQty, Row, tmp 'Qty
'  spItem.GetText eItem.mRate, Row, tmp1  'Rate
'  spItem.SetText eItem.mGV, Row, Format(Val(tmp) * Val(tmp1), "0.00") 'GV
'  spItem.GetText eItem.mARate, Row, tmp1 'ARate
'  spItem.SetText eItem.mAGV, Row, Format(Val(tmp) * Val(tmp1), "0.00") 'AGV
'
'  OrderCalulation  'to calculate the Order Value

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mInvQty, eItem.mRate, eItem.mARate
          spItem.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
'            Call StartCalcFormula(IIf(Trim$(txtGrandTotal.Text) = "", True, False))
          Else
            txtGoodsValue.Text = ""
          End If
          
          Select Case Col
            Case eItem.mInvQty
              If Val(retVal) > 0 Then
'                If optquotation = True Then
'                  .SetActiveCell eItem.mItem, Row + 1
'                Else
                  spItem.SetActiveCell eItem.mRate, Row
'                End If
              End If
            Case eItem.mRate
              If Val(retVal) > 0 Then
                spItem.SetActiveCell eItem.mARate, Row
              End If
            Case eItem.mARate
              If Val(retVal) > 0 Then
                spItem.SetActiveCell eItem.mItemCode, Row + 1
              End If
          End Select
      End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 spItem_DblClick spItem.ActiveCol, spItem.ActiveRow
End If
End Sub

Private Sub spItem_KeyUp(sender as object, e as KeyUpEventArgs)
 
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyDelete And Shift = 1 Then
      spItem.DeleteRows spItem.ActiveRow, 1
      spItem.MaxRows = spItem.DataRowCnt + 1
    Call SetTotalGoodsValue
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim tmp As Variant
  
  Try
  
  If Row <> NewRow And NewRow > 0 Then 'fraJORefDc.Visible = True
    Call SetSheet(sprJODC, NewRow)
    Call SetSheet(SprRecdItem, NewRow)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

'Private Sub sprGoods_DblClick(sender as object, e as DblClickEventArgs)
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  With sprGoods
'    Select Case Col
'      Case eItemGoods.mHeight
'          If AddNewRow(Row) Then
'            .SetFocus
'            .SetActiveCell eItemGoods.mContainerNo, Row
'          End If
'    End Select
'  End With
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'
'
'End Sub
'
Private Sub sprGoods_EditMode(sender as object, e as EditModeEventArgs)
    If Mode = 0 Then
      sprGoods.MaxRows = sprGoods.DataRowCnt + 1
    End If
End Sub

Private Function AddNewRow(ByVal fRow As Long) As Boolean

  Try
  
    If fRow = sprGoods.MaxRows Then
       sprGoods.MaxRows = sprGoods.MaxRows + 1
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub sprJODC_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  Dim i As Long
  Dim sRow As Long
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date

  Try
  Me.MousePointer = vbHourglass

  spItem.GetText eItem.mInvQty, sprJODC.ActiveSheet, tmp
  If Val(tmp) = 0 Then
    sShowMessage "Pls Enter Invoice Qty"
    Call Spread_Set_Focus(spItem, sprJODC.ActiveSheet, eItem.mInvQty)
    GoTo ExitHere
  End If

  If Col = eDc.mDCNo Then
   
    sprJODC.GetText Col, Row, tmp
    If Trim$(tmp) <> "" Then
      GoTo ExitHere
    End If
    
    'Set ServerDate to dtpAsOnDate datepicker
    CurDate = Format(ServerDate, "dd/MM/yyyy HH:mm:ss")
  
    Sql = " Execute SAL_JobworkDCRef_Hlp "
    spItem.GetText eItem.mOType, spItem.ActiveRow, tmp
    Sql = Sql & " @DocType='" & Trim$(tmp) & "',"
    
    spItem.GetText eItem.mONo, spItem.ActiveRow, tmp
    Sql = Sql & " @DocNo='" & Val(tmp) & "',"
    
    spItem.GetText eItem.mODate, spItem.ActiveRow, tmp
    Sql = Sql & " @DocDate='" & Format(tmp, "yyyy-MM-dd") & "',"
    
    spItem.GetText eItem.mSetNo, spItem.ActiveRow, tmp
    Sql = Sql & " @ItemGroupKey='" & Trim$(tmp) & "',"
    
    Sql = Sql & " @InvType='" & Trim$(cmbinvtype) & "',"
    Sql = Sql & " @InvNo='" & Val(txtinvnum) & "',"
    Sql = Sql & " @InvDate='" & Format(dpinvdate, "yyyy-MM-dd") & "',"
    spItem.GetText eItem.mItemCode, spItem.ActiveRow, tmp
    Sql = Sql & " @SentItem='" & Trim$(tmp) & "',"
    Sql = Sql & " @AsonDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " @BranchID='" & gBranchID & "',"
    Sql = Sql & " @CompID='" & gCompID & "'"
    
    If FetchValue_New(True, db.ConnectionString, Sql, "111101", "List of Customers", "Dc No", "Dc Date", "Item Code", "Dc Qty", "Bal Qty") = True Then
'      SprRecdItem.Visible = True
        sprJODC.MaxRows = sprJODC.DataRowCnt + UBound(getvalue()) + 1
        sRow = sprJODC.DataRowCnt + 1
        For i = 1 To UBound(getvalue())
            sprJODC.SetText eDc.mDCNo, sRow, getvalue(i, 1)
            sprJODC.SetText eDc.mDCDate, sRow, getvalue(i, 2)
            sprJODC.SetText eDc.mrecdItem, sRow, getvalue(i, 3)
            
          If Spread_Clear_Duplicate(sprJODC, sRow, eDc.mDCNo, eDc.mDCDate, eDc.mrecdItem) = False Then
            sprJODC.SetText eDc.mDCNo, sRow, getvalue(i, 1)
            sprJODC.SetText eDc.mDCDate, sRow, getvalue(i, 2)
            sprJODC.SetText eDc.mrecdItem, sRow, getvalue(i, 3)
            
            Call Spread_Set_Decimal(sprJODC, NoOfDecimal(CDbl(getvalue(i, 5))), eDc.mDCQty, sRow, eDc.mDCQty, sRow, sprJODC.ActiveSheet)
            
            sprJODC.SetText eDc.mDCQty, sRow, getvalue(i, 4)
            sprJODC.SetText eDc.mNorms, sRow, getvalue(i, 5)
            
            Call Spread_Set_Decimal(sprJODC, NoOfDecimal(CDbl(getvalue(i, 5))), eDc.mSelectQty, sRow, eDc.mSelectQty, sRow, sprJODC.ActiveSheet)
            sprJODC.SetText eDc.mBalQty, sRow, getvalue(i, 6)
            
            Call Spread_Set_Max(sprJODC, CDbl(getvalue(i, 6)), eDc.mSelectQty, sRow, eDc.mSelectQty, sRow)
            sRow = sRow + 1
          End If
          
        Next i
        sprJODC.MaxRows = sprJODC.DataRowCnt + 1
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprJODC_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  
  If Shift = 1 And KeyCode = vbKeyDelete Then
      sprJODC.DeleteRows sprJODC.ActiveRow, 1
      sprJODC.MaxRows = sprJODC.DataRowCnt + 1
  End If
End Sub

Private Sub sprJODC_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim tmp As Variant

  Try

  If (Row <> NewRow And NewRow > 0) Or (Row = NewRow And Row = 1) Then
    sprJODC.GetText eDc.mrecdItem, NewRow, tmp
    txtitemCode = tmp
    sprJODC.GetText eDc.mNorms, NewRow, tmp
    txtitemCode.Tag = tmp
    If Val(txtitemCode.Tag) > 0 Then
      txtsentqty.Enabled = True
      txtrecdqty.Enabled = True
    Else
      txtsentqty.Enabled = False
      txtrecdqty.Enabled = False
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprJODC_Validate(sender as object, e as ValidateEventArgs)
  Dim i As Long
  Dim tmp As Variant
  Dim norms As Variant

  Try
  
  SprRecdItem.sheet = sprJODC.ActiveSheet
  
  If sprJODC.DataRowCnt > 0 Then
    For i = 1 To SprRecdItem.DataRowCnt
      SprRecdItem.GetText 1, i, tmp
      sprJODC.GetText eDc.mNorms, Spread_Search_Col(sprJODC, eDc.mrecdItem, Trim$(tmp), mValue, 0), norms
      If Val(norms) > 0 Then
        SprRecdItem.SetText 2, i, Floor(Spread_Col_Total(sprJODC, eDc.mSelectQty, 1, sprJODC.DataRowCnt, eDc.mrecdItem, Trim$(tmp), mEquals) / norms)
      Else
        sShowMessage "Please Select All Child Code"
        Cancel = True
        Exit For
      End If
    Next i
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
 
Private Sub txtLCNo_DblClick()
  Dim Sql As String
  Try
  Me.MousePointer = vbHourglass

  Sql = " select a.LCNo, a.LCDate, a.LCAmount, (a.LCAmount -"
  Sql = Sql & " isnull((select "
  Sql = Sql & " sum(case when c.AcceptedQty > 0 then ((c.AcceptedQty/c.InvoiceQty)*b.InvAmount) else b.InvAmount End) InvAmount"
  Sql = Sql & " from SAL_Invoice_LCRef b"
  Sql = Sql & " left outer join SAL_Invoice_Dtl c"
  Sql = Sql & " on c.InvoiceNo=b.InvNo and c.InvoiceDate=b.InvDate and c.InvoiceType=b.InvType"
  Sql = Sql & " and c.Branchid=b.Branchid and c.Compid=b.Compid "
  Sql = Sql & " where b.CustomerCode=a.CustomerCode"
  Sql = Sql & " and b.LCNo=a.LCNo and b.LCDate=a.LCDate and b.Branchid=a.Branchid and b.Compid=a.Compid),0)) as BalanceAmt"
  Sql = Sql & " from SAL_LC_Dtl a"
  Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and '" & Format(dpinvdate, "yyyy-MM-dd") & "' between a.ValidFrom and a.ValidTo"
  Sql = Sql & " and a.Branchid ='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "1110", "List of LC's", "LC No", "LC Date", "LC Amount") = True Then
    txtLCNo.Text = getvalue(1)
    dtpLCDate.Value = getvalue(2)
    txtLCAmt.Text = getvalue(3)
    txtLCBalAmt.Text = getvalue(4)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtLCNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  If KeyCode = vbKeyF2 Then
 txtLCNo_DblClick
End If
End Sub

Private Sub txtsentqty_Validate(ByVal Cancel As Boolean)
' On Error Resume Next
  txtrecdqty = Val(txtsentqty) * Val(txtitemCode.Tag)
End Sub

Private Sub txtrecdqty_Validate(ByVal Cancel As Boolean)
' On Error Resume Next
  txtsentqty = Floor(Val(txtrecdqty) / Val(txtitemCode.Tag))
End Sub

Private Sub txtcustcd_DblClick()
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Sql = " Select distinct (Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.CustomerCode "
  Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.CustomerCode From SAL_Customer_ItemMaster a"
  Sql = Sql & " Where a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " Union (select distinct (Select b1.PartyName From FAS_PARTY_MASTER b1 Where b1.PartyCode=a1.CustomerCode "
  Sql = Sql & " and b1.Branchid=a1.Branchid and b1.Compid=a1.Compid) PartyName,a1.CustomerCode From SAL_Jobwork_ItemMaster a1"
  Sql = Sql & " Where a1.Branchid='" & gBranchID & "' and a1.Compid='" & gCompID & "')"
  If optdirect Then
    Sql = Sql & " Union select distinct (Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.VendorCode "
    Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.VendorCode From PUR_Vendor_ItemMaster a"
    Sql = Sql & " Where a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  End If
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Name", "Code") = True Then
    txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
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

Private Sub DC()
  Dim Sql As String
  Dim i As Long
  Dim Item As String
  Dim desc As String
  Dim UOMDescimal As Integer
  Dim Qty As Double
  Dim Rate As Double
  Dim ARate As Double
  Dim sRow As Long
  Dim sRow1 As Long
  Dim tot As Integer
  
  Dim PrevFrmCode As Integer
  Dim CurDate As Date
  Dim Uom As String
  
  CurDate = ServerDate
  
  Try
  
'  Sql = "select a.DCType,a.DCNo,a.DCDate,a.ItemCode,b.ItemDescription,"
'  Sql = Sql + " b.UomDecimal,a.DCQty,a.Rate,a.AssessableRate,a.OrderType,a.OrderNo,a.OrderDate, a.DCSource, b.UOM"
'  Sql = Sql + " from SalesDC a left outer join MATERIAL_MASTER b on a.ItemCode=b.ItemCode "
'  Sql = Sql + " and a.Branchid=b.Branchid and a.Compid=b.Compid  where a.InvoiceType='' and b.Status = 1 "
'  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
'  Sql = Sql + " and a.CustomerCode=" & Val(txtcustcd.Tag) & " order by a.ItemCode"

  Sql = " Select a.DCType, a.DCNo, a.DCDate, a.ItemCode, b.ItemDescription, b.UomDecimal, a.DCQty, a.Rate,"
  Sql = Sql & " A.AssessableRate, A.OrderType, A.OrderNo, A.OrderDate, c.FormulaCode, d.FrmLName, a.DCSource, b.UOM"
  Sql = Sql & " From SalesDC a"
  Sql = Sql & " Left Outer Join [dbo].[fn_SAL_Order_Hdr] ('" & Format(CurDate, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')   c"
  Sql = Sql & " On a.OrderType = c.OrderType and a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate"
  Sql = Sql & " and a.BranchID = c.BranchID and a.CompID = c.CompID"
  Sql = Sql & " Left Outer Join (Select * From COM_FormulaMast_Hdr Where ModuleCode = " & gModuleCode & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "')d"
  Sql = Sql & " On c.FormulaCode = d.FrmID and c.BranchID = d.BranchID and c.CompID = d.CompID"
  Sql = Sql & " Left Outer Join MATERIAL_MASTER b"
  Sql = Sql & " On a.ItemCode = b.ItemCode and a.BranchID = b.BranchID and a.CompID = b.CompID"
  Sql = Sql & " Where a.InvoiceType='' and b.Status = 1 and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.CustomerCode= " & Val(txtcustcd.Tag) & ""
  Sql = Sql & " Order by a.ItemCode"
  
  If FetchValue_New(True, db.ConnectionString, Sql, "1111101000000100", "List of DC's", "DC Type", "DC No", "DC Date", "Item Code", "Item Description", "DC Qty", "FrmName") = True Then
    spDCDetails.MaxRows = UBound(getvalue, 1)
    Qty = 0
Rate = 0
ARate = 0
    tot = UBound(getvalue, 1)
    spItem.MaxRows = tot
    
    If ValidateItemCount(MaxItemCount, spItem.MaxRows) = True Then
      For i = 1 To tot
        If getvalue(i, 1) <> "" And (getvalue(1, 10) = getvalue(i, 10)) And (getvalue(1, 11) = getvalue(i, 11)) And (getvalue(1, 12) = getvalue(i, 12)) Then
          
          If getvalue(i, 15) = "D" Then
            txtFormula.Enabled = True
            txtFormula.Text = ""
            txtFormula.Tag = ""
            lblFormulaName.Enabled = True
          Else
            If i = 1 Then
              PrevFrmCode = getvalue(i, 13)
              txtFormula.Text = getvalue(i, 14)
              txtFormula.Tag = getvalue(i, 13)
              Call Spread_HideCol(spItem, True, 4, 6)
            End If
          End If
          
  '        If Val(txtFormula.Tag) = getvalue(i, 13) Then
            If Item <> Trim$(getvalue(i, 4)) And Qty > 0 Then
              sRow = spItem.DataRowCnt + 1
              spItem.SetText eItem.mItemCode, sRow, Trim(Item)
              spItem.SetText eItem.mItemDesc, sRow, Trim(desc)
              Call Spread_Set_Decimal(spDCDetails, UOMDescimal, eItem.mStockQty, i, eItem.mInvQty)
              spItem.SetText eItem.mInvQty, sRow, Val(Qty)
              spItem.SetText eItem.mRate, sRow, Val(Rate)
              spItem.SetText eItem.mARate, sRow, Val(ARate)
              spItem.SetText eItem.mGV, sRow, Val(Qty) * Val(Rate)
              spItem.SetText eItem.mAGV, sRow, Val(Qty) * Val(ARate)
              spItem.SetText eItem.mUOM, sRow, Trim$(Uom)
              Call Spread_Lock(spItem, False, eItem.mInvQty, sRow)
              Qty = 0
Rate = 0
ARate = 0
            End If
         
            sRow1 = spDCDetails.DataRowCnt + 1
            spDCDetails.SetText 1, sRow1, Trim(getvalue(i, 1))
            spDCDetails.SetText 2, sRow1, getvalue(i, 2)
            spDCDetails.SetText 3, sRow1, getvalue(i, 3)
            spDCDetails.SetText 4, sRow1, Trim(getvalue(i, 4))
            spDCDetails.SetText 5, sRow1, Trim(getvalue(i, 6))
            
            Item = Trim(getvalue(i, 4))
            desc = getvalue(i, 5)
          
            Call Spread_Lock(spItem, True, eItem.mInvQty, sRow1)
    '        Spread_Decimal_Set spDCDetails, 5, 5, SRow1, SRow1, Val(getvalue(i, 6))
            Call Spread_Set_Decimal(spDCDetails, Val(getvalue(i, 6)), 5, sRow1)
            
            UOMDescimal = Val(getvalue(i, 6))
            spDCDetails.SetText 6, sRow1, Val(getvalue(i, 7))
            Qty = Qty + Val(getvalue(i, 7))
            Rate = Val(getvalue(i, 8))
            ARate = Val(getvalue(i, 9))
            PrevFrmCode = getvalue(i, 13)
            Uom = getvalue(i, 14)
  '        End If
        End If
      Next i
    End If
    
    If i = (tot + 1) And Qty > 0 Then
      sRow = spItem.DataRowCnt + 1
      spItem.SetText eItem.mItemCode, sRow, Trim(Item)
      spItem.SetText eItem.mItemDesc, sRow, Trim(desc)
'      Spread_Decimal_Set spDCDetails, eItem.mStockQty, eItem.mInvQty, SRow, SRow, UOMDescimal
      Call Spread_Set_Decimal(spDCDetails, UOMDescimal, eItem.mStockQty, sRow)
      spItem.SetText eItem.mInvQty, sRow, Val(Qty)
      spItem.SetText eItem.mRate, sRow, Val(Rate)
      spItem.SetText eItem.mARate, sRow, Val(ARate)
      spItem.SetText eItem.mGV, sRow, Val(Qty) * Val(Rate)
      spItem.SetText eItem.mAGV, sRow, Val(Qty) * Val(ARate)
      Qty = 0
Rate = 0
ARate = 0
      
      Call spItem_EditMode(eItem.mInvQty, 1, 0, True)
    End If
    
    TermsDetails   'To display the Terms & Formula
    fraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
    cmdCalc.Enabled = True
    cmdSave.Enabled = True
    spItem.SetFocus
    Spread_Set_Focus spItem, 1, eItem.mWhsName
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtFormula_Change()
  
  Try
  Me.MousePointer = vbHourglass
  
  If spItem.DataRowCnt <= 0 Then
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
  
  Sql = "Select FrmLName, FrmID from COM_FORMULAMAST_HDR where isActive='1'"
  Sql = Sql & " and ModuleCode=" & gModuleCode & ""
  Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
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

Private Sub txtMode_DblClick()
  Dim Sql As String
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description,Code From TERMSMASTER where typeflg = 'Mode' and Compid = '" & gCompID & "'"
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

Private Sub txtMode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtMode_DblClick
End If
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retVal As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  Dim scfRate As Double
  
  Try
  StartCalcFormula = False
  
  If Val(txtFormula.Tag) <= 0 Then
    sShowMessage "No formula was found to Calculate. Check the Reference Document."
    GoTo ExitHere
  Else
    scfRate = 0
    
    If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
        ReDim ItemList(spItem.DataRowCnt - 1)
        ReDim SCValList(spItem.DataRowCnt * 2 - 1)
        
        For rCnt = 1 To spItem.DataRowCnt
          spItem.GetText eItem.mRate, rCnt, retVal
scfRate = Val(retVal)
          
          'Item List
          spItem.GetText eItem.mItemCode, rCnt, retVal
          Item = Trim$(retVal)
          ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
          
          'SCValue List
          spItem.GetText eItem.mInvQty, rCnt, retVal
          Qty = Val(retVal)
  '        .GetText eItem.mRatePer, rCnt, retVal
          RatePer = 1
          
'          .GetText eItem.mRate, rCnt, retVal
          SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * scfRate
          
          spItem.GetText eItem.mARate, rCnt, retVal
          retVal = IIf(Val(retVal) = 0, scfRate, Val(retVal))
          SCValList((rCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        Next rCnt
    Else
      ReDim ItemList(0)
      ItemList(0) = "0" & Chr(145) & "0"
      
      ReDim SCValList(1)
      SCValList(0) = eParamID.mGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Text)
      SCValList(1) = eParamID.mAGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Tag)
    End If
  
    'Insert Tmp_FormulaCalc_Ref
    If optorder.Value = True Then
      If InsertTmpFormulaCalcRef = False Then
        GoTo ExitHere
      End If
    End If
    ' RefUserKey added on 2007-12-15
    If CalculateFormula(Val(txtcustcd.Tag), Val(txtFormula.Tag), txtUserKey, SCValList, ItemList, txtGrandTotal, sprCalc, eCalc.mDesc, eCalc.mValue, , "Sales Order", False, ShowPromptScreen, , , , , , , IIf(optorder.Value = True, Trim$(RefUserKey), "")) = False Then

    
      GoTo ExitHere
    End If
    
    ' Delete Formula Calc TmpTable
    If optorder.Value = True Then
      Call DeleteTmpFormulaCalcRef
    End If
    
    Call CalcButtonStatus(Me, False)
    
    StartCalcFormula = True
 End If
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
'      .cmdCalc.Enabled = False
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

  If Val(txtGoodsValue.Text) > 0 Then
    If StartCalcFormula(True) = False Then
      GoTo ExitHere
    End If
    
    Call CalcButtonStatus(Me, False)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetSprGoodsValue(ByVal fRow As Long)
  Dim retVal(1 To 3) As Variant
  Dim tmpTot As Double
  Dim gvRate As Double
    
  Try
  
    gvRate = 0
    
    spItem.GetText eItem.mInvQty, fRow, retVal(1)
    retVal(2) = 1
    
    spItem.GetText eItem.mRate, fRow, retVal(3)
gvRate = Val(retVal(3))
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    spItem.SetText eItem.mGV, fRow, Format(tmpTot, "0.00")
    
    spItem.GetText eItem.mARate, fRow, retVal(3)
    retVal(3) = IIf(Val(retVal(3)) > 0, gvRate, Val(retVal(3)))
    
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    spItem.SetText eItem.mAGV, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(spItem, eItem.mGV), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(spItem, eItem.mAGV), "0.00")
End Sub

Private Function SprTotCalc(ByRef fSpread As AceSpread,ByVal  fCol As Long) As Double
  Dim rCnt As Long
  Dim retVal As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, rCnt, retVal
      Total = Total + Val(retVal)
    Next rCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function FASAccountPosting(ByVal CurDate As Date) As Boolean
  Dim Sql As String
  Dim retVal As Variant
  Dim rs As AceADODB.Recordset
  Dim OrderAc As Integer
  
  Try
  FASAccountPosting = False
  
  If cmbType = "Rejection" Then
    Sql = " Select a.PostingAc from dbo.fn_PUR_PurchaseOrder_Hdr('" & Format(dpinvdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dpinvdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
    Sql = Sql + " left outer join dbo.fn_PUR_PurchaseOrder_Dtl('" & Format(dpinvdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dpinvdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') b"
    Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate and a.Branchid = b.Branchid and a.Compid = b.Compid "
    spItem.GetText eItem.mItemCode, 1, retVal
    Sql = Sql + " where a.Status = 'A' and b.ItemCode = '" + Trim$(retVal) + "'"
  Else
    Sql = " Select PostingAc"
    Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " Where OrderType='" & Trim$(txtRefType.Text) & "' and OrderNo=" & Val(txtRefNo.Text) & " and OrderDate='" & Format(dtpRefDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  End If
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    OrderAc = rs("PostingAc")
  End If
  
  If OrderAc = 0 Then
    Sql = " Select LedgerName, LedgerCode from FAS_LedgerMaster"
    Sql = Sql & " Where isActive='1' and Left(LrRoute, 2) in ('06', '0D', '0E')"
    Sql = Sql & " and CompCode='" & gFasCompCode & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "1", "List of Accounts for Posting", "Accounts") = False Then
      GoTo ExitHere
    End If
    
    If getvalue(1) = 0 Then
      GoTo ExitHere
    Else
      OrderAc = getvalue(2)
    End If
  End If
  
  'For Sales Invoice
  If cmbType = "Rejection" Then
    If FASPosting(CurDate, eVrType.mSales, Trim$(TxtVrRoute), DtpVrDate, Val(txtcustcd.Tag), OrderAc, Val(txtFormula.Tag), sprCalc, txtGrandTotal, , , Trim$(cmbinvtype.Text), Val(txtinvnum.Text), dpinvdate, Trim$(txtinvnum.Text), dpinvdate, Trim$(cmbinvtype.Text), Val(txtinvnum.Text), dpinvdate, 1, 1, 2) = False Then

      GoTo ExitHere
    End If
  Else
    If FASPosting(CurDate, eVrType.mSales, Trim$(TxtVrRoute.Text), DtpVrDate, Val(txtcustcd.Tag), OrderAc, Val(txtFormula.Tag), sprCalc, txtGrandTotal, , , Trim$(cmbinvtype.Text), Val(txtinvnum.Text), dpinvdate, Trim$(txtinvnum.Text), dpinvdate, Trim$(txtRefType.Text), Val(txtRefNo.Text), dtpRefDate, 1, 1, 2) = False Then

      GoTo ExitHere
    End If
  End If
  
  FASAccountPosting = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub FillOrder()
  Dim Sql As String
  Dim retVal As Variant
  Dim CurDate As Date
  
  Try

  CurDate = ServerDate
  
  If optdc.Value = True Then
    Sql = " Select a.DCType, a.DCNo, a.DCDate, a.DCQty, A.OrderType, A.OrderNo, A.OrderDate "
    Sql = Sql & " From SalesDC a"
    Sql = Sql & " Left Outer Join dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') c"
    Sql = Sql & " On a.OrderType = c.OrderType and a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate"
    Sql = Sql & " and a.BranchID = c.BranchID and a.CompID = c.CompID"
    Sql = Sql & " Where a.InvoiceType = '' and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and a.CustomerCode= " & txtcustcd.Tag & ""
    Sql = Sql & " Order by a.ItemCode"
  ElseIf optorder.Value = True Then
    Sql = " Select x.OrderType, x.OrderNo, x.OrderDate, x.OrderQty, x.OrderingType, x.InvoiceQty"
    Sql = Sql & " From (Select a.OrderType,a.OrderNo, a.OrderDate, a.OrderQty, a.OrderingType,"
    Sql = Sql & " isNull((Select sum(d.InvoiceQty) from SAL_INVOICE_DTL d"
    Sql = Sql & " Where d.ItemCode = A.ItemCode And d.OrderType = A.OrderType And d.OrderNo = A.OrderNo"
    Sql = Sql & " and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate  and d.Branchid=a.Branchid"
    Sql = Sql & " and d.Compid=a.Compid),0) InvoiceQty"
    Sql = Sql & " From dbo.fn_SAL_SalesOrder ('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
    Sql = Sql & " Where a.CustomerCode = " & txtcustcd.Tag & " and a.Branchid = '" & gBranchID & "'"
    Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Status = 'A'  and '" & Format(CurDate, "yyyy-MM-dd") & "' between a.ValidFrom and a.ValidTo"
    Sql = Sql & " Union"
    Sql = Sql & " (select a1.OrderType, a1.OrderNo, a1.OrderDate, a1.OrderQty, a1.OrderingType,"
    Sql = Sql & " isNull((Select sum(d1.InvoiceQty) from SAL_INVOICE_DTL d1"
    Sql = Sql & " Where d1.ItemCode = a1.SentItemCode And d1.OrderType = a1.OrderType And d1.OrderNo = a1.OrderNo"
    Sql = Sql & " and d1.OrderDate=a1.OrderDate and d1.Branchid=a1.Branchid and d1.Compid=a1.Compid),0) InvoiceQty"
    Sql = Sql & " From dbo.fn_SAL_JobworkOrder ('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')  a1"
    Sql = Sql & " Where a1.CustomerCode = " & txtcustcd.Tag & " and a1.Branchid = '" & gBranchID & "' and a1.Compid = '" & gCompID & "' and a1.Status='A'"
    Sql = Sql & " and a1.Rate >0 and '" & Format(CurDate, "yyyy-MM-dd") & "' between a1.ValidFrom and a1.ValidTo)"
    Sql = Sql & " ) x Where ( x.OrderingType = 'Specific' and (x.OrderQty-IsNull(x.InvoiceQty, 0)) > 0 ) or x.OrderingType = 'Open'"
    Sql = Sql & " Order by y.OrderType, y.OrderNo, y.OrderDate"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub TxtVoucher_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select VoucherName, VoucherCode, (VrRoute+VrWidth) as Path From FAS_VoucherMaster"
  Sql = Sql & " Where (VrRoute+VrWidth) Like '0k%' and IsActive = 1 and CompCode = '" & gFasCompCode & "'"
  Sql = Sql & " Order by VrRoute+VrWidth"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Voucher", "Voucher Name") = True Then
    TxtVoucher.Text = getvalue(1)
    TxtVoucher.Tag = getvalue(2)
    TxtVrRoute.Text = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub TxtVoucher_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
    TxtVoucher_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprJODC(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprJODC


sprJODC.DblClick +=  new EventHandler(sprJODC_DblClick)
sprJODC.KeyDown +=  new EventHandler(sprJODC_KeyDown)
sprJODC.LeaveCell +=  new EventHandler(sprJODC_LeaveCell)
sprJODC.Validate +=  new EventHandler(sprJODC_Validate)

End Sub

Private Sub Design_spDCDetails(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDCDetails

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub Design_sprGoods(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprGoods


sprGoods.DblClick +=  new EventHandler(sprGoods_DblClick)
sprGoods.EditMode +=  new EventHandler(sprGoods_EditMode)

End Sub

Private Sub Design_spOrder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spOrder

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.EditMode +=  new EventHandler(spItem_EditMode)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)
spItem.KeyUp +=  new EventHandler(spItem_KeyUp)
spItem.LeaveCell +=  new EventHandler(spItem_LeaveCell)

End Sub

Private Sub Design_SprRecdItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprRecdItem

End Sub

Private Sub InitializeSpreads()

	Call Design_sprJODC("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", sprJODC, "")

	Call Design_spDCDetails("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", spDCDetails, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", sprCalc, "")

	Call Design_sprGoods("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", sprGoods, "")

	Call Design_spOrder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", spOrder, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", spItem, "")

	Call Design_SprRecdItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesInvoice.frm", SprRecdItem, "")

End Sub
