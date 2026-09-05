VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSalesOrderItemUpdate 
   Caption         =   "Sales Order Item Update"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
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
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5325
      TabIndex        =   51
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4035
      TabIndex        =   50
      ToolTipText     =   "Click to Save the Data"
      Top             =   5805
      Width           =   1300
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   8040
      TabIndex        =   57
      Top             =   5790
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   46
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
         TabIndex        =   49
         Top             =   1845
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   58
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
         TabIndex        =   47
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   0
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   1485
         Left            =   45
         TabIndex        =   48
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
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   15
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin TabDlg.SSTab tabSales 
      Height          =   3855
      Left            =   270
      TabIndex        =   61
      Top             =   1905
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Order List"
      TabPicture(0)   =   "SalesOrderItemUpdate.frx":033E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "spOrder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTerms"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Packing Details"
      TabPicture(1)   =   "SalesOrderItemUpdate.frx":035A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraBuyer"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Packing List"
      TabPicture(2)   =   "SalesOrderItemUpdate.frx":0376
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtReference"
      Tab(2).Control(1)=   "sprGoods"
      Tab(2).Control(2)=   "fraCarriage"
      Tab(2).Control(3)=   "lblRef"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "General Details / Departments"
      TabPicture(3)   =   "SalesOrderItemUpdate.frx":0392
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtGeneralDetails"
      Tab(3).Control(1)=   "spDept"
      Tab(3).Control(2)=   "Label26"
      Tab(3).Control(3)=   "Label25"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txtGeneralDetails 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3195
         Left            =   -74865
         MaxLength       =   3000
         MultiLine       =   -1  'True
         TabIndex        =   113
         Top             =   540
         Width           =   6900
      End
      Begin VB.Frame fraBuyer 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1710
         Left            =   -74910
         TabIndex        =   63
         Top             =   315
         Width           =   11220
         Begin VB.TextBox txtBuyerName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   50
            TabIndex        =   19
            Top             =   315
            Width           =   3660
         End
         Begin VB.TextBox txtBAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7425
            TabIndex        =   21
            Top             =   315
            Width           =   3660
         End
         Begin VB.TextBox txtBAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   20
            Top             =   315
            Width           =   3660
         End
         Begin VB.TextBox txtBCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            TabIndex        =   23
            Top             =   810
            Width           =   1890
         End
         Begin VB.TextBox txtBPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5655
            TabIndex        =   24
            ToolTipText     =   "Sales Quote Date"
            Top             =   810
            Width           =   885
         End
         Begin VB.TextBox txtBAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            TabIndex        =   22
            Top             =   810
            Width           =   3660
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8430
            MaxLength       =   50
            TabIndex        =   26
            ToolTipText     =   "Sales Quote Date"
            Top             =   810
            Width           =   2640
         End
         Begin VB.TextBox txtBPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6555
            MaxLength       =   20
            TabIndex        =   25
            Top             =   810
            Width           =   1860
         End
         Begin VB.TextBox txtBCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   28
            ToolTipText     =   "Sales Quote Date"
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox txtBState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   30
            TabIndex        =   27
            ToolTipText     =   "Sales Quote Date"
            Top             =   1320
            Width           =   2790
         End
         Begin VB.Label lblBuyerName 
            AutoSize        =   -1  'True
            Caption         =   "Buyer Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   71
            Top             =   135
            Width           =   885
         End
         Begin VB.Label lblBAdd2 
            AutoSize        =   -1  'True
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7425
            TabIndex        =   73
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblBAdd1 
            AutoSize        =   -1  'True
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   72
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblBAdd3 
            AutoSize        =   -1  'True
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   70
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblBCity 
            AutoSize        =   -1  'True
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   69
            Top             =   630
            Width           =   270
         End
         Begin VB.Label lblBPinCode 
            AutoSize        =   -1  'True
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5655
            TabIndex        =   68
            Top             =   630
            Width           =   585
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   8430
            TabIndex        =   67
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6555
            TabIndex        =   66
            Top             =   630
            Width           =   450
         End
         Begin VB.Label lblcountry 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2880
            TabIndex        =   65
            Top             =   1125
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   64
            Top             =   1140
            Width           =   375
         End
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
         Height          =   1680
         Left            =   90
         TabIndex        =   74
         Top             =   315
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
            TabIndex        =   17
            ToolTipText     =   "Enter the Remarks"
            Top             =   1305
            Width           =   1176
         End
         Begin VB.ComboBox cmbScheduleType 
            Height          =   330
            ItemData        =   "SalesOrderItemUpdate.frx":03AE
            Left            =   6540
            List            =   "SalesOrderItemUpdate.frx":03B8
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Select the Sales Order Type"
            Top             =   1305
            Width           =   1410
         End
         Begin VB.TextBox txtDelIns 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Select the Delivery Instruction. Press F2 or Double Click for Help"
            Top             =   300
            Width           =   3675
         End
         Begin VB.TextBox txtFreight 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtInsurance 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
            Top             =   300
            Width           =   3675
         End
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
            Top             =   1305
            Width           =   3675
         End
         Begin VB.TextBox txtMode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7455
            Locked          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtForm 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   11
            ToolTipText     =   "Select the Form. Press F2 or Double Click for Help"
            Top             =   810
            Width           =   3675
         End
         Begin VB.TextBox txtPack 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   8
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
            TabIndex        =   15
            ToolTipText     =   "Select the Formula Name. Press F2 / DblClick Key for Help."
            Top             =   1305
            Width           =   2760
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Revision Level"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   7950
            TabIndex        =   116
            Top             =   1125
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Schedule Type"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   6540
            TabIndex        =   83
            Top             =   1125
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Mode"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   7440
            TabIndex        =   82
            Top             =   630
            Width           =   390
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Insurance"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   3765
            TabIndex        =   81
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Form"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   75
            TabIndex        =   80
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Packing && Forwarding"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   75
            TabIndex        =   79
            Top             =   120
            Width           =   1590
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Freight"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   3765
            TabIndex        =   78
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Terms of Payment"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   75
            TabIndex        =   77
            Top             =   1125
            Width           =   1305
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Delivery Instruction"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   7455
            TabIndex        =   76
            Top             =   120
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Formula Name"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   3765
            TabIndex        =   75
            Top             =   1125
            Width           =   1020
         End
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73590
         TabIndex        =   62
         ToolTipText     =   "Sales Quotation Type"
         Top             =   405
         Width           =   2430
      End
Begin AceSpread sprGoods
         Height          =   1455
         Left            =   -74895
         TabIndex        =   39
         Top             =   795
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
         Left            =   90
         TabIndex        =   18
         ToolTipText     =   "Enter the Sales Item Details"
         Top             =   2025
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
         MaxCols         =   11
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SelectBlockOptions=   0
      End
      Begin VB.Frame fraCarriage 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1260
         Left            =   -74910
         TabIndex        =   94
         Top             =   2220
         Width           =   7725
         Begin VB.TextBox txtCarriage 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   40
            ToolTipText     =   "Sales Quotation Type"
            Top             =   180
            Width           =   2430
         End
         Begin VB.TextBox txtCarriageNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   41
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtRecName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   43
            ToolTipText     =   "Select the Sales Quotation. Press F2 or Double Click for Help"
            Top             =   180
            Width           =   2430
         End
         Begin VB.TextBox txtLoadName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   44
            ToolTipText     =   "Sales Quote Date"
            Top             =   525
            Width           =   2430
         End
         Begin VB.TextBox txtDesName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5205
            TabIndex        =   45
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.TextBox txtPortName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1380
            TabIndex        =   42
            ToolTipText     =   "Sales Quote Date"
            Top             =   870
            Width           =   2430
         End
         Begin VB.Label lblCarriageNo 
            AutoSize        =   -1  'True
            Caption         =   "Vessel / Filght No"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   100
            Top             =   577
            Width           =   1260
         End
         Begin VB.Label lblReceivingName 
            AutoSize        =   -1  'True
            Caption         =   "Place Of Receipt"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3990
            TabIndex        =   99
            Top             =   232
            Width           =   1215
         End
         Begin VB.Label lblCarriage 
            AutoSize        =   -1  'True
            Caption         =   "Pre Carriage by"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   150
            TabIndex        =   98
            Top             =   232
            Width           =   1125
         End
         Begin VB.Label lblPortName 
            AutoSize        =   -1  'True
            Caption         =   "Port of Discharge"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   90
            TabIndex        =   97
            Top             =   922
            Width           =   1260
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
         Begin VB.Label lblDesName 
            AutoSize        =   -1  'True
            Caption         =   "Final Destination"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   4020
            TabIndex        =   95
            Top             =   922
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1770
         Left            =   -74895
         TabIndex        =   101
         Top             =   1935
         Width           =   11220
         Begin VB.TextBox txtState 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   30
            TabIndex        =   37
            ToolTipText     =   "Sales Quote Date"
            Top             =   1365
            Width           =   2790
         End
         Begin VB.TextBox txtCountry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   38
            ToolTipText     =   "Sales Quote Date"
            Top             =   1365
            Width           =   2790
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6555
            MaxLength       =   20
            TabIndex        =   35
            Top             =   855
            Width           =   1860
         End
         Begin VB.TextBox txtBEmail 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8430
            MaxLength       =   50
            TabIndex        =   36
            ToolTipText     =   "Sales Quote Date"
            Top             =   855
            Width           =   2640
         End
         Begin VB.TextBox txtAdd3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            TabIndex        =   32
            Top             =   855
            Width           =   3660
         End
         Begin VB.TextBox txtPinCode 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5655
            TabIndex        =   34
            ToolTipText     =   "Sales Quote Date"
            Top             =   855
            Width           =   885
         End
         Begin VB.TextBox txtCity 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            TabIndex        =   33
            Top             =   855
            Width           =   1890
         End
         Begin VB.TextBox txtAdd1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   30
            Top             =   345
            Width           =   3660
         End
         Begin VB.TextBox txtAdd2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7425
            TabIndex        =   31
            Top             =   345
            Width           =   3660
         End
         Begin VB.TextBox txtConsignName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   75
            MaxLength       =   50
            TabIndex        =   29
            Top             =   345
            Width           =   3660
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "State"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   111
            Top             =   1185
            Width           =   375
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2880
            TabIndex        =   110
            Top             =   1170
            Width           =   570
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Phone"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   6555
            TabIndex        =   109
            Top             =   675
            Width           =   450
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "EMail"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   8430
            TabIndex        =   108
            Top             =   675
            Width           =   360
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "PinCode"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   5655
            TabIndex        =   107
            Top             =   675
            Width           =   585
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   106
            Top             =   675
            Width           =   270
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Address3"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   105
            Top             =   675
            Width           =   720
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Consignee Name"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   104
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Address1"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3750
            TabIndex        =   103
            Top             =   165
            Width           =   720
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Address2"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7410
            TabIndex        =   102
            Top             =   150
            Width           =   720
         End
      End
Begin AceSpread spDept
         Height          =   3135
         Left            =   -67935
         TabIndex        =   112
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
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "General Details"
         Height          =   210
         Left            =   -74865
         TabIndex        =   115
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Departments"
         Height          =   210
         Left            =   -67935
         TabIndex        =   114
         Top             =   375
         Width           =   915
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         Caption         =   "Other Reference"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   84
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   720
      Left            =   270
      TabIndex        =   53
      Top             =   -15
      Width           =   3660
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   3090
         TabIndex        =   118
         Top             =   330
         Width           =   480
      End
      Begin VB.TextBox txtOrderType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   117
         ToolTipText     =   "Sales Order No"
         Top             =   315
         Width           =   870
      End
      Begin MSComCtl2.DTPicker dporddate 
         Height          =   315
         Left            =   1890
         TabIndex        =   1
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
         Format          =   67108867
         CurrentDate     =   37250
      End
      Begin VB.TextBox txtordnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Sales Order No"
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1890
         TabIndex        =   56
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1230
         TabIndex        =   55
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Order Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   54
         Top             =   135
         Width           =   825
      End
   End
   Begin VB.Frame fraCustomer 
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
      ForeColor       =   &H80000006&
      Height          =   1230
      Left            =   270
      TabIndex        =   85
      Top             =   630
      Width           =   10680
      Begin VB.TextBox txtPostingAc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   92
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   825
         Width           =   4620
      End
      Begin VB.TextBox txtrefno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   7
         ToolTipText     =   "Enter the Remarks"
         Top             =   825
         Width           =   4605
      End
      Begin MSComCtl2.DTPicker dprefdate 
         Height          =   315
         Left            =   6645
         TabIndex        =   4
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
         Format          =   67108867
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   8085
         TabIndex        =   5
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
         Format          =   67108867
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   9360
         TabIndex        =   6
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
         Format          =   67108867
         CurrentDate     =   37250
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Posting Account"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4725
         TabIndex        =   93
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   105
         TabIndex        =   91
         Top             =   645
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valid Till"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   9360
         TabIndex        =   90
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8085
         TabIndex        =   89
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer PO Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6645
         TabIndex        =   88
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer PO No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4740
         TabIndex        =   87
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   105
         TabIndex        =   86
         Top             =   135
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmSalesOrderItemUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SKey As String

Public mCallMenu As eSalesCallType

'Column Nos for Order Spread
Private Enum eItem
  mItem = 1
  mDesc = 2
  mPartNo = 3
  mUOMCode = 4
  mUOM = 5
  mSchType = 6
  mQty = 7
  mRate = 8
  mAssRate = 9
  mGoodsValue = 10
  mAssGoodsValue = 11
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

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Sub LoadDept()
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  Set snap = New AceADODB.Recordset
  
  Sql = "select a.Code, a.Description,isnull((select 1 from SAL_Order_Department b "
  Sql = Sql + " where a.Code = b.DepartmentCode and b.OrderType = '" + Trim(txtOrderType) + "'"
  Sql = Sql + " and b.OrderNo = " & Val(txtordnum) & " and b.OrderDate = '" + Format(dporddate, "yyyy-MM-dd") + "'"
  Sql = Sql + " and b.Branchid = '" & gBranchID & "' and b.Compid = '" & gCompID & "' "
  Sql = Sql + " and b.Finyear = '" & gFinyear & "'),0) Sel"
  Sql = Sql + " from TermsMaster a where a.TypeFlg = 'Department'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " order by a.Description"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spDept.MaxRows = snap.RecordCount
    For i = 1 To snap.RecordCount
      spDept.SetText 1, i, snap("Code")
      spDept.SetText 2, i, snap("Description")
      spDept.SetText 3, i, snap("Sel")
      snap.MoveNext
    Next i
  End If
  snap.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmbScheduleType_Click()
  Dim i As Long
  For i = 1 To spOrder.DataRowCnt
    If cmbScheduleType = "Yearly" Then
      Call Spread_Lock(spOrder, True, mQty, i)
      spOrder.SetText 6, i, 0
    Else
      Call Spread_Lock(spOrder, False, mQty, i)
      spOrder.SetText mQty, i, 0
    End If
  Next i
End Sub

Private Sub cmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  Form_Clear Me
  CurrDate Me
  fraOrder.Enabled = True
  spOrder.Enabled = False
  cmdSave.Enabled = False
  spOrder.MaxRows = 1
  cmbScheduleType.Enabled = True
  cmbScheduleType.ListIndex = 0
  txtcustcd.Enabled = True
  txtOrderType.SetFocus
  CurrDate Me
  LoadDept
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim tot As Double
  Dim i As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  If Trim(txtOrderType) = "" Then
    sShowMessage "Please select the Order Type"
    txtOrderType.SetFocus
    GoTo ExitHere
  End If
  
  CurDate = ServerDate
  
  Set snap = New AceADODB.Recordset
  
  Sql = "select distinct a.CustomerCode, c.PartyName, a.ReferenceNo,a.ReferenceDate, a.ValidFrom, a.ValidTo,"
  Sql = Sql + " a.Remarks, a.RevisionLevel, a.GeneralDetails,b.PackingDesc, b.InsuranceDesc, b.DeliveryDesc,"
  Sql = Sql + " b.PaymentDesc,  b.FreightDesc, b.ModeDesc, b.FormDesc,a.formulaCode, b.FormulaName, b.PostingAcName"
  Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join "
  Sql = Sql + " dbo.fn_SAL_SO_Terms('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b"
  Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate"
  Sql = Sql + " and a.Branchid = b.Branchid and a.Compid = b.Compid and a.Finyear = b.Finyear"
  Sql = Sql + " left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
  Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql + " where a.OrderType = '" + Trim(txtOrderType) + "' and a.OrderNo = " & Val(txtordnum) & " "
  Sql = Sql + " and a.OrderDate = '" + Format(dporddate, "yyyy-MM-dd") + "' and a. ItemCode = '' "
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
  Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    txtcustcd.Tag = snap("CustomerCode")
    txtcustcd.Text = snap("PartyName")
    txtrefno = snap("ReferenceNo")
    dprefdate = snap("ReferenceDate")
    dpvalidfrom = snap("ValidFrom")
    dpvalidto = snap("ValidTo")
    txtremarks = snap("Remarks")
    txtRevisionLevel = snap("RevisionLevel")
    txtGeneralDetails = snap("GeneralDetails")
    txtPack = snap("PackingDesc")
    txtInsurance = snap("InsuranceDesc")
    txtDelIns = snap("DeliveryDesc")
    txtTerms = snap("PaymentDesc")
    txtFreight = snap("FreightDesc")
    txtMode = snap("ModeDesc")
    txtForm = snap("FormDesc")
    txtFormula.Tag = snap("FormulaCode")
    txtFormula = snap("FormulaName")
    txtPostingAc = snap("PostingAcName")
  End If
  
  snap.Close
  
  Sql = "select a.ItemDescription, a.PartNo, a.UOM, b.SDesc, a.OrderingType, a.OrderQty,"
  Sql = Sql + " a.Rate, a.AssessableRate, a.Status from dbo.fn_SAL_Order_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join "
  Sql = Sql + " INV_UOM_Master b on a.UOM = b.Code and a.Branchid = b.Branchid "
  Sql = Sql + " and a.Compid = b.Compid where a.OrderType = '" + Trim(txtOrderType) + "'"
  Sql = Sql + " and a.OrderNo = " & Val(txtordnum) & " and a.OrderDate = '" + Format(dporddate, "yyyy-MM-dd") + "' "
  Sql = Sql + " and a. ItemCode = '' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    cmbScheduleType = IIf(Trim$(snap("OrderingType")) = "S", "Specific", "Yearly")
    tot = 0
    spOrder.MaxRows = snap.RecordCount
    For i = 1 To snap.RecordCount
      spOrder.SetText 2, i, snap("ItemDescription")
      spOrder.SetText 3, i, snap("PartNo")
      spOrder.SetText 4, i, snap("UOM")
      spOrder.SetText 5, i, snap("SDesc")
      spOrder.SetText 7, i, snap("OrderQty")
      spOrder.SetText 8, i, snap("Rate")
      spOrder.SetText 9, i, snap("AssessableRate")
      spOrder.SetText 10, i, Format((snap("OrderQty") * snap("Rate")), "0.00")
      spOrder.SetText 11, i, Format((snap("OrderQty") * snap("AssessableRate")), "0.00")
      tot = tot + Format((snap("OrderQty") * snap("Rate")), "0.00")
      snap.MoveNext
    Next i
  End If
  txtGoodsValue = Format(tot, "0.00")
  
  LoadDept
  
  fraOrder.Enabled = False
  spOrder.Enabled = True
  cmdSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim snap As AceADODB.Recordset
  Dim i As Long
  Dim tmp As Variant
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Set cmd = New AceADODB.Command
  Set snap = New AceADODB.Recordset
  
  If Chkbefsave = False Then GoTo ExitHere
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  
    For i = 1 To spOrder.DataRowCnt
      spOrder.GetText mItem, i, tmp
      If Trim$(tmp) <> "" Then
        commClear cmd
        cmd.CommandText = "SAL_UP_ORDER_DTL"
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrderType))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
        spOrder.GetText mItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        spOrder.GetText mPartNo, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
        
        Sql = " select a.ItemCode from SAL_Customer_ItemMaster a where a.CustomerCode = " & Val(txtcustcd.Tag) & ""
        spOrder.GetText mItem, i, tmp
        Sql = Sql + " and a.ItemCode = '" + Trim$(tmp) + "' and a.Branchid = '" & gBranchID & "' "
        spOrder.GetText mPartNo, i, tmp
        Sql = Sql + " and a.PartNo = '" + Trim$(tmp) + "' and a.Compid = '" & gCompID & "' "
        If RSOpen(snap, Sql) = False Then GoTo ExitHere
        If snap.EOF Then
          cmd.CommandText = "SAL_INS_CUSTOMER_ITEMMASTER"
          commClear cmd
          cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
          spOrder.GetText mItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          spOrder.GetText mPartNo, i, tmp
           cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
          spOrder.GetText mRate, i, tmp
           cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(tmp))
          spOrder.GetText mAssRate, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(tmp))
          cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(dpvalidfrom, "yyyy-MM-dd"))
          cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(dpvalidto, "yyyy-MM-dd"))
          cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
          cmd.Parameters.Append cmd.CreateParameter("QCCheck", adBoolean, adParamInput, , "0")
          cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, , "1")
          Call AddCommonCmdParameters(cmd, , False)
          cmd.Execute
        End If
        
        spOrder.GetText mItem, i, tmp
        Sql = "update SAL_Order_Schedule set ItemCode = '" + Trim$(tmp) + "'"
        Sql = Sql + " where OrderType = '" + Trim(txtOrderType) + "' and OrderNo = " & Val(txtordnum) & " "
        spOrder.GetText mPartNo, i, tmp
        Sql = Sql + " and OrderDate = '" + Format(dporddate, "yyyy-MM-dd") + "' and PartNo = '" + Trim$(tmp) + "'"
        Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' "
        Sql = Sql + " and Finyear = '" & gFinyear & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        
      End If
    Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  sShowMessage "Order Number :" & Trim$(txtOrderType) & Trim$(txtordnum) & " Items are Updated"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set cmd = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub dporddate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub dprefdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub
Private Sub dpvalidfrom_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub dpvalidto_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
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

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try

  CurrDate Me
  
  cmbScheduleType.ListIndex = 0
  SKey = Format(Now, "yyyyMMddHHmmss") & gUserID
  If mCallMenu = mExportOrder Then
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = True
    tabSales.TabVisible(2) = True
  Else
    tabSales.TabVisible(0) = True
    tabSales.TabVisible(1) = False
    tabSales.TabVisible(2) = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub spOrder_DblClick(sender as object, e as DblClickEventArgs)
  Try
  Me.MousePointer = vbHourglass

  Sprhlp
  Avoid_Duplicate spOrder, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub spOrder_EditMode(sender as object, e as EditModeEventArgs)
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mQty, eItem.mRate, eItem.mAssRate
          spOrder.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
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
  
  Try
  Me.MousePointer = vbHourglass
  
  If KeyCode = vbKeyF2 Then
    If spOrder.ActiveCol = 1 Then Sprhlp '''Procedure for Help
    If spOrder.ActiveCol = 5 Then spOrder_DblClick spOrder.ActiveCol, spOrder.ActiveRow
  ElseIf spOrder.ActiveCol = 11 And KeyCode = 13 Then
    spOrder.MaxRows = spOrder.DataRowCnt + 1
  ElseIf KeyCode = 46 And spOrder.ActiveCol = 1 Then
    Spread_Clear_Row spOrder, 1, spOrder.MaxCols, spOrder.ActiveRow, spOrder.ActiveRow
    spOrder.DeleteRows spOrder.ActiveRow, 1
    spOrder.MaxRows = spOrder.DataRowCnt + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub sprGoods_EditMode(sender as object, e as EditModeEventArgs)
    If Mode = 0 Then
      sprGoods.MaxRows = sprGoods.DataRowCnt + 1
    End If
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
    fraTerms.Enabled = True
    spOrder.Enabled = True
    cmdSave.Enabled = True
    txtrefno.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtcustcd_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
     txtcustcd_DblClick
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtDelIns_DblClick()
  Dim Sql As String
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Delivery Instruction' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Delivery Instruction", "Description") = True Then
    txtDelIns.Text = getvalue(1)
    txtDelIns.Tag = getvalue(2)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
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

  Me.MousePointer = 0
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtOrderType_DblClick()
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate

  Sql = "select distinct a.OrderType, a.OrderNo, a.OrderDate,a.CustomerCode, c.PartyName "
  Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
  Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql + " where a. ItemCode = '' and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "11101", "List of Sales Order", "Order Type", "Order No", "Order Date", "Customer Name") = True Then
    txtOrderType = getvalue(1)
    txtordnum = getvalue(2)
    dporddate = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtOrderType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  If KeyCode = vbKeyF2 Then
 txtOrderType_DblClick
End If
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
      
  Sql = "Select LedgerName, LedgerCode from FAS_LedgerMaster where left(LrRoute,4) not in ('0206', '0303') "
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

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Insurance' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Insurance", "Description") = True Then
    txtInsurance.Text = getvalue(1)
    txtInsurance.Tag = getvalue(2)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtrefno_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtremarks_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim tmp As Variant
  Dim flg As Boolean
  
  Try
  
  flg = False
    For i = 1 To spOrder.DataRowCnt
      spOrder.GetText eItem.mItem, i, tmp
      If Trim$(tmp) <> "" Then
 flg = True
End If
    Next
  
  If flg = False Then
    sShowMessage "Please enter the Item Code"
    spOrder.SetFocus
    Spread_Set_Focus spOrder, 1, eItem.mItem
    Chkbefsave = False
    Exit Function
  End If
  
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Sprhlp()
  Dim Sql As String
  Dim Row As Long
  
  Try
  
  Row = spOrder.ActiveRow
  
  If spOrder.ActiveCol = 1 Then
    Sql = "select a.ItemCode,a.ItemDescription from INV_MATERIAL_MASTER a "
    Sql = Sql + " where a.Branchid = '" & gBranchID & "' and a.compid='" & gCompID & "' and a.Status = 1"
    If FetchValue_New(False, db.ConnectionString, Sql, "11", "List of Items", "Item Code", "Item Description") = True Then
      spOrder.SetText eItem.mItem, Row, getvalue(1) 'ItemCode
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim RCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retVal As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(spOrder.DataRowCnt - 1)
      ReDim SCValList(spOrder.DataRowCnt * 2 - 1)
      
      For RCnt = 1 To spOrder.DataRowCnt
        'Item List
        spOrder.GetText eItem.mItem, RCnt, retVal
        Item = Trim$(retVal)
        ItemList(RCnt - 1) = RCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        spOrder.GetText eItem.mQty, RCnt, retVal
        Qty = Val(retVal)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        spOrder.GetText eItem.mRate, RCnt, retVal
        SCValList((RCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        
        spOrder.GetText eItem.mAssRate, RCnt, retVal
        SCValList((RCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
      Next RCnt
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

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CalcButtonStatus(ByVal FrmObj As Form,ByVal  Enable As Boolean) As Boolean
  On Error GoTo 0
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
  Dim retVal(1 To 3) As Variant
  Dim tmpTot As Double
    
  Try
  
    spOrder.GetText eItem.mQty, fRow, retVal(1)
    retVal(2) = 1
    
    spOrder.GetText eItem.mRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    spOrder.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    spOrder.GetText eItem.mAssRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
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
  Dim RCnt As Long
  Dim retVal As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For RCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, RCnt, retVal
      Total = Total + Val(retVal)
    Next RCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub Design_sprGoods(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprGoods


sprGoods.EditMode +=  new EventHandler(sprGoods_EditMode)

End Sub

Private Sub Design_spOrder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spOrder


spOrder.DblClick +=  new EventHandler(spOrder_DblClick)
spOrder.EditMode +=  new EventHandler(spOrder_EditMode)
spOrder.KeyDown +=  new EventHandler(spOrder_KeyDown)

End Sub

Private Sub Design_spDept(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDept

End Sub

Private Sub InitializeSpreads()

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderItemUpdate.frm", sprCalc, "")

	Call Design_sprGoods("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderItemUpdate.frm", sprGoods, "")

	Call Design_spOrder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderItemUpdate.frm", spOrder, "")

	Call Design_spDept("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderItemUpdate.frm", spDept, "")

End Sub
