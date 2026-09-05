VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInvoiceChange 
   Caption         =   "Invoice Change"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRefNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   70
      ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
      Top             =   8130
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtRefType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   69
      ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
      Top             =   8130
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frahide 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   11895
      TabIndex        =   49
      Top             =   2670
      Width           =   3285
      Begin MSComCtl2.DTPicker dtpInvTime 
         Height          =   315
         Left            =   1890
         TabIndex        =   60
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   138805251
         CurrentDate     =   38845
      End
      Begin VB.TextBox txtFileRef 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1065
         TabIndex        =   54
         Top             =   1785
         Width           =   810
      End
      Begin VB.TextBox txtVrDate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1065
         TabIndex        =   53
         Top             =   1395
         Width           =   1260
      End
      Begin VB.TextBox txtVrNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1065
         TabIndex        =   52
         Top             =   990
         Width           =   765
      End
      Begin VB.TextBox txtVrType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1065
         TabIndex        =   51
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox txtCust 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1065
         TabIndex        =   50
         Top             =   210
         Width           =   2100
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "VrType"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   59
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "VrNo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   510
         TabIndex        =   58
         Top             =   1050
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "VrDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   345
         TabIndex        =   57
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FileRef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   330
         TabIndex        =   56
         Top             =   1785
         Width           =   585
      End
      Begin VB.Label lblCust 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   55
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame fraInvoice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3435
      TabIndex        =   35
      Top             =   75
      Width           =   2910
      Begin VB.TextBox txtinvnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   37
         ToolTipText     =   "Invoice No"
         Top             =   315
         Width           =   795
      End
      Begin VB.ComboBox cmbinvtype 
         Height          =   330
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   36
         ToolTipText     =   "Select the Invoice Type"
         Top             =   315
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpInvDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
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
         Format          =   141033473
         CurrentDate     =   37250
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Inv. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1695
         TabIndex        =   41
         Top             =   135
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inv. Type"
         Height          =   210
         Left            =   90
         TabIndex        =   40
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Inv. No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   870
         TabIndex        =   39
         Top             =   135
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6930
      Width           =   1300
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source of Invoice"
      Enabled         =   0   'False
      Height          =   705
      Left            =   6345
      TabIndex        =   28
      ToolTipText     =   "Select the Invoice Source"
      Top             =   75
      Width           =   2250
      Begin VB.OptionButton optdirect 
         Caption         =   "&Direct"
         Height          =   210
         Left            =   75
         TabIndex        =   31
         Top             =   285
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.OptionButton optorder 
         Caption         =   "&Order"
         Height          =   210
         Left            =   1470
         TabIndex        =   30
         Top             =   300
         Width           =   720
      End
      Begin VB.OptionButton optdc 
         Caption         =   "D&C"
         Height          =   210
         Left            =   825
         TabIndex        =   29
         Top             =   300
         Width           =   510
      End
   End
Begin AceSpread spDCDetails
      Height          =   1920
      Left            =   240
      TabIndex        =   0
      Top             =   8205
      Width           =   5625
      _Version        =   458752
      _ExtentX        =   9922
      _ExtentY        =   3387
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
Begin AceSpread spItem
      Height          =   2010
      Left            =   195
      TabIndex        =   42
      ToolTipText     =   "Enter the Invoice Details"
      Top             =   1515
      Width           =   11565
      _Version        =   458752
      _ExtentX        =   20399
      _ExtentY        =   3545
      _StockProps     =   64
      Enabled         =   0   'False
      ColsFrozen      =   2
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
      MaxCols         =   19
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5327
      TabIndex        =   33
      ToolTipText     =   "Click to Clear the Data"
      Top             =   6930
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4035
      TabIndex        =   34
      ToolTipText     =   "Click to Save the Data"
      Top             =   6930
      Width           =   1300
   End
   Begin MSComCtl2.DTPicker dtpRefDate 
      Height          =   315
      Left            =   10410
      TabIndex        =   71
      ToolTipText     =   "Enter the Quotation Date"
      Top             =   8130
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   141230083
      CurrentDate     =   37250
   End
   Begin VB.Frame fraDetails 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   195
      TabIndex        =   1
      Top             =   3480
      Width           =   7815
      Begin VB.TextBox txtLCNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5970
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   83
         ToolTipText     =   "Enter the PLA Sl No"
         Top             =   255
         Width           =   1710
      End
      Begin VB.TextBox txtLCAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5970
         MaxLength       =   20
         TabIndex        =   82
         ToolTipText     =   "Enter the No of Packets"
         Top             =   1035
         Width           =   1710
      End
      Begin VB.TextBox txtLCBalAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5970
         MaxLength       =   20
         TabIndex        =   81
         ToolTipText     =   "Enter the No of Packets"
         Top             =   1425
         Width           =   1710
      End
      Begin VB.TextBox txtInsSlNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   12
         ToolTipText     =   "Enter the Insurance Sl No"
         Top             =   2970
         Width           =   1245
      End
      Begin VB.TextBox txtInsFor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   11
         ToolTipText     =   "Enter the Insured Amount"
         Top             =   2970
         Width           =   1125
      End
      Begin VB.TextBox txtInsPolicyNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1485
         MaxLength       =   20
         TabIndex        =   10
         ToolTipText     =   "Insurance Policy No"
         Top             =   2445
         Width           =   1860
      End
      Begin VB.TextBox txtPackets 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4755
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Enter the No of Packets"
         Top             =   2445
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtSlNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5235
         MaxLength       =   10
         TabIndex        =   8
         ToolTipText     =   "Enter the PLA Sl No"
         Top             =   1965
         Width           =   960
      End
      Begin VB.TextBox txtVehicleNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3795
         MaxLength       =   15
         TabIndex        =   7
         ToolTipText     =   "Enter the Vehicle No"
         Top             =   1965
         Width           =   1425
      End
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   105
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Enter the Narration"
         Top             =   1305
         Width           =   4320
      End
      Begin VB.TextBox txtAI2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Enter the Additional Information1"
         Top             =   825
         Width           =   4320
      End
      Begin VB.TextBox txtAI1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Enter the Additional Information1"
         Top             =   315
         Width           =   4335
      End
      Begin VB.TextBox txtMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   1965
         Width           =   3675
      End
      Begin VB.TextBox txtFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help."
         Top             =   2970
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker dtpPLADate 
         Height          =   330
         Left            =   105
         TabIndex        =   13
         ToolTipText     =   "Enter the PLA Date"
         Top             =   2445
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   138936321
         CurrentDate     =   38238
      End
      Begin MSComCtl2.DTPicker dtpInsDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         ToolTipText     =   "Enter the Policy Date"
         Top             =   2445
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   138936321
         CurrentDate     =   38238
      End
      Begin MSComCtl2.DTPicker dtpLCDate 
         Height          =   330
         Left            =   5970
         TabIndex        =   84
         ToolTipText     =   "Enter the PLA Date"
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   138936321
         CurrentDate     =   41759.5153240741
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "LC No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5475
         TabIndex        =   88
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "LC Date"
         Height          =   210
         Left            =   5340
         TabIndex        =   87
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "LC Amt."
         Height          =   210
         Left            =   5340
         TabIndex        =   86
         Top             =   1110
         Width           =   570
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Balance LC Amt."
         Height          =   210
         Left            =   4710
         TabIndex        =   85
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Insurance SlNo"
         Height          =   210
         Left            =   1260
         TabIndex        =   27
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Policy Dated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3360
         TabIndex        =   26
         Top             =   2265
         Width           =   1065
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Insured For"
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   2790
         Width           =   825
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Policy No"
         Height          =   210
         Left            =   1485
         TabIndex        =   24
         Top             =   2265
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No of Packets"
         Height          =   210
         Left            =   4815
         TabIndex        =   23
         Top             =   2265
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PLA Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   2265
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PLA Sl No"
         Height          =   210
         Left            =   5235
         TabIndex        =   21
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle No"
         Height          =   210
         Left            =   3795
         TabIndex        =   20
         Top             =   1770
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         Height          =   210
         Left            =   105
         TabIndex        =   19
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Additional Information2"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1635
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Additional Information1"
         Height          =   210
         Left            =   105
         TabIndex        =   17
         Top             =   135
         Width           =   1635
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Mode"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   16
         Top             =   1785
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2520
         TabIndex        =   15
         Top             =   2790
         Width           =   1125
      End
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   8130
      TabIndex        =   61
      Top             =   3465
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   65
         Top             =   135
         Width           =   900
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   64
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   150
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   63
         Top             =   150
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   62
         Top             =   2475
         Visible         =   0   'False
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   1965
         Left            =   45
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   3570
         _Version        =   458752
         _ExtentX        =   6297
         _ExtentY        =   3466
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   68
         Top             =   165
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   210
         Left            =   840
         TabIndex        =   67
         Top             =   2535
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Frame fraCustomer 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   43
      Top             =   750
      Width           =   11385
      Begin VB.TextBox TxtVrRoute 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   77
         ToolTipText     =   "Invoice No"
         Top             =   315
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox TxtVoucher 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5895
         TabIndex        =   75
         Top             =   315
         Width           =   2790
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   4335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   330
         Left            =   10740
         TabIndex        =   45
         ToolTipText     =   "Click to OK"
         Top             =   285
         Width           =   420
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "InvoiceChange.frx":106B
         Left            =   4425
         List            =   "InvoiceChange.frx":108D
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   "Select the Type of Invoice"
         Top             =   315
         Width           =   1440
      End
      Begin MSComCtl2.DTPicker DtpVrDate 
         Height          =   330
         Left            =   8715
         TabIndex        =   76
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Format          =   140902401
         CurrentDate     =   39483
      End
      Begin VB.Label lblVrRoute 
         AutoSize        =   -1  'True
         Caption         =   "Route"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   10185
         TabIndex        =   80
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5880
         TabIndex        =   79
         Top             =   135
         Width           =   1080
      End
      Begin VB.Label lblVoucherDate 
         AutoSize        =   -1  'True
         Caption         =   "Voucher Date"
         Height          =   210
         Left            =   8700
         TabIndex        =   78
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Type of Invoice"
         Height          =   210
         Left            =   4425
         TabIndex        =   48
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label lblcustcode 
         AutoSize        =   -1  'True
         Caption         =   "* Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   47
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.Label lblOrdDate 
      AutoSize        =   -1  'True
      Caption         =   "Ord. Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   10410
      TabIndex        =   74
      Top             =   7935
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblRefType 
      AutoSize        =   -1  'True
      Caption         =   "Ord. Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   8655
      TabIndex        =   73
      Top             =   7920
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblRefNo 
      AutoSize        =   -1  'True
      Caption         =   "Ord. No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   9810
      TabIndex        =   72
      Top             =   7935
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmInvoiceChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

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
  mGroupKey = 19
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Dim SKey As String

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

Private Sub Orders()
  Dim Sql As String
  Dim i As Integer
  Dim CurDate As Date
  Dim sRow As Long
  
  Try
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
  Sql = "execute SalesOrderList '" + Format(dtpInvDate, "yyyy-MM-dd") + "', " & Val(txtcustcd.Tag) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "'"
  If FetchValue_New(True, db.ConnectionString, Sql, "111000101110000", "List of Item's", "Part No", "Item Code", "Item Description", "Rate", "Order Type", "Order No", "Order Date") = True Then
    
    spItem.MaxRows = UBound(getvalue, 1)
    
    If ValidateItemCount(MaxItemCount, spItem.MaxRows) = True Then
      For i = 1 To UBound(getvalue, 1)
        If Trim(getvalue(i, 9)) = Trim(getvalue(1, 9)) And Val(getvalue(i, 10)) = Val(getvalue(1, 10)) And Trim(getvalue(i, 11)) = Trim(getvalue(1, 11)) Then
          txtFormula.Tag = Val(getvalue(i, 14)) 'Formula Code
          txtFormula.Text = Trim$(getvalue(i, 15)) 'Formula Name
          sRow = spItem.DataRowCnt + 1
          spItem.SetText eItem.mItemCode, sRow, Trim(getvalue(i, 2)) 'Item Code
          spItem.SetText eItem.mItemDesc, sRow, getvalue(i, 3) 'Item Desc
          spItem.SetText eItem.mUOM, sRow, getvalue(i, 4) ' UomDesc
          Call Spread_Set_Decimal(spItem, Val(getvalue(i, 5)), eItem.mStockQty, sRow, eItem.mInvQty)
          Call Spread_Set_Decimal(spItem, Val(getvalue(i, 5)), eItem.mBalQty, sRow)
          spItem.SetText eItem.mRate, sRow, getvalue(i, 7)  'Rate
          spItem.SetText eItem.mARate, sRow, getvalue(i, 8) 'Assessable Rate
          spItem.SetText eItem.mOType, sRow, getvalue(i, 9) 'Order Type
          spItem.SetText eItem.mONo, sRow, getvalue(i, 10) 'Order No
          spItem.SetText eItem.mODate, sRow, getvalue(i, 11) 'Order Date
          spItem.SetText eItem.mBalQty, sRow, getvalue(i, 12) 'Bal Qty
          spItem.SetText eItem.mSchType, sRow, getvalue(i, 13)  ' Schedule Type
          spItem.SetText eItem.mGroupKey, sRow, getvalue(i, 17)  ' Item GroupKey
          Call Spread_Lock(spItem, True, eItem.mInvQty, sRow)
        End If
      Next i
    End If
    
    TermsDetails   'To display the Terms & Formula
    FraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
    cmdSave.Enabled = True
    spItem.SetFocus
    Spread_Set_Focus spItem, 1, 4
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub TermsDetails()
  Dim snap As New AceADODB.Recordset
  Dim Sql As String

  Try
 
  Sql = " select a.CustomerFormulaCode,(select distinct b.FrmLName from COM_FormulaMast_Hdr b where"
  Sql = Sql + " b.ModuleCode = '" & gModuleCode & "' and b.FrmId = a.CustomerFormulaCode and b.Branchid = a.Branchid "
  Sql = Sql + " and b.Compid = a.Compid) FrmLName, a.ModeCode,"
  Sql = Sql + " (select b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid and b5.Compid=a.Compid) ModeDesc"
  Sql = Sql + " from FAS_PARTY_MASTER a where a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.PartyCode = " & Val(txtcustcd.Tag) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    If Not IsNull(snap("CustomerFormulaCode")) = True Then
 txtFormula.Tag = snap("CustomerFormulaCode")
End If
    If Not IsNull(snap("FrmLName")) = True Then
 txtFormula = snap("FrmLName")
End If
    If Not IsNull(snap("ModeCode")) = True Then
 txtMode.Tag = snap("ModeCode")
End If
    If Not IsNull(snap("ModeDesc")) = True Then
 txtMode.Text = snap("ModeDesc")
End If
  End If
  snap.Close
'  Set snap = Nothing
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Sub

Private Sub Direct()
  Dim Sql As String
  Dim i As Long

  Try
  
  Sql = " select a.ItemCode,c.ItemDescription,a.PartNo,c.UOMDecimal,a.Rate,"
  Sql = Sql + " a.AssessableRate from SAL_Customer_ItemMaster a left outer join"
  Sql = Sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid"
  Sql = Sql + " and a.Compid=C.Compid where a.CustomerCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and c.Status = 1"
  Sql = Sql + " and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " union (select a.ItemCode,c.ItemDescription,a.ItemCode,c.UOMDecimal,a.Rate, "
  Sql = Sql + " a.AssessableRate from PUR_Vendor_ItemMaster a left outer join "
  Sql = Sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid "
  Sql = Sql + " and a.Compid=C.Compid where a.VendorCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and c.Status = 1"
  Sql = Sql + " and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111))"
'  Sql = Sql + " select a.ItemCode,c.ItemDescription,a.PartNo,c.UOMDecimal,a.Rate,"
'  Sql = Sql + " a.AssessableRate from SAL_Customer_ItemMaster a left outer join"
'  Sql = Sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid"
'  Sql = Sql + " and a.Compid=C.Compid where a.CustomerCode=" & Val(txtcustcd.Tag) & ""
'  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
'  Sql = Sql + " and a.ValidTo >=convert(datetime,'" + Format(Serverdate, "yyyy-MM-dd") + "',111)"
  If FetchValue_New(True, db.ConnectionString, Sql, "111000", "List of Item's", "Item Code", "Item Description", "Part No") = True Then
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
        If cmbType = "Supplementary" Then
          Call Spread_Lock(spItem, False, eItem.mInvQty, i, eItem.mARate)
        Else
          Call Spread_Lock(spItem, True, eItem.mInvQty, i, eItem.mARate)
        End If
      Next i
    End If
    
    TermsDetails   'To display the Terms & Formula
    FraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
    cmdSave.Enabled = True
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
  
  If optdc = False Then
    For i = 1 To spItem.DataRowCnt
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
      
    Next i
  End If
  
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

Private Sub cmdClear_Click()
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  Call LoadDocType(mInvChange)
  FraInvoice.Enabled = True
  fraSource.Enabled = True
  fraCustomer.Enabled = True
  spItem.Enabled = False
  fraDetails.Enabled = False
  cmdSave.Enabled = False
  cmbType.ListIndex = 0
'  LoadInvtype
  CmbInvType.SetFocus
  fraSource.Enabled = False
  fraCustomer.Enabled = False
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

Private Sub CmdOk_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If Trim$(CmbInvType) = "" Then
    sShowMessage "Please Select the Invoice Type"
    CmbInvType.SetFocus
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
  If optOrder Then
 Orders 'list the Orders
End If
  
  If cmbType = "Rejection" Then
    gModuleCode = 9
  Else
    gModuleCode = 4
  End If

  If cmbType.Text = "Supplementary" Then
    Call Spread_Lock(spItem, False, eItem.mInvQty, 1, eItem.mARate, -1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim Sql As String
  Dim tmp As Variant
  Dim sRate As Double
  Dim Stime As Date
  Dim i As Long
  Dim CurDate As Date
  
  Try
  
  Set cmd = New AceADODB.Command
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
  sRate = 0
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  commClear cmd
  cmd.CommandText = "SAL_Invoice_Delete"
  cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(CmbInvType))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("Branchid", adVarChar, adParamInput, 3, gBranchID)
  cmd.Parameters.Append cmd.CreateParameter("Compid", adVarChar, adParamInput, 3, gCompID)
  cmd.Parameters.Append cmd.CreateParameter("finyear", adVarChar, adParamInput, 20, gFinyear)
  cmd.Execute
  
  commClear cmd
  cmd.CommandText = "SAL_INS_INVOICE_HDR"
  cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(CmbInvType))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("InvoiceTime", adDBTimeStamp, adParamInput, , Format(dtpInvTime, "HH:mm"))
  If optdirect Then
    cmd.Parameters.Append cmd.CreateParameter("InvoicSource", adChar, adParamInput, 1, "D")
  ElseIf optOrder Then
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
  cmd.Parameters.Append cmd.CreateParameter("RemovelTime", adDBTimeStamp, adParamInput, , Format(DateAdd("n", 30, dtpInvTime), "HH:mm"))
  cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , Format(Val(txtGrandTotal), "0.00"))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute
  
    For i = 1 To spItem.DataRowCnt
      spItem.GetText eItem.mRate, i, tmp
sRate = Val(tmp)
      
      commClear cmd
      cmd.CommandText = "SAL_INS_INVOICE_DTL"
      cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(CmbInvType))
      cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(txtinvnum))
      cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-MM-dd"))
      spItem.GetText eItem.mGroupKey, i, tmp
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
          cmd.CommandText = "INV_INS_MATERIALTRAN_DTL"
          cmd.Parameters.Append cmd.CreateParameter("DOCType", adVarChar, adParamInput, 3, Trim$(CmbInvType))
          cmd.Parameters.Append cmd.CreateParameter("DOCNo", adInteger, adParamInput, , Val(txtinvnum))
          cmd.Parameters.Append cmd.CreateParameter("DOCDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-mm-dd"))
          cmd.Parameters.Append cmd.CreateParameter("TransactionDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-MM-dd"))
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
          spItem.GetText eItem.mGroupKey, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
          cmd.Parameters.Append cmd.CreateParameter("FromItem", adVarChar, adParamInput, 20, "")
          spItem.GetText eItem.mItemCode, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ToItem", adVarChar, adParamInput, 20, "")
          cmd.Parameters.Append cmd.CreateParameter("Justification", adVarChar, adParamInput, 4, "")
          Call AddCommonCmdParameters(cmd)
          cmd.Execute
        End If
      End If
    Next i
  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If SaveCalcDetails = False Then
    GoTo ExitHere
  End If
  
  'To FAS Posting
  If FASAccountPosting(dtpInvDate) = False Then
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
    cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(CmbInvType))
    cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(Trim$(txtinvnum)))
    cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(dtpInvDate, "yyyy-MM-dd"))
    Call AddCommonCmdParameters(cmd)
    cmd.Execute
  Next i
    
  ' Update Itemdescription, Qty and Uom as Narration in FAS_VoucherTranMast Table
  If UpdateNarrationToFASVrMast = False Then
    GoTo ExitHere
  End If
  
  If Trim$(txtLCNo.Text) <> "" Then     'Val(txtcustcd.Tag) = 11
    Sql = " Execute SAL_Invoice_LCRef_Ins"
    Sql = Sql & " @InvType = '" & Trim$(CmbInvType) & "',"
    Sql = Sql & " @InvNo = " & Val(Trim$(txtinvnum.Text)) & ","
    Sql = Sql & " @InvDate = '" & Format(dtpInvDate, "yyyy-MM-dd") & "',"
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
  Sql = Sql & " Where VrRoute = '0K' and CompCode = '" & gFasCompCode & "' and RefType = '" & (CmbInvType.Text) & "'"
  Sql = Sql & " and RefNo = " & (txtinvnum.Text) & " and RefDate = '" & Format(dtpInvDate, "yyyy-MM-dd") & "'"
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
    
    If Trim$(Narr) <> "" Then
      Narr = Trim$(Narr) & "."
         
      Sql = " Update FAS_VoucherTranMast Set Narration = '" & Trim$(Narr) & "'"
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
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    CurDate = ServerDate
    
    
    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(CmbInvType) & "' DocType," & Val(txtinvnum) & " DocNo,"
    Sql = Sql & "'" & Format(dtpInvDate, "yyyy/MM/dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,'',"
    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear,'" & Format(CurDate, "yyyy-MM-dd HH:MM:SS  ") & "'"
    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
'''    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(cmbinvtype) & "' DocType," & Val(txtinvnum) & " DocNo,"
'''    Sql = Sql & "'" & Format(dtpInvDate, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
'''    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,"
'''    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
'''    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
'''    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
'''    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
'''    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    
    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(CmbInvType) & "', " & Val(txtinvnum) & ", '" & Format(dtpInvDate, "yyyy-MM-dd") & "',"
    Sql = Sql & " " & gModuleCode & " ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, a.IsValueEditable, a.ValueUsed, a.IsValueEditable, '',a.UserID, a.BranchID, a.CompID, a.Finyear,'" & Format(CurDate, "yyyy-MM-dd HH:MM:SS  ") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      
'''    Sql = "Insert into COM_FormulaCalc_Dtl"
'''    Sql = Sql & " Select '" & Trim$(cmbinvtype) & "', " & Val(txtinvnum) & ", '" & Format(dtpInvDate, "yyyy-MM-dd") & "',"
'''    Sql = Sql & " " & gModuleCode & " ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, "
'''    Sql = Sql & " a.IsValueEditable, a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
'''    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
'''    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
'''    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  SaveCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub dtpInvDate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
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
  
  Call LoadInitialData
  
'  LoadInvtype
  
  'frahide.Visible = False
'  If gReportID = "KUS" Then
'    dtpInvDate.Enabled = True
'  Else
'    dtpInvDate.Enabled = False
'  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub LoadInitialData()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
 
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

'Private Function GetMax(TblName As String, FldName As String, Optional cmpCode As String = "", Optional ReqCond As String = "") As String
'  Dim Sql As String
'  Dim adodbrs As ADODB.Recordset
'  Dim MaxNo As String
'
'  On Error GoTo ErrHandler
'  Set adodbrs = New ADODB.Recordset
'
'  Sql = "Select Max(" & Trim$(FldName) & ") from " & Trim$(TblName)
'  If Trim$(cmpCode) <> "" And Trim$(ReqCond) <> "" Then
'    Sql = Sql & " Where CompCode='" & Trim$(cmpCode) & "' and " & Trim$(ReqCond)
'  ElseIf Trim$(cmpCode) = "" And Trim$(ReqCond) <> "" Then
'    Sql = Sql & " Where " & Trim$(ReqCond)
'  ElseIf Trim$(cmpCode) <> "" And Trim$(ReqCond) = "" Then
'    Sql = Sql & " Where CompCode='" & Trim$(cmpCode) & "'"
'  End If
'  If RSOpen(adodbrs, Sql) = False Then Exit Function
'  If adodbrs.EOF Or IsNull(adodbrs(0)) Then
'    MaxNo = "0"
'  Else
'    MaxNo = adodbrs(0)
'  End If
'  GetMax = MaxNo
'
'ExitHere:
'  On Error GoTo 0
'  Set adodbrs = Nothing
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Function

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  Dim fCurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  'Get ServerDate
  fCurrDate = ServerDate
  
  If Col = eItem.mWhsName And Row > 0 And optdc = False And cmbType <> "Supplementary" Then
    spItem.GetText 1, Row, tmp
    
    Sql = "Select y.WareHouseName, x.WareHouseCode, x.RackBinCode, Sum(x.CBQty)"
    Sql = Sql & " From ("
    Sql = Sql & " select b.WarehouseCode, '' as RackBinCode, a.CBQty"
    Sql = Sql & " From dbo.fn_INV_AsOnDate_Stock_V2('" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "',  '" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(fCurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
    'Sql = Sql & " from INV_STOCK_MASTER a "
    Sql = Sql & " Right Outer Join INV_WAREHOUSE_MASTER b"
    Sql = Sql & " On a.WarehouseCode=b.WarehouseCode and a.Branchid=b.Branchid  and a.Compid=b.Compid"
    Sql = Sql & " where b.WarehouseStatus = 1 and CBQty >0"
    Sql = Sql & " and a.ItemCode = '" & Trim$(tmp) & "'"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Sql = Sql & " Union All"
    Sql = Sql & " Select FromWareHouseCode, '' as RackBinCode, TransactionQty"
    Sql = Sql & " From INV_MaterialTran_Dtl"
    Sql = Sql & " Where DocType='" & Trim$(CmbInvType.Text) & "' and DocNo=" & Val(txtinvnum.Text) & " and DocDate='" & Format(dtpInvDate, "yyyy-MM-dd") & "'"
    
    spItem.GetText 1, Row, tmp
    Sql = Sql & " and TransactionItem='" & Trim$(tmp) & "'"
    Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Sql = Sql & " ) x"
    Sql = Sql & " Left Outer Join (Select WareHouseName, WareHouseCode"
    Sql = Sql & " from INV_WAREHOUSE_MASTER "
    Sql = Sql & " Where WareHouseStatus = 1 and BranchID='" & gBranchID & "' and CompID='" & gCompID & "') y"
    Sql = Sql & " On x.WarehouseCode=y.WarehouseCode"
    Sql = Sql & " Group by y.WareHouseName, x.WareHouseCode, x.RackBinCode"
    
    If FetchValue_New(False, db.ConnectionString, Sql, "1011", "List of Stock Details", "Warehouse Name", "Rack & Bin", "Stock Qty") = True Then
      spItem.SetText eItem.mWhsCode, Row, getvalue(2)
      spItem.SetText eItem.mWhsName, Row, getvalue(1)
      spItem.SetText eItem.mRackCode, Row, getvalue(3)
      spItem.SetText eItem.mStockQty, Row, getvalue(4)
      spItem.GetText eItem.mSchType, Row, tmp
      Call Spread_Lock(spItem, False, eItem.mInvQty, Row)
      If Trim$(tmp) = "S" And optOrder Then
        spItem.GetText eItem.mBalQty, Row, tmp
'        Spread_Max_Set spItem, eItem.mInvQty, eItem.mInvQty, Row, IIf(Val(getvalue(4)) > Val(tmp), Val(tmp), Val(getvalue(4)))
        Call Spread_Set_Max(spItem, IIf(Val(getvalue(4)) > Val(tmp), Val(tmp), Val(getvalue(4))), eItem.mInvQty, Row)
      Else
'        Spread_Max_Set spItem, eItem.mInvQty, eItem.mInvQty, Row, Val(getvalue(4))
        Call Spread_Set_Max(spItem, Val(getvalue(4)), eItem.mInvQty, Row)
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

Private Sub txtcustcd_DblClick()
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Sql = " Select distinct (Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.CustomerCode "
  Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.CustomerCode From SAL_Customer_ItemMaster a"
  Sql = Sql & " Where a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " Union (Select distinct (Select b1.PartyName From FAS_PARTY_MASTER b1 Where b1.PartyCode=a1.CustomerCode "
  Sql = Sql & " and b1.Branchid=a1.Branchid and b1.Compid=a1.Compid) PartyName,a1.CustomerCode From SAL_Jobwork_ItemMaster a1"
  Sql = Sql & " Where a1.Branchid='" & gBranchID & "' and a1.Compid='" & gCompID & "')"
  If optdirect Then
    Sql = Sql & " Union Select distinct (Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.VendorCode "
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
  
  Try

  Sql = "select a.DCType,a.DCNo,a.DCDate,a.ItemCode,b.ItemDescription,"
  Sql = Sql + " b.UomDecimal,a.DCQty,a.Rate,a.AssessableRate,a.OrderType,a.OrderNo,a.OrderDate "
  Sql = Sql + " from SalesDC a left outer join MATERIAL_MASTER b on a.ItemCode=b.ItemCode "
  Sql = Sql + " and a.Branchid=b.Branchid and a.Compid=b.Compid  where a.InvoiceType='' and b.Status = 1 "
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql + " and a.CustomerCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql + " order by a.ItemCode"
  If FetchValue_New(True, db.ConnectionString, Sql, "1111101000000", "List of DC's", "DC Type", "DC No", "DC Date", "Item Code", "Item Description", "DC Qty") = True Then
    spDCDetails.MaxRows = UBound(getvalue, 1)
    Qty = 0
Rate = 0
ARate = 0
    tot = UBound(getvalue, 1)
    spItem.MaxRows = tot
    
    If ValidateItemCount(MaxItemCount, spItem.MaxRows) = True Then
      For i = 1 To tot
        If getvalue(i, 1) <> "" And (getvalue(1, 10) = getvalue(i, 10)) And (getvalue(1, 11) = getvalue(i, 11)) And (getvalue(1, 12) = getvalue(i, 12)) Then
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
        End If
      Next i
    End If
    
    If i = (tot + 1) And Qty > 0 Then
      sRow = spItem.DataRowCnt + 1
      spItem.SetText eItem.mItemCode, sRow, Trim(Item)
      spItem.SetText eItem.mItemDesc, sRow, Trim(desc)
'      Spread_Decimal_Set spDCDetails, eItem.mStockQty, eItem.mInvQty, SRow, SRow, UOMDescimal
      Call Spread_Set_Decimal(spDCDetails, UOMDescimal, eItem.mStockQty, sRow, eItem.mInvQty)
      spItem.SetText eItem.mInvQty, sRow, Val(Qty)
      spItem.SetText eItem.mRate, sRow, Val(Rate)
      spItem.SetText eItem.mARate, sRow, Val(ARate)
      spItem.SetText eItem.mGV, sRow, Val(Qty) * Val(Rate)
      spItem.SetText eItem.mAGV, sRow, Val(Qty) * Val(ARate)
      Qty = 0
Rate = 0
ARate = 0
    End If
    TermsDetails   'To display the Terms & Formula
    FraInvoice.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    fraDetails.Enabled = True
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

Private Sub txtinvnum_DblClick()
  Dim Sql As String
  Dim cmpCode As String
  Dim InvTypeRec As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass

  Sql = " Select InvoiceNo,InvoiceDate From Sal_Invoice_hdr Where Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "11", "List of Invoice", "Invoice No", "Invoice Date") = True Then
    txtinvnum.Text = getvalue(1)
    dtpInvDate.Value = getvalue(2)
    fraSource.Enabled = True
    fraCustomer.Enabled = True
  
  cmpCode = gCompID & gBranchID & Mid(gFinyear, 4, 4) + Right(gFinyear, 4)
  
  Sql = " Select a.InvoiceTime, a.CustomerCode, d.Doctype, d.DocNo, d.DocDate, e.FileRef"
  Sql = Sql & " FROM SAL_INVOICE_HDR a"
  Sql = Sql & " Left Outer join FAS_VoucherTranRef d on d.RefType=a.InvoiceType and d.RefNo=a.InvoiceNo"
  Sql = Sql & " and d.RefDate=a.InvoiceDate and d.Compcode='" & Left(cmpCode, 4) & "'"
  Sql = Sql & " Left Outer join FAS_VoucherTranMast e on e.DocType=d.DocType and e.DocNo=d.DocNo"
  Sql = Sql & " and e.DocDate=d.DocDate and e.CompCode='" & Left(cmpCode, 4) & "'"
  Sql = Sql & " Where a.InvoiceNo='" & txtinvnum & "' and a.InvoiceDate='" & Format(dtpInvDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.InvoiceType='" & CmbInvType & "' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  
  Set InvTypeRec = New AceADODB.Recordset
  If RSOpen(InvTypeRec, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not InvTypeRec.EOF Then
    dtpInvTime = Format(InvTypeRec("InvoiceTime"), "HH:mm:ss")
    txtcust = InvTypeRec("CustomerCode")
    txtVrType = CheckNull(InvTypeRec("DocType"))
    txtVrNo = InvTypeRec("DocNo")
    txtVrDate = Format(InvTypeRec("Docdate"), "yyyy-MM-dd")
    txtFileRef = InvTypeRec("FileRef")
  End If
 End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set InvTypeRec = Nothing
  Me.MousePointer = vbDefault
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
  Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and '" & Format(dtpInvDate, "yyyy-MM-dd") & "' between a.ValidFrom and a.ValidTo"
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

Private Sub txtMode_DblClick()
  Dim Sql As String
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description,Code From TERMSMASTER Where typeflg = 'Mode' and Compid = '" & gCompID & "'"
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
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(retVal)
        
        'SCValue List
        spItem.GetText eItem.mInvQty, rCnt, retVal
        Qty = Val(retVal)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
'        .GetText eItem.mRate, rCnt, retVal
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
  
  SprTotCalc = Total
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
  
  CurDate = ServerDate
  
  If cmbType = "Rejection" Then
    Sql = " select a.PostingAc from dbo.fn_PUR_PurchaseOrder_Hdr('" & Format(dtpInvDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dtpInvDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
    Sql = Sql + " left outer join dbo.fn_PUR_PurchaseOrder_Dtl('" & Format(dtpInvDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dtpInvDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') b"
    Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate and a.Branchid = b.Branchid and a.Compid = b.Compid "
    spItem.GetText eItem.mItemCode, 1, retVal
    Sql = Sql + " where a.Status = 'A' and b.ItemCode = '" + Trim$(retVal) + "' order by postingac desc"
  Else
    Sql = "Select PostingAc"
    Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " Where OrderType='" & Trim$(txtRefType.Text) & "' and OrderNo=" & Val(txtRefNo.Text) & " and OrderDate='" & Format(dtpRefdate, "yyyy-MM-dd") & "'"
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
    Sql = "Select LedgerName, LedgerCode from FAS_LedgerMaster"
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
  
  If cmbType = "Rejection" Then
    If FASPosting(dtpInvDate, eVrType.mSales, Trim$(TxtVrRoute), DtpVrDate, Val(txtcustcd.Tag), OrderAc, Val(txtFormula.Tag), sprCalc, txtGrandTotal, , , Trim$(CmbInvType.Text), Val(txtinvnum.Text), dtpInvDate, Trim$(txtinvnum.Text), dtpInvDate, Trim$(CmbInvType.Text), Val(txtinvnum.Text), dtpInvDate, 1, 1, 2) = False Then

      GoTo ExitHere
    End If
  Else
    If FASPosting(dtpInvDate, eVrType.mSales, Trim$(TxtVrRoute), DtpVrDate, Val(txtcustcd.Tag), OrderAc, Val(txtFormula.Tag), sprCalc, txtGrandTotal, , , Trim$(CmbInvType.Text), Val(txtinvnum.Text), dtpInvDate, Trim$(txtinvnum.Text), dtpInvDate, Trim$(txtRefType.Text), Val(txtRefNo.Text), dtpRefdate, 1, 1, 2) = False Then

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

Public Sub OldInvoiceCalcuation()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim i As Integer
  Dim ItemList() As String
  Dim SCValList() As String

  Try

'  Set snap = New ADODB.Recordset
  
  Sql = "select invoicetype, invoiceno, invoicedate, customercode, formulacode, sum((invoiceqty*rate)) gv"
  Sql = Sql + " from salesinvoice where formulacode > 0 "
  Sql = Sql + " and invoiceno not in (select docno from COM_formulacalc_Hdr where ModuleCode = " & gModuleCode & ")"
  Sql = Sql + " group by invoicetype, invoiceno, invoicedate, customercode, formulacode order by invoiceno"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    For i = 1 To snap.RecordCount
      ReDim ItemList(0)
      ItemList(0) = "0" & Chr(145) & "0"
      
      ReDim SCValList(1)
      SCValList(0) = eParamID.mGV & "-0-" & Val(snap("GV"))
      SCValList(1) = eParamID.mAGV & "-0-" & Val(snap("GV"))
      
      txtFormula.Tag = Val(snap("formulacode"))
      
      If CalculateFormula(Val(snap("customercode")), Val(snap("formulacode")), txtUserKey, SCValList, ItemList, txtGrandTotal, sprCalc, eCalc.mDesc, eCalc.mValue, , "Sales Order", False, False) = False Then

                  
        GoTo ExitHere
      End If
      
      db.BeginTrans
      txtinvnum = snap("invoiceno")
      dtpInvDate = snap("invoicedate")
      'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
      If SaveCalcDetails = False Then
        GoTo ExitHere
      End If
      db.CommitTrans
      snap.MoveNext
    Next i
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)
  db.RollbackTrans

Finally

  Set snap = Nothing
End Try
End Sub

Private Sub TxtVoucher_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select VoucherName, VoucherCode, (VrRoute+VrWidth) as Path From FAS_VoucherMaster"
  Sql = Sql & " Where (VrRoute+VrWidth) Like '0k%' and IsActive = 1 and CompCode = '" & gFasCompCode & "'"
  Sql = Sql & " Order by VrRoute+VrWidth"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Voucher", "Voucher Name", "Path") = True Then
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

Private Sub Design_spDCDetails(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDCDetails

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.EditMode +=  new EventHandler(spItem_EditMode)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_spDCDetails("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InvoiceChange.frm", spDCDetails, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InvoiceChange.frm", spItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InvoiceChange.frm", sprCalc, "")

End Sub
