VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobworkOrder 
   Caption         =   "Jobwork Order"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
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
   ScaleHeight     =   8265
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSch 
      Height          =   1875
      Left            =   255
      TabIndex        =   63
      Top             =   6030
      Visible         =   0   'False
      Width           =   7590
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   64
         Top             =   195
         Width           =   2175
      End
Begin AceSpread spSchedule
         Height          =   1185
         Left            =   60
         TabIndex        =   65
         Top             =   600
         Width           =   7440
         _Version        =   458752
         _ExtentX        =   13123
         _ExtentY        =   2090
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
         Caption         =   "Item Code"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   66
         Top             =   240
         Width           =   855
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   720
      Left            =   630
      TabIndex        =   52
      Top             =   0
      Width           =   3150
      Begin MSComCtl2.DTPicker dporddate 
         Height          =   315
         Left            =   1890
         TabIndex        =   5
         ToolTipText     =   "Jobwork Order Date"
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
         Format          =   20185091
         CurrentDate     =   37250
      End
      Begin VB.TextBox txtordnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Jobwork Order No"
         Top             =   330
         Width           =   930
      End
      Begin VB.ComboBox cmbordtype 
         Height          =   330
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select the Jobwork Order Type"
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1905
         TabIndex        =   4
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1230
         TabIndex        =   2
         Top             =   150
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Type"
         Height          =   210
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5565
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5235
      TabIndex        =   48
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5565
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3945
      TabIndex        =   47
      ToolTipText     =   "Click to Save the Data"
      Top             =   5565
      Width           =   1300
   End
   Begin VB.Frame fraQuotation 
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   5775
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtQuoType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Jobwork Quotation Type"
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
         MultiLine       =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Jobwork Quotation Date"
         Top             =   330
         Width           =   1140
      End
      Begin VB.TextBox txtQuoNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Select the Jobwork Quotation. Press F2 or Double Click for Help"
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quot. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1905
         TabIndex        =   12
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo. No."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1275
         TabIndex        =   10
         Top             =   150
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   105
         TabIndex        =   8
         Top             =   150
         Width           =   705
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Order Source"
      Height          =   720
      Left            =   3795
      TabIndex        =   6
      ToolTipText     =   "Select the Jobwork Order Source"
      Top             =   0
      Width           =   1965
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
         TabIndex        =   51
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
         TabIndex        =   50
         ToolTipText     =   "Select if it is a Direct Order"
         Top             =   300
         Value           =   -1  'True
         Width           =   720
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
      Left            =   630
      TabIndex        =   53
      Top             =   660
      Width           =   10515
      Begin VB.TextBox txtPostingAc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4725
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   825
         Width           =   4620
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   25
         ToolTipText     =   "Enter the Remarks"
         Top             =   825
         Width           =   4605
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   4620
      End
      Begin VB.TextBox txtrefno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   17
         ToolTipText     =   "Enter the Ref. No"
         Top             =   315
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker dprefdate 
         Height          =   315
         Left            =   6645
         TabIndex        =   19
         ToolTipText     =   "Enter Ref. Date"
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
         Format          =   20185091
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   7920
         TabIndex        =   21
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
         Format          =   20185091
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   9195
         TabIndex        =   23
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
         Format          =   20185091
         CurrentDate     =   37250
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Account"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4725
         TabIndex        =   26
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   14
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref No."
         Height          =   210
         Left            =   4740
         TabIndex        =   16
         Top             =   135
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Date"
         Height          =   210
         Left            =   6645
         TabIndex        =   18
         Top             =   135
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid From"
         Height          =   210
         Left            =   7920
         TabIndex        =   20
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         Height          =   210
         Left            =   9195
         TabIndex        =   22
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   210
         Left            =   105
         TabIndex        =   24
         Top             =   645
         Width           =   630
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
      Height          =   1710
      Left            =   270
      TabIndex        =   54
      Top             =   1830
      Width           =   11220
      Begin VB.ComboBox cmbScheduleType 
         Height          =   330
         ItemData        =   "JobworkOrder.frx":03CB
         Left            =   7455
         List            =   "JobworkOrder.frx":03D5
         Style           =   2  'Dropdown List
         TabIndex        =   45
         ToolTipText     =   "Select the Sales Order Type"
         Top             =   1305
         Width           =   1410
      End
      Begin VB.TextBox txtDelIns 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "Select theDelivery Instruction. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtFreight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   37
         ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtInsurance 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtTerms 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3675
      End
      Begin VB.TextBox txtMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtForm 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Select the Form. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   29
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
         TabIndex        =   43
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Type"
         Height          =   210
         Left            =   7455
         TabIndex        =   44
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7470
         TabIndex        =   38
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3780
         TabIndex        =   30
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   34
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing && Forwarding"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   28
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3780
         TabIndex        =   36
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   40
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
         TabIndex        =   32
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3780
         TabIndex        =   42
         Top             =   1125
         Width           =   1020
      End
   End
Begin AceSpread sprItem
      Height          =   1920
      Left            =   285
      TabIndex        =   46
      ToolTipText     =   "Display the Jobwork Item Details"
      Top             =   3570
      Width           =   11205
      _Version        =   458752
      _ExtentX        =   19764
      _ExtentY        =   3387
      _StockProps     =   64
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7875
      TabIndex        =   55
      Top             =   5400
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   56
         Top             =   135
         Width           =   900
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   59
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   150
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   2355
         Visible         =   0   'False
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   1845
         Left            =   45
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   3570
         _Version        =   458752
         _ExtentX        =   6297
         _ExtentY        =   3254
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
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   62
         Top             =   165
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   975
         TabIndex        =   61
         Top             =   2400
         Visible         =   0   'False
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmJobworkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim FrmLName As String
Dim SKey As String

'Item Spread Column Numbers
Private Enum eItem
  mFirstRow = 1
  mrecdItem = 2
  mRecdDesc = 3
  mRecdUOM = 4
  msentItem = 5
  mSentDesc = 6
  mPartNo = 7
  mSentUOM = 8
  mSentQty = 9
  mRatePer = 10
  mRate = 11
  mAssRate = 12
  mGoodsValue = 13
  mAssGoodsValue = 14
  mRelation = 15
  mItemGroupKey = 16
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Enum eSch
  mSchDate = 1
  mSchQty = 2
  mLocation = 3
End Enum

Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim CurDate As Date
  Dim srow As Long
  Dim ret As Integer
  Dim sCnt As Long
  Dim retVal As Variant
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
  Try
  
  Chkbefsave = False
  
  
  If Trim$(cmbordtype) = "" Then
    sShowMessage "Select the Jobwork Order Type"
    cmbordtype.Enabled = True
    cmbordtype.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtcustcd) = "" Then
    sShowMessage "Select the Customer"
    txtcustcd.Enabled = True
    txtcustcd.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtRefNo) = "" Then
    sShowMessage "Enter Reference No"
    txtRefNo.Enabled = True
    txtRefNo.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtFormula.Text) = "" Then
    sShowMessage "Select the ForMula Name"
    cmbordtype.Enabled = True
    txtFormula.SetFocus
    GoTo ExitHere
  End If
  
  If dprefdate > CurDate Then
    sShowMessage "Reference Date should be less the or equal to Current Date"
    dprefdate.Enabled = True
    dprefdate.SetFocus
    GoTo ExitHere
  End If
  
  If dpvalidfrom < CurDate Then
    sShowMessage "Valid From Date should be greater than or equal to Current Date"
    dpvalidfrom.Enabled = True
    dpvalidfrom.SetFocus
    GoTo ExitHere
  End If

  If dpvalidto < dpvalidfrom Then
    sShowMessage "Valid Till Date should be greater than or equal to Valid From Date"
    dpvalidto.Enabled = True
    dpvalidto.SetFocus
    GoTo ExitHere
  End If
  
  If Not Trim$(txtQuoDate.Tag) = "" Then
    If dporddate > CDate(txtQuoDate.Tag) Then
      ret = MsgBox("Jobwork Order Date is greater than Valid till Date of Quotation" & vbCrLf & " Do you want to continue creating the Order ?", vbYesNo + vbQuestion, App.Title)

      If Val(ret) <> 6 Then
'        fraOrder.Enabled = True
'        txtQuoNo = "": txtQuoNo.SetFocus:
        GoTo ExitHere
      End If
    End If
  End If

  If sprItem.DataRowCnt = 0 Then
    sShowMessage "Select atleast one item"
    sprItem.Enabled = True
    sprItem.SetFocus
    Spread_Set_Focus sprItem, 1, 1
    GoTo ExitHere
  End If
  
  For i = 1 To sprItem.DataRowCnt
    sprItem.Row = i
sprItem.Col = eItem.mrecdItem
    If Trim$(sprItem.Text) = "" Then
      sShowMessage "Select the Item Code"
      Spread_Set_Focus sprItem, i, eItem.mrecdItem
      GoTo ExitHere
    End If
    If Trim$(cmbScheduleType.Text) = "Specific" Then
      sprItem.Row = i
sprItem.Col = eItem.mFirstRow
srow = Val(sprItem.Text)

      sprItem.Row = srow
sprItem.Col = eItem.mSentQty
      If Val(sprItem.Text) = 0 Then
        sShowMessage "Select the Order Qty"
        Spread_Set_Focus sprItem, srow, eItem.mSentQty
        GoTo ExitHere
      End If
    End If
    sprItem.Row = i
sprItem.Col = eItem.mFirstRow
      srow = Val(sprItem.Text)
    sprItem.GetText eItem.mRate, srow, retVal
    If Val(retVal) = 0 Then
      sShowMessage "Enter The Rate"
      Spread_Set_Focus sprItem, srow, eItem.mRate
      GoTo ExitHere
    End If
     spSchedule.Sheet = srow
      If fraSch.Visible = True Then
        If spSchedule.DataRowCnt = 0 Then
          sShowMessage "Enter the Schedule Details"
          spSchedule.ActiveSheet = srow
          GoTo ExitHere
        End If
        For sCnt = 1 To spSchedule.DataRowCnt
          spSchedule.GetText eSch.mSchDate, sCnt, retVal
          If Len(retVal) = 0 Then
            sShowMessage "Enter the ScheduleDate"
            spSchedule.ActiveSheet = srow
            Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
            GoTo ExitHere
          End If
          spSchedule.GetText eSch.mSchDate, sCnt, retVal
          If IsDate(retVal) = 0 Then
            sShowMessage "invalid ScheduleDate"
            spSchedule.ActiveSheet = srow
            Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
            GoTo ExitHere
          End If
          spSchedule.GetText eSch.mSchDate, sCnt, retVal
          If sCnt > spSchedule.GetColItemData(eSch.mSchDate) Then
            If Format(retVal, "yyyyMMdd") < Format(CurDate, "yyyyMMdd") Then
              sShowMessage "invalid ScheduleDate"
              spSchedule.ActiveSheet = srow
              Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
              GoTo ExitHere
            End If
          End If
          spSchedule.GetText eSch.mSchQty, sCnt, retVal
          If Val(retVal) = 0 Then
            sShowMessage "Enter the ScheduleQty"
            spSchedule.ActiveSheet = srow
            Spread_Set_Focus spSchedule, sCnt, eSch.mSchQty
            GoTo ExitHere
          End If
        Next sCnt
        sprItem.GetText eItem.mSentQty, srow, retVal
        If Val(retVal) <> Spread_Cell_Total(spSchedule, eSch.mSchQty, 1, eSch.mSchQty, spSchedule.DataRowCnt) Then
          sShowMessage "Schedule TotalQty Should be OrderQty"
          spSchedule.ActiveSheet = srow
          Spread_Set_Focus spSchedule, 1, eSch.mSchQty
          GoTo ExitHere
        End If
     End If
    Next i
  
'  If spCalc.DataRowCnt = 0 Then
'    sShowMessage "Please check the Calculation"
'    txtFormula.SetFocus
'    Chkbefsave = False
'    GoTo ExitHere
'  End If
  
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function Group_Child(ByVal fRow As Long,ByVal  ParentCode As String) As String
  Dim i As Integer
  Dim retVal As Variant
  Dim Row As Long
  Dim srow As Long
  Dim ChildCode As String
  
  Try
  Me.MousePointer = vbHourglass
  
    srow = sprItem.SearchCol(eItem.mFirstRow, 0, sprItem.MaxRows, Val(fRow), SearchFlagsValue)
    ChildCode = ""
    Do While srow <> -1
      sprItem.GetText eItem.mrecdItem, srow, retVal
      ChildCode = ChildCode + retVal
      srow = sprItem.SearchCol(eItem.mFirstRow, srow, sprItem.MaxRows, Trim$(fRow), SearchFlagsValue)
    Loop
  
  Group_Child = ChildCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

Private Function ManyToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim sCnt As Long
  Dim retVal As Variant
  Dim sNo As Integer
  Dim sParentCode As String
  Dim sStageCode As String
  Dim sChildCode As String
  Dim DubLicate As Boolean
  Dim srow As Long
  Dim i As Long
  Dim j As Long
  
  Try
  DubLicate = False
  
  
    If fRow > sprItem.DataRowCnt Then     'already saved itemcode should not be changed
      'Fill the Child Items for the selected Item Structure Details in SentItem
        
      Sql = "Select Distinct b.SentItemCode, c.ItemDescription as ChildDesc,b.ItemGroupKey,b.PartNo,b.Rate,c.UOMDesc as UOM,c.UOMDecimal from SAl_JobWork_ItemMaster b"
      Sql = Sql & " Left Outer Join MATERIAL_MASTER c  On b.SentItemCode=c.ItemCode and b.BranchID=c.BranchID"
      Sql = Sql & " and b.CompID=c.CompID where CustomerCode=" & txtcustcd.Tag & " and RelationType='N-1'"
      Sql = Sql & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "' Order By b.ItemGroupKey"
      If FetchValue_New(True, db.ConnectionString, Sql, "111000", "Select Received Item", "Item Code", "Description", "Set No") = True Then
        For i = 1 To UBound(getvalue, 1)
          sCnt = sprItem.DataRowCnt + 1
          sprItem.SetText eItem.msentItem, sCnt, getvalue(i, 1)
          sprItem.SetText eItem.mSentDesc, sCnt, CheckNull(getvalue(i, 2))
          sprItem.SetText eItem.mSentUOM, sCnt, CheckNull(getvalue(i, 6))
          sprItem.SetText eItem.mRatePer, sCnt, 1
          sprItem.SetText eItem.mFirstRow, sCnt, Val(sCnt)
          
          If cmbScheduleType.Text <> "Open" Then
            Spread_Lock sprItem, False, eItem.mSentQty, sCnt, eItem.mAssRate, sCnt
          Else
            Spread_Lock sprItem, False, eItem.mRate, sCnt, eItem.mAssRate, sCnt
          End If
          
          Call Spread_Set_Decimal(sprItem, Val(CheckNull(getvalue(i, 7), 0)), eItem.mRatePer, sCnt)
        
          Sql = "Select Distinct b.ReceivedItemCode, c.ItemDescription as ChildDesc,b.ItemGroupKey, c.UOMDesc as UOM from SAl_JobWork_ItemMaster b"
          Sql = Sql & " Left Outer Join MATERIAL_MASTER c  On b.ReceivedItemCode=c.ItemCode and b.BranchID=c.BranchID"
          Sql = Sql & " and b.CompID=c.CompID where ItemGroupKey =" & getvalue(i, 3) & " "
          Sql = Sql & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "' Order By b.ItemGroupKey"
          Set rs = New AceADODB.Recordset
          If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
          sprItem.MaxRows = sprItem.DataRowCnt + rs.RecordCount
          sChildCode = ""
          For j = 1 To rs.RecordCount
            sChildCode = sChildCode + rs("ReceivedItemCode")
            rs.MoveNext
          Next j
            
          srow = sprItem.SearchCol(eItem.msentItem, 0, sprItem.MaxRows, Trim$(getvalue(i, 1)), SearchFlagsValue)
          Do While srow <> -1
            If Group_Child(srow, Trim$(getvalue(i, 1))) = sChildCode Then
              DubLicate = True
              sprItem.DeleteRows sCnt, 1
              Exit Do
            End If
            srow = sprItem.SearchCol(eItem.msentItem, srow, sprItem.MaxRows, Trim$(getvalue(i, 1)), SearchFlagsValue)
          Loop
          If DubLicate = False Then
            rs.MoveLast
rs.MoveFirst
            rCnt = sprItem.DataRowCnt
            sprItem.SetText eItem.mRate, rCnt, getvalue(i, 5)
            sprItem.SetText eItem.mPartNo, rCnt, getvalue(i, 4)
            For rCnt = sprItem.DataRowCnt To rs.RecordCount + sprItem.DataRowCnt - 1
              sprItem.SetText eItem.mFirstRow, rCnt, Val(sCnt)
              sprItem.SetText eItem.mrecdItem, rCnt, rs("ReceivedItemCode")
              sprItem.SetText eItem.mRecdDesc, rCnt, rs("ChildDesc")
              sprItem.SetText eItem.mRecdUOM, rCnt, rs("UOM")
              sprItem.SetText eItem.mRelation, rCnt, "N-1"
              sprItem.SetText eItem.mItemGroupKey, rCnt, getvalue(i, 3)
              rs.MoveNext
            Next rCnt
          Else
            sprItem.MaxRows = sprItem.DataRowCnt + 1
          End If
        Next i
      End If
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub SetSheet(ByVal fRow As Long)
' On Error Resume Next
    If spSchedule.SheetCount < fRow Then
      spSchedule.SheetCount = fRow
      Call CopySpreadStructure(spSchedule, 1, , fRow)
    End If
    spSchedule.Sheet = fRow
    spSchedule.ActiveSheet = fRow
End Sub

Private Function HelpDirectItem(ByVal fCol As Long,ByVal  fRow As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim sFirstRow As Long
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim tmp2 As Variant
  Dim tmp3 As Variant
  
  Try
  HelpDirectItem = False
  
  If optquotation Then GoTo ExitHere
  
  If fCol = eItem.mrecdItem Then
    'Show 1 to 1 Item for the Vendor
    Sql = " Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
    Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
    Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey,a.PartNo,a.AssessableRate from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode = b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.ReceivedItemCode = c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode = " & Val(txtcustcd.Tag) & " and a.RelationType='1-1' and a.isActive='1'"
    Sql = Sql & " and a.ValidFrom <= '" & Format(dporddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dporddate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "11011", "List of Item", "Recd Item", "Description", "Sent Item", "Description") = True Then
        If fRow <= sprItem.DataRowCnt Then
          rCnt = fRow
        Else
          sprItem.MaxRows = sprItem.DataRowCnt + 1
          rCnt = sprItem.DataRowCnt + 1
        End If
        If cmbScheduleType.Text <> "Open" Then
          SetSheet (rCnt)
        End If
        
        'Key Values First Set in Spread
        sprItem.SetText eItem.mItemGroupKey, rCnt, getvalue(10)
        sprItem.SetText eItem.mrecdItem, rCnt, getvalue(1)
        sprItem.SetText eItem.msentItem, rCnt, getvalue(4)
        
        If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.mrecdItem, eItem.msentItem) = False Then
          sprItem.SetText eItem.mFirstRow, rCnt, rCnt
          
          sprItem.SetText eItem.mRecdDesc, rCnt, getvalue(2)
          sprItem.SetText eItem.mRecdUOM, rCnt, getvalue(3)
          
          sprItem.SetText eItem.mSentDesc, rCnt, getvalue(5)
          sprItem.SetText eItem.mSentUOM, rCnt, getvalue(6)
          sprItem.SetText eItem.mPartNo, rCnt, getvalue(11)
          Call Spread_Set_Decimal(sprItem, Val(getvalue(7)), eItem.mSentQty, rCnt)
          If Trim$(cmbScheduleType) = "Open" Then
            Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          Else
            Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          End If
          sprItem.SetText eItem.mRatePer, rCnt, getvalue(8)
          sprItem.SetText eItem.mRate, rCnt, getvalue(9)
          sprItem.SetText eItem.mAssRate, rCnt, getvalue(12)
          
          sprItem.SetText eItem.mRelation, rCnt, "1-1"
          Call Spread_Lock(sprItem, False, eItem.mRatePer, rCnt, eItem.mRate, rCnt)
          sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
          sprItem.GetText eItem.mRate, rCnt, tmp1  'Rate
               sprItem.GetText eItem.mRatePer, rCnt, tmp3   'Rateper
          If Val(tmp3) > 0 Then
            sprItem.SetText eItem.mGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
            sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
            sprItem.GetText eItem.mAssRate, rCnt, tmp2 'ARate
            sprItem.SetText eItem.mAssGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp2), "0.00")
          End If
        Else
          spSchedule.DeleteSheets fRow, 1
        End If
        sprItem.MaxRows = sprItem.DataRowCnt + 1    Else
      GoTo ExitHere
    End If
  Else
    'Show N to 1 Item for the Vendor
    
    Sql = "Select distinct a.ReceivedItemCode, c.ItemDescription as RecdDesc,a.SentItemCode, b.ItemDescription as SentDesc,a.ItemGroupKey,a.PartNo"
    Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
    Sql = Sql & " and (a.ValidFrom <= '" & Format(dporddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dporddate, "yyyy-MM-dd") & "')"
    Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    Sql = Sql & " Order by a.ReceivedItemCode, a.SentItemCode, a.ItemGroupKey"
    If FetchValue_New(False, db.ConnectionString, Sql, "11110", "List of Item", "Recd Item", "Description", "Sent Item", "Description") = True Then
    
      Sql = "Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
      Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
      Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey"
      Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
      Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
      Sql = Sql & " Left Outer Join Material_Master c"
      Sql = Sql & " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
      Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
      Sql = Sql & " and a.ItemGroupKey=" & getvalue(5)
      Sql = Sql & " and a.ValidFrom <= '" & Format(dporddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dporddate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If rs.EOF Then
 GoTo ExitHere
End If
        rs.MoveLast
rs.MoveFirst
        If fRow <= sprItem.DataRowCnt Then
          sFirstRow = fRow
          sprItem.MaxRows = rs.RecordCount + 1
        Else
          sFirstRow = sprItem.DataRowCnt + 1
          sprItem.MaxRows = sprItem.DataRowCnt + rs.RecordCount + 1
        End If
        For rCnt = sFirstRow To sprItem.MaxRows - 1
          'Key Values First Set
          sprItem.SetText eItem.mItemGroupKey, rCnt, rs("ItemGroupKey")
          sprItem.SetText eItem.mrecdItem, rCnt, rs("ReceivedItemCode")
          If sFirstRow = rCnt Then
            sprItem.SetText eItem.msentItem, rCnt, rs("SentItemCode")
            sprItem.SetText eItem.mSentDesc, rCnt, rs("SentDesc")
            sprItem.SetText eItem.mSentUOM, rCnt, rs("SentUOM")
          End If
          
          If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.msentItem, eItem.mrecdItem) = True Then
            sprItem.MaxRows = sprItem.DataRowCnt + 1
            Exit For
          End If
          
          sprItem.SetText eItem.mRecdDesc, rCnt, rs("RecdDesc")
          sprItem.SetText eItem.mRecdUOM, rCnt, rs("RecdUOM")

          sprItem.SetText eItem.mFirstRow, rCnt, sFirstRow
          
          If sFirstRow = rCnt Then
            If cmbScheduleType.Text <> "Open" Then
              SetSheet (rCnt)
            End If
            sprItem.SetText eItem.mPartNo, rCnt, getvalue(6)
            sprItem.SetText eItem.mRatePer, rCnt, rs("RatePer")
            sprItem.SetText eItem.mRate, rCnt, rs("Rate")
            sprItem.SetText eItem.mAssRate, rCnt, rs("Rate")
            Call Spread_Set_Decimal(sprItem, Val(rs("SentNoOfDec")), eItem.mSentQty, rCnt)
            If Trim$(cmbScheduleType) = "Open" Then
              Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            Else
              Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            End If
            sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
            sprItem.GetText eItem.mRate, rCnt, tmp1  'Rate
            sprItem.GetText eItem.mRatePer, rCnt, tmp3   'Rateper
            If Val(tmp3) > 0 Then
              sprItem.SetText eItem.mGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
              sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
              sprItem.GetText eItem.mAssRate, rCnt, tmp1 'ARate
              sprItem.SetText eItem.mAssGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
            End If
          Else
            Call Spread_Lock(sprItem, True, eItem.mPartNo, rCnt, eItem.mSentQty, rCnt)
          End If
          sprItem.SetText eItem.mRelation, rCnt, "N-1"
          rs.MoveNext
        Next rCnt
    Else
      GoTo ExitHere
    End If
  End If
  
  OrderCalulation  'to calculate the Jobwork Order Value
  
  HelpDirectItem = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub OrderCalulation()
'  Dim Tot As Double
'  Dim atot As Double
'  Dim i As Long
'  Dim Sql As String
'  Dim snap As ADODB.Recordset
'
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  If sprItem.DataRowCnt = 0 Then: GoTo ExitHere
'
'  Tot = 0: atot = 0
'  For i = 1 To sprItem.DataRowCnt
'    Tot = Tot + Val(SpVal(sprItem, i, eItem.mGoodsValue))
'    atot = atot + Val(SpVal(sprItem, i, eItem.mAssGoodsValue))
'  Next i
'
'  If txtFormula <> "" Then
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
'  End If
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

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

  txtordnum = GetDocNo("SAL_JOBWORKORDER_HDR", "OrderType", "OrderNo", Trim$(cmbordtype), "J")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmbScheduleType_Click()
  Dim i As Long
  
  Try
  
  For i = 1 To sprItem.DataRowCnt
    If cmbScheduleType = "Open" Then
      Call Spread_Lock(sprItem, True, eItem.mSentQty, i, eItem.mSentQty, i)
      sprItem.SetText eItem.mSentQty, i, ""
    Else
      If Trim(SpVal(sprItem, i, eItem.mSentUOM)) <> "" Then
        Call Spread_Lock(sprItem, False, eItem.mSentQty, i, eItem.mSentQty, i)
'        sprItem.SetText eItem.mSentQty, i, ""
      Else
        Call Spread_Lock(sprItem, True, eItem.mSentUOM, i, eItem.mSentQty)
      End If
    End If
  Next i
  If cmbScheduleType = "Open" Then
    fraSch.Visible = False
  Else
    fraSch.Visible = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub CmdClear_Click()
  Dim Sql As String
 
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  CurrDate Me
  Call LoadDocType(mJWOrder)
  fraOrder.Enabled = True
  fraSource.Enabled = True
  fraQuotation.Enabled = True
  fraTerms.Enabled = False
  sprItem.Enabled = False
  cmdsave.Enabled = False
  cmbScheduleType.Enabled = True
  cmbScheduleType.ListIndex = 0
  sprItem.MaxRows = 1
  optdirect.Value = True
  fraQuotation.Visible = False
  txtcustcd.Enabled = True
  txtcustcd.SetFocus
  CurrDate Me
'  Loadordtype
'  Sql = "delete from COM_FORMULA_TMP where SKey = '" + SKey + "'"
'  Sql = Sql + " and Userid ='" & gUserID & "' and Branchid='" & gBranchID & "' "
'  Sql = Sql + " and Compid='" & gCompID & "'"
'  db.Execute Sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim ordnorec As AceADODB.Recordset
  Dim i As Long
  Dim sCnt As Long
  Dim Sql As String
  Dim tmp As Variant
  Dim sRelation As String
  Dim sFirstRow As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  Set cmd = New AceADODB.Command
  Set ordnorec = New AceADODB.Recordset
  
  If Chkbefsave = False Then GoTo ExitHere
  
  CurDate = ServerDate
  
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
  cmd.CommandText = "SAL_INS_JOBWORKORDER_HDR"
  cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
  cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
  cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("OrderSource", adChar, adParamInput, 1, IIf(optdirect, "D", "Q"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("ReferanceNo", adVarChar, adParamInput, 20, Trim$(txtRefNo))
  cmd.Parameters.Append cmd.CreateParameter("ReferanceDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("OrderingType", adVarChar, adParamInput, 10, Trim$(cmbScheduleType.Text))
  cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(dpvalidfrom, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(dpvalidto, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtremarks))
  cmd.Parameters.Append cmd.CreateParameter("FormCode", adInteger, adParamInput, , Val(txtForm.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PackingCode", adInteger, adParamInput, , Val(txtPack.Tag))
  cmd.Parameters.Append cmd.CreateParameter("DeliveryCode", adInteger, adParamInput, , Val(txtDelIns.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PaymentCode", adInteger, adParamInput, , Val(txtTerms.Tag))
  cmd.Parameters.Append cmd.CreateParameter("InsuranceCode", adInteger, adParamInput, , Val(txtInsurance.Tag))
  cmd.Parameters.Append cmd.CreateParameter("ModeCode", adInteger, adParamInput, , Val(txtMode.Tag))
  cmd.Parameters.Append cmd.CreateParameter("FreightCode", adInteger, adParamInput, , Val(txtFreight.Tag))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , Trim$(txtGrandTotal))
  cmd.Parameters.Append cmd.CreateParameter("PostingAc", adInteger, adParamInput, , Val(txtPostingAc.Tag))
  cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "N")
  cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))

  Call AddCommonCmdParameters(cmd)
  cmd.Execute

  If Save_Jobwork_ItemMaster = False Then
    GoTo ExitHere
  End If
  
    For i = 1 To sprItem.DataRowCnt
      sprItem.GetText eItem.mFirstRow, i, tmp
      sFirstRow = Val(tmp)
      commClear cmd
      cmd.CommandText = "SAL_INS_JOBWORKORDER_DTL"
      cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
      cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
      cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
      sprItem.GetText eItem.mItemGroupKey, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      sprItem.GetText eItem.mRelation, i, tmp
      sRelation = Trim$(tmp)
      If sRelation = "1-1" Or sRelation = "N-1" Then
        sprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        sprItem.GetText eItem.msentItem, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      Else
        sprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        sprItem.GetText eItem.msentItem, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      End If
      cmd.Parameters.Append cmd.CreateParameter("ScheduleType", adChar, adParamInput, 1, Left(cmbScheduleType.Text, 1))
      If cmbScheduleType.Text = "Open" Then
        cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , "0")
      Else
        sprItem.GetText eItem.mSentQty, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
      End If
      sprItem.GetText eItem.mRatePer, sFirstRow, tmp
      cmd.Parameters.Append cmd.CreateParameter("RatePer", adDouble, adParamInput, , Val(tmp))
      sprItem.GetText eItem.mRate, sFirstRow, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(tmp))
'      cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adVarChar, adParamInput, 5, "0")
      cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "N")
      cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
      Call AddCommonCmdParameters(cmd, , True)
      cmd.Execute
      
      If optquotation Then 'Updating Quotation table
        commClear cmd
        cmd.CommandText = "SAL_UP_JOBWORKQUOTATION_DTL"
        cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuotype))
        cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
        cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(txtQuoDate, "yyyy-MM-dd"))
        sprItem.GetText eItem.mRelation, i, tmp
        sRelation = Trim$(tmp)
        If sRelation = "1-1" Or sRelation = "N-1" Then
          sprItem.GetText eItem.mrecdItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          sprItem.GetText eItem.msentItem, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        Else
          sprItem.GetText eItem.mrecdItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          sprItem.GetText eItem.msentItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        End If
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(dporddate, "yyyy-MM-dd"))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
      If cmbScheduleType.Text <> "Open" And i = sFirstRow Then
        With spSchedule
          sprItem.Sheet = sFirstRow
          For sCnt = 1 To sprItem.DataRowCnt
            cmd.CommandText = "Sal_Ins_JobWorkOrder_Schedule"
            commClear cmd
            cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype.Text))
            cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
            cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
            sprItem.GetText eItem.mItemGroupKey, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
            sprItem.GetText eItem.msentItem, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
            sprItem.GetText eItem.mPartNo, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
            sprItem.GetText eSch.mSchDate, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(tmp, "yyyy-MM-dd"))
            sprItem.GetText eSch.mSchQty, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(tmp))
            sprItem.GetText eSch.mLocation, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(tmp))
            Call AddCommonCmdParameters(cmd)
            cmd.Execute
          Next sCnt
          spSchedule.Sheet = sFirstRow
          For sCnt = 1 To spSchedule.DataRowCnt
            cmd.CommandText = "Sal_Ins_JobWorkOrder_Schedule"
            commClear cmd
            cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(cmbordtype.Text))
            cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
            cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
            sprItem.GetText eItem.mItemGroupKey, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
            sprItem.GetText eItem.msentItem, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
            sprItem.GetText eItem.mPartNo, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
            spSchedule.GetText eSch.mSchDate, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(tmp, "yyyy-MM-dd"))
            spSchedule.GetText eSch.mSchQty, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(tmp))
            spSchedule.GetText eSch.mLocation, sCnt, tmp
            cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(tmp))
            Call AddCommonCmdParameters(cmd)
            cmd.Execute
          Next sCnt
      End If
    Next i
  End With
  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If Trim$(cmbScheduleType.Text) = "Specific" Then
    If SaveCalcDetails = False Then
      GoTo ExitHere
    End If
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Jobwork Order Number :" & Trim$(cmbordtype) & Trim$(txtordnum) & "  Is Generated"
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
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,"
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
    Sql = Sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, "
    Sql = Sql & " a.IsTypePercentage, a.IsValueEditable, a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
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
  cmbScheduleType.ListIndex = 0
  SKey = Format(Now, "yyyyMMddHHmmss") & gUserID
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
  cmbScheduleType.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optquotation_Click()
  Try
  
  fraQuotation.Visible = True
  txtcustcd.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprItem_DblClick(sender as object, e as DblClickEventArgs)
  
  Try
  Me.MousePointer = vbHourglass
  
  If Col = eItem.mrecdItem Then    'Select Sent ItemCode means One To One / One To Many
    Call HelpDirectItem(Col, Row)
  ElseIf Col = eItem.msentItem Then 'Select Recd ItemCode means Many to One
    Call ManyToOneItemHelp(Row)
  End If
  
'  If Col = eItem.mrecdItem Or Col = eItem.msentItem Then
'    If Row = sprItem.MaxRows Then
'      Call HelpDirectItem(Col, Row)
'    End If
'  End If
  cmbScheduleType.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 Then
      Select Case Col
        Case eItem.mSentQty, eItem.mRate, eItem.mAssRate
          sprItem.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
'            Call StartCalcFormula(IIf(Trim$(txtGrandTotal.Text) = "", True, False))
          End If
          
          Select Case Col
            Case eItem.mSentQty
              If Val(retVal) > 0 Then
                If optquotation = True Then
                  sprItem.SetActiveCell eItem.mrecdItem, Row + 1
                Else
                  sprItem.SetActiveCell eItem.mRate, Row
                End If
              End If
            Case eItem.mRate
              If Val(retVal) > 0 Then
                sprItem.SetActiveCell eItem.mAssRate, Row
              End If
            Case eItem.mAssRate
              If Val(retVal) > 0 Then
                sprItem.SetActiveCell eItem.mrecdItem, Row + 1
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

Private Sub sprItem_KeyDown(sender as object, e as KeyDownEventArgs)
  
  Try
  Me.MousePointer = vbHourglass
  
    Select Case KeyCode
      Case vbKeyF2
        sprItem_DblClick sprItem.ActiveCol, sprItem.ActiveRow
      Case 46
        If sprItem.ActiveCol = 1 Then
          Spread_Clear_Row sprItem, 1, sprItem.MaxCols, sprItem.ActiveRow, sprItem.ActiveRow
          sprItem.DeleteRows sprItem.ActiveRow, 1
          sprItem.MaxRows = sprItem.DataRowCnt + 1
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retVal As Variant
  Dim tmp As Variant
  
  Try
  
  sprItem.GetText eItem.mFirstRow, NewRow, retVal
  If cmbScheduleType.Text <> "Open" Then
    If Row <> NewRow And NewRow >= 1 And Val(retVal) = NewRow Then
      SetSheet (NewRow)
      sprItem.GetText eItem.msentItem, NewRow, retVal
      txtItemCode.Text = Trim$(retVal)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spSchedule_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retVal As Variant
  
  Try
  
    If Col = eSch.mSchDate Then
      Spread_Clear_Duplicate spSchedule, Row, eSch.mSchDate
      spSchedule.MaxRows = spSchedule.DataRowCnt + 1
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

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
  Sql = Sql + " and b.Compid=a.Compid) PartyName, a.CustomerCode from SAL_Jobwork_ItemMaster a "
  Sql = Sql + " where a.Branchid = '" & gBranchID & "' and a.compid='" & gCompID & "'"
  Sql = Sql & " order by a.CustomerCode"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer") = True Then
    txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
    fraOrder.Enabled = False
    fraSource.Enabled = False
    fraQuotation.Enabled = False
    fraTerms.Enabled = True
    sprItem.Enabled = True
    cmdsave.Enabled = True
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
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='J'"
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
'    txtordnum = GetNo("SAL_JOBWORKORDER_HDR", "OrderType", "OrderNo", Trim$(cmbordtype), "J", "SAL")
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
  
  If sprItem.DataRowCnt <= 0 Then
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

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'FreightCode' "
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

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Insurance' "
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
  Dim snap As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  Dim tot As Double
  Dim tot1 As Double
  Dim rCnt As Long
  Dim PrevItemGroup As Integer
  Dim ValueRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  Set snap = New AceADODB.Recordset
  Set snap1 = New AceADODB.Recordset
  
  Sql = " select distinct a.QuotationType,a.QuotationNo,a.QuotationDate,Revisionno,isnull(case when a.CustomerCode=0 then (select PartyCode from Fas_Party_MAster where PartyName=a.CustomerName and "
  Sql = Sql + " Branchid=a.Branchid and Compid=a.Compid) else a.CustomerCode end,0) CustomerCode,a.CustomerName,"
  Sql = Sql + " ValidTo,a.FormulaCode, (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where"
  Sql = Sql + " b7.ModuleCode = '" & gModuleCode & "' and b7.FrmId=a.FormulaCode and b7.Branchid=a.Branchid and b7.Compid=a.Compid) FrmLName "
  Sql = Sql + " from JobworkQuotation a where OrderNo = 0 and a.compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "111101000", "List of Quotation Nos.", "Quotation Type", "Quotation No", "Quotation Date", "Revision no", "Customer") = True Then
    If Val(getvalue(4)) = 0 Then
      If MsgBox("Please create the Entry in the Customer Master for the Customer " & getvalue(5) & vbLf & "Do you Create Now?", vbYesNo, App.Title) = vbYes Then
        frmPartyMaster.Show
        frmPartyMaster.tbPartyName = getvalue(5)
      End If
      GoTo ExitHere
    End If
    txtQuotype = getvalue(1)
    txtQuoNo = getvalue(2)
    txtQuoDate = getvalue(3)
    dpvalidfrom.MinDate = txtQuoDate
    dpvalidto.MinDate = txtQuoDate
    txtcustcd.Tag = getvalue(5)
    txtcustcd = getvalue(6)
    txtQuoDate.Tag = getvalue(7)
    txtFormula.Tag = getvalue(8)
    txtFormula.Text = getvalue(9)
    
    tot = 0
tot1 = 0
    Sql = " Select a.ItemGroupKey, a.SentItemCode, b.ItemDescription as SentDesc, b.UOMDesc as SentUOM,"
    Sql = Sql + " b.UOMDecimal as NoOfDec, a.ReceivedItemCode,c.ItemDescription as RecdDesc, c.UOMDesc as RecdUOM, "
    Sql = Sql + " a.RatePer,a.rate,a.AssRate,a.SendQty,a.RelationType,"
    Sql = Sql + " isnull((select d.PartNo from SAL_Jobwork_ItemMaster d where d.ItemGroupKey = a.ItemGroupKey and d.Branchid = a.Branchid"
    Sql = Sql + " and d.Compid = a.Compid and d.ReceivedItemCode=a.ReceivedItemCode and d.SentItemCode=a.SentItemCode),'')PartNo"
    Sql = Sql + " from SAL_VW_JOBWORKQUOTATION_DTL a Left Outer Join Material_Master b On a.SentItemCode=b.ItemCode"
    Sql = Sql + " and a.BranchID=b.BranchID and a.CompID=b.CompID Left Outer Join Material_Master c"
    Sql = Sql + " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID where"
    Sql = Sql + " a.QuotationType='" + Trim(txtQuotype) + "' and a.QuotationNo = " & Trim(txtQuoNo) & ""
    Sql = Sql + " and a.QuotationDate = convert(datetime,'" + Format(Trim(txtQuoDate), "yyyy-MM-dd") + "',111)"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
      If Not snap.EOF Then
        sprItem.Enabled = True
        snap.MoveLast
snap.MoveFirst
        For rCnt = 1 To snap.RecordCount
          If IsNull(snap("RecdDesc")) = True Then
            sShowMessage "Please Create the ItemCode (" & snap("ReceivedItemCode") & ") in the Item Master"
            Call CmdClear_Click
            GoTo ExitHere
          End If
          If IsNull(snap("SentDesc")) = True Then
            sShowMessage "Please Create the ItemCode (" & snap("SentItemCode") & ") in the Item Master"
            Call CmdClear_Click
            GoTo ExitHere
          End If
          snap.MoveNext
        Next rCnt
        PrevItemGroup = -1
        snap.MoveLast
snap.MoveFirst
        sprItem.MaxRows = snap.RecordCount
        For rCnt = 1 To snap.RecordCount
          If PrevItemGroup <> snap("ItemGroupKey") Then
            PrevItemGroup = snap("ItemGroupKey")
            ValueRow = rCnt
          End If
          sprItem.SetText eItem.mFirstRow, rCnt, ValueRow
          
          sprItem.SetText eItem.mrecdItem, rCnt, snap("ReceivedItemCode")
          sprItem.SetText eItem.mRecdDesc, rCnt, CheckNull(snap("RecdDesc"))
          sprItem.SetText eItem.mRecdUOM, rCnt, CheckNull(snap("RecdUOM"))
          
          sprItem.SetText eItem.mRelation, rCnt, CheckNull(snap("RelationType"))
          sprItem.SetText eItem.mItemGroupKey, rCnt, CheckNull(snap("ItemGroupKey"))
          
          If snap("RelationType") = "1-1" Or (snap("RelationType") = "N-1" And ValueRow = rCnt) Then
            sprItem.SetText eItem.msentItem, rCnt, snap("SentItemCode")
            sprItem.SetText eItem.mSentDesc, rCnt, CheckNull(snap("SentDesc"))
            sprItem.SetText eItem.mSentUOM, rCnt, CheckNull(snap("SentUOM"))
            sprItem.SetText eItem.mPartNo, rCnt, CheckNull(snap("PartNo"))
            Call Spread_Set_Decimal(sprItem, CheckNull(snap("NoOfDec"), 0), eItem.mSentQty, rCnt)
            sprItem.SetText eItem.mRatePer, rCnt, CheckNull(snap("RatePer"))
            sprItem.SetText eItem.mRate, rCnt, CheckNull(snap("Rate"))
            sprItem.SetText eItem.mAssRate, rCnt, CheckNull(snap("AssRate"))
            If CheckNull(snap("SendQty")) > 0 Then
              cmbScheduleType.Text = "Specific"
              cmbScheduleType.Enabled = False
              Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            Else
              cmbScheduleType.Enabled = True
              cmbScheduleType = "Open"
              Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            End If
            sprItem.SetText eItem.mSentQty, rCnt, CheckNull(snap("SendQty"))
            sprItem.SetText eItem.mGoodsValue, rCnt, (CheckNull(snap("SendQty")) * CheckNull(snap("Rate")))
            tot = tot + (CheckNull(snap("SendQty")) * CheckNull(snap("Rate")))
            sprItem.SetText eItem.mAssGoodsValue, rCnt, (CheckNull(snap("SendQty")) * CheckNull(snap("AssRate")))
            tot1 = tot1 + (CheckNull(snap("SendQty")) * CheckNull(snap("AssRate")))
            Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mAssRate, rCnt)
          Else
            Call Spread_Lock(sprItem, True, eItem.mSentUOM, rCnt, eItem.mItemGroupKey, rCnt)
          End If
          Call SetSprGoodsValue(rCnt)
          snap.MoveNext
        Next rCnt
        sprItem.Tag = snap.RecordCount
      Else
        sprItem.MaxRows = 1
        sprItem.Tag = 0
      End If
    Call SetTotalGoodsValue
    If txtFormula <> "" Then
      Call CalcButtonStatus(Me, True)
    End If
    
    fraOrder.Enabled = False
    fraSource.Enabled = False
    fraQuotation.Enabled = False
    fraTerms.Enabled = True
    sprItem.Enabled = True
    cmdsave.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
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

Private Function Save_Jobwork_ItemMaster() As Boolean
  Dim i As Long
  Dim CurDate As Date
  Dim srow As Long
  Dim ret As Integer
  Dim sCnt As Long
  Dim retVal As Variant
  Dim Sql As String
  Dim Group As Long
  Dim Relation As String
  Dim rs As AceADODB.Recordset
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
  Try
  
  Save_Jobwork_ItemMaster = False
    Group = -1
     For i = 1 To sprItem.DataRowCnt
      sprItem.Row = i
sprItem.Col = eItem.mFirstRow
srow = Val(sprItem.Text)
      sprItem.GetText eItem.mItemGroupKey, srow, retVal
      If Val(retVal) <> Group Then
        Sql = "select Case when exists(select * from SAL_Jobwork_ItemMaster where CusTomerCode=" & txtcustcd.Tag & " and ItemGroupKey =" & Val(retVal) & " ) then"
        Sql = Sql & " 'Y' else 'N' end"
        Set rs = New AceADODB.Recordset
        If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      End If
        If rs(0) = "N" Then
          Sql = "Execute SAL_INS_JOBWORK_ITEMMASTER"
          sprItem.GetText eItem.mItemGroupKey, srow, retVal
          Sql = Sql & " @ItemGroupKey=" & Val(retVal) & ","
          Sql = Sql & " @CustomerCode=" & txtcustcd.Tag & ","
          sprItem.GetText eItem.mRelation, i, retVal
          Relation = Trim$(retVal)
          sprItem.GetText eItem.mrecdItem, i, retVal
          Sql = Sql & " @ReceivedItemCode='" & Trim$(retVal) & "',"
          
          If Relation = "1-1" Then
            sprItem.GetText eItem.msentItem, i, retVal
          Else
            sprItem.GetText eItem.msentItem, srow, retVal
          End If
          Sql = Sql & " @SentItemCode='" & Trim$(retVal) & "',"
          
          Sql = Sql & " @RelationType='" & Relation & "',"
          
          sprItem.GetText eItem.mPartNo, srow, retVal
          Sql = Sql & " @PartNo='" & Trim$(retVal) & "',"
          sprItem.GetText eItem.mRatePer, srow, retVal
          Sql = Sql & " @RatePer=" & IIf(Val(retVal) = 0, 1, Val(retVal)) & ","
          sprItem.GetText eItem.mRate, srow, retVal
          Sql = Sql & " @Rate=" & Val(retVal) & ","
          Sql = Sql & " @AssessableRate=" & Val(retVal) & ","
          Sql = Sql & " @ValidFrom='" & Format(dpvalidfrom, "yyyy-MM-dd") & "',"
          Sql = Sql & " @ValidTo='" & Format(dpvalidto, "yyyy-MM-dd") & "',"
          Sql = Sql & " @QcCheck='0',"
          Sql = Sql & " @FormulaCode=" & Val(txtFormula.Tag) & ","
          Sql = Sql & " @isActive='1',"
          Sql = Sql & " @UserID='" & gUserID & "',"
          Sql = Sql & " @BranchID='" & gBranchID & "',"
          Sql = Sql & " @CompID='" & gCompID & "'"
          If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        End If
      sprItem.GetText eItem.mItemGroupKey, srow, retVal
      Group = Val(retVal)
    Next i
 
  Save_Jobwork_ItemMaster = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retVal As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(sprItem.DataRowCnt - 1)
      ReDim SCValList(sprItem.DataRowCnt * 2 - 1)
      
      For rCnt = 1 To sprItem.DataRowCnt
        'Item List
        sprItem.GetText eItem.mrecdItem, rCnt, retVal
        Item = Trim$(retVal)
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        sprItem.GetText eItem.mSentQty, rCnt, retVal
        Qty = Val(retVal)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        sprItem.GetText eItem.mRate, rCnt, retVal
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        
        sprItem.GetText eItem.mAssRate, rCnt, retVal
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
    
  Try
  
    sprItem.GetText eItem.mSentQty, fRow, retVal(1)
    retVal(2) = 1
    
    sprItem.GetText eItem.mRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    sprItem.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    sprItem.GetText eItem.mAssRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    sprItem.SetText eItem.mAssGoodsValue, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(sprItem, eItem.mGoodsValue), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(sprItem, eItem.mAssGoodsValue), "0.00")
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

Private Sub Design_spSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spSchedule


spSchedule.LeaveCell +=  new EventHandler(spSchedule_LeaveCell)

End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.DblClick +=  new EventHandler(sprItem_DblClick)
sprItem.EditMode +=  new EventHandler(sprItem_EditMode)
sprItem.KeyDown +=  new EventHandler(sprItem_KeyDown)
sprItem.LeaveCell +=  new EventHandler(sprItem_LeaveCell)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_spSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrder.frm", spSchedule, "")

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrder.frm", sprItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrder.frm", sprCalc, "")

End Sub
