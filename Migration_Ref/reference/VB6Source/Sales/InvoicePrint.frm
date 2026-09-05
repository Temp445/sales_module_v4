VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptInvoicePrint 
   Caption         =   "Invoice Print"
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPreRequest 
      Height          =   750
      Left            =   195
      TabIndex        =   32
      Top             =   7650
      Visible         =   0   'False
      Width           =   2010
      Begin MSComCtl2.DTPicker dtpAsOnDate 
         Height          =   315
         Left            =   90
         TabIndex        =   33
         Top             =   345
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   108724227
         CurrentDate     =   39743
      End
      Begin VB.Label lblAsOnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AsOnDate"
         Height          =   210
         Left            =   90
         TabIndex        =   34
         Top             =   165
         Width           =   750
      End
   End
   Begin VB.Frame fraView 
      Height          =   990
      Left            =   60
      TabIndex        =   20
      Top             =   -75
      Width           =   1155
      Begin VB.CommandButton cmdexit1 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         Height          =   360
         Left            =   9855
         TabIndex        =   22
         Top             =   7680
         Width           =   1695
      End
Begin AceSpread sprView
         Height          =   7530
         Left            =   60
         TabIndex        =   21
         Top             =   150
         Width           =   11475
         _Version        =   458752
         _ExtentX        =   20241
         _ExtentY        =   13282
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         MaxCols         =   9
         MaxRows         =   50
      End
   End
   Begin VB.Frame fraPlaceHolder 
      BorderStyle     =   0  'None
      Height          =   7620
      Left            =   2100
      TabIndex        =   23
      Top             =   810
      Width           =   11775
      Begin VB.Frame fraPackingSlipPrint 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4080
         Left            =   8220
         TabIndex        =   30
         Top             =   2445
         Width           =   2850
         Begin VB.TextBox txtSavedBoxQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   900
            Width           =   1140
         End
         Begin VB.CommandButton cmdPackingSlipPrint 
            BackColor       =   &H80000003&
            Caption         =   "Packing &Slip Label Print"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3540
            Width           =   2700
         End
         Begin VB.TextBox txtBoxQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            TabIndex        =   11
            Top             =   900
            Width           =   1110
         End
         Begin VB.CheckBox chkPackingSlipPrint 
            Appearance      =   0  'Flat
            Caption         =   "Packing Slip Print"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   9
            Top             =   225
            Width           =   2040
         End
         Begin VB.Frame fraLabel 
            Caption         =   "Select Label"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   90
            TabIndex        =   36
            Top             =   1275
            Width           =   2685
            Begin VB.CheckBox chkAll 
               Appearance      =   0  'Flat
               Caption         =   "All"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   150
               TabIndex        =   42
               Top             =   300
               Width           =   660
            End
            Begin VB.CheckBox chkPosition 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   1500
               TabIndex        =   41
               Top             =   1335
               Width           =   240
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Proceed"
               Height          =   390
               Left            =   6705
               TabIndex        =   40
               Top             =   795
               Width           =   960
            End
            Begin VB.CheckBox chkPosition 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1500
               TabIndex        =   39
               Top             =   945
               Width           =   330
            End
            Begin VB.CheckBox chkPosition 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   915
               TabIndex        =   38
               Top             =   945
               Width           =   195
            End
            Begin VB.CheckBox chkPosition 
               Appearance      =   0  'Flat
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   915
               TabIndex        =   37
               Top             =   1335
               Width           =   240
            End
            Begin VB.Label lblPosition 
               Caption         =   "Bottom Right"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   540
               Index           =   3
               Left            =   1500
               TabIndex        =   46
               Top             =   1620
               Width           =   645
            End
            Begin VB.Label lblPosition 
               Caption         =   "Top Right"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   345
               Index           =   1
               Left            =   1455
               TabIndex        =   45
               Top             =   660
               Width           =   975
            End
            Begin VB.Label lblPosition 
               Caption         =   "Bottom Left"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   540
               Index           =   2
               Left            =   435
               TabIndex        =   44
               Top             =   1620
               Width           =   720
            End
            Begin VB.Label lblPosition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Top Left"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   240
               Index           =   0
               Left            =   330
               TabIndex        =   43
               Top             =   660
               Width           =   780
            End
         End
         Begin VB.Label lblBoxQty 
            AutoSize        =   -1  'True
            Caption         =   "Box Qty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1410
            TabIndex        =   35
            Top             =   630
            Width           =   750
         End
         Begin VB.Label lblSavedBoxQty 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saved Qty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   90
            TabIndex        =   31
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   450
         Left            =   5505
         TabIndex        =   19
         Top             =   6075
         Width           =   1500
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   450
         Left            =   4035
         TabIndex        =   18
         Top             =   6075
         Width           =   1485
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   450
         Left            =   2565
         TabIndex        =   17
         Top             =   6075
         Width           =   1485
      End
      Begin VB.Frame fraSelect 
         Caption         =   "Select"
         Height          =   1365
         Left            =   135
         TabIndex        =   25
         Top             =   135
         Width           =   7845
         Begin VB.CheckBox chkInvoiceNo 
            Appearance      =   0  'Flat
            Caption         =   "By Invoice No"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1665
            TabIndex        =   4
            Top             =   840
            Width           =   1365
         End
         Begin VB.CheckBox chkcancel 
            Appearance      =   0  'Flat
            Caption         =   "Cancel invoice"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   135
            TabIndex        =   0
            Top             =   360
            Width           =   1425
         End
         Begin VB.CheckBox chkInvoiceDate 
            Appearance      =   0  'Flat
            Caption         =   "By Invoice Date"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1665
            TabIndex        =   1
            Top             =   360
            Width           =   1440
         End
         Begin VB.CommandButton cmdProceed 
            Caption         =   "&Proceed"
            Height          =   390
            Left            =   6705
            TabIndex        =   7
            Top             =   795
            Width           =   960
         End
         Begin VB.Frame fraInvoiceDate 
            Height          =   585
            Left            =   3270
            TabIndex        =   28
            Top             =   165
            Visible         =   0   'False
            Width           =   3330
            Begin MSComCtl2.DTPicker dtpInvoiceTo 
               Height          =   330
               Left            =   1905
               TabIndex        =   3
               Top             =   180
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   109379587
               CurrentDate     =   38257
            End
            Begin MSComCtl2.DTPicker dtpInvoiceFrom 
               Height          =   330
               Left            =   75
               TabIndex        =   2
               Top             =   180
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   109379587
               CurrentDate     =   38257
            End
            Begin VB.Label lblTo 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1500
               TabIndex        =   29
               Top             =   210
               Width           =   315
            End
         End
         Begin VB.Frame fraInvoiceNo 
            Height          =   585
            Left            =   3270
            TabIndex        =   26
            Top             =   645
            Visible         =   0   'False
            Width           =   3330
            Begin VB.TextBox txtInvoiceFrom 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   90
               TabIndex        =   5
               Top             =   180
               Width           =   1335
            End
            Begin VB.TextBox txtInvoiceTo 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1905
               TabIndex        =   6
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label lblTo 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1500
               TabIndex        =   27
               Top             =   195
               Width           =   315
            End
         End
      End
      Begin VB.Frame fraPack 
         Height          =   525
         Left            =   1860
         TabIndex        =   24
         Top             =   5325
         Visible         =   0   'False
         Width           =   4425
         Begin VB.OptionButton optInvoice 
            Caption         =   "Invoice Only"
            Height          =   210
            Left            =   195
            TabIndex        =   13
            Top             =   210
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   210
            Left            =   3630
            TabIndex        =   15
            Top             =   210
            Width           =   750
         End
         Begin VB.OptionButton optPack 
            Caption         =   "Packing List Only"
            Height          =   210
            Left            =   1725
            TabIndex        =   14
            Top             =   210
            Width           =   1590
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Enabled         =   0   'False
         Height          =   450
         Left            =   1095
         TabIndex        =   16
         Top             =   6075
         Width           =   1485
      End
Begin AceSpread sprInvoice
         Height          =   3630
         Left            =   120
         TabIndex        =   8
         Top             =   1620
         Width           =   7860
         _Version        =   458752
         _ExtentX        =   13864
         _ExtentY        =   6403
         _StockProps     =   64
         Enabled         =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         SelectBlockOptions=   0
      End
   End
End
Attribute VB_Name = "rptInvoicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Private Enum anFindBy
  anByDate = 0
  anByNo = 1
  anByBoth = 2
  anByNone = 3
End Enum

Private Enum eInv
  mInvType = 1
  mInvNo = 2
  mInvDate = 3
  mCustomer = 4
  mCustCode = 5
  mExport = 6
  mBoxQty = 7
End Enum

Private Enum ecol
  mCol1 = 1
  mCol2 = 2
  mCol3 = 3
  mcol4 = 4
  mcol5 = 5
  mcol6 = 6
  mcol7 = 7
  mcol8 = 8
  mcol9 = 9
End Enum

Private k As Integer
Private i As Integer
Private EDAmt As Double
Private CessAmt As Double
Private FetchLabelData As String

Private Net_PkSlipParam() As String

Private Function Get_TotalDutyPaid(ByVal ExprName As String,ByVal  ResultVal As Double) As Double
' On Error Resume Next
  
  If (Left(ExprName, 4) = "CGST") Or (Left(ExprName, 4) = "SGST") Or (Left(ExprName, 4) = "IGST") Then
    Get_TotalDutyPaid = ResultVal
  Else
    Get_TotalDutyPaid = 0
  End If
  
End Function

Private Sub INFANT()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim EccNo As String
  Dim Rule As String
  Dim Place As String 'Postal Address Range and Division printing
  Dim Range As String
  Dim Division As String
  Dim City As String
  
  Try

  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12 * 1440
  Printer.PrintQuality = vbPRPQDraft

  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select top 1 ReferenceNo from SAL_ORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'), (select top 1 ReferenceNo from SAL_JOBWORKORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "')) RefNo, a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select (Address1+', '+Address2) as add1, (City+'-'+Pincode)as Add2,"
  Sql = Sql & " ('TNGST No. '+ LST+' CST No. : ' + Cst +' Dt : '+ convert(varchar(10), CstDate, 105)+' '+'AreaCode No.: '+AreaCode)as ApprovalDet,"
  Sql = Sql & " NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
 
  'EccNo = "CE RC No. M /050605 / 1129 / ECC No. AAA CI 847 4P XM 001"
  'Rule = " Rule (52A & 173C of the Central Excise Rules 1994)"
  
  Dumrow 2
  'Printer.FontName = "Courier New"
  Printer.FontName = "Arial"
  Printer.FontBold = True
  Printer.FontSize = 15
  Printer.Print Tab(53); Left(CStr(snap("InvoiceNo")), 4)
  Printer.FontBold = False
  Printer.FontSize = 8
  'Printer.Print Tab(2); Left(rs("Add1"), 80)
  'Printer.Print Tab(2); Left(rs("Add2"), 80)
  'Printer.Print Tab(2); Left(EccNo, 75)
  'Printer.Print Tab(2); Left(Rule, 50);
  Dumrow 2
  Printer.FontSize = 14
  Printer.Print Tab(64); Left(Format(snap("InvoiceDate"), "dd/mm/yyyy"), 10) 'ApprovalDet
  Printer.FontName = "Draft 10cpi"
  Dumrow 2
  Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(34); snap("CustomerCode"); Tab(41.5); snap("ECCNo"); Tab(58); IIf(IsNull(snap("LST")), "", snap("LST")); Tab(70); IIf(IsNull(snap("CST")), "", snap("CST"))
  Printer.Print Tab(2); snap("add1")
  Printer.Print Tab(2); IIf(IsNull(snap("add2")), "", snap("add2"))
  Printer.Print Tab(2); IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city"); Tab(41); IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
  Printer.Print Tab(2); snap("pincode")
  Dumrow 2
  Printer.Print Tab(41); snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")
  Dumrow 2
  Place = "OFFICE OF THE SUPERINTENDENT OF CENTRAL EXCISE"
  Range = "RANGE - Poonamallee - V"
  Division = "Division - Poonamallee, Commissionerate - IV  Chennai."
  Printer.FontName = "Arial"
  Printer.FontBold = True
  Printer.FontSize = 12 'Create on 20-07-06 only these 6 lines For Printing add, Range, Division
  'Printer.Print Tab(2); Left(Place, 48);
  'Printer.Print Tab(2); Left(Range, 48); Tab(66); snap("mode") & "  " & snap("VehicleNo")
  Printer.Print Tab(45); Left(snap("mode"), 66) & "  " & snap("VehicleNo")
  'Printer.Print Tab(2); Left(Division, 59);
  'Printer.Print Tab(2); Left(City, 48);
  Printer.FontName = "Draft 10cpi" 'Create on 20-07-06
  Printer.FontSize = 8
  Dumrow 4
  Printer.Print Tab(41); Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm"); Tab(61); Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("RemovalTime"), "HH:mm")
  Dumrow 2
  Printer.Print Tab(2); snap("Commodity"); Tab(33); snap("TariffNo"); Tab(54); rs("NotificationNo") & "   " & rs("NotificationDate")
  Dumrow 3
  Printer.Print 'Tab(83); IIf(IsNull(Snap("eddesc")), "", Snap("eddesc"))
  Printer.Print Tab(2); "1"; Tab(6); snap("ItemDescription"); Tab(38); snap("NoofPackets"); Tab(48 - Len(snap("InvoiceQty"))); snap("InvoiceQty"); Tab(56 - Len(Format(snap("Rate"), "0.00"))); Format(snap("Rate"), "0.00"); Tab(67 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Printer.Print Tab(6); IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))
  X = 0
  If Trim$(snap("Type")) = "Rejection" Then
    WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      Printer.Print Tab(5); WSplit(1)
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      Printer.Print Tab(5); WSplit(2)
X = X + 1
    End If
  End If
  
  If Trim$(snap("Type")) = "Supplementary" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(15); "SUPPLEMENTARY INVOICE"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 3
  'ElseIf Trim(snap("Type")) = "OE" Then
  '  Dumrow 10 - X
  ElseIf Trim(snap("Type")) = "Labour" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(16); "LABOUR CHARGES ONLY"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  ElseIf Trim(snap("Type")) = "OE" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      Printer.Print Tab(16); "VAT INVOICE"
    Else
      Printer.Print Tab(16); ""
    End If
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  ElseIf Trim(snap("Type")) = "EOU" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      Printer.Print Tab(16); "VAT INVOICE"
    Else
      Printer.Print Tab(16); ""
    End If
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  Else
    Dumrow 7 - X
     Dumrow 3
  End If
  Printer.Print
  Printer.Print Tab(78 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Dumrow 3
  
  Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
    For i = 1 To 14
      If i = 7 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        End If
        If Not rs.EOF Then
          Printer.Print Tab(2); WSplit(1); Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          rs.MoveNext
        End If
        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2); Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          rs.MoveNext
        End If
      ElseIf i >= 10 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          Printer.Print Tab(2); WSplit(1);
        Else
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2);
          End If
        End If
        If Not rs.EOF Then
          If rs("ExprName") = "GRAND TOTAL" Then
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          rs.MoveNext
        End If
      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          If CheckNull(rs("ExprName")) = "GRAND TOTAL" Then
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          If rs("ExprName") = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If rs("ExprName") = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        Else
          Printer.Print
        End If
      Else
        Printer.Print
      End If
    Next i
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub INFANT_CancelInvoice()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim EccNo As String
  Dim Rule As String
  Dim Place As String 'Postal Address Range and Division printing
  Dim Range As String
  Dim Division As String
  Dim City As String
  
  Try

  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12 * 1440
  Printer.PrintQuality = vbPRPQDraft

  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select top 1 ReferenceNo from SAL_ORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'), (select top 1 ReferenceNo from SAL_JOBWORKORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "')) RefNo, a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " Sales_CancelInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select (Address1+', '+Address2) as add1, (City+'-'+Pincode)as Add2,"
  Sql = Sql & " ('TNGST No. '+ LST+' CST No. : ' + Cst +' Dt : '+ convert(varchar(10), CstDate, 105)+' '+'AreaCode No.: '+AreaCode)as ApprovalDet,"
  Sql = Sql & " NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
 
  'EccNo = "CE RC No. M /050605 / 1129 / ECC No. AAA CI 847 4P XM 001"
  'Rule = " Rule (52A & 173C of the Central Excise Rules 1994)"
  
  Dumrow 2
  'Printer.FontName = "Courier New"
  Printer.FontName = "Arial"
  Printer.FontBold = True
  Printer.FontSize = 15
  Printer.Print Tab(53); Left(CStr(snap("InvoiceNo")), 4)
  Printer.FontBold = False
  Printer.FontSize = 8
  'Printer.Print Tab(2); Left(rs("Add1"), 80)
  'Printer.Print Tab(2); Left(rs("Add2"), 80)
  'Printer.Print Tab(2); Left(EccNo, 75)
  'Printer.Print Tab(2); Left(Rule, 50);
  Dumrow 2
  Printer.FontSize = 14
  Printer.Print Tab(64); Left(Format(snap("InvoiceDate"), "dd/mm/yyyy"), 10) 'ApprovalDet
  Printer.FontName = "Draft 10cpi"
  Dumrow 2
  Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(34); snap("CustomerCode"); Tab(41.5); snap("ECCNo"); Tab(58); IIf(IsNull(snap("LST")), "", snap("LST")); Tab(70); IIf(IsNull(snap("CST")), "", snap("CST"))
  Printer.Print Tab(2); snap("add1")
  Printer.Print Tab(2); IIf(IsNull(snap("add2")), "", snap("add2"))
  Printer.Print Tab(2); IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city"); Tab(41); IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
  Printer.Print Tab(2); snap("pincode")
  Dumrow 2
  Printer.Print Tab(41); snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")
  Dumrow 2
  Place = "OFFICE OF THE SUPERINTENDENT OF CENTRAL EXCISE"
  Range = "RANGE - Poonamallee - V"
  Division = "Division - Poonamallee, Commissionerate - IV  Chennai."
  Printer.FontName = "Arial"
  Printer.FontBold = True
  Printer.FontSize = 12 'Create on 20-07-06 only these 6 lines For Printing add, Range, Division
  'Printer.Print Tab(2); Left(Place, 48);
  'Printer.Print Tab(2); Left(Range, 48); Tab(66); snap("mode") & "  " & snap("VehicleNo")
  Printer.Print Tab(45); Left(snap("mode"), 66) & "  " & snap("VehicleNo")
  'Printer.Print Tab(2); Left(Division, 59);
  'Printer.Print Tab(2); Left(City, 48);
  Printer.FontName = "Draft 10cpi" 'Create on 20-07-06
  Printer.FontSize = 8
  Dumrow 4
  Printer.Print Tab(41); Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm"); Tab(61); Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("RemovalTime"), "HH:mm")
  Dumrow 2
  Printer.Print Tab(2); snap("Commodity"); Tab(33); snap("TariffNo"); Tab(54); rs("NotificationNo") & "   " & rs("NotificationDate")
  Dumrow 3
  Printer.Print 'Tab(83); IIf(IsNull(Snap("eddesc")), "", Snap("eddesc"))
  Printer.Print Tab(2); "1"; Tab(6); snap("ItemDescription"); Tab(38); snap("NoofPackets"); Tab(48 - Len(Format(snap("InvoiceQty"), "0"))); Format(snap("InvoiceQty"), "0"); Tab(56 - Len(Format(snap("Rate"), "0.00"))); Format(snap("Rate"), "0.00"); Tab(67 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Printer.Print Tab(6); IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))
  X = 0
  If Trim$(snap("Type")) = "Rejection" Then
    WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      Printer.Print Tab(5); WSplit(1)
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      Printer.Print Tab(5); WSplit(2)
X = X + 1
    End If
  End If
  
  If Trim$(snap("Type")) = "Supplementary" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(15); "SUPPLEMENTARY INVOICE"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 3
  'ElseIf Trim(snap("Type")) = "OE" Then
  '  Dumrow 10 - X
  ElseIf Trim(snap("Type")) = "Labour" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(16); "LABOUR CHARGES ONLY"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  ElseIf Trim(snap("Type")) = "CANCEL" Then
    Dumrow 4 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 35
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      Printer.Print Tab(6); "CANCELED INVOICE"
    Else
      Printer.Print Tab(16); ""
    End If
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  ElseIf Trim(snap("Type")) = "EOU" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      Printer.Print Tab(16); "VAT INVOICE"
    Else
      Printer.Print Tab(16); ""
    End If
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 2
  End If
  Printer.Print
  Printer.Print Tab(78 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Dumrow 3
  
  Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
    For i = 1 To 14
      If i = 7 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        End If
        If Not rs.EOF Then
          Printer.Print Tab(2); WSplit(1); Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          rs.MoveNext
        End If
        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2); Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          rs.MoveNext
        End If
      ElseIf i >= 10 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          Printer.Print Tab(2); WSplit(1);
        Else
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2);
          End If
        End If
        If Not rs.EOF Then
          If rs("ExprName") = "GRAND TOTAL" Then
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          rs.MoveNext
        End If
      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          If CheckNull(rs("ExprName")) = "GRAND TOTAL" Then
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(58); rs("ExprName"); Tab(78 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          If rs("ExprName") = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If rs("ExprName") = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        Else
          Printer.Print
        End If
      Else
        Printer.Print
      End If
    Next i
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub Dwaraka()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim tot As Double
  Dim CurDate As Date
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If
  
  CurDate = ServerDate

  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12.1 * 1440
  Printer.PrintQuality = vbPRPQDraft
  Printer.FontName = "Draft 10cpi"

  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isnull((select ReferenceNo from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'),(select ReferenceNo from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'))  RefNo, a.AInformation1,a.Type from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join INV_Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and f.status = 1 and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  Else
    'Continue
  End If
  
  Dumrow 4
  Printer.FontBold = True
  Printer.Print Tab(50); snap("InvoiceNo")
  Printer.FontBold = False
  Dumrow 2
  Printer.Print Tab(47); Format(snap("InvoiceDate"), "dd/mm/yyyy")
  Dumrow 3
  Printer.Print Tab(2); Left(snap("PartyName"), 32); Tab(35); snap("CustomerCode"); Tab(48); IIf(IsNull(snap("LST")), "", snap("LST")); Tab(65); IIf(IsNull(snap("CST")), "", snap("CST"))
  Printer.Print Tab(2); snap("add1")
  Printer.Print Tab(2); IIf(IsNull(snap("add2")), "", snap("add2"))
  Printer.Print Tab(2); IIf(IsNull(snap("add3")), "", snap("add3")); Tab(35); snap("ECCNo"); Tab(58); snap("Refno")
  Printer.Print Tab(2); snap("city")
  Printer.Print Tab(2); snap("Pincode")
  Printer.Print
  Printer.Print Tab(48); IIf(IsNull(snap("Mode")), "", snap("Mode")) & "  " & snap("VehicleNo")
  Dumrow 10
  tot = 0
k = 0

  snap.MoveLast
snap.MoveFirst
  For i = 1 To snap.RecordCount
    Printer.Print Tab(2); i; Tab(5); snap("ItemDescription"); Tab(56 - Len(Format(snap("InvoiceQty"), "0.00"))); Format(snap("InvoiceQty"), "0.00"); Tab(66 - Len(Format(snap("Rate"), "0.00"))); Format(snap("Rate"), "0.00"); Tab(77 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
    Printer.Print Tab(5); IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))
    Printer.Print
    tot = tot + Format(snap("GoodsValue"), "0.00")
    k = k + 1
    snap.MoveNext
  Next i
  If snap.RecordCount < 8 Then
    Dumrow (8 - k) * 3
  End If
  Printer.Print
  Printer.Print Tab(77 - Len(Format(tot, "0.00"))); Format(tot, "0.00")
  Dumrow 2
  
  Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & gModuleCode & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  snap.MoveFirst
  k = 0
tot = snap("TotalAmount")
EDAmt = 0
  If Not rs.EOF Then
    rs.MoveLast
rs.MoveFirst
    For i = 1 To 8
      If i = 6 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
          If Not rs.EOF Then
            If rs("ExprName") = "GRAND TOTAL" Then
              Printer.Print Tab(2); WSplit(1);
              Printer.FontBold = True
              Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
              Printer.FontBold = False
            Else
              Printer.Print Tab(2); WSplit(1); Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
            End If
          Else
            Printer.Print
          End If
          rs.MoveNext
        End If
      ElseIf i = 7 Then
        If EDAmt > 0 Then
          If UBound(WSplit) >= 2 Then
            If Not rs.EOF Then
              If rs("ExprName") = "GRAND TOTAL" Then
                Printer.Print Tab(2); WSplit(2);
                Printer.FontBold = True
                Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
                Printer.FontBold = False
              Else
                Printer.Print Tab(2); WSplit(2); Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
              End If
              rs.MoveNext
            Else
              Printer.Print
            End If
          Else
            If Not rs.EOF Then
              If rs("ExprName") = "GRAND TOTAL" Then
                Printer.FontBold = True
                Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
                Printer.FontBold = False
              Else
                Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
              End If
              rs.MoveNext
            Else
              Printer.Print
            End If
          End If
        Else
          Printer.Print
        End If
      Else
        If Not rs.EOF Then
          If rs("ExprName") = "GRAND TOTAL" Then
            Printer.FontBold = True
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
            Printer.FontBold = False
          Else
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          If rs("ExprName") = "EXCISE DUTY" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If rs("ExprName") = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        Else
          Printer.Print
        End If
      End If
    Next i
  End If
  Printer.Print
  Printer.Print
  tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
  WordSplit tmp, 48
  Printer.Print Tab(2); WSplit(1);
  If UBound(WSplit) >= 2 Then
    Printer.Print Tab(2); WSplit(2);
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set snap = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Function Infant_ViewPrint() As Boolean
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim EccNo As String
  Dim Rule As String
  Dim Place As String             'Postal Address Range and Division printing
  Dim Range As String
  Dim Division As String
  Dim City As String
  
  Try
  
  Infant_ViewPrint = False
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Record's Found"
    GoTo ExitHere
  End If

  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select top 1 ReferenceNo from SAL_ORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'), (select top 1 ReferenceNo from SAL_JOBWORKORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "')) RefNo, a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select (Address1+', '+Address2) as add1, (City+'-'+Pincode)as Add2,"
  Sql = Sql & " ('TNGST No. '+ LST+' CST No. : ' + Cst +' Dt : '+ convert(varchar(10), CstDate, 105)+' '+'AreaCode No.: '+AreaCode)as ApprovalDet,"
  Sql = Sql & " NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
    sprView.SetText ecol.mCol1, 2, gCompName
    sprView.SetText ecol.mcol7, 3, CStr(snap("InvoiceNo"))
    sprView.SetText ecol.mcol7, 6, Format(snap("InvoiceDate"), "dd/mm/yyyy")
    sprView.SetText ecol.mCol1, 9, Space(3) & snap("PartyName")
    sprView.SetText ecol.mCol1, 10, Space(3) & snap("add1")
    sprView.SetText ecol.mCol1, 11, Space(3) & IIf(IsNull(snap("add2")), "", snap("add2"))
    sprView.SetText ecol.mCol1, 12, Space(3) & IIf(IsNull(snap("add3")), "", snap("add3")) & Space(2) & snap("city")
    sprView.SetText ecol.mCol1, 13, Space(3) & CStr(snap("Pincode"))
    sprView.SetText ecol.mcol4, 9, snap("CustomerCode")
    sprView.SetText ecol.mcol6, 9, snap("ECCNo")
    sprView.SetText ecol.mcol8, 9, IIf(IsNull(snap("LST")), "", snap("LST"))
    sprView.SetText ecol.mcol9, 9, IIf(IsNull(snap("CST")), "", snap("CST"))
    sprView.SetText ecol.mcol4, 11, IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
    sprView.SetText ecol.mcol4, 13, snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")
    sprView.SetText ecol.mcol4, 15, Left(snap("mode"), 66) & Space(2) & snap("VehicleNo")
'    Place = "Office of the Superintent of central Excise,"
'    Range = "Range 1C,Brindavanam,Puducherry-605013."
'    Division = "Division - Poonamallee, Commissionerate - IV  Chennai."
'    .SetText ecol.mCol1, 15, Place
'    .SetText ecol.mCol1, 16, Range
'    .SetText ecol.mCol1, 17, Division
    sprView.SetText ecol.mcol5, 19, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm")
    sprView.SetText ecol.mcol8, 19, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("RemovalTime"), "HH:mm")
    sprView.SetText ecol.mCol1, 21, snap("Commodity")
    sprView.SetText ecol.mcol4, 21, snap("TariffNo")
    sprView.SetText ecol.mcol7, 21, rs("NotificationNo") & "   " & rs("NotificationDate")
    sprView.Col = ecol.mCol1
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mCol1, 23, 1 & "."
    sprView.Col = ecol.mCol2
sprView.Row = 23
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mCol2, 23, snap("ItemDescription")
    sprView.Col = ecol.mcol5
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol5, 23, snap("NoofPackets")
    sprView.Col = ecol.mcol6
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol6, 23, snap("InvoiceQty")
    sprView.Col = ecol.mcol7
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol7, 23, Format(snap("Rate"), "0.00")
    sprView.Col = ecol.mcol8
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol8, 23, Format(snap("GoodsValue"), "0.00")
    For X = 1 To sprView.MaxCols
      sprView.Col = X
sprView.Row = 23
sprView.BackColor = RGB(255, 255, 0)

      sprView.Col = X
sprView.Row = 24
sprView.BackColor = RGB(255, 255, 0)
    Next X
    sprView.SetText ecol.mCol2, 24, IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))
  
    If Trim$(snap("Type")) = "Supplementary" Then
      sprView.SetText ecol.mCol1, 32, "SUPPLEMENTARY INVOICE"
    ElseIf Trim(snap("Type")) = "Labour" Then
      sprView.SetText ecol.mCol1, 32, "LABOUR CHARGES ONLY"
    ElseIf Trim(snap("Type")) = "OE" Then
      Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
      Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If Not rs.EOF Then
        sprView.SetText ecol.mCol1, 32, "VAT INVOICE"
      Else
        sprView.SetText ecol.mCol1, 32, ""
      End If
    ElseIf Trim(snap("Type")) = "EOU" Then
      Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
      Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If Not rs.EOF Then
        sprView.SetText ecol.mCol1, 32, "VAT INVOICE"
      Else
        sprView.SetText ecol.mCol1, 32, ""
      End If
    End If
    sprView.SetText ecol.mcol9, 34, Format(snap("GoodsValue"), "0.00")
  
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " order by a.ExprSeq"
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      snap.MoveLast
snap.MoveFirst

      k = 0
tot = snap("TotalAmount")
EDAmt = 0
X = 1

      rs.MoveLast
rs.MoveFirst
      sprView.ClearRange ecol.mcol8, 36, sprView.MaxCols, 44, True
      sprView.ClearRange ecol.mCol1, 42, 1, 44, True
      For i = 1 To 14
        If i = 7 Then
          If Not rs.EOF Then
            sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
            sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            rs.MoveNext
            X = X + 1
          End If
          If Not rs.EOF Then
            tmp1 = RupeesInWords(rs("ResultValue"), "Rupees", "Paise")
            sprView.SetText ecol.mCol1, 42, tmp1
            sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
            sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            rs.MoveNext
            X = X + 1
          End If
        ElseIf i >= 10 And i <= 11 Then
          tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")

          If Not rs.EOF Then
            If rs("ExprName") = "GRAND TOTAL" Then
              sprView.SetText ecol.mCol1, 44, tmp
              sprView.SetText ecol.mcol8, 44, rs("ExprName")
              sprView.SetText ecol.mcol9, 44, Format(rs("ResultValue"), "0.00")
            ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
              sprView.SetText ecol.mcol8, X + 35, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 35, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "VAT" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "CST" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            Else
              sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            End If
            rs.MoveNext
          End If
        ElseIf i Mod 2 = 1 Then
          If Not rs.EOF Then
            tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
            If CheckNull(rs("ExprName")) = "GRAND TOTAL" Then
              sprView.SetText ecol.mCol1, 44, tmp
              sprView.SetText ecol.mcol8, 44, rs("ExprName")
              sprView.SetText ecol.mcol9, 44, Format(rs("ResultValue"), "0.00")
            ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
              sprView.SetText ecol.mcol8, X + 35, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 35, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "VAT" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "CST" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            Else
              sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
              X = X + 1
            End If
            If rs("ExprName") = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
            If rs("ExprName") = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
            rs.MoveNext
          End If
        End If
      Next i
    End If
    
  FraPlaceHolder.Visible = False
  fraView.Visible = True
  fraView.Height = 8070
  fraView.Width = 11595
  fraView.Left = 90
  fraView.Top = -90
'  Call Spread_View(sprView)

  Infant_ViewPrint = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Function

Public Sub PrintPackingList(ByVal OrdType As String,ByVal  ORDNo As Integer,ByVal  OrdDate As Date,ByVal  CustomerCode As Integer)
  Dim rs As AceADODB.Recordset
  Dim rt As AceADODB.Recordset
  Dim rd As AceADODB.Recordset
  Dim rsOrd As AceADODB.Recordset
  Dim Sql As String
  Dim rCnt As Long
  Dim Qty As Double
  Dim NetWeight As Double
  Dim TotalPack As String
  Dim Container As String
  Dim CompName As String
  Dim CurDate As Date
  
  Try
   
  k = 0
  iPageNo = 1
  iPageWidth = 80
  pheight = 86
  
  CurDate = ServerDate

  Printer.PaperSize = vbPRPSA4
  Printer.FontName = "Draft 10cpi"
  If Printer.FontName <> "Draft 10cpi" Then
    Printer.FontName = "Roman 10cpi"
    If Printer.FontName <> "Roman 10cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If

  Sql = " select CompanyName, Address1, Address2, City, Pincode from company_master"
  Sql = Sql & " where compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Dumrow 5
  
  Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate From Sal_Invoice_Dtl a"
  Sql = Sql & " Where a.InvoiceType='" & OrdType & "' and a.InvoiceNo=" & ORDNo & ""
  Sql = Sql & " and a.InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "' and a.Compid='" & gCompID & "' and a.Branchid='" & gBranchID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Set rsOrd = New AceADODB.Recordset
  If RSOpen(rsOrd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = " Select a.OrderNo, a.ReferenceNo, a.ReferenceDate,b.CarriageName,b.CarriageNo,"
  Sql = Sql + " b.DischargePort,b.ReceivingPlace,b.LoadingPort,b.DestinationPlace,b.Reference From dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " Left Outer Join SAL_PACKINGLIST b"
  Sql = Sql & " On b.OrderType = a.OrderType and b.OrderNo = a.OrderNo and b.OrderDate = a.OrderDate"
  Sql = Sql & " and b.BranchID = a.BranchID and b.CompID = a.CompID and b.Finyear = a.Finyear"
  Sql = Sql & " Where a.OrderType= '" & rsOrd("OrderType") & "' and a.OrderNo=" & rsOrd("OrderNo") & " and a.OrderDate='" & Format(rsOrd("OrderDate"), "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.Compid='" & gCompID & "' and a.Branchid='" & gBranchID & "' and a.finyear='" & gFinyear & "'"
  Set rd = New AceADODB.Recordset
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  Printer.FontBold = True
  Printer.Print Space(1) & CAlign(80, "PACKING LIST")
  Printer.FontBold = False
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(17, "Packing List No.:") & Space(1) & LAlign(10, CStr(IIf(ORDNo = 0, "", ORDNo))) & Space(20) & LAlign(18, "Packing List Date:") & Space(1) & LAlign(10, CStr(OrdDate))
  Printer.Print Space(1) & LAlign(17, "Customer PO No. :") & Space(1) & LAlign(10, rd("ReferenceNo")) & Space(20) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, rd("ReferenceDate"))
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(40, "Exporter:")
  
  Printer.Print Space(1) & LAlign(40, rs("CompanyName"))
  CompName = rs("CompanyName")
  Printer.Print Space(1) & LAlign(60, rs("Address1"))
  Printer.Print Space(1) & LAlign(60, rs("Address2"))
  Printer.Print Space(1) & LAlign(60, rs("City") & "-" & rs("PinCode"))
  
  Sql = " Select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag,Country"
  Sql = Sql & " from Sal_PackingList_Address where Type='Consignee' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
 
  Sql = " Select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag, Country "
  Sql = Sql & " from Sal_PackingList_Address where Type='Buyer' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(38, "Consignee:") & Space(1) & LAlign(38, "BUYER (if Other than Consignee):")
  Printer.Print Space(1) & LAlign(38, rs("Name")) & Space(1) & LAlign(38, rd("Name"))
  Printer.Print Space(1) & LAlign(38, rs("Address1") + rs("Address2")) & Space(1) & LAlign(38, rd("Address1") + rd("Address2"))
  If rs("Address3") <> "" Or rd("Address3") <> "" Then
    Printer.Print Space(1) & LAlign(38, rs("Address3")) & Space(1) & LAlign(38, rd("Address3"))
  End If
  Printer.Print Space(1) & LAlign(38, IIf(rs("City") = "", "", rs("City") & "-") + rs("PinCode") & IIf(rs("Country") = "", "", "," & rs("Country"))) & Space(1) & LAlign(38, IIf(rd("City") = "", "", rd("City") & "-") + rd("PinCode") & IIf(rd("Country") = "", "", "," & rd("Country")))
  Printer.Print Space(1) & String(iPageWidth, "-")
  
  Sql = " Select CarriageName, CarriageNo, DischargePort, ReceivingPlace , LoadingPort,"
  Sql = Sql & " DestinationPlace, Reference, Buyer, Consignee, OrderFlag"
  Sql = Sql & " from Sal_PackingList Where OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  Set rt = New AceADODB.Recordset
  If RSOpen(rt, Sql) = False Then
    GoTo ExitHere
  End If
 
  Sql = " Select c.OrderType,c.OrderNo,c.OrderDate,b.DeliveryDesc,b.PaymentDesc, ModeDesc "
  Sql = Sql & " from sal_Invoice_dtl c"
  Sql = Sql & " left outer join ("
  Sql = Sql & " select a.OrderType,a.OrderNo,a.OrderDate, isNull((select b3.Description from TERMSMASTER b3"
  Sql = Sql & " where b3.TypeFlg='Delivery Instruction' and b3.Code=a.DeliveryCode"
  Sql = Sql & " and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') DeliveryDesc , "
  Sql = Sql & " isNull((select b4.Description from TERMSMASTER b4 "
  Sql = Sql & " where b4.TypeFlg='Payment Terms' and b4.Code=a.PaymentCode"
  Sql = Sql & " and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') PaymentDesc , "
  Sql = Sql & " isNull((select b4.Description from TERMSMASTER b4 where b4.TypeFlg='Mode'"
  Sql = Sql & " and b4.Code=a.ModeCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '')ModeDesc"
  Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " ) b on b.OrderType=c.OrderType and b.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate "
  Sql = Sql & " where c.InvoiceType='" & OrdType & "' and c.InvoiceNo=" & ORDNo & " and c.InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and c.Compid='" & gCompID & "' and c.Branchid='" & gBranchID & "' and  c.finyear='" & gFinyear & "'"
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = " Select ContainerNo, PackageKind, GoodsDescription,"
  Sql = Sql & "Quantity, NetWeight, Length, Width, Height,OrderFlag from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrdType & "' and OrderNo='" & ORDNo & "' and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If

  
  Printer.Print Space(1) & LAlign(26, "Pre-Carriage by          :") & Space(1) & LAlign(25, rt("CarriageName"))
  Printer.Print Space(1) & LAlign(26, "Place of receipt         :") & Space(1) & LAlign(25, rt("ReceivingPlace"))
  Printer.Print Space(1) & LAlign(26, "Origin                   :") & Space(1) & LAlign(25, "India")
  Printer.Print Space(1) & LAlign(26, "Destination              :") & Space(1) & LAlign(25, "USA")
  Printer.Print Space(1) & LAlign(26, "Vessel/Flight No         :") & Space(1) & LAlign(25, rt("CarriageNo"))
  Printer.Print Space(1) & LAlign(26, "Port of Loading          :") & Space(1) & LAlign(25, rt("LoadingPort"))
  Printer.Print Space(1) & LAlign(26, "Port Of Discharge        :") & Space(1) & LAlign(25, rt("DischargePort"))
  Printer.Print Space(1) & LAlign(26, "Final Destination        :") & Space(1) & LAlign(25, rt("DestinationPlace"))
  Printer.Print Space(1) & LAlign(26, "Terms of Delivery        :") & Space(1) & LAlign(25, rd("DeliveryDesc"))
  Printer.Print Space(1) & LAlign(26, "Mode                     :") & Space(1) & LAlign(25, rd("ModeDesc"))
  Printer.Print Space(1) & LAlign(26, "Payment Terms            :") & Space(1) & LAlign(25, rd("PaymentDesc"))
  Printer.Print Space(1) & LAlign(26, "Marks & Nos./Container No:") & Space(1) & LAlign(25, rs("ContainerNo"))
  Printer.Print Space(1) & LAlign(26, "No. & Kind of Package    :") & Space(1) & LAlign(25, rs("PackageKind"))
  Printer.Print Space(1) & String(iPageWidth, "-")
  
  Printer.Print Space(1) & LAlign(5, "S.No") & Space(1) & LAlign(35, "Description of Goods") & Space(1) & RAlign(10, "Quantity") & Space(1) & RAlign(10, "Nt.Wt") & Space(1) & RAlign(10, "Gr.Wt")
  Printer.Print Space(1) & Space(6) & LAlign(35, "Mechanical(Machinery) Spares") & Space(1) & RAlign(10, "in Pcs") & Space(1) & RAlign(10, "inKGS") & Space(1) & RAlign(10, "in KGS")
  Printer.Print Space(1) & String(iPageWidth, "-")
  Container = ""
  If Not rs.EOF Then
    For rCnt = 1 To rs.RecordCount
      Qty = Qty + rs("Quantity")
      NetWeight = NetWeight + rs("NetWeight")
      TotalPack = rs("PackageKind")
      If rCnt < rs.RecordCount Then
        Printer.Print Space(1) & LAlign(5, CStr(rCnt)) & Space(1) & LAlign(35, rs("GoodsDescription")) & Space(1) & RAlign(10, rs("Quantity")) & Space(1) & RAlign(10, rs("Netweight"))
      Else
        Printer.Print Space(1) & LAlign(5, CStr(rCnt)) & Space(1) & LAlign(35, rs("GoodsDescription")) & Space(1) & RAlign(10, rs("Quantity")) & Space(1) & RAlign(10, rs("Netweight")) & Space(1) & RAlign(10, Val(NetWeight) + 1)
      End If
      Container = rs("ContainerNo")
      rs.MoveNext
    Next rCnt
  End If
  
  If rs.RecordCount < 5 Then
    Dumrow 3 - rs.RecordCount
  End If
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & Space(6) & LAlign(35, "TOTAL") & Space(1) & RAlign(10, Val(Qty)) & Space(1) & RAlign(10, Val(NetWeight)) & Space(1) & RAlign(10, Val(NetWeight) + 1)
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(10, "Total Pack:") & Space(1) & LAlign(60, Trim$(TotalPack))
  Printer.Print
  Printer.Print Space(1) & Space(26) & LAlign(10, "Length") & Space(5) & LAlign(10, "Width") & Space(5) & LAlign(10, "Height")
  Printer.Print Space(1) & Space(26) & LAlign(10, "------") & Space(5) & LAlign(10, "-----") & Space(5) & LAlign(10, "------")
  Printer.Print
  rs.MoveFirst
  
  Sql = " Select distinct ContainerNo, Length, Width, Height"
  Sql = Sql & " from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrdType & "' and OrderNo='" & ORDNo & "' and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    Do While Not rs.EOF
      Printer.Print Space(1) & Space(5) & LAlign(11, rs("ContainerNo")) & Space(10) & LAlign(10, CStr(rs("Length")) & """") & Space(5) & LAlign(10, CStr(rs("Width")) & """") & Space(5) & LAlign(10, CStr(rs("Height")) & """")
      rs.MoveNext
    Loop
  End If
  
  Dumrow 2
  
  Printer.Print Space(1) & Space(44) & CAlign(30, "Signature & Date")
  Printer.Print Space(1) & Space(44) & CAlign(30, gCompName)
  Dumrow 3
  Printer.Print Space(1) & Space(48) & CAlign(30, "Shayam Todi")          'Hard code theDirector name according to Company.
  Printer.Print Space(1) & Space(48) & CAlign(30, "Director")
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set rd = Nothing
  Set rt = Nothing
End Try
End Sub

Public Sub RowProcess()
' On Error Resume Next
  If k >= pheight Then
    Printer.NewPage
    If gReportID = "KUS" Then
      Dumrow 11
      k = 11
    Else
      Dumrow 5
      k = 5
    End If
    iPageNo = iPageNo + 1
    sPageHeaderNew
  End If
End Sub

Public Sub PrintPackingListOld(ByVal OrdType As String,ByVal  ORDNo As Integer,ByVal  OrdDate As Date,ByVal  CustomerCode As Integer)
  Dim rs As AceADODB.Recordset
  Dim rt As AceADODB.Recordset
  Dim rd As AceADODB.Recordset
  Dim rsOrd As AceADODB.Recordset
  Dim Sql As String
  Dim rCnt As Long
  Dim Qty As Double
  Dim NetWeight As Double
  Dim TotalPack As String
  Dim Container As String
  Dim CompName As String
  Dim CurDate As Date
  
  Try
   
  k = 0
  iPageNo = 1
  iPageWidth = 80
  pheight = 86

  CurDate = ServerDate
  
  Printer.PaperSize = vbPRPSA4
  Printer.FontName = "Draft 10cpi"
  If Printer.FontName <> "Draft 10cpi" Then
    Printer.FontName = "Roman 10cpi"
    If Printer.FontName <> "Roman 10cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If

  Sql = " select CompanyName, (Address1+', '+Address2) as Add1, City+'-'+Pincode as Add2, Convert(Varchar, Phone) as Add3 "
  Sql = Sql & " from company_master where compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Dumrow 5
  
  Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate From Sal_Invoice_Dtl a"
  Sql = Sql & " Where a.InvoiceType='" & OrdType & "' and a.InvoiceNo=" & ORDNo & ""
  Sql = Sql & " and a.InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "' and a.Compid='" & gCompID & "' and a.Branchid='" & gBranchID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Set rsOrd = New AceADODB.Recordset
  If RSOpen(rsOrd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = " Select a.OrderNo, a.ReferenceNo, a.ReferenceDate, b.Reference From dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " Left Outer Join SAL_PACKINGLIST b"
  Sql = Sql & " On b.OrderType = a.OrderType and b.OrderNo = a.OrderNo and b.OrderDate = a.OrderDate"
  Sql = Sql & " and b.BranchID = a.BranchID and b.CompID = a.CompID and b.Finyear = a.Finyear"
  Sql = Sql & " Where a.OrderType= '" & rsOrd("OrderType") & "' and a.OrderNo=" & rsOrd("OrderNo") & " and a.OrderDate='" & Format(rsOrd("OrderDate"), "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.Compid='" & gCompID & "' and a.Branchid='" & gBranchID & "' and a.finyear='" & gFinyear & "'"
  Set rd = New AceADODB.Recordset
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Printer.Print Space(1) & CAlign(80, "Packing List")
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(17, "Packing List No.:") & Space(1) & LAlign(10, CStr(IIf(ORDNo = 0, "", ORDNo & "/"))) & Space(25) & LAlign(18, "Packing List Date:") & Space(1) & LAlign(10, CStr(OrdDate))
  Printer.Print Space(1) & LAlign(17, "Customer PO No. :") & Space(1) & LAlign(10, rd("ReferenceNo")) & Space(25) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, rd("ReferenceDate"))
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(40, "Exporter:") & Space(1) & LAlign(38, "Packing List No & Date")
  
  Printer.Print Space(1) & LAlign(40, rs("CompanyName"))
  CompName = rs("CompanyName")
  Printer.Print Space(1) & LAlign(40, rs("Add1"))
  Printer.Print Space(1) & LAlign(40, rs("Add2"))
  Printer.Print Space(1) & LAlign(40, rs("Add3"))
  
  Sql = " Select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag,Country"
  Sql = Sql & " from Sal_PackingList_Address where Type='Consignee' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
 
  Sql = " Select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag, Country "
  Sql = Sql & " from Sal_PackingList_Address where Type='Buyer' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Printer.Print Space(1) & String(iPageWidth, "-")
  
  Printer.Print Space(1) & LAlign(38, "CONSIGNEE") & Space(1) & LAlign(38, "BUYER (if Other than Consignee)")
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(38, rs("Name")) & Space(1) & LAlign(38, rd("Name"))
  Printer.Print Space(1) & LAlign(38, rs("Address1") + rs("Address2")) & Space(1) & LAlign(38, rd("Address1") + rd("Address2"))
  If rs("Address3") <> "" Or rd("Address3") <> "" Then
    Printer.Print Space(1) & LAlign(38, rs("Address3")) & Space(1) & LAlign(38, rd("Address3"))
  End If
  Printer.Print Space(1) & LAlign(38, IIf(rs("City") = "", "", rs("City") & "-") + rs("PinCode") & IIf(rs("Country") = "", "", "," & rs("Country"))) & Space(1) & LAlign(38, IIf(rd("City") = "", "", rd("City") & "-") + rd("PinCode") & IIf(rd("Country") = "", "", "," & rd("Country")))
  Printer.Print Space(1)
  
  Printer.Print Space(1) & String(iPageWidth, "-")
  
  Sql = " Select CarriageName, CarriageNo, DischargePort, ReceivingPlace , LoadingPort,"
  Sql = Sql & " DestinationPlace, Reference, Buyer, Consignee, OrderFlag"
  Sql = Sql & " from Sal_PackingList Where OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  Set rt = New AceADODB.Recordset
  If RSOpen(rt, Sql) = False Then
    GoTo ExitHere
  End If
 
  Printer.Print Space(1) & LAlign(18, "Pre-Carriage") & Space(1) & LAlign(19, "Place of receipt") & Space(1) & LAlign(19, "Country of Origin") & Space(1) & LAlign(20, "Country of Final")
  Printer.Print Space(1) & LAlign(18, "By") & Space(1) & LAlign(19, "byPre-Carrier") & Space(1) & LAlign(19, "Of Goods") & Space(1) & LAlign(20, "Destination")
  Printer.Print Space(1) & String(18, "-") & Space(1) & String(19, "-") & Space(1) & String(19, "-") & Space(1) & String(20, "-")
    
  Sql = " Select c.OrderType,c.OrderNo,c.OrderDate,b.DeliveryDesc,b.PaymentDesc, ModeDesc "
  Sql = Sql & " from sal_Invoice_dtl c"
  Sql = Sql & " left outer join ("
  Sql = Sql & " select a.OrderType,a.OrderNo,a.OrderDate, isNull((select b3.Description from TERMSMASTER b3"
  Sql = Sql & " where b3.TypeFlg='Delivery Instruction' and b3.Code=a.DeliveryCode"
  Sql = Sql & " and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') DeliveryDesc , "
  Sql = Sql & " isNull((select b4.Description from TERMSMASTER b4 "
  Sql = Sql & " where b4.TypeFlg='Payment Terms' and b4.Code=a.PaymentCode"
  Sql = Sql & " and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') PaymentDesc , "
  Sql = Sql & " isNull((select b4.Description from TERMSMASTER b4 where b4.TypeFlg='Mode'"
  Sql = Sql & " and b4.Code=a.ModeCode and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '')ModeDesc"
  Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " ) b on b.OrderType=c.OrderType and b.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate "
  Sql = Sql & " where c.InvoiceType='" & OrdType & "' and c.InvoiceNo=" & ORDNo & " and c.InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and c.Compid='" & gCompID & "' and c.Branchid='" & gBranchID & "' and  c.finyear='" & gFinyear & "'"
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
    
  Printer.Print Space(1) & LAlign(18, rt("CarriageName")) & Space(1) & LAlign(19, rt("ReceivingPlace")) & Space(1) & LAlign(19, "India") & Space(1) & LAlign(20, CheckNull(rs("Country")))
  Printer.Print
  Printer.Print Space(1) & LAlign(18, "Vessel/Flight No") & Space(1) & LAlign(19, "Port of Loading")
  Printer.Print Space(1) & String(18, "-") & Space(1) & String(19, "-")
  Printer.Print Space(1) & LAlign(18, rt("CarriageNo")) & Space(1) & LAlign(19, rt("LoadingPort"))
  Printer.Print
  Printer.Print Space(1) & LAlign(18, "Port Of Discharge") & Space(1) & LAlign(19, "Final Destination") & Space(1) & LAlign(40, "Terms of Delivery and Payment")
  Printer.Print Space(1) & String(18, "-") & Space(1) & String(19, "-") & Space(1) & String(40, "-")
  Printer.Print Space(1) & LAlign(18, rt("DischargePort")) & Space(1) & LAlign(19, rs("Address2")) & Space(1) & LAlign(40, rd("DeliveryDesc") & " -" & "By " & rd("ModeDesc"))
  Printer.Print Space(1) & Space(20) & LAlign(19, rs("City") + "," + rs("Pincode")) & Space(1) & LAlign(40, rd("PaymentDesc"))
  
  Sql = " Select ContainerNo, PackageKind, GoodsDescription,"
  Sql = Sql & "Quantity, NetWeight, Length, Width, Height,OrderFlag from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrdType & "' and OrderNo='" & ORDNo & "' and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(9, "Marks & ") '& Space(63) &CAlign(9, "ReMarks")
  Printer.Print Space(1) & LAlign(9, "Nos./  ") & Space(1) & LAlign(12, "No. & Kind") & Space(1) & LAlign(29, "Description of Goods") & Space(1) & LAlign(9, "Quantity") & Space(1) & LAlign(9, "Nt.Wt") & Space(1) & LAlign(5, "Gr.Wt")
  Printer.Print Space(1) & LAlign(9, "Container") & Space(1) & LAlign(12, "Of Package") & Space(31) & LAlign(9, "in Pcs") & Space(1) & LAlign(9, "inKGS") & Space(1) & LAlign(5, "in KGS")
  Printer.Print Space(1) & String(iPageWidth, "-")
  Dumrow 1
  Container = ""
  If Not rs.EOF Then
    For rCnt = 1 To rs.RecordCount
      Qty = Qty + rs("Quantity")
      NetWeight = NetWeight + rs("NetWeight")
      TotalPack = rs("PackageKind")
      If rCnt < rs.RecordCount Then
        If Container <> rs("ContainerNo") Then
          Printer.Print Space(1) & LAlign(9, rs("ContainerNo")) & Space(1) & LAlign(12, rs("PackageKind")) & Space(1) & LAlign(29, rs("GoodsDescription")) & Space(1) & LAlign(9, rs("Quantity")) & Space(1) & LAlign(9, rs("Netweight"))
        Else
          Printer.Print Space(1) & Space(23) & LAlign(29, rs("GoodsDescription")) & Space(1) & LAlign(9, rs("Quantity")) & Space(1) & LAlign(9, rs("Netweight"))
        End If
      Else
        Printer.Print Space(1) & Space(23) & LAlign(29, rs("GoodsDescription")) & Space(1) & LAlign(9, rs("Quantity")) & Space(1) & LAlign(9, rs("Netweight")) & Space(1) & LAlign(9, Val(NetWeight) + 1)
      End If
      Container = rs("ContainerNo")
      rs.MoveNext
    Next rCnt
  End If
  
  If rs.RecordCount < 5 Then
    Dumrow 3 - rs.RecordCount
  End If
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & Space(23) & LAlign(29, "TOTAL") & Space(1) & LAlign(9, Val(Qty)) & Space(1) & LAlign(9, Val(NetWeight)) & Space(1) & LAlign(9, Val(NetWeight) + 1)
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(10, "Total Pack") & Space(1) & LAlign(60, Trim$(TotalPack))
  
  Dumrow 1
  Printer.Print Space(1) & Space(26) & LAlign(10, "Length") & Space(5) & LAlign(10, "Width") & Space(5) & LAlign(10, "Height")
  Printer.Print Space(1) & Space(26) & LAlign(10, "------") & Space(5) & LAlign(10, "-----") & Space(5) & LAlign(10, "------")
  Dumrow 1
  rs.MoveFirst
  
  Sql = " Select distinct ContainerNo, Length, Width, Height"
  Sql = Sql & " from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrdType & "' and OrderNo='" & ORDNo & "' and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    Do While Not rs.EOF
      Printer.Print Space(1) & Space(5) & LAlign(11, rs("ContainerNo")) & Space(10) & LAlign(10, CStr(rs("Length")) & """") & Space(5) & LAlign(10, CStr(rs("Width")) & """") & Space(5) & LAlign(10, CStr(rs("Height")) & """")
      rs.MoveNext
    Loop
  End If
  
  Dumrow 2
  
  Printer.Print Space(1) & Space(44) & CAlign(30, "Signature & Date")
  Printer.Print Space(1) & Space(44) & CAlign(30, gCompName)
  Dumrow 3
  Printer.Print Space(1) & Space(48) & CAlign(30, "Shayam Todi")          'Hard code theDirector name according to Company.
  Printer.Print Space(1) & Space(48) & CAlign(30, "Director")
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set rd = Nothing
  Set rt = Nothing
End Try
End Sub
Public Sub sPageHeaderNew()

  Try
  Printer.Print Space(1) & LAlign(5, "S.No") & Space(1) & LAlign(38, "Part Description") & Space(1) & RAlign(8, "QTY") & Space(1) & CAlign(4, "Unit") & Space(1) & RAlign(10, "Price") & Space(1) & RAlign(10, "Value(USD)")
  Printer.Print Space(1) & Space(6) & LAlign(38, "Part No")
  Printer.Print Space(1) & String(iPageWidth, "-")
  k = k + 3
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Function Dumrow(ByVal rownos As Integer)
  Dim rowcnt As Integer
  For rowcnt = 1 To rownos
    Printer.Print
  Next rowcnt
End Function

Private Sub Unicorn()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim CurDate As Date
  Dim Total_DutyPaid As Double
  Dim Cnt As Long
  
  Dim Y_Pos As Single
  
  Dim PPRN As CLS_PREPRINT
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  CurDate = ServerDate
  
  Set PPRN = New CLS_PREPRINT
  
  If PPRN.PPRN_SetFont("Arial", 10, PPRNFSNone) = False Then
 GoTo ExitHere
End If
  If PPRN.PPRN_SetPaper(PPRNPSUser, PPRNSMCentimeters, 21.59, 30.48, PPRNORPortrait, PPRNPQDraft) = False Then GoTo ExitHere  '21.59, 30.48
  
  'sales invoice
  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_CUSTOMER_ITEMMASTER g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select ReferenceNo from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'),"
  Sql = Sql + " (select ReferenceNo from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "')) RefNo,"
  Sql = Sql + " a.AInformation1,a.Type,f.UOMDecimal, g.LCNo, g.LCDate, e.GSTRegNo"
  Sql = Sql + " from SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  Sql = Sql + " left outer join SAL_Invoice_LCRef g"
  Sql = Sql + " on g.InvNo=a.InvoiceNo and g.InvDate=a.InvoiceDate and g.InvType=a.InvoiceType and g.Branchid=a.Branchid and g.Compid=a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceDate = '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  Else
    'Continue
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  Else
    'Continue
  End If

  Sql = " select NotificationNo, NotificationDate, GSTRegNo from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  PPRN.PPRN_PrintText LAlign(15, snap("InvoiceNo")), 13.5, 1.1, , 16, PPRNFSBold
  
  'GSTIN Number
  'PPRN.PPRN_PrintText LAlign(10, "GSTIN: 34AAACU4426D1ZR"), 13.5, 2.1, , 12, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(10, Format(snap("InvoiceDate"), "dd/mm/yyyy")), 13.5, 2.1, , 16, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(35, snap("PartyName")), 0.5, 3.3, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(9, snap("CustomerCode")), 8.2, 3.3, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(44, String$(35, "=")), 10.5, 2.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(44, "GSTIN: " & snap("GSTRegNo")), 10.5, 3.3, , 10, PPRNFSNone
  'PPRN.PPRN_PrintText LAlign(44, snap("ECCNo") & "/" & IIf(IsNull(snap("LST")), "", snap("LST")) & "/" & IIf(IsNull(snap("CST")), "", snap("CST"))), 10.5, 3.3, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, snap("add1")), 0.5, 3.7, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("add2")), "", snap("add2"))), 0.5, 4.1, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city")), 0.5, 4.5, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("RefNo")), "", snap("RefNo"))), 10.3, 4.5, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(32, snap("pincode")), 0.5, 4.9, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(50, snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")), 10.3, 6.1, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText CAlign(50, snap("mode") & "  " & snap("VehicleNo")), 11, 7.6, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(20, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm")), 10.4, 9, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(36, snap("Commodity")), 0.5, 10.2, , 10, PPRNFSNone
  
  'HSN Code
  PPRN.PPRN_PrintText CAlign(22, "===================="), 8.7, 9.7, , 10, PPRNFSNone
  PPRN.PPRN_PrintText CAlign(22, "HSN Code: " & snap("TariffNo")), 8.7, 10.2, , 10, PPRNFSNone
  
  'PPRN.PPRN_PrintText LAlign(30, rs("NotificationNo") & "   " & rs("NotificationDate")), 14, 10.2, , 10, PPRNFSNone

  PPRN.PPRN_PrintText LAlign(30, "1"), 0.05, 11.8, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(37, snap("ItemDescription")), 1.3, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(4, snap("NoofPackets")), 9.2, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(8, snap("InvoiceQty")), 10.8, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(8, Format(snap("Rate"), "0.00")), 12.7, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(12, Format(snap("GoodsValue"), "0.00")), 15, 11.8, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(36, IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))), 1.3, 12.2, , 10, PPRNFSNone
  
  If CheckNull(snap("LCNo")) <> "" Then
    PPRN.PPRN_PrintText LAlign(45, "PAYMENT TERMS-90DAYS LC FROM THE DATE OF GRN"), 1.3, 13.2, , 10, PPRNFSNone
    PPRN.PPRN_PrintText LAlign(45, "LC NO :" & snap("LCNo") & " / " & snap("LCDate")), 1.3, 13.6, , 10, PPRNFSNone
  Else
    PPRN.PPRN_PrintText LAlign(45, "PAYMENT TERMS: " & snap("AInformation1")), 1.3, 13.2, , 10, PPRNFSNone
  End If
  
  X = 0
  
  If Trim$(snap("Type")) = "Rejection" Then

    WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      PPRN.PPRN_PrintText WSplit(1), 1.3, 15.4, , 10, PPRNFSNone
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      PPRN.PPRN_PrintText WSplit(2), 1.3, 15.8, , 10, PPRNFSNone
X = X + 1
    End If
  End If

'  If Trim$(snap("Type")) = "Supplementary" Then
'    PPRN.PPRN_PrintText "SUPPLEMENTARY INVOICE", 2.5, 15.6, , 20, PPRNFSNone
'  ElseIf Trim(snap("Type")) = "OE" Then
'    PPRN.PPRN_PrintText "VAT INVOICE", 2.5, 15.6, , 20, PPRNFSNone
'   ElseIf Trim(snap("Type")) = "Scrap" Then
'    PPRN.PPRN_PrintText "VAT INVOICE", 2.5, 15.6, , 20, PPRNFSNone
'  ElseIf Trim(snap("Type")) = "Labour" Then
'    PPRN.PPRN_PrintText "LABOUR CHARGES ONLY", 2.5, 15.6, , 20, PPRNFSNone
'  End If
  
  PPRN.PPRN_PrintText rs("GSTRegNo"), 2.5, 15.6, , 20, PPRNFSNone

  PPRN.PPRN_PrintText RAlign(13, Format(snap("GoodsValue"), "0.00")), 17.5, 17, , 10, PPRNFSNone
  
  Sql = " Select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults (" & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " where a.DocType='" & Trim$(sprInvoice.Text) & "' "
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocNo= " & Val(sprInvoice.Text) & " "
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocDate = convert(datetime, '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "',111) order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If Not rs.EOF Then
    Total_DutyPaid = 0
    Total_DutyPaid = Total_DutyPaid + Get_TotalDutyPaid(rs("ExprName"), rs("ResultValue"))
    
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
    Y_Pos = 18.4
   For i = 1 To 14

      If i = 7 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        Else
          WordSplit "", 48
        End If

        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(57, WSplit(1)), 0.5, 20.9, , 10, PPRNFSNone
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
          If rs.EOF = False Then
 Total_DutyPaid = Total_DutyPaid + Get_TotalDutyPaid(rs("ExprName"), rs("ResultValue"))
End If
        End If

        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            PPRN.PPRN_PrintText LAlign(57, WSplit(2)), 0.5, 21.3, , 10, PPRNFSNone
            PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
            PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          Else
            PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
            PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          End If

          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
          If rs.EOF = False Then
 Total_DutyPaid = Total_DutyPaid + Get_TotalDutyPaid(rs("ExprName"), rs("ResultValue"))
End If
        End If

      ElseIf i >= 10 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          PPRN.PPRN_PrintText LAlign(58, WSplit(1)), 0.5, 22, , 10, PPRNFSNone
          
        Else
          If UBound(WSplit) >= 2 Then
            PPRN.PPRN_PrintText LAlign(58, WSplit(2)), 0.5, 22.4, , 10, PPRNFSNone
          End If
        End If

        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
          If rs.EOF = False Then
 Total_DutyPaid = Total_DutyPaid + Get_TotalDutyPaid(rs("ExprName"), rs("ResultValue"))
End If
        End If

      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS 2%" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
          If rs.EOF = False Then
 Total_DutyPaid = Total_DutyPaid + Get_TotalDutyPaid(rs("ExprName"), rs("ResultValue"))
End If
        Else
'          Printer.Print
        End If
      Else
'        Printer.Print
      End If
    Next i
  End If
  
  Y_Pos = 20.8
  tmp = RupeesInWords(Format(Total_DutyPaid, "0.00"), "Rupees", "Paise")
  WordSplit tmp, 48
  
  For Cnt = 1 To UBound(WSplit)
    PPRN.PPRN_PrintText LAlign(58, Trim$(WSplit(Cnt))), 0.5, Y_Pos, , 10, PPRNFSNone
    Y_Pos = Y_Pos + 0.4
  Next Cnt
  
  
  PPRN.PPRN_EndDocument
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  PPRN.PPRN_KillDocument

Finally

  Set rs = Nothing
  Set snap = Nothing
   Set PPRN = Nothing
End Try
End Sub

Private Sub Unicorn_Old_BeforeGST()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim CurDate As Date
  
  Dim Y_Pos As Single
  
  Dim PPRN As CLS_PREPRINT
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  CurDate = ServerDate
  
  Set PPRN = New CLS_PREPRINT
  
  If PPRN.PPRN_SetFont("Arial", 10, PPRNFSNone) = False Then
 GoTo ExitHere
End If
  If PPRN.PPRN_SetPaper(PPRNPSUser, PPRNSMCentimeters, 21.59, 30.48, PPRNORPortrait, PPRNPQDraft) = False Then GoTo ExitHere  '21.59, 30.48
  
  'sales invoice
  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_CUSTOMER_ITEMMASTER g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select ReferenceNo from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'),"
  Sql = Sql + " (select ReferenceNo from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "')) RefNo,"
  Sql = Sql + " a.AInformation1,a.Type,f.UOMDecimal, g.LCNo, g.LCDate"
  Sql = Sql + " from SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  Sql = Sql + " left outer join SAL_Invoice_LCRef g"
  Sql = Sql + " on g.InvNo=a.InvoiceNo and g.InvDate=a.InvoiceDate and g.InvType=a.InvoiceType and g.Branchid=a.Branchid and g.Compid=a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceDate = '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  PPRN.PPRN_PrintText LAlign(15, snap("InvoiceNo")), 13.5, 1.1, , 16, PPRNFSBold
  PPRN.PPRN_PrintText LAlign(10, Format(snap("InvoiceDate"), "dd/mm/yyyy")), 13.5, 2.1, , 16, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(35, snap("PartyName")), 0.5, 3.3, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(9, snap("CustomerCode")), 8.2, 3.3, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(44, snap("ECCNo") & "/" & IIf(IsNull(snap("LST")), "", snap("LST")) & "/" & IIf(IsNull(snap("CST")), "", snap("CST"))), 10.5, 3.3, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, snap("add1")), 0.5, 3.7, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("add2")), "", snap("add2"))), 0.5, 4.1, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city")), 0.5, 4.5, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(32, IIf(IsNull(snap("RefNo")), "", snap("RefNo"))), 10.3, 4.5, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(32, snap("pincode")), 0.5, 4.9, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(50, snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")), 10.3, 6.1, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText CAlign(50, snap("mode") & "  " & snap("VehicleNo")), 11, 7.6, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(20, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm")), 10.4, 9, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(36, snap("Commodity")), 0.5, 10.2, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(22, snap("TariffNo")), 8.7, 10.2, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(30, rs("NotificationNo") & "   " & rs("NotificationDate")), 14, 10.2, , 10, PPRNFSNone

  PPRN.PPRN_PrintText LAlign(30, "1"), 0.05, 11.8, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(37, snap("ItemDescription")), 1.3, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText LAlign(4, snap("NoofPackets")), 9.2, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(8, snap("InvoiceQty")), 10.8, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(8, Format(snap("Rate"), "0.00")), 12.7, 11.8, , 10, PPRNFSNone
  PPRN.PPRN_PrintText RAlign(12, Format(snap("GoodsValue"), "0.00")), 15, 11.8, , 10, PPRNFSNone
  
  PPRN.PPRN_PrintText LAlign(36, IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))), 1.3, 12.2, , 10, PPRNFSNone
  
  If CheckNull(snap("LCNo")) <> "" Then
    PPRN.PPRN_PrintText LAlign(45, "PAYMENT TERMS-90DAYS LC FROM THE DATE OF GRN"), 1.3, 13.2, , 10, PPRNFSNone
    PPRN.PPRN_PrintText LAlign(45, "LC NO :" & snap("LCNo") & " / " & snap("LCDate")), 1.3, 13.6, , 10, PPRNFSNone
  Else
    PPRN.PPRN_PrintText LAlign(45, "PAYMENT TERMS: " & snap("AInformation1")), 1.3, 13.2, , 10, PPRNFSNone
  End If
  
  X = 0
  
  If Trim$(snap("Type")) = "Rejection" Then

    WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      PPRN.PPRN_PrintText WSplit(1), 1.3, 15.4, , 10, PPRNFSNone
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      PPRN.PPRN_PrintText WSplit(2), 1.3, 15.8, , 10, PPRNFSNone
X = X + 1
    End If
  End If

  If Trim$(snap("Type")) = "Supplementary" Then
    PPRN.PPRN_PrintText "SUPPLEMENTARY INVOICE", 2.5, 15.6, , 20, PPRNFSNone
  ElseIf Trim(snap("Type")) = "OE" Then
    PPRN.PPRN_PrintText "VAT INVOICE", 2.5, 15.6, , 20, PPRNFSNone
   ElseIf Trim(snap("Type")) = "Scrap" Then
    PPRN.PPRN_PrintText "VAT INVOICE", 2.5, 15.6, , 20, PPRNFSNone
  ElseIf Trim(snap("Type")) = "Labour" Then
    PPRN.PPRN_PrintText "LABOUR CHARGES ONLY", 2.5, 15.6, , 20, PPRNFSNone
  End If

  PPRN.PPRN_PrintText RAlign(13, Format(snap("GoodsValue"), "0.00")), 17.5, 17, , 10, PPRNFSNone
  
  Sql = " Select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults (" & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " where a.DocType='" & Trim$(sprInvoice.Text) & "' "
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocNo= " & Val(sprInvoice.Text) & " "
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocDate = convert(datetime, '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "',111) order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If Not rs.EOF Then
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
    Y_Pos = 18.4
   For i = 1 To 14

      If i = 7 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        End If

        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(57, WSplit(1)), 0.5, 20.9, , 10, PPRNFSNone
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
        End If

        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            PPRN.PPRN_PrintText LAlign(57, WSplit(2)), 0.5, 21.3, , 10, PPRNFSNone
            PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
            PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          Else
            PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
            PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          End If

          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
        End If

      ElseIf i >= 10 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          PPRN.PPRN_PrintText LAlign(58, WSplit(1)), 0.5, 22, , 10, PPRNFSNone
          
        Else
          If UBound(WSplit) >= 2 Then
            PPRN.PPRN_PrintText LAlign(58, WSplit(2)), 0.5, 22.4, , 10, PPRNFSNone
          End If
        End If

        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
        End If

      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          PPRN.PPRN_PrintText LAlign(20, rs("ExprName")), 13, Y_Pos, , 10, PPRNFSNone
          PPRN.PPRN_PrintText RAlign(13, Format(rs("ResultValue"), "0.00")), 17.5, Y_Pos, , 10, PPRNFSNone
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS 2%" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          
          Y_Pos = Y_Pos + 0.6
          rs.MoveNext
        Else
'          Printer.Print
        End If
      Else
'        Printer.Print
      End If
    Next i
  End If
  
  PPRN.PPRN_EndDocument
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  PPRN.PPRN_KillDocument

Finally

  Set rs = Nothing
  Set snap = Nothing
   Set PPRN = Nothing
End Try
End Sub
Private Sub Unicorn_WithSpaces()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim CurDate As Date
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  CurDate = ServerDate
  
  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12 * 1440
  Printer.PrintQuality = vbPRPQDraft
  
  'sales invoice
  
  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_CUSTOMER_ITEMMASTER g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select ReferenceNo from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'),"
  Sql = Sql + " (select ReferenceNo from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "')) RefNo,"
  Sql = Sql + " a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceDate = '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Dumrow 3
  Printer.FontName = "Courier New"
  Printer.FontSize = 16
  Printer.FontBold = True
  Printer.Print Space(40) & LAlign(15, snap("InvoiceNo")) '50
  Printer.FontBold = False
  Dumrow 1
  
  Printer.Print Space(38) & LAlign(10, Format(snap("InvoiceDate"), "dd/mm/yyyy"))
  Printer.FontSize = 10
  Dumrow 2
  
  Printer.Print Space(2) & LAlign(35, snap("PartyName")) & Space(2) & LAlign(9, snap("CustomerCode")) & Space(1) & LAlign(44, snap("ECCNo") & "/" & IIf(IsNull(snap("LST")), "", snap("LST")) & "/" & IIf(IsNull(snap("CST")), "", snap("CST")))
  Printer.Print Space(2) & LAlign(32, snap("add1"))
  Printer.Print Space(2) & LAlign(32, IIf(IsNull(snap("add2")), "", snap("add2")))
  Printer.Print Space(2) & LAlign(32, IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city")) & Space(10) & IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
  Printer.Print Space(2) & LAlign(32, snap("pincode"))
  Dumrow 1
  
  Printer.Print Space(43) & LAlign(50, snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy"))
  Dumrow 2
  
  Printer.Print Space(48) & LAlign(50, snap("mode") & "  " & snap("VehicleNo"))
  Dumrow 3
  
  Printer.Print Space(49) & LAlign(20, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm"))
  Dumrow 2
  
  Printer.Print Space(2) & LAlign(36, snap("Commodity")) & Space(1) & LAlign(22, snap("TariffNo")) & Space(1) & LAlign(30, rs("NotificationNo") & "   " & rs("NotificationDate"))
  Dumrow 3
  
  Printer.Print Space(1) & LAlign(4, "1") & Space(1) & LAlign(37, snap("ItemDescription")) & Space(1) & LAlign(4, snap("NoofPackets")) & Space(1) & RAlign(8, Format(snap("InvoiceQty"), "0")) & Space(1) & RAlign(8, Format(snap("Rate"), "0.00")) & Space(1) & RAlign(12, Format(snap("GoodsValue"), "0.00"))
  Printer.Print Space(5) & LAlign(36, IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo")))

  X = 0

  If Trim$(snap("Type")) = "Rejection" Then
    Dumrow 8
    
    WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      Printer.Print Space(5) & WSplit(1)
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      Printer.Print Space(5) & WSplit(2)
X = X + 1
    End If
  End If

  If Trim$(snap("Type")) = "Supplementary" Then
    Dumrow 7
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Space(7) & "SUPPLEMENTARY INVOICE"
    Printer.FontSize = 10
    Dumrow 1
  ElseIf Trim(snap("Type")) = "OE" Then
    Dumrow 7
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Space(7) & "VAT INVOICE"
    Printer.FontSize = 10
    Dumrow 1
   ElseIf Trim(snap("Type")) = "Scrap" Then
    Dumrow 7
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Space(7) & "VAT INVOICE"
    Printer.FontSize = 10
    Dumrow 1
  ElseIf Trim(snap("Type")) = "Labour" Then
    Dumrow 7 '7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Space(7) & "LABOUR CHARGES ONLY"
    Printer.FontSize = 10
    Dumrow 1
  End If
  
  Printer.Print Space(80) & RAlign(13, Format(snap("GoodsValue"), "0.00")) '73
  Dumrow 3

  Sql = " Select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults (" & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " where a.DocType='" & Trim$(sprInvoice.Text) & "' "
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocNo= " & Val(sprInvoice.Text) & " "
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocDate = convert(datetime, '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "',111) order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If Not rs.EOF Then
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
   For i = 1 To 14
  '''''''' For i = 1 To rs.RecordCount
      If i = 7 Then

''''     For i = 1 To 20
''''      If i = 19 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        End If
        
        If Not rs.EOF Then
          Printer.Print Space(2) & LAlign(57, WSplit(1)) & Space(1) & LAlign(20, rs("ExprName")) & Space(1) & RAlign(13, Format(rs("ResultValue"), "0.00"))
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        End If
        
        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            Printer.Print Space(2) & LAlign(57, WSplit(2)) & Space(1) & LAlign(20, rs("ExprName")) & Space(1) & RAlign(13, Format(rs("ResultValue"), "0.00"))
          Else
            Printer.Print Space(60) & LAlign(20, rs("ExprName")) & Space(1) & RAlign(13, Format(rs("ResultValue"), "0.00"))
            
          End If
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        End If
        
      ElseIf i >= 10 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          Printer.Print Space(2) & LAlign(58, WSplit(1))
        Else
          If UBound(WSplit) >= 2 Then
            Printer.Print Space(2) & LAlign(58, WSplit(2))
          End If
        End If
        
        If Not rs.EOF Then
'          If rs("ExprName") = "GRAND TOTAL" Then
'            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
'          Else
'            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
'          End If
          
          Printer.Print Space(60) & LAlign(20, rs("ExprName")) & Space(1) & RAlign(13, Format(rs("ResultValue"), "0.00"))
          
          rs.MoveNext
        End If
      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          
          Printer.Print Space(60) & LAlign(20, rs("ExprName")) & Space(1) & RAlign(13, Format(rs("ResultValue"), "0.00"))
          
'          If rs("ExprName") = "GRAND TOTAL" Then
'            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
'          Else
'            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
'          End If
          
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS 2%" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        Else
          Printer.Print
        End If
      Else
        Printer.Print
      End If
    Next i
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub Unicorn_Old()

Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim CurDate As Date
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If

  CurDate = ServerDate
  
  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12 * 1440
  Printer.PrintQuality = vbPRPQDraft
  
  

  'sales invoice
  
  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_CUSTOMER_ITEMMASTER g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select ReferenceNo from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'),"
  Sql = Sql + " (select ReferenceNo from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "')) RefNo,"
  Sql = Sql + " a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceDate = '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Dumrow 2
  Printer.FontName = "Courier New"
  Printer.FontSize = 16
  Printer.FontBold = True
  Printer.Print Tab(40); snap("InvoiceNo")
  Printer.FontBold = False
  Dumrow 1
  Printer.Print Tab(38); Format(snap("InvoiceDate"), "dd/mm/yyyy")
  'Printer.FontName = "Draft 10cpi"
  Printer.FontSize = 10
  Dumrow 2
  'Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(34); snap("CustomerCode"); Tab(41.5); snap("ECCNo"); Tab(57); IIf(IsNull(snap("LST")), "", snap("LST")); Tab(68); IIf(IsNull(snap("CST")), "", snap("CST"))
   'Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(34); snap("CustomerCode"); Tab(39.5); snap("ECCNo"); "/"; Tab(55); IIf(IsNull(snap("LST")), "", snap("LST")); "/"; Tab(66); IIf(IsNull(snap("CST")), "", snap("CST"))
  'Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(33); snap("CustomerCode"); Tab(37.5); snap("ECCNo"); "/"; Tab(55); IIf(IsNull(snap("LST")), "", snap("LST")); "/"; Tab(66); IIf(IsNull(snap("CST")), "", snap("CST"))
  Printer.Print Tab(2); Left(snap("PartyName"), 31); Tab(37); snap("CustomerCode") & Space(3); snap("ECCNo"); "/"; IIf(IsNull(snap("LST")), "", snap("LST")); "/"; IIf(IsNull(snap("CST")), "", snap("CST"))
  Printer.Print Tab(2); snap("add1")
  Printer.Print Tab(2); IIf(IsNull(snap("add2")), "", snap("add2"))
  Printer.Print Tab(2); IIf(IsNull(snap("add3")), "", snap("add3")) & " " & snap("city"); Tab(41); IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
  Printer.Print Tab(2); snap("pincode")
  Dumrow 2
  Printer.Print Tab(41); snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")
  Dumrow 3
  Printer.Print Tab(48); snap("mode") & "  " & snap("VehicleNo")
  Dumrow 3
  Printer.Print Tab(41); Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm"); Tab(61); ' Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("RemovalTime"), "HH:mm")
  Dumrow 3
  Printer.Print Tab(2); snap("Commodity"); Tab(33); snap("TariffNo"); Tab(54); rs("NotificationNo") & "   " & rs("NotificationDate")
  Dumrow 3
  Printer.Print 'Tab(83); IIf(IsNull(Snap("eddesc")), "", Snap("eddesc"))
  Dumrow 1
  Printer.Print Tab(2); "1"; Tab(5); snap("ItemDescription"); Tab(38); snap("NoofPackets"); Tab(48 - Len(Format(snap("InvoiceQty"), "0"))); Format(snap("InvoiceQty"), "0"); Tab(56 - Len(Format(snap("Rate"), "0.00"))); Format(snap("Rate"), "0.00"); Tab(67 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Printer.Print Tab(5); IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))

  X = 0

  If Trim$(snap("Type")) = "Rejection" Then

  WordSplit snap("AInformation1"), 30
    If UBound(WSplit) >= 1 Then
      Printer.Print Tab(5); WSplit(1)
X = X + 1
    End If
    If UBound(WSplit) >= 2 Then
      Printer.Print Tab(5); WSplit(2)
X = X + 1
    End If
  End If

  If Trim$(snap("Type")) = "Supplementary" Then
    'Dumrow 7 - X
    Dumrow 6
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(6); "SUPPLEMENTARY INVOICE"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 1
  ElseIf Trim(snap("Type")) = "OE" Then
    Dumrow 5
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(6); "VAT INVOICE"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 1
   ElseIf Trim(snap("Type")) = "Scrap" Then
    Dumrow 6
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(6); "VAT INVOICE"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 1
  ElseIf Trim(snap("Type")) = "Labour" Then
    Dumrow 7 - X
    Printer.FontName = "Courier New"
    Printer.FontSize = 20
    Printer.Print Tab(16); "LABOUR CHARGES ONLY"
    Printer.FontName = "Draft 10cpi"
    Printer.FontSize = 10
    Dumrow 5
  End If
  Printer.Print
  Printer.Print Tab(78 - Len(Format(snap("GoodsValue"), "0.00"))); Format(snap("GoodsValue"), "0.00")
  Dumrow 5

  Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
  Sql = Sql & " from dbo.fn_COM_GetExprResults (" & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " where a.DocType='" & Trim$(sprInvoice.Text) & "' "
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocNo= " & Val(sprInvoice.Text) & " "
  sprInvoice.Col = 3
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql & " and a.DocDate = convert(datetime, '" & Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") & "',111) order by a.ExprSeq"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    snap.MoveLast
snap.MoveFirst

    k = 0
tot = snap("TotalAmount")
EDAmt = 0

    rs.MoveLast
rs.MoveFirst
   For i = 1 To 14
  '''''''' For i = 1 To rs.RecordCount
      If i = 7 Then

''''     For i = 1 To 20
''''      If i = 19 Then
        If EDAmt > 0 Then
          tmp1 = RupeesInWords(EDAmt, "Rupees", "Paise")
          WordSplit tmp1, 48
        End If
        If Not rs.EOF Then
          Printer.Print Tab(2); WSplit(1); Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        End If
        If Not rs.EOF Then
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2); Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        End If
      ElseIf i >= 10 And i <= 11 Then
     '''ElseIf i >= 9 And i <= 11 Then
        tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
        WordSplit tmp, 48
        If i = 10 Then
          Printer.Print Tab(2); WSplit(1);
        Else
          If UBound(WSplit) >= 2 Then
            Printer.Print Tab(2); WSplit(2);
          End If
        End If
        If Not rs.EOF Then
          If rs("ExprName") = "GRAND TOTAL" Then
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          rs.MoveNext
        End If
      ElseIf i Mod 2 = 1 Then
        If Not rs.EOF Then
          If rs("ExprName") = "GRAND TOTAL" Then
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          Else
            Printer.Print Tab(52); rs("ExprName"); Tab(77 - Len(Format(rs("ResultValue"), "0.00"))); Format(rs("ResultValue"), "0.00")
          End If
          If Left(rs("ExprName"), 11) = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
          If Left(rs("ExprName"), 4) = "CESS 2%" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
          rs.MoveNext
        Else
          Printer.Print
        End If
      Else
        Printer.Print
      End If
    Next i
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub UMT()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim Sql As String
  Dim tot As Double
  Dim A(1 To 20) As String
  Dim CurDate As Date
  
  Try

  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If
  
  CurDate = ServerDate

  Printer.PaperSize = vbPRPSFanfoldUS
  Printer.FontName = "Draft 10cpi"
  If Printer.FontName <> "Draft 10cpi" Then
    Printer.FontName = "Roman 10cpi"
    If Printer.FontName <> "Roman 10cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If

  Sql = " Select  a.InvoiceType,a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, f.UOMDesc, f.UOmDecimal, "
  Sql = Sql + " a.InvoiceQty, a.Rate, a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull(d.ReferenceNo,i.ReferenceNo) RefNo, isNull(d.ReferenceDate,i.ReferenceDate) RefDate,"
  Sql = Sql + " a.AInformation1,a.Type, a.InsurancePolicyNo, a.InsurancePolicyDate, a.InsuranceFor, a.InsuranceSlNo from "
  Sql = Sql + " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  Sql = Sql + " left outer join dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') d on d.OrderType = a.OrderType "
  Sql = Sql + " and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate  and d.Branchid = a.Branchid"
  Sql = Sql + " and d.Compid = a.Compid and d.finyear = a.Finyear left outer join dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') i"
  Sql = Sql + " on i.OrderType = a.OrderType and i.OrderNo = a.OrderNo and i.OrderDate = a.OrderDate"
  Sql = Sql + " and i.Branchid = a.Branchid and i.Compid = a.Compid and i.finyear = a.Finyear"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
  
  Sql = "select CompanyName, Address1, Address2, City, Pincode, Phone, Fax, EMail,Web, LST, LSTDate, Cst,CstDate,AreaCode "
  Sql = Sql & " from company_master where compid='" & gCompID & "' and BranchID='" & gBranchID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Trim$(snap("InvoiceType")) = "LI" Then
    Dumrow 6
    Printer.Print Space(15) & LAlign(50, rs("Address1"))
    Printer.Print Space(15) & LAlign(50, rs("Address2"))
    Dumrow 6
    Printer.Print Space(67) & LAlign(10, snap("InvoiceType") & snap("InvoiceNo"))
    Printer.Print Space(4) & LAlign(50, snap("PartyName"))
    Printer.Print Space(4) & LAlign(50, snap("add1")) & Space(13) & LAlign(10, Format(snap("InvoiceDate"), "dd/mm/yyyy"))
    Printer.Print Space(4) & LAlign(50, IIf(IsNull(snap("add2")), "", snap("add2")))
    Printer.Print Space(4) & LAlign(50, IIf(IsNull(snap("add3")), "", snap("add3")))
    Printer.Print Space(4) & LAlign(50, snap("city"))
    Printer.Print Space(4) & LAlign(50, snap("PinCode")) & Space(13) & LAlign(20, IIf(IsNull(snap("Refno")), "", snap("Refno") & " & " & snap("RefDate")))
    Dumrow 2
    Printer.Print Space(67) & LAlign(20, IIf(IsNull(snap("Mode")), "", snap("Mode")))
    Dumrow 4
    tot = 0
k = 0

    snap.MoveLast
snap.MoveFirst
    For i = 1 To snap.RecordCount
      Printer.Print Space(4) & LAlign(4, CStr(i)) & Space(1) & LAlign(12, snap("PartNo")) & Space(1) & LAlign(31, snap("ItemDescription")) & Space(1) & RAlign(7, snap("InvoiceQty")) & Space(1) & RAlign(12, Format(snap("Rate"), "0.00")) & Space(1) & RAlign(13, Format(snap("GoodsValue"), "0.00"))
      Printer.Print
      tot = tot + Format(snap("GoodsValue"), "0.00")
      k = k + 2
      snap.MoveNext
    Next i
    
    i = 0
    For i = k To 40
      Printer.Print
    Next i
    Printer.Print Space(75) & RAlign(13, Format(tot, "0.00"))
    Printer.Print
    Printer.Print
    
    tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
    WordSplit tmp, 45
    Printer.Print Space(43) & LAlign(45, WSplit(1))
    For i = 2 To UBound(WSplit)
      Printer.Print Space(43) & LAlign(45, WSplit(i))
    Next i
  ElseIf Trim$(snap("InvoiceType")) = "INV" Then
    Dumrow 6
    Printer.Print Space(78) & LAlign(10, snap("InvoiceType") & snap("InvoiceNo"))
    Printer.Print
    Printer.Print Space(15) & LAlign(50, rs("Address1"))
    Printer.Print Space(15) & LAlign(50, rs("Address2")) & Space(13) & LAlign(10, Format(snap("InvoiceDate"), "dd/mm/yyyy"))
    Dumrow 2
    Printer.Print Space(78) & LAlign(10, Format(snap("InvoiceTime"), "HH:mm"))
    Dumrow 2
    Printer.Print Space(78) & LAlign(10, Format(snap("RemovalTime"), "HH:mm"))
    Dumrow 2
    Printer.Print Space(26) & LAlign(43, snap("Commodity"))
    Dumrow 3
    Printer.Print Space(26) & LAlign(19, snap("TariffNo"))
    Printer.Print
    Printer.Print Space(75) & LAlign(10, snap("Mode"))
    Printer.Print Space(33) & LAlign(27, IIf(IsNull(snap("Refno")), "", snap("Refno") & " & " & snap("RefDate")))
    Printer.Print Space(3) & LAlign(29, snap("PartyName")) & Space(39) & LAlign(10, snap("VehicleNo"))
    Printer.Print Space(3) & LAlign(29, snap("add1"))
    Printer.Print Space(3) & LAlign(29, IIf(IsNull(snap("add2")), "", snap("add2")))
    Printer.Print Space(3) & LAlign(29, IIf(IsNull(snap("add3")), "", snap("add3"))) & Space(6) & LAlign(30, IIf(IsNull(snap("LST")), "", snap("LST")))
    Printer.Print Space(3) & LAlign(29, snap("city") & " - " & snap("PinCode")) & Space(6) & LAlign(30, IIf(IsNull(snap("CST")), "", snap("CST")))
    Dumrow 3
    tot = 0
k = 0

    snap.MoveLast
snap.MoveFirst
    For i = 1 To snap.RecordCount
      Printer.Print Space(3) & LAlign(13, snap("PartNo")) & Space(1) & LAlign(22, snap("ItemDescription")) & Space(1) & RAlign(6, snap("InvoiceQty")) & Space(1) & LAlign(4, snap("UOMDesc")) & Space(1) & RAlign(8, Format(snap("AssessableRate"), "0.00")) & Space(1) & RAlign(9, Format(snap("Rate"), "0.00")) & Space(1) & RAlign(8, Format(snap("AssessableGoods"), "0.00")) & Space(1) & RAlign(8, Format(snap("GoodsValue"), "0.00"))
      Printer.Print
      tot = tot + Format(snap("GoodsValue"), "0.00")
      k = k + 2
      snap.MoveNext
    Next i
    
    i = 0
    For i = k To 16
      Printer.Print
    Next i
    
    For i = 1 To 20
      A(i) = Space(89)
    Next i
    
    snap.MoveFirst
    
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " order by a.ExprSeq"
    If RSOpen(rs, Sql) = False Then Exit Sub
    If Not rs.EOF Then
      rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        A(i) = Space(60) & LAlign(19, CheckNull(rs("ExprName"))) & Space(1) & RAlign(8, Format(rs("Resultvalue"), "0.00"))
        If CheckNull(rs("ExprName")) = "EXCISE DUTY" Or CheckNull(rs("ExprName")) = "CENVAT" Then
 EDAmt = Format(rs("Resultvalue"), "0.00")
End If
        If CheckNull(rs("ExprName")) = "CESS" Or CheckNull(rs("ExprName")) = "CESS ON CENVAT" Then
 CessAmt = Format(rs("Resultvalue"), "0.00")
End If
        If CheckNull(rs("ExprName")) = "Total" Then
 tot = Format(rs("Resultvalue"), "0.00")
End If
        rs.MoveNext
      Next i
    End If
    A(1) = Space(41) & LAlign(10, CheckNull(snap("InsurancePolicyNo"))) & Right(A(1), 37)
    A(4) = Space(41) & LAlign(10, CheckNull(snap("InsurancePolicyDate"))) & Right(A(4), 37)
    A(6) = Space(41) & LAlign(10, CheckNull(snap("InsuranceFor"))) & Right(A(6), 37)
    A(8) = Space(41) & LAlign(10, CheckNull(snap("InsuranceSlNo"))) & Right(A(8), 37)
    
    If EDAmt > 0 Then
      tmp = RupeesInWords(Format(EDAmt, "0.00"), "Rupees", "Paise")
      WordSplit tmp, 38
      For i = 1 To UBound(WSplit)
        A(8 + i) = Space(10) & LAlign(38, WSplit(i)) & Right(A(8 + i), 37)
      Next i
    End If
    
    If CessAmt > 0 Then
      tmp = RupeesInWords(Format(CessAmt, "0.00"), "Rupees", "Paise")
      WordSplit tmp, 36
      For i = 1 To UBound(WSplit)
        A(12 + i) = Space(12) & LAlign(36, WSplit(i)) & Right(A(12 + i), 37)
      Next i
    End If
    
    If tot > 0 Then
      tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
      WordSplit tmp, 38
      For i = 1 To UBound(WSplit)
        A(16 + i) = Space(10) & LAlign(38, WSplit(i)) & Right(A(16 + i), 37)
      Next i
    End If
    
    For i = 1 To 20
      Printer.Print A(i)
    Next i
  End If
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub PMI()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim snap2 As AceADODB.Recordset
  Dim snap3 As AceADODB.Recordset
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim CurDate As Date
  
  Try
 
  CurDate = ServerDate
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
  Set snap1 = New AceADODB.Recordset
  Set snap2 = New AceADODB.Recordset
  Set snap3 = New AceADODB.Recordset
       
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Invoice to Print"
    GoTo ExitHere
  End If
  
  k = 0
  iPageNo = 1
  iPageWidth = 80
  pheight = 86
  
  Printer.PaperSize = vbPRPSA4
  Printer.FontName = "Draft 10cpi"
  If Printer.FontName <> "Draft 10cpi" Then
    Printer.FontName = "Roman 10cpi"
    If Printer.FontName <> "Roman 10cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If
  
  Sql = "Select a.InvoiceType, a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode, "
  Sql = Sql + " a.ItemCode,  b.PartNo, c.ItemDescription, c.UOMDesc, c.UOMDecimal,a.InvoiceQty , a.Rate,  "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue, "
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, "
  Sql = Sql + " a.OrderType, a.OrderNo,a.OrderDate, b.ReferenceNo, b.ReferenceDate,"
  Sql = Sql + " d.DeliveryDesc, d.ModeDesc, d.PaymentDesc, e.CarriageName,"
  Sql = Sql + " e.CarriageNo,e.DischargePort,e.ReceivingPlace,e.LoadingPort,e.DestinationPlace"
  Sql = Sql + " from  SalesInvoice a left outer join dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b"
  Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate"
  Sql = Sql + " and a.ItemCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
  Sql = Sql + " and a.Finyear = b.Finyear left outer join Material_Master c on a.ItemCode = c.ItemCode"
  Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join"
  Sql = Sql + " dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') d "
  Sql = Sql + " on a.OrderNo = d.OrderNo and a.OrderDate = d.OrderDate and a.Branchid = d.Branchid  "
  Sql = Sql + " and a.Compid = d.Compid and a.Finyear = d.Finyear left outer join SAL_PackingList e "
  Sql = Sql + " on a.InvoiceType = e.OrderType and a.InvoiceNo = e.OrderNo and a.InvoiceDate = e.OrderDate and a.Branchid = e.Branchid  "
  Sql = Sql + " and a.Compid = e.Compid and a.Finyear = e.Finyear"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = "select CompanyName, Address1, Address2, City, Pincode, Phone, Fax, EMail,Web, LST, LSTDate, Cst,CstDate,AreaCode "
  Sql = Sql & " from company_master where compid='" & gCompID & "' and BranchID='" & gBranchID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = "select e.PartyName, f.Add1, f.Add2, f.Add3, f.City, f.Pincode"
  Sql = Sql + " from FAS_Party_Master e left outer join FAS_Party_Address f on e.PartyCode = f.PartyCode "
  Sql = Sql + " and e.Branchid = f.Branchid and e.Compid = f.Compid where e.PartyCode = " & Val(snap("CustomerCode")) & ""
  Sql = Sql + " and e.BranchID='" & gBranchID & "' and e.Compid='" & gCompID & "'"
  If RSOpen(snap1, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = "select ContainerNo, PackageKind, Sum(Quantity) Quantity, sum(NetWeight) NetWeight, Length,Width, Height "
  Sql = Sql + " from SAL_PackingList_Goods a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.OrderType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.OrderNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql + "group by ContainerNo,PackageKind,Length,Width, Height"
  If RSOpen(snap2, Sql) = False Then
    GoTo ExitHere
  End If
  
  Sql = " select a.Name,a.Address1,a.Address2,a.Address3,a.City,a.PinCode,a.State,a.Country"
  Sql = Sql + " from SAL_PackingList_Address a"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.OrderType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.OrderNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap3, Sql) = False Then
    GoTo ExitHere
  End If

  Dumrow 5
  k = k + 5
RowProcess
  Printer.FontBold = True
  Printer.Print Space(1) & CAlign(iPageWidth, "INVOICE")
  Printer.FontBold = False
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(16, "Invoice No.    :") & Space(1) & LAlign(10, snap("InvoiceType") & snap("InvoiceNo")) & Space(25) & LAlign(17, "Invoice Date    :") & Space(1) & LAlign(10, snap("InvoiceDate"))
  Printer.Print Space(1) & LAlign(16, "Customer PO No.:") & Space(1) & LAlign(10, snap("ReferenceNo")) & Space(25) & LAlign(17, "Customer PO Date:") & Space(1) & LAlign(10, snap("ReferenceDate"))
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(33, "Exporter:") & Space(10) & LAlign(40, "Buyer:")
  Printer.Print Space(1) & LAlign(33, gCompName) & Space(10) & LAlign(40, snap1("PartyName"))
  Printer.Print Space(1) & LAlign(33, IIf(rs("Address1") = "", "", rs("Address1"))) & Space(10) & LAlign(40, IIf(snap1("Add1") = "", "", snap1("Add1")))
  Printer.Print Space(1) & LAlign(33, IIf(rs("Address2") = "", "", rs("Address2"))) & Space(10) & LAlign(40, IIf(snap1("Add2") = "", "", snap1("Add2")))
  Printer.Print Space(1) & LAlign(25, IIf(rs("City") = "", "", rs("City")) + IIf(rs("pincode") = "", "", "-" + rs("Pincode") + ".")) & Space(18) & LAlign(35, snap1("Add3") + ", " + snap1("city") + IIf(snap1("Pincode") = "", "", "-" + snap1("Pincode") + "."))
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & LAlign(26, "Pre-Carriage by          :") & Space(1) & LAlign(25, snap("CarriageName")) & Space(1) & LAlign(30, "Consignee(if other than Buyer")
  Printer.Print Space(1) & LAlign(26, "Place of Receipt by pre- :") & Space(1) & LAlign(25, snap("ReceivingPlace")) & Space(1) & LAlign(30, IIf(IsNull(snap3("Name")), "", snap3("Name")))
  Printer.Print Space(1) & LAlign(26, "Vessel / Flight No       :") & Space(1) & LAlign(25, snap("CarriageNo")) & Space(1) & LAlign(30, IIf(IsNull(snap3("Address1")), "", snap3("Address1")))
  Printer.Print Space(1) & LAlign(26, "Port of Loading          :") & Space(1) & LAlign(25, snap("LoadingPort")) & Space(1) & LAlign(30, IIf(IsNull(snap3("Address2")), "", snap3("Address2")))
  Printer.Print Space(1) & LAlign(26, "Port of Discharge        :") & Space(1) & LAlign(25, snap("DischargePort")) & Space(1) & LAlign(30, IIf(IsNull(snap3("Address3")), "", snap3("Address3")))
  Printer.Print Space(1) & LAlign(26, "Final Distination        :") & Space(1) & LAlign(25, snap("DestinationPlace")) & Space(1) & LAlign(30, IIf(IsNull(snap3("City")), "", snap3("City")))
  Printer.Print Space(1) & LAlign(26, "Country of Origin        :") & Space(1) & LAlign(25, "India") & Space(1) & LAlign(30, IIf(IsNull(snap3("PinCode")), "", snap3("PinCode")))
  Printer.Print Space(1) & LAlign(26, "Country of Destination   :") & Space(1) & LAlign(25, "USA") & Space(1) & LAlign(30, IIf(IsNull(snap3("State")), "", snap3("State")))
  Printer.Print Space(1) & LAlign(26, "Terms of Delivery        :") & Space(1) & LAlign(25, snap("DeliveryDesc")) & Space(1) & LAlign(30, IIf(IsNull(snap3("Country")), "", snap3("Country")))
  Printer.Print Space(1) & LAlign(26, "Mode                     :") & Space(1) & LAlign(25, snap("ModeDesc"))
  Printer.Print Space(1) & LAlign(26, "Payment Terms            :") & Space(1) & LAlign(25, snap("PaymentDesc"))
  Printer.Print Space(1) & LAlign(26, "Marks & Nos./Container No:") & Space(1) & LAlign(25, snap2("ContainerNo"))
  Printer.Print Space(1) & LAlign(26, "No. & Kind of Package    :") & Space(1) & LAlign(25, snap2("PackageKind"))
  
  Printer.Print Space(1) & String(iPageWidth, "-")
  k = k + 25
RowProcess
  
  Call sPageHeaderNew
RowProcess
  
  snap.MoveFirst
  For i = 1 To snap.RecordCount
    Printer.Print Space(1) & LAlign(5, CStr(i)) & Space(1) & LAlign(38, snap("ItemDescription")) & Space(1) & RAlign(8, snap("InvoiceQty")) & Space(1) & CAlign(4, snap("UOMDesc")) & Space(1) & RAlign(10, Format(snap("Rate"), "0.00")) & Space(1) & RAlign(10, Format(snap("GoodsValue"), "0.00"))
    Printer.Print Space(1) & Space(6) & LAlign(38, snap("PartNo"))
    Printer.Print
    tot = tot + Format(snap("GoodsValue"), "0.00")
    k = k + 3
RowProcess
    snap.MoveNext
  Next i
  Printer.Print Space(1) & String(iPageWidth, "-")
  Printer.Print Space(1) & Space(6) & LAlign(38, "Total") & Space(29) & Format(tot, "0.00")
  Printer.Print Space(1) & String(iPageWidth, "-")
  tmp1 = RupeesInWords(tot, "USD", "Cents")
  WordSplit tmp1, 70
  Printer.Print Space(1) & LAlign(15, "Total in Words:") & Space(1) & WSplit(1)
k = k + 1
RowProcess
  For X = 2 To UBound(WSplit)
    Printer.Print Space(1) & Space(17) & WSplit(X)
k = k + 1
RowProcess
  Next X
  Printer.Print Space(1) & String(iPageWidth, "-")
  
  If Not snap2.EOF Then
    Printer.Print Space(1) & LAlign(16, "I.E.C No     :") & Space(1) & LAlign(10, "3805000057") & Space(20) & LAlign(30, "Our Banker's Name and Address"): k = k + 1: RowProcess

    Printer.Print Space(1) & LAlign(16, "Total Qty    :") & Space(1) & RAlign(10, snap2("Quantity")) & Space(20) & LAlign(30, "INDIAN BANK MEPZ Branch")
k = k + 1
RowProcess

    Printer.Print Space(1) & LAlign(16, "Total Pkg    :") & Space(1) & RAlign(10, snap2("ContainerNo")) & Space(20) & LAlign(30, "Tambaram, Chennai - 600 045")
k = k + 1
RowProcess

    Printer.Print Space(1) & LAlign(16, "Total Net Wt.:") & Space(1) & RAlign(10, snap2("NetWeight") & " Kgs") & Space(20) & LAlign(30, "INDIA")
k = k + 1
RowProcess

    Printer.Print Space(1) & LAlign(30, "Dim. of the Corrugated Boxes:") & Space(17) & LAlign(30, "AD Code : 02611749030666")
    Printer.Print Space(1) & Space(10) & LAlign(10, "Length") & Space(1) & LAlign(10, "Width") & Space(1) & LAlign(10, "Height")
    Printer.Print Space(1) & Space(10) & LAlign(10, snap2("Length")) & Space(1) & LAlign(10, snap2("Width")) & Space(1) & LAlign(10, snap2("Height"))
  End If
  Printer.Print Space(1) & LAlign(75, "We declare that this invoice shows the actual price of the goods described")
  Printer.Print Space(1) & LAlign(75, "and that all the particulers are true and correct")
  Printer.Print Space(1) & Space(44) & CAlign(30, "Signature & Date")
  Printer.Print Space(1) & Space(44) & CAlign(30, gCompName)
  Dumrow 2
  Printer.Print Space(1) & Space(48) & CAlign(30, "Shayam Todi")          'Hard code theDirector name according to Company.
  Printer.Print Space(1) & Space(48) & CAlign(30, "Director")
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
  Set snap1 = Nothing
  Set snap2 = Nothing
  Set snap3 = Nothing
End Try
End Sub

Private Function Validate() As Boolean
  Dim FromNo As Long        'Invoice No Lower Range
  Dim ToNo As Long          'Invoice No Upper Range
  
  Try
  
  Validate = False
  
  If chkInvoiceDate.Value = 1 Then
    If ValidateDateFromTo(dtpInvoiceFrom.Value, dtpInvoiceTo.Value, dtpInvoiceFrom.MinDate, dtpInvoiceTo.MaxDate) = False Then
      GoTo ExitHere
    End If
  End If
  
  If chkInvoiceNo.Value = 1 Then
  
    FromNo = Val(txtInvoiceFrom.Text)
    ToNo = Val(txtInvoiceTo.Text)
    
    If Len(Trim$(txtInvoiceFrom.Text)) = 0 And Len(Trim$(txtInvoiceTo.Text)) = 0 Then 'From & To No not given
      sShowMessage "Please Enter Invoice No. to restrict."
      GoTo ExitHere
    ElseIf Len(Trim$(txtInvoiceFrom.Text)) = 0 Then     'From No not given
      FromNo = Val(txtInvoiceFrom.Tag)                  'set to Min No
      txtInvoiceFrom.Text = Val(txtInvoiceFrom.Tag)
    ElseIf Len(Trim$(txtInvoiceTo.Text)) = 0 Then       'To No not given
      ToNo = Val(txtInvoiceTo.Tag)                      'set to Max No
      txtInvoiceTo.Text = Val(txtInvoiceTo.Tag)
    End If
    
    If ValidateNumberFromTo(FromNo, ToNo, Val(txtInvoiceFrom.Tag), Val(txtInvoiceTo.Tag)) = False Then
      GoTo ExitHere
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateDateFromTo(ByVal fromdate As Date,ByVal  toDate As Date,ByVal  MinDate As Date,ByVal  MaxDate As Date) As Boolean
  
  Try
  
  ValidateDateFromTo = False  'Initialize to return false when validation fails
  
    If fromdate < MinDate Then
      sShowMessage "From Date should be greater than equal to the permissible Minimum Date '" & Format(MinDate, "dd-MM-yyyy") & "'"
      GoTo ExitHere
    ElseIf fromdate > MaxDate Then
      sShowMessage "From Date should be lesser than equal to the permissible Maximum Date '" & Format(MaxDate, "dd-MM-yyyy") & "'"
      GoTo ExitHere
    ElseIf toDate < MinDate Then
      sShowMessage "To Date should be greater than equal to the permissible Minimum Date '" & Format(MinDate, "dd-MM-yyyy") & "'"
      GoTo ExitHere
    ElseIf toDate > MaxDate Then
      sShowMessage "From Date should be lesser than equal to the permissible Maximum Date '" & Format(MaxDate, "dd-MM-yyyy") & "'"
      GoTo ExitHere
    ElseIf fromdate > toDate Then
      sShowMessage "From Date should be lesser than equal to the To Date"
      GoTo ExitHere
    End If
    
  ValidateDateFromTo = True     'Validation success. so returns true
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateNumberFromTo(ByVal FromNo As Long,ByVal  ToNo As Long,ByVal  MinNo As Long,ByVal  MaxNo As Long) As Boolean
  
  Try
  
  ValidateNumberFromTo = False  'Initialize to return false when validation fails
  
    If FromNo < MinNo Then
      sShowMessage "From No should be greater than equal to the permissible Minimum No " & MinNo
      GoTo ExitHere
    ElseIf FromNo > MaxNo Then
      sShowMessage "From No should be lesser than equal to the permissible Maximum No " & MaxNo
      GoTo ExitHere
    ElseIf ToNo < MinNo Then
      sShowMessage "To No should be greater than equal to the permissible Minimum No " & MinNo
      GoTo ExitHere
    ElseIf ToNo > MaxNo Then
      sShowMessage "From No should be lesser than equal to the permissible Maximum No " & MaxNo
      GoTo ExitHere
    ElseIf FromNo > ToNo Then
      sShowMessage "From No should be lesser than equal to the To No"
      GoTo ExitHere
    End If
    
  ValidateNumberFromTo = True     'Validation success. so returns true
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub chkInvoiceDate_Click()
  If chkInvoiceDate.Value = 1 Then
    fraInvoiceDate.Visible = True
    dtpInvoiceFrom.SetFocus
  Else
    fraInvoiceDate.Visible = False
  End If
End Sub

Private Sub chkInvoiceNo_Click()
  If chkInvoiceNo.Value = 1 Then
    fraInvoiceNo.Visible = True
    txtInvoiceFrom.SetFocus
  Else
    fraInvoiceNo.Visible = False
  End If
End Sub

Private Sub chkPackingSlipPrint_Click()
  Dim retVal As Variant
  
' On Error Resume Next
  txtSavedBoxQty.Text = ""
txtBoxQty.Text = ""
  
  If chkPackingSlipPrint.Value = 1 Then
    txtBoxQty.Enabled = True
lblBoxQty.Enabled = True
fraLabel.Enabled = True
    cmdPackingSlipPrint.Enabled = True
    CmdView.Enabled = False
cmdprint.Enabled = False
    
      sprInvoice.GetText eInv.mBoxQty, sprInvoice.ActiveRow, retVal
      txtSavedBoxQty.Text = Val(retVal)
    
    If fraPackingSlipPrint.Enabled = True And txtBoxQty.Enabled = True Then
 txtBoxQty.SetFocus
End If
  Else
    txtBoxQty.Enabled = False
lblBoxQty.Enabled = False
fraLabel.Enabled = False
    cmdPackingSlipPrint.Enabled = False
    CmdView.Enabled = True
cmdprint.Enabled = True
  End If
End Sub

Private Sub cmdPackingSlipPrint_Click()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRetVal As Variant
  Dim sInvType As String
  Dim sInvNo As Long
  Dim sInvDate As Date
  Dim sCustomerCode As Long
  Dim sLblCnt As Long
  Dim sRow As Long
  Dim sCurDate As Date
  
  Try
  
    sInvType = ""
sInvNo = 0
sCustomerCode = 0
    sRow = sprInvoice.ActiveRow
  
    If chkPackingSlipPrint.Value = 1 Then
      'Validate - Box Qty
      If Val(txtBoxQty.Text) <= 0 Then
        sShowMessage "Please Enter the Box Qty."
        If fraPackingSlipPrint.Enabled = True And txtBoxQty.Enabled = True Then
 txtBoxQty.SetFocus
End If
        GoTo ExitHere
      Else
        'Continue
      End If
      
      If Val(txtSavedBoxQty.Text) > 0 And Val(txtSavedBoxQty.Text) <> Val(txtBoxQty.Text) Then
      
        If MsgBox("BoxQty and Already SavedQty Not Matched. Do you want to continue?", vbYesNo) = vbNo Then
          If fraPackingSlipPrint.Enabled = True And txtBoxQty.Enabled = True Then
 txtBoxQty.SetFocus
End If
          GoTo ExitHere
        Else
          'Continue
        End If
      Else
        'Continue
      End If
      
      If ((Val(txtSavedBoxQty.Text) <= 0 And Val(txtBoxQty.Text) > 0) Or (Val(txtBoxQty.Text) > 0 And Val(txtSavedBoxQty.Text) <> Val(txtBoxQty.Text))) Then
        'Getting Server Date
        sCurDate = ServerDate
  
        'Save PackingSlipQty
        Sql = " Execute SAL_PackingSlipQty_Dtl_Ins"
        
        sprInvoice.GetText eInv.mInvType, sRow, sRetVal
        Sql = Sql & " @DocType = '" & Trim$(sRetVal) & "',"
        
        sprInvoice.GetText eInv.mInvNo, sRow, sRetVal
        Sql = Sql & " @DocNo = " & Val(sRetVal) & ","
        
        sprInvoice.GetText eInv.mInvDate, sRow, sRetVal
        Sql = Sql & " @DocDate = '" & Format(Trim$(sRetVal), "yyyy-MM-dd") & "',"
        
        sprInvoice.GetText eInv.mCustCode, sRow, sRetVal
        Sql = Sql & " @CustomerCode = " & Val(sRetVal) & ","
        
        Sql = Sql & " @Qty = " & Val(txtBoxQty.Text) & ","
        Sql = Sql & " @EffStatus = 'AA',"
        Sql = Sql & " @EffDate = '" & Format(sCurDate, "yyyy-MM-dd") & "',"
        Sql = Sql & " @UserID = '" & gUserID & "',"
        Sql = Sql & " @BranchID = '" & gBranchID & "',"
        Sql = Sql & " @CompID = '" & gCompID & "',"
        Sql = Sql & " @Finyear = '" & gFinyear & "',"
        Sql = Sql & " @TranStamp = '" & Format(sCurDate, "yyyy-MM-dd") & "'"
        
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        
        'Set BoxQty into Spreadsheet Relevant Column
        sprInvoice.SetText eInv.mBoxQty, sRow, Val(txtBoxQty.Text)

        txtSavedBoxQty.Text = Val(txtBoxQty.Text)
      Else
        'Continue
      End If
                
      'InvType
      sprInvoice.GetText eInv.mInvType, sRow, sRetVal
sInvType = Trim$(sRetVal)
      
      'InvNo
      sprInvoice.GetText eInv.mInvNo, sRow, sRetVal
 sInvNo = Val(sRetVal)
      
      'InvDate
      sprInvoice.GetText eInv.mInvDate, sRow, sRetVal
sInvDate = Format(Trim$(sRetVal), "yyyy-MM-dd")
      
      'CustomerCode
      sprInvoice.GetText eInv.mCustCode, sRow, sRetVal
sCustomerCode = Val(sRetVal)

      'Packing Slip Label Print (DOTNET Application)
      chkAll.Tag = -2     'Start
      FetchLabelData = Set_LabelPosition(False, True)
      sLblCnt = chkPosition.count
      
      If Trim$(FetchLabelData) = "" Or Len(FetchLabelData) < sLblCnt Then
        FetchLabelData = FetchLabelData & String((sLblCnt - Len(FetchLabelData)), "0")
      Else
        'Continue
      End If
      
      chkAll.Tag = 0      'End
  
      If PrintNetDocuments(eNetReportTypes.PackingSlipPrint, sInvType, sInvNo, sInvDate, Trim$(FetchLabelData)) = False Then
        GoTo ExitHere
      Else
        'Continue
      End If
      
'      If PrintNet_SalesInvoicePackingSlipLabel(eNetReportTypes.PackingSlipPrint, fInvType, fInvNo, fInvDate, fCustomerCode, fCustomerName, fSupplierName, '                                              fPartName, fPartNo, fInvQty, fLot_HeatNo, fPart_PacketQty, fBoxQty) = False Then

'        GoTo ExitHere
'       Else
'        'Continue
'      End If
                                            
    Else
      'Continue
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmdClear_Click()
  Dim CurDate As Date
  
  Try
    
  'Getting Server Date
  CurDate = ServerDate

  'Assigning Server Date
  dtpAsOnDate.Value = Format(CurDate, "dd/MM/yyyy HH:mm:ss")
  
  chkInvoiceDate.Value = 0
  dtpInvoiceFrom.Value = dtpInvoiceFrom.MinDate
  dtpInvoiceTo.Value = dtpInvoiceTo.MaxDate

  chkInvoiceNo.Value = 0
  txtInvoiceFrom.Text = ""
  txtInvoiceTo.Text = ""
  sprInvoice.MaxRows = 1
  sprInvoice.ClearRange 1, 1, -1, -1, True
  sprInvoice.Enabled = False
  
  chkPackingSlipPrint.Value = 0
'  txtBoxQty.Text = "":
  
  chkAll.Tag = -2
  chkAll.Value = 1
  Call Set_LabelPosition(True, False)
  chkAll.Tag = 0
  FetchLabelData = ""
  fraPackingSlipPrint.Enabled = False
chkPackingSlipPrint.Enabled = False

  CmdView.Enabled = False
cmdprint.Enabled = False

'  chkPackingSlipPrint.Enabled = False: lblBoxQty.Enabled = False

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdExcel_Click()
Dim X As Boolean
  Dim Temp As Variant
  
  
  Try
  Me.MousePointer = vbHourglass

  Temp = Year(ServerDate) & Month(ServerDate) & Day(ServerDate) & ".xls"
  X = sprView.ExportToExcel(App.Path & "\" & Temp, "Test Sheet 1", "C:Program Files\Spread30\Samples\LOGFILE.TXT")
  ' Display result to user based on true/false value of x
  If X = True Then
    MsgBox "Export complete.", , "Result"
  Else
    MsgBox "Export did not succeed.", , "Result"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit1_Click()

  
  Try
  Me.MousePointer = vbHourglass

  fraView.Visible = False
  FraPlaceHolder.Visible = True
  fraView.Top = -90
  fraView.Left = 90
  fraView.Height = 675
  fraView.Width = 1035
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdProceed_Click()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  
  Try
  
  If Validate = False Then
    GoTo ExitHere
  Else
    'Continue
  End If
  
'  Sql = "Select a.InvoiceType [Inv. Type], a.InvoiceNo [Inv. No], a.InvoiceDate [Inv. Date], b.PartyName [Customer], a.CustomerCode, y.OrderFlag "
'  Sql = Sql & " from SAL_INVOICE_HDR a Left Outer Join FAS_PARTY_MASTER b On a.CustomerCode = b.PartyCode and a.CompId=b.Compid"
'  Sql = Sql & " left Outer join Sal_PackingList y On a.InvoiceType=y.OrderType"
'  Sql = Sql & " and a.InvoiceNo=y.OrderNo and a.InvoiceDate=y.OrderDate"
'  Sql = Sql & " and a.Branchid=y.Branchid and a.Compid=y.Compid"
'  Sql = Sql & " where a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"

'  If chkInvoiceDate.Value = 1 Then
'    Sql = Sql & " and a.InvoiceDate between '" & Format(dtpInvoiceFrom, "yyyy-MM-dd") & "' and '" & Format(dtpInvoiceTo, "yyyy-MM-dd") & "'"
'  End If
'  If chkInvoiceNo.Value = 1 Then
'    Sql = Sql & " and a.InvoiceNo between " & Val(Trim$(txtInvoiceFrom.Text)) & " and " & Val(Trim$(txtInvoiceTo.Text))
'  End If
'
'  Sql = Sql & " Order by a.InvoiceNo Desc"
  
  Sql = " Execute SAL_InvoicePrint_Fil"
  
  If chkInvoiceDate.Value = 1 Then
    Sql = Sql & " @IsInvDate = '1',"
    Sql = Sql & " @FromInvDate = '" & Format(dtpInvoiceFrom.Value, "yyyy-MM-dd") & "',"
    
    If Format(dtpInvoiceTo.Value, "yyyy-MM-dd") < Format(dtpInvoiceFrom.Value, "yyyy-MM-dd") Then
      Sql = Sql & " @ToInvDate = NULL,"
    Else
      Sql = Sql & " @ToInvDate = '" & Format(dtpInvoiceTo.Value, "yyyy-MM-dd") & "',"
    End If
  Else
    Sql = Sql & " @IsInvDate = '0',"
    Sql = Sql & " @FromInvDate = NULL,"
    Sql = Sql & " @ToInvDate = NULL,"
  End If
  
  If chkInvoiceNo.Value = 1 Then
    Sql = Sql & " @IsInvNo = '1',"
    Sql = Sql & " @FromInvNo = " & Val(txtInvoiceFrom.Text) & ","
    
    If Val(txtInvoiceTo.Text) < Val(txtInvoiceFrom.Text) Then
      Sql = Sql & " @ToInvNo = 0,"
    Else
       Sql = Sql & " @ToInvNo = " & Val(txtInvoiceTo.Text) & ","
    End If
  Else
    Sql = Sql & " @IsInvNo = '0',"
    Sql = Sql & " @FromInvNo = 0,"
    Sql = Sql & " @ToInvNo = 0,"
  End If
   
  Sql = Sql & " @AsOnDate = '" & Format(dtpAsOnDate.Value, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @RptDate = '" & Format(dtpAsOnDate.Value, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @BranchID = '" & gBranchID & "', @CompID = '" & gCompID & "', @Finyear = '" & gFinyear & "'"
    
'  Call sprInvoice_Click(1, 1)
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  Else
    sprInvoice.Enabled = True
sprInvoice.Visible = True
    sprInvoice.ClearSelection
    
    Call ADORS_Spread_DataFill(rs, sprInvoice, True)
    
    Call Spread_HideCol(sprInvoice, True, eInv.mCustCode, eInv.mBoxQty)
    sprInvoice.MaxRows = sprInvoice.DataRowCnt
    Call Spread_Set_Focus(sprInvoice, 1, 1)
  End If
  
  fraPackingSlipPrint.Enabled = True
chkPackingSlipPrint.Enabled = True
  chkPackingSlipPrint.Value = 0
  chkAll.Tag = -2     'Start
  chkAll.Value = 1
  Call Set_LabelPosition(True, False)
  chkAll.Tag = 0      'End
  CmdView.Enabled = True
cmdprint.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  If Trim$(gReportID) = "unicorn" Then
      sprInvoice.Col = eInv.mInvDate
sprInvoice.Row = sprInvoice.ActiveRow
    
    If Format(Trim$(sprInvoice.Text), "yyyy-MM-dd") > Format(gGSTStartDate, "yyyy-MM-dd") Then
      If NET_Unicorn_Reports = False Then
        GoTo ExitHere
      End If
    Else
      fraView.Visible = True
      
      If chkcancel = 0 Then
        Call Infant_ViewPrint
      Else
        Call Infant_CancelViewPrint
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function NET_Unicorn_Reports() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  Dim fInvType As String
  Dim fInvNo As Long
  Dim fInvDate As Date
  Dim fCustomerCode As Long
  Dim fInvCount As Long
  Dim fRow As Long
  
  Try
  
  NET_Unicorn_Reports = False
    
    fRow = sprInvoice.ActiveRow
    
    ' CustomerCode
    sprInvoice.GetText eInv.mCustCode, fRow, retVal
fCustomerCode = Val(retVal)
      
    ' InvType
    sprInvoice.GetText eInv.mInvType, fRow, retVal
fInvType = Trim$(retVal)
    
    ' InvNo
    sprInvoice.GetText eInv.mInvNo, fRow, retVal
 fInvNo = Val(retVal)
    
    ' InvDate
    sprInvoice.GetText eInv.mInvDate, fRow, retVal
fInvDate = Format(Trim$(retVal), "yyyy-MM-dd")
  
  fInvCount = 6     'Only 6 Invoice can taken for Users
  
  If gUserID = "A0007" Or gUserID = "G0024" Then
    'No change
  Else
    Sql = "Select Count(*) as cnt from SAL_InvoicePrint_Dtl where InvoiceType='" & Trim$(fInvType) & "'"
    Sql = Sql & " and Invoiceno=" & Val(fInvNo) & " and InvoiceDate='" & Format(fInvDate, "yyyy-MM-dd") & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    
    If Val(rs("cnt")) >= fInvCount Then  'Only 6 Invoice can taken for Users
      sShowMessage "Already Printed this Invoice"
      GoTo ExitHere
    End If
  End If
    
  'Print RDLC Format
'  If IsShowRdlcReport = True Then
    NET_Unicorn_Reports = PrintNetDocuments(eNetReportTypes.mInvoicePrint, fInvType, fInvNo, fInvDate)
    
    Sql = " Insert into SAL_InvoicePrint_Dtl Select Distinct '" & fInvType & "'," & Val(fInvNo) & ",'" & Format(fInvDate, "yyyy-MM-dd") & "','" & Format(ServerDate, "yyyy-MM-dd HH:mm:ss") & "','" & gUserID & "','" & gBranchID & "','" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    GoTo ExitHere
'  Else
'    'Continue
'  End If
  
  NET_Unicorn_Reports = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function Infant_CancelViewPrint() As Boolean
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim tot As Double
  Dim EccNo As String
  Dim Rule As String
  Dim Place As String             'Postal Address Range and Division printing
  Dim Range As String
  Dim Division As String
  Dim City As String
  
  Try
  
  Infant_CancelViewPrint = False
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If sprInvoice.DataRowCnt = 0 Then
    sShowMessage "No Record's Found"
    GoTo ExitHere
  End If

  Sql = " Select a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount, a.CustomerCode,"
  Sql = Sql + " (select b.PartyName from FAS_PArty_Master b where b.PartyCode = a.CustomerCode"
  Sql = Sql + " and b.Branchid = a.Branchid and b.Compid = a.Compid) PartyName, a.ModeCode, "
  Sql = Sql + " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql + " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql + " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql + " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql + " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql + " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql + " a.AssessableRate, (a.InvoiceQty*a.Rate) GoodsValue,"
  Sql = Sql + " (a.InvoiceQty*a.AssessableRate) AssessableGoods, a.NoofPackets, a.OrderType,"
  Sql = Sql + " a.OrderNo,a.OrderDate, e.Add1, e.Add2,e.Add3,e.City,e.Pincode, e.LST, "
  Sql = Sql + " e.CST, e.ECCNo, isNull((select top 1 ReferenceNo from SAL_ORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "'), (select top 1 ReferenceNo from SAL_JOBWORKORDER_HDR d where "
  Sql = Sql + " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql + " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql + " and d.finyear = '" & gFinyear & "')) RefNo, a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql + " Sales_CancelInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql + " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql + " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql + " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  sprInvoice.Col = 1
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " where a.InvoiceType = '" + Trim$(sprInvoice.Text) + "'"
  sprInvoice.Col = 2
sprInvoice.Row = sprInvoice.ActiveRow
  Sql = Sql + " and a.InvoiceNo = " & Trim$(sprInvoice.Text) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = " select (Address1+', '+Address2) as add1, (City+'-'+Pincode)as Add2,"
  Sql = Sql & " ('TNGST No. '+ LST+' CST No. : ' + Cst +' Dt : '+ convert(varchar(10), CstDate, 105)+' '+'AreaCode No.: '+AreaCode)as ApprovalDet,"
  Sql = Sql & " NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
    sprView.SetText ecol.mCol1, 2, gCompName
    sprView.SetText ecol.mcol7, 3, CStr(snap("InvoiceNo"))
    sprView.SetText ecol.mcol7, 6, Format(snap("InvoiceDate"), "dd/mm/yyyy")
    sprView.SetText ecol.mCol1, 9, Space(3) & snap("PartyName")
    sprView.SetText ecol.mCol1, 10, Space(3) & snap("add1")
    sprView.SetText ecol.mCol1, 11, Space(3) & IIf(IsNull(snap("add2")), "", snap("add2"))
    sprView.SetText ecol.mCol1, 12, Space(3) & IIf(IsNull(snap("add3")), "", snap("add3")) & Space(2) & snap("city")
    sprView.SetText ecol.mCol1, 13, Space(3) & CStr(snap("Pincode"))
    sprView.SetText ecol.mcol4, 9, snap("CustomerCode")
    sprView.SetText ecol.mcol6, 9, snap("ECCNo")
    sprView.SetText ecol.mcol8, 9, IIf(IsNull(snap("LST")), "", snap("LST"))
    sprView.SetText ecol.mcol9, 9, IIf(IsNull(snap("CST")), "", snap("CST"))
    sprView.SetText ecol.mcol4, 11, IIf(IsNull(snap("RefNo")), "", snap("RefNo"))
    sprView.SetText ecol.mcol4, 13, snap("InvoiceNo") & " & " & Format(snap("InvoiceDate"), "dd/mm/yyyy")
    sprView.SetText ecol.mcol4, 15, Left(snap("mode"), 66) & Space(2) & snap("VehicleNo")
'    Place = "OFFICE OF THE SUPERINTENDENT OF CENTRAL EXCISE"
'    Range = "RANGE  - Poonamallee - V"
'    Division = "Division - Poonamallee, Commissionerate - IV  Chennai."
'    .SetText ecol.mCol1, 15, Place
'    .SetText ecol.mCol1, 16, Range
'    .SetText ecol.mCol1, 17, Division
    sprView.SetText ecol.mcol5, 19, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("InvoiceTime"), "HH:mm")
    sprView.SetText ecol.mcol8, 19, Format(snap("InvoiceDate"), "dd/mm/yy") & " " & Format(snap("RemovalTime"), "HH:mm")
    sprView.SetText ecol.mCol1, 21, snap("Commodity")
    sprView.SetText ecol.mcol4, 21, snap("TariffNo")
    sprView.SetText ecol.mcol7, 21, rs("NotificationNo") & "   " & rs("NotificationDate")
    sprView.Col = ecol.mCol1
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mCol1, 23, 1 & "."
    sprView.Col = ecol.mCol2
sprView.Row = 23
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mCol2, 23, snap("ItemDescription")
    sprView.Col = ecol.mcol5
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol5, 23, snap("NoofPackets")
    sprView.Col = ecol.mcol6
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol6, 23, Format(snap("InvoiceQty"), "0")
    sprView.Col = ecol.mcol7
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol7, 23, Format(snap("Rate"), "0.00")
    sprView.Col = ecol.mcol8
sprView.Row = 23
sprView.TypeVAlign = TypeVAlignCenter
sprView.TypeHAlign = TypeHAlignCenter
sprView.BackColor = RGB(255, 255, 0)
    sprView.SetText ecol.mcol8, 23, Format(snap("GoodsValue"), "0.00")
    For X = 1 To sprView.MaxCols
      sprView.Col = X
sprView.Row = 23
sprView.BackColor = RGB(255, 255, 0)

      sprView.Col = X
sprView.Row = 24
sprView.BackColor = RGB(255, 255, 0)
    Next X
    sprView.SetText ecol.mCol2, 24, IIf(IsNull(snap("PartNo")), snap("ItemCode"), snap("PartNo"))
  
    If Trim$(snap("Type")) = "Supplementary" Then
      sprView.SetText ecol.mCol1, 32, "SUPPLEMENTARY INVOICE"
    ElseIf Trim(snap("Type")) = "Labour" Then
      sprView.SetText ecol.mCol1, 32, "LABOUR CHARGES ONLY"
    ElseIf Trim(snap("Type")) = "CANCEL" Then
      Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
      Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If Not rs.EOF Then
        sprView.SetText ecol.mCol1, 32, "CANCELED INVOICE"
      Else
        sprView.SetText ecol.mCol1, 32, ""
      End If
    ElseIf Trim(snap("Type")) = "EOU" Then
      Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
      Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where ExprName like 'VAT%'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If Not rs.EOF Then
        sprView.SetText ecol.mCol1, 32, "VAT INVOICE"
      Else
        sprView.SetText ecol.mCol1, 32, ""
      End If
    End If
    sprView.SetText ecol.mcol9, 34, Format(snap("GoodsValue"), "0.00")
  
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Cancel ('" & Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 1)) & "'," & Val(SpVal(sprInvoice, sprInvoice.ActiveRow, 2)) & ", '" & Format(Trim$(SpVal(sprInvoice, sprInvoice.ActiveRow, 3)), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " order by a.ExprSeq"
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      snap.MoveLast
snap.MoveFirst

      k = 0
tot = snap("TotalAmount")
EDAmt = 0
X = 1

      rs.MoveLast
rs.MoveFirst
      sprView.ClearRange ecol.mcol8, 36, sprView.MaxCols, 44, True
      sprView.ClearRange ecol.mCol1, 42, 1, 44, True
      For i = 1 To 14
        If i = 7 Then
          If Not rs.EOF Then
            sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
            sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            rs.MoveNext
            X = X + 1
          End If
          If Not rs.EOF Then
            tmp1 = RupeesInWords(rs("ResultValue"), "Rupees", "Paise")
            sprView.SetText ecol.mCol1, 42, tmp1
            sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
            sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            rs.MoveNext
            X = X + 1
          End If
        ElseIf i >= 10 And i <= 11 Then
          tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")

          If Not rs.EOF Then
            If rs("ExprName") = "GRAND TOTAL" Then
              sprView.SetText ecol.mCol1, 44, tmp
              sprView.SetText ecol.mcol8, 44, rs("ExprName")
              sprView.SetText ecol.mcol9, 44, Format(rs("ResultValue"), "0.00")
            ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
              sprView.SetText ecol.mcol8, X + 35, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 35, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "VAT" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "CST" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            Else
              sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
            End If
            rs.MoveNext
          End If
        ElseIf i Mod 2 = 1 Then
          If Not rs.EOF Then
            tmp = RupeesInWords(Format(tot, "0.00"), "Rupees", "Paise")
            If CheckNull(rs("ExprName")) = "GRAND TOTAL" Then
              sprView.SetText ecol.mCol1, 44, tmp
              sprView.SetText ecol.mcol8, 44, rs("ExprName")
              sprView.SetText ecol.mcol9, 44, Format(rs("ResultValue"), "0.00")
            ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
              sprView.SetText ecol.mcol8, X + 35, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 35, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "VAT" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            ElseIf CheckNull(rs("ExprName")) = "CST" Then
              sprView.SetText ecol.mcol8, 43, rs("ExprName")
              sprView.SetText ecol.mcol9, 43, Format(rs("ResultValue"), "0.00")
              X = X + 1
            Else
              sprView.SetText ecol.mcol8, X + 36, rs("ExprName")
              sprView.SetText ecol.mcol9, X + 36, Format(rs("ResultValue"), "0.00")
              X = X + 1
            End If
            If rs("ExprName") = "EXCISE DUTY" Or rs("ExprName") = "CENVAT CLAIMS" Then
 EDAmt = Format(rs("ResultValue"), "0.00")
End If
            If rs("ExprName") = "CESS" Then
 EDAmt = EDAmt + Format(rs("ResultValue"), "0.00")
End If
            rs.MoveNext
          End If
        End If
      Next i
    End If
    
  FraPlaceHolder.Visible = False
  fraView.Visible = True
  fraView.Height = 8070
  fraView.Width = 11595
  fraView.Left = 90
  fraView.Top = -90
'  Call Spread_View(sprView)

  Infant_CancelViewPrint = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Function

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
    
  'Getting Server Date
  CurDate = ServerDate
 
  'Assigning Server Date
  dtpAsOnDate.Value = Format(CurDate, "dd/MM/yyyy HH:mm:ss")
  
  Call InitializeForm
  
  sprInvoice.MaxRows = 1
  
  If gUserID = "A0007" Then
    chkcancel.Visible = True
  Else
    chkcancel.Visible = False
  End If
  
  chkAll.Tag = -2
  chkAll.Value = 1
  Call Set_LabelPosition(True, False)
  chkAll.Tag = 0
  FetchLabelData = ""
  txtSavedBoxQty.Text = ""
txtBoxQty.Text = ""
  fraView.Visible = False
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

Private Sub sprInvoice_Click(sender as object, e as ClickEventArgs)
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  Dim fRow As Long
  
  Try
  Me.MousePointer = vbHourglass

  fRow = Row
  
    Spread_BkColour sprInvoice, vbWhite, 1, 1, 4, sprInvoice.DataRowCnt
    Spread_BkColour sprInvoice, vbYellow, 1, fRow, 4, fRow
'    sprInvoice.Col = 4: sprInvoice.Row = fRow

    sprInvoice.GetText eInv.mExport, fRow, retVal
    If Trim$(retVal) = "I" Then
      fraPack.Visible = True
    Else
      fraPack.Visible = False
      optInvoice.Value = True
    End If
    
    'BoxQty
    If chkPackingSlipPrint.Value = 1 Then
      sprInvoice.GetText eInv.mBoxQty, fRow, retVal
      txtSavedBoxQty.Text = Val(retVal)
      txtBoxQty.Text = ""
      If fraPackingSlipPrint.Enabled = True And txtBoxQty.Enabled = True Then
 txtBoxQty.SetFocus
End If
    Else
      txtSavedBoxQty.Text = ""
txtBoxQty.Text = ""
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = 0
End Try
End Sub

'Private Sub sprInvoice_Click(sender as object, e as ClickEventArgs)
'  Dim snap As ADODB.Recordset
'  Dim Sql As String
'  Dim pRetVal As Variant
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  With sprInvoice
'    'Get PackingSlip - Qty Data
'    Sql = " Execute SAL_PackingSlipQty_Dtl_Fil"
'
'    Sql = Sql & " @ModuleCode = " & gModuleCode & ","
'
'    .GetText eInv.mInvType, fRow, pRetVal
'    Sql = Sql & " @DocType = '" & Trim$(pRetVal) & "',"
'
'    .GetText eInv.mInvNo, fRow, pRetVal
'    Sql = Sql & " @DocNo = " & Val(pRetVal) & ","
'
'    .GetText eInv.mInvDate, fRow, pRetVal
'    Sql = Sql & " @DocDate = '" & Format(Trim$(pRetVal), "yyyy-MM-dd") & "',"
'
'    .GetText eInv.mCustCode, fRow, pRetVal
'    Sql = Sql & " @CustomerCode = " & Val(pRetVal) & ","
'
'    Sql = Sql & " @AsOnDate = '" & Format(dtpAsOnDate.Value, "yyyy-MM-dd HH:mm:ss") & "',"
'    Sql = Sql & " @RptDate = '" & Format(dtpAsOnDate.Value, "yyyy-MM-dd HH:mm:ss") & "',"
'    Sql = Sql & " @BranchID = '" & gBranchID & "', @CompID = '" & gCompID & "', @Finyear = '" & gFinyear & "'"
'
'    Set snap = New ADODB.Recordset
'    If RSOpen(snap, Sql) = False Then: GoTo ExitHere
'
'    If Not rs.EOF() Then
'      'InvQty
'      txtExistQty.Text = Val(rs("Qty"))
'    Else
'      'Continue
'    End If
'
'
'ExitHere:
'  On Error GoTo 0
'  Set snap = Nothing
'  Me.MousePointer = 0
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub
Private Sub cmdprint_Click()
  Dim OrdType As String
  Dim ORDNo As Integer
  Dim OrdDate As Date
  Dim Custcode As Integer
  Dim retVal As Variant
  Dim Sql As String
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass
  
    sprInvoice.GetText 1, sprInvoice.ActiveRow, retVal
     OrdType = Trim$(retVal)
    sprInvoice.GetText 2, sprInvoice.ActiveRow, retVal
     ORDNo = retVal
    sprInvoice.GetText 3, sprInvoice.ActiveRow, retVal
     OrdDate = Format(retVal, "yyyy-MM-dd")
    sprInvoice.GetText 4, sprInvoice.ActiveRow, retVal
     Custcode = Val(retVal)
    
  If gUserID = "A0007" Or gUserID = "G0024" Then
    'No change
  Else
    Sql = "Select Count(*) as cnt from SAL_InvoicePrint_Dtl where InvoiceType='" & Trim$(OrdType) & "'"
    Sql = Sql & " and Invoiceno=" & Val(ORDNo) & " and InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    
    If Val(rs("cnt")) >= 2 Then
      sShowMessage "Already Printed This Invoice"
      GoTo ExitHere
    End If
  End If
  
  If optInvoice.Value = True Then
      Call PrintInvoice
  ElseIf optPack.Value = True Then
    Call PrintPackingList(OrdType, ORDNo, OrdDate, Custcode)
  Else    'For Both Option
    Call PrintInvoice
    Call PrintPackingList(OrdType, ORDNo, OrdDate, Custcode)
  End If
  
  Sql = "Insert into SAL_InvoicePrint_Dtl Select Distinct '" & OrdType & "'," & Val(ORDNo) & ",'" & Format(OrdDate, "yyyy-MM-dd") & "','" & Format(ServerDate, "yyyy-MM-dd HH:mm:ss") & "','" & gUserID & "','" & gBranchID & "','" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub PrintInvoice()
' On Error Resume Next
  
  Select Case gReportID
    Case "INFANT"
      If chkcancel = 0 Then
        Call INFANT
      Else
        Call INFANT_CancelInvoice
      End If
      
    Case "UNICORN"
      Call Unicorn
      
    Case "DWARAKA"
      Call Dwaraka
      
    Case "PMI"
      Call PMI
      
    Case "UMT"
      Call UMT
  End Select
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub InitializeForm()
  Call FindFillMinMaxValue(anByNone)
End Sub

Private Sub FindFillMinMaxValue(Optional ByVal ChooseBy As anFindBy = anFindBy.anByNone,Optional ByVal fromdate As Date,Optional ByVal toDate As Date,Optional ByVal FromNo As Long,Optional ByVal ToNo As Long)
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  
  Try

  Sql = "Select Max(InvoiceDate) as MaxDt, Min(InvoiceDate) as MinDt, Max(InvoiceNo) as MaxNo, Min(InvoiceNo) as MinNo from SAL_INVOICE_HDR" & " where Compid='" & gCompID & "' and finyear='" & gFinyear & "'"

  Select Case ChooseBy
    Case anFindBy.anByBoth
      Sql = Sql & " and InvoiceDate between Convert(datetime, '" & fromdate & "', 120) and Convert(datetime, '" & toDate & "', 120)" & " and InvoiceNo between " & FromNo & " and " & ToNo

    Case anFindBy.anByDate
      Sql = Sql & " and InvoiceDate between Convert(datetime, '" & fromdate & "', 120) and Convert(datetime, '" & toDate & "', 120)"
    Case anFindBy.anByNo
      Sql = Sql & " and InvoiceNo between " & FromNo & " and " & ToNo
  End Select
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    Select Case ChooseBy
      Case anFindBy.anByNone
        dtpInvoiceFrom.MinDate = rs("MinDt")
        dtpInvoiceFrom.MaxDate = rs("MaxDt")
        dtpInvoiceFrom.Value = rs("MaxDt")
        dtpInvoiceFrom.Tag = rs("MinDt")
        
        dtpInvoiceTo.MinDate = rs("MinDt")
        dtpInvoiceTo.MaxDate = rs("MaxDt")
        dtpInvoiceTo.Value = rs("MaxDt")
        dtpInvoiceTo.Tag = rs("MaxDt")
        
        txtInvoiceFrom.Tag = rs("MinNo")
        txtInvoiceTo.Tag = rs("MaxNo")
    End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtBoxQty_KeyPress(ByVal KeyAscii As Integer)

  Try
'  Me.MousePointer = vbHourglass
    
  KeyAscii = NumValKeyDownPress(txtBoxQty, KeyAscii, True, 0, 1000000000, 0, 0)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

'  Me.MousePointer = vbDefault
End Try
End Sub
Private Sub txtInvoiceFrom_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = NumValInt(txtInvoiceFrom, KeyAscii, Val(txtInvoiceTo.Tag), Val(txtInvoiceFrom.Tag))
End Sub

Private Sub txtInvoiceTo_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = NumValInt(txtInvoiceTo, KeyAscii, Val(txtInvoiceTo.Tag), Val(txtInvoiceFrom.Tag))
End Sub

Private Sub chkAll_Click()
' On Error Resume Next
  If chkAll.Tag = -1 Or chkAll.Tag = -2 Then
 Exit Sub
End If
  
  chkAll.Tag = 1
  Call Set_LabelPosition(True, False)
  chkAll.Tag = 0
End Sub

Private Sub chkPosition_Click(ByVal Index As Integer)
' On Error Resume Next
  If chkAll.Tag = 1 Or chkAll.Tag = -2 Then
 Exit Sub
End If
  
  chkAll.Tag = -1
  Call Set_LabelPosition(False, False)
  chkAll.Tag = 0
End Sub

Private Function Set_LabelPosition(ByVal isAll As Boolean,ByVal  IsLebelCnt As Boolean) As String
  Dim fRetStr As String
  Dim fIsAllPos As Boolean
  Dim Idx As Long
    
' On Error Resume Next
  Set_LabelPosition = ""
  fRetStr = ""

  If IsLebelCnt = True Then
    For Idx = 0 To chkPosition.count - 1
      If chkPosition(Idx).Value = 0 Then
        fRetStr = fRetStr & (Idx + 1)         'Label Print Hide (No Need)
      Else
        fRetStr = fRetStr & "0"               'Label Print Show (Need)
      End If
    Next Idx
  Else
    If isAll = True Then      'All
      For Idx = 0 To chkPosition.count - 1
        chkPosition(Idx).Value = IIf(chkAll.Value = 1, 1, 0)
      Next Idx
    Else                      'Selective
      fIsAllPos = True
      
      For Idx = 0 To chkPosition.count - 1
        If chkPosition(Idx).Value = 0 Then
          fIsAllPos = False
          Exit For
        Else
          'Continue
        End If
      Next Idx
      
      If fIsAllPos = True Then
        chkAll.Value = 1
      Else
        chkAll.Value = 0
      End If
    End If
  End If
  
  'Return
  Set_LabelPosition = fRetStr
End Function

Private Sub Design_sprView(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprView

End Sub

Private Sub Design_sprInvoice(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprInvoice


sprInvoice.Click +=  new EventHandler(sprInvoice_Click)
sprInvoice.Click +=  new EventHandler(sprInvoice_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprView("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InvoicePrint.frm", sprView, "")

	Call Design_sprInvoice("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InvoicePrint.frm", sprInvoice, "")

End Sub
