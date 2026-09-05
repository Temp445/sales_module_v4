VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesRegisterColumnarReport 
   Caption         =   "Sales Register Columnar Report "
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13080
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
   ScaleHeight     =   8955
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   15
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton cmdexit1 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9480
         TabIndex        =   36
         Top             =   7215
         Width           =   1300
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8190
         TabIndex        =   35
         Top             =   7215
         Width           =   1300
      End
Begin AceSpread spview
         Height          =   300
         Left            =   60
         TabIndex        =   37
         Top             =   150
         Width           =   630
         _Version        =   458752
         _ExtentX        =   1111
         _ExtentY        =   529
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
         MaxCols         =   11
         MaxRows         =   501
      End
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   0
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
      Begin VB.CommandButton cmdexit2 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9600
         TabIndex        =   33
         Top             =   7110
         Width           =   1300
      End
      Begin VB.CommandButton cmdprint1 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8310
         TabIndex        =   32
         Top             =   7110
         Width           =   1300
      End
Begin AceSpread spviewsum
         Height          =   360
         Left            =   270
         TabIndex        =   31
         Top             =   150
         Width           =   735
         _Version        =   458752
         _ExtentX        =   1296
         _ExtentY        =   635
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
         MaxCols         =   4
         MaxRows         =   501
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer RR"
      Height          =   585
      Index           =   4
      Left            =   4140
      TabIndex        =   29
      Top             =   2955
      Width           =   2685
      Begin VB.OptionButton optRRAll 
         Caption         =   "All"
         Height          =   225
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optWRR 
         Caption         =   "with RR"
         Height          =   225
         Left            =   630
         TabIndex        =   12
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton optWORR 
         Caption         =   "Without RR"
         Height          =   225
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Type Of Invoice"
      Enabled         =   0   'False
      Height          =   690
      Index           =   3
      Left            =   4140
      TabIndex        =   28
      Top             =   2265
      Width           =   2685
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "SalesRegisterColumnarReport.frx":6B474
         Left            =   90
         List            =   "SalesRegisterColumnarReport.frx":6B496
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   2505
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4140
      TabIndex        =   25
      Top             =   360
      Width           =   2685
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108527617
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108527617
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   120
         TabIndex        =   27
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1395
         TabIndex        =   26
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   4140
      TabIndex        =   24
      Top             =   1125
      Width           =   2685
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   255
         Left            =   1605
         TabIndex        =   4
         Top             =   225
         Width           =   960
      End
   End
   Begin VB.Frame fraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4140
      TabIndex        =   23
      Top             =   3555
      Width           =   2685
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "SalesRegisterColumnarReport.frx":6B4FB
         Left            =   90
         List            =   "SalesRegisterColumnarReport.frx":6B505
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   210
         Width           =   2505
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Products"
      Enabled         =   0   'False
      Height          =   585
      Index           =   2
      Left            =   4140
      TabIndex        =   22
      Top             =   1680
      Width           =   2685
      Begin VB.OptionButton OptPAll 
         Caption         =   "All"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton OptPrdrandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1590
         TabIndex        =   7
         Top             =   225
         Width           =   960
      End
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
      Index           =   3
      Left            =   6915
      TabIndex        =   10
      Top             =   2460
      Width           =   420
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
      Left            =   6915
      TabIndex        =   5
      Top             =   1260
      Width           =   420
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
      Left            =   6915
      TabIndex        =   2
      Top             =   615
      Width           =   420
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
      Index           =   4
      Left            =   6915
      TabIndex        =   14
      Top             =   3105
      Width           =   420
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
      Index           =   2
      Left            =   6915
      TabIndex        =   8
      Top             =   1830
      Width           =   420
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6120
      TabIndex        =   19
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4832
      TabIndex        =   18
      Top             =   4590
      Width           =   1300
   End
Begin AceSpread spCust
      Height          =   1290
      Left            =   7485
      TabIndex        =   20
      Top             =   1215
      Visible         =   0   'False
      Width           =   4185
      _Version        =   458752
      _ExtentX        =   7382
      _ExtentY        =   2275
      _StockProps     =   64
      DAutoSizeCols   =   1
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
      OperationMode   =   4
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SelectBlockOptions=   0
   End
Begin AceSpread spItem
      Height          =   1545
      Left            =   7485
      TabIndex        =   21
      Top             =   2550
      Visible         =   0   'False
      Width           =   4185
      _Version        =   458752
      _ExtentX        =   7382
      _ExtentY        =   2725
      _StockProps     =   64
      DAutoSizeCols   =   1
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
      OperationMode   =   4
      RowHeaderDisplay=   0
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   3570
      TabIndex        =   17
      Top             =   4590
      Width           =   1275
   End
End
Attribute VB_Name = "rptSalesRegisterColumnarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sprow As Long

Private Erpno As String
Private TsNo As String

Public Enum eview
  mInvNos = 1
  mItem = 2
  mUOM = 3
  mQty = 4
  mgval = 5
  med = 6
  mcess = 7
  mscess = 8
  msubtot = 9
  mvat = 10
  mInvAmt = 11
End Enum

Dim j As Long

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
  Dim Sql As String
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim sRow As Long
  Dim tmp As String
  Dim tmp1 As String
  Dim iType As String
  Dim iNo As Integer
  Dim iDate As Date
  Dim INVType As String
  Dim tot As Double
  Dim Gtot As Double
  Dim Gtot1 As Double
  Dim etot As Double
  Dim ctot As Double
  Dim Cnt As Integer
  Dim Spr As AceSpread
  
  Dim rsDetails As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
 
  Try
  
  spview.Tag = 0
  Sql = " select a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql & " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount,"
  Sql = Sql & " a.ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode "
  Sql = Sql & " and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) ItemDescription,"
  Sql = Sql & " (select c.RevisionNo from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode "
  Sql = Sql & " and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) RevisionNo,"
  Sql = Sql & " a.InvoiceQty,a.Rate,a.AssessableRate, a.Type,"
  Sql = Sql & " (select c.Uomdesc from MATERIAL_MASTER c where c.ItemCode=a.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid "
  Sql = Sql & " and c.status = 1)Uom From SalesInvoice a "
  Sql = Sql & " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and InvoiceDate BETWEEN convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  If cmbType <> "All" Then
    Sql = Sql & " and a.type = '" & Trim$(cmbType) & "'"
  End If
  If optWRR Then
    Sql = Sql & " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql & " and a.RRNo=''"
  End If
  
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Customer"
GoTo ExitHere
End If
    Sql = Sql & " and CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql & "'" & Trim$(spCust.Text) & "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql & ")"
  End If
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.ItemCode in("
      sRow = 0
      For i = 1 To spItem.SelectionCount
        sRow = spItem.GetMultiSelItem(sRow)
        spItem.Row = sRow
spItem.Col = 1
        Sql = Sql + "'" + Trim$(spItem.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  Sql = Sql & " order by PartyName,InvoiceDate,InvoiceNo"
  
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_sPrintcurrent = False
    GoTo ExitHere
  End If
  
  Set Spr = frmPrint.sprPrint
              
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Sales Register Columnar"
  
  Call Header
      
  Do While Not rsDetails.EOF
    If tmp <> (rsDetails("InvoiceType") & rsDetails("InvoiceNo")) Then
      If tmp <> "" Then
        Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
        Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim(iType) & "', " & Val(iNo) & ",'" & Format(iDate, "yyyy-MM-dd") & "', " & IIf(Trim$(INVType) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        Sql = Sql & " order by a.ExprSeq"
        If RSOpen(snap, Sql) = False Then GoTo ExitHere
        If Not snap.EOF Then
          tot = 0
          snap.MoveLast
snap.MoveFirst
          For j = 1 To snap.RecordCount
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
'            frmPrint.sprPrint.AddCellSpan 2, SpRow, 2, 1
            spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(snap("ExprName")), False
'            frmPrint.sprPrint.AddCellSpan 3, SpRow, 2, 1
            spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(snap("Resultvalue"), "0.00"), False
            sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  '          sSendToPort Space(11) & LAlign(50, CheckNull(snap("ExprName"))) & Space(10) & RAlign(9, Format(snap("Resultvalue"), "0.00"))
            snap.MoveNext
          Next j
        End If
      End If
'      SpRow = SpRow + 1: GridCount = GridCount + 1: RowProcess Me, SpRow
      If tmp1 <> rsDetails("PartyName") Then
          If optRandom = True Then
'            If i > 1 Then: SpRow = SpRow + 1: GridCount = GridCount + 1: RowProcess Me, SpRow
          End If
          frmPrint.sprPrint.AddCellSpan 1, sprow, 4, 1
          spLAlign frmPrint.sprPrint, 1, sprow, 15, "Customer :" & rsDetails("PartyName"), True
          sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
'          sSendToPort LAlign(10, "Customer :") & Space(1) & LAlign(50, CheckNull(rsDetails("PartyName")))
          
        End If
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        If GridCount = 53 Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        End If
        spLAlign frmPrint.sprPrint, 1, sprow, 15, rsDetails("InvoiceType") & rsDetails("InvoiceNo"), False
        spLAlign frmPrint.sprPrint, 2, sprow, 40, rsDetails("ItemCode"), False
        spCAlign frmPrint.sprPrint, 3, sprow, 5, rsDetails("UOM"), False
        spRAlign frmPrint.sprPrint, 4, sprow, 10, rsDetails("InvoiceQty"), False
        spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(rsDetails("Rate"), "0.00"), False
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        
        spLAlign frmPrint.sprPrint, 1, sprow, 15, rsDetails("InvoiceDate"), False
        spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(rsDetails("ItemDescription")), False
        spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(rsDetails("AssessableRate"), "0.00"), False
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    Else
      spLAlign frmPrint.sprPrint, 2, sprow, 40, rsDetails("ItemCode"), False
      spCAlign frmPrint.sprPrint, 3, sprow, 5, rsDetails("UOM"), False
      spRAlign frmPrint.sprPrint, 4, sprow, 10, rsDetails("InvoiceQty"), False
      spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(rsDetails("Rate"), "0.00"), False
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      
      spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(rsDetails("ItemDescription")), False
      spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(rsDetails("AssessableRate"), "0.00"), False
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
    End If
    tmp = (rsDetails("InvoiceType") & rsDetails("InvoiceNo"))
    iType = rsDetails("InvoiceType")
    iNo = rsDetails("InvoiceNo")
    iDate = rsDetails("InvoiceDate")
    INVType = rsDetails("Type")
    tmp1 = rsDetails("PartyName")
    rsDetails.MoveNext
  Loop
  If rsDetails.EOF Then
  
    If tmp <> "" Then
        Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
        Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim(iType) & "', " & Val(iNo) & ",'" & Format(iDate, "yyyy-MM-dd") & "', " & IIf(Trim$(INVType) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        Sql = Sql & " order by a.ExprSeq"
        If RSOpen(snap, Sql) = False Then GoTo ExitHere
        If Not snap.EOF Then
          tot = 0
          snap.MoveLast
snap.MoveFirst
          For j = 1 To snap.RecordCount
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
'            frmPrint.sprPrint.AddCellSpan 2, SpRow, 2, 1
            spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(snap("ExprName")), False
'            frmPrint.sprPrint.AddCellSpan 3, SpRow, 2, 1
            spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(snap("Resultvalue"), "0.00"), False
            sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  '          sSendToPort Space(11) & LAlign(50, CheckNull(snap("ExprName"))) & Space(10) & RAlign(9, Format(snap("Resultvalue"), "0.00"))
            snap.MoveNext
          Next j
        End If
      End If
    Sql = " Select a.ExprName, Sum(ResultValue) as ResultValue, a.GrandTotalFlag from "
    Sql = Sql & " SalesCalculationDetails a"
    Sql = Sql & " right Outer Join SalesInvoice b On b.InvoiceType = a.DocType "
    Sql = Sql & " and b.InvoiceNo = a.DocNo and b.InvoiceDate = a.DocDate "
    Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER c"
    Sql = Sql & " On c.PartyCode = b.CustomerCode and c.Branchid = b.Branchid "
    Sql = Sql & " and c.Compid = b.Compid where a.GrandTotalFlag <> 0"
    Sql = Sql & " and a.DocDate Between  convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111) "
    Sql = Sql & " and  convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111) "
    If optRandom.Value = True Then
      Sql = Sql & " and b.CustomerCode in("
        sRow = 0
        For i = 1 To spCust.SelectionCount
          sRow = spCust.GetMultiSelItem(sRow)
          spCust.Row = sRow
spCust.Col = 1
          Sql = Sql & "'" & Trim$(spCust.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
    End If
    If OptPrdrandom Then
      If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and b.ItemCode in("
        sRow = 0
        For i = 1 To spItem.SelectionCount
          sRow = spItem.GetMultiSelItem(sRow)
          spItem.Row = sRow
spItem.Col = 1
          Sql = Sql + "'" + Trim$(spItem.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)
      Sql = Sql + ")"
    End If
    If cmbType <> "All" Then
      Sql = Sql & " and b.type = '" & Trim$(cmbType) & "'"
    End If
    If optWRR Then
      Sql = Sql & " and b.RRNo <> ''"
    ElseIf optWORR Then
      Sql = Sql & " and b.RRNo = ''"
    End If
    Sql = Sql & " Group By a.ExprName, a.GrandTotalFlag,a.exprseq order by a.exprseq"
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
    
    If Not snap.EOF Then
       snap.MoveLast
snap.MoveFirst
       Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
       frmPrint.sprPrint.AddCellSpan 1, sprow, 2, 1
       spLAlign frmPrint.sprPrint, 1, sprow, 15, "Cumulative Total", True
       sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      tot = 0
      For j = 1 To snap.RecordCount
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(snap("ExprName")), False
        spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(snap("Resultvalue"), "0.00"), False
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        tot = tot + Format(snap("ResultValue"), "0.00")
        snap.MoveNext
      Next j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      spLAlign frmPrint.sprPrint, 2, sprow, 40, "Invoice Amount", False
      spRAlign frmPrint.sprPrint, 5, sprow, 10, Format(CStr(tot), "0.00"), False
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
  End If
  totalpages = iPageNo
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  frmPrint.Show
  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
  Set snap = Nothing
End Try
End Function
Public Sub Header()
  
  Try
  
  If SumColrFlag <> "CR" Then
    
    Call RptStandard_Header(4, sprow, Erpno, TsNo, chkPrintdate)
    GridCount = GridCount + 6
  
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    
    If optRandom = True Then
      spCAlign frmPrint.sprPrint, 1, sprow, 10, "S.No", True
      spCAlign frmPrint.sprPrint, 2, sprow, 15, "Invoice No", True
      spCAlign frmPrint.sprPrint, 3, sprow, 10, "Invoice Date", True
      spCAlign frmPrint.sprPrint, 4, sprow, 20, "Invoice Amount", True
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    Else
      spCAlign frmPrint.sprPrint, 1, sprow, 15, "Invoice No", True
      spCAlign frmPrint.sprPrint, 2, sprow, 10, "Invoice Date", True
      spCAlign frmPrint.sprPrint, 3, sprow, 15, "Invoice Amount", True
      spCAlign frmPrint.sprPrint, 4, sprow, 40, "Customer Name", True
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
  Else
    Call RptStandard_Header(5, sprow, Erpno, TsNo, chkPrintdate)
    GridCount = GridCount + 6
  
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    spCAlign frmPrint.sprPrint, 1, sprow, 15, "Invoice No", True
    spCAlign frmPrint.sprPrint, 2, sprow, 40, "Item Code", True
    spCAlign frmPrint.sprPrint, 3, sprow, 5, "UOM", True
    spCAlign frmPrint.sprPrint, 4, sprow, 10, "Qty", True
    spCAlign frmPrint.sprPrint, 5, sprow, 10, "Rate", True
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    spCAlign frmPrint.sprPrint, 1, sprow, 15, "Invoice Date", True
    spCAlign frmPrint.sprPrint, 2, sprow, 40, "Item Description", True
    spCAlign frmPrint.sprPrint, 5, sprow, 10, "Ass.Rate", True
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sSpread_sPrintcurrentSummary() As Boolean
  Dim Sql As String
  Dim i As Long
  Dim Cnt As Integer
  Dim sRow As Long
  Dim tamt As Double
  Dim tmp As String
  Dim subtot As Double
  Dim rsDetails As AceADODB.Recordset
  Dim Spr As AceSpread
  Dim setstr As String
  
  Try
  
  Set rsDetails = New AceADODB.Recordset
  
  Cnt = 0
  Sql = " select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount"
  Sql = Sql + " from SAL_Invoice_Hdr a "
  Sql = Sql + " Left outer join SAL_Invoice_Dtl d on a.InvoiceType=d.InvoiceType and a.InvoiceNo=d.InvoiceNo and a.InvoiceDate=d.InvoiceDate"
  Sql = Sql + " and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql + " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.InvoiceDate BETWEEN convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
  If cmbType <> "All" Then
    Sql = Sql + " and a.type = '" + Trim$(cmbType) + "'"
  End If
  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo=''"
  End If
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Customer"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql + "'" + Trim$(spCust.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atleast one item"
GoTo ExitHere
End If
    Sql = Sql & " and d.ItemCode in("
      sRow = 0
      For i = 1 To spItem.SelectionCount
        sRow = spItem.GetMultiSelItem(sRow)
        spItem.Row = sRow
spItem.Col = 1
        Sql = Sql + "'" + Trim$(spItem.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If optRandom.Value = 1 Then
    Sql = Sql + " order by PartyName,a.InvoiceDate,a.InvoiceNo"
  Else
    Sql = Sql + " order by a.InvoiceDate,a.InvoiceNo"
  End If
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_sPrintcurrentSummary = False
    GoTo ExitHere
  End If
  
  Set Spr = frmPrint.sprPrint
              
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Sales Register Summary"
  
  Call Header
  
    i = 1
    Do While Not rsDetails.EOF
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      If optRandom = True Then
        If tmp <> rsDetails("PartyName") Then
          If i > 1 Then
            frmPrint.sprPrint.AddCellSpan 1, sprow, 3, 1
            spRAlign frmPrint.sprPrint, 1, sprow, 10, "Sub Total", True
            spRAlign frmPrint.sprPrint, 4, sprow, 14, Format(subtot, "0.00"), True
            subtot = 0
            sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
          End If
          frmPrint.sprPrint.AddCellSpan 1, sprow, 4, 1
          spLAlign frmPrint.sprPrint, 1, sprow, 50, "Customer Name :" & rsDetails("PartyName"), True
          sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        End If
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        spLAlign frmPrint.sprPrint, 1, sprow, 10, Val(i)
        spLAlign frmPrint.sprPrint, 2, sprow, 15, CheckNull(rsDetails("InvoiceType") & " -" & rsDetails("InvoiceNo"))
        spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("InvoiceDate"))
        spRAlign frmPrint.sprPrint, 4, sprow, 20, Format(rsDetails("TotalAmount"), "0.00")
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Else
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
        spLAlign frmPrint.sprPrint, 1, sprow, 10, CheckNull(rsDetails("InvoiceType") & " -" & rsDetails("InvoiceNo"))
        spLAlign frmPrint.sprPrint, 2, sprow, 10, CheckNull(rsDetails("InvoiceDate"))
        spRAlign frmPrint.sprPrint, 3, sprow, 20, Format(rsDetails("TotalAmount"), "0.00")
        spLAlign frmPrint.sprPrint, 4, sprow, 40, rsDetails("PartyName")
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
'        sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00")) & Space(3) & LAlign(39, rsDetails("PartyName"))
      End If
      
      tamt = tamt + rsDetails("TotalAmount")
      subtot = subtot + rsDetails("TotalAmount")
      i = i + 1
      tmp = rsDetails("PartyName")
      rsDetails.MoveNext
      If rsDetails.EOF Then
        If optRandom = False Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
          frmPrint.sprPrint.AddCellSpan 1, sprow, 2, 1
          spRAlign frmPrint.sprPrint, 1, sprow, 10, "Total", True
          spRAlign frmPrint.sprPrint, 3, sprow, 14, Format(subtot, "0.00"), True
        Else
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
          frmPrint.sprPrint.AddCellSpan 1, sprow, 3, 1
          spRAlign frmPrint.sprPrint, 1, sprow, 10, "Sub Total", True
          spRAlign frmPrint.sprPrint, 4, sprow, 14, Format(subtot, "0.00"), True
          sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
          frmPrint.sprPrint.AddCellSpan 1, sprow, 3, 1
          spRAlign frmPrint.sprPrint, 1, sprow, 10, "Total", True
          spRAlign frmPrint.sprPrint, 4, sprow, 14, Format(tamt, "0.00"), True
        End If
      End If
    Loop
  
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  totalpages = iPageNo
  frmPrint.Show
  
  sSpread_sPrintcurrentSummary = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

'Private Sub Header()
'  Dim mstr As String
'
'  On Error GoTo ErrHandler
'
'  With frmPrint.sprPrint
'
'    .MaxCols = 6
'
'    .AddCellSpan 1, j, 3, 1
'    .SetText 1, j, "Page Number :" + str(iPageNo)
'
'    .AddCellSpan 4, j, .MaxCols, 1
'    .Row = j: .Col = 4

'    If chkPrintDate.Value = 1 Then: .SetText 4, j, "Printed Date: " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss"): .TypeHAlign = TypeHAlignRight

'    j = j + 1: GridCount = GridCount + 1
'
'    Spread_ERPnoPrint frmPrint.sprPrint, 1, j
'    j = j + 1: GridCount = GridCount + 1
'
'    If iPageNo = 1 Then
'      .Row = j: .Col = 1
'      .AddCellSpan 1, j, .MaxCols, 1
'      .SetText 1, j, gCompName: .FontBold = True: .FontSize = 12: .FontUnderline = True: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: j = j + 2: GridCount = GridCount + 2
'    End If
'
'    .Row = j: .Col = 1
'    .AddCellSpan 1, j, .MaxCols, 1
'    .SetText 1, j, "Sales Register between " & dpfromdt & " and " & dptodt: .Font = "Arial": .FontBold = True: .FontSize = 10: .FontUnderline = True: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: j = j + 2: GridCount = GridCount + 2
'
'    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, .MaxCols, j
'    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
'
'    .Row = j: .Col = 1

'    .SetText 1, j, "Inv No": .ColWidth(1) = 10: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 2

'    .SetText 2, j, "Item Code": .ColWidth(2) = 30: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 3

'    .SetText 3, j, "Inv Qty": .ColWidth(3) = 10: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 4

'    .SetText 4, j, "Received": .ColWidth(4) = 10: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 5

'    .SetText 5, j, "Accepted": .ColWidth(5) = 10: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 6

'    .SetText 6, j, "RR No": .ColWidth(6) = 10: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
'    j = j + 1: GridCount = GridCount + 1
'
'    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, .MaxCols, j
'    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, .MaxCols, j
'
'    .Row = j: .Col = 1

'    .SetText 1, j, "Inv Date": .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 2

'    .SetText 2, j, "Item Description": .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 3

'    .SetText 3, j, "Qty": .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 4

'    .SetText 4, j, "Qty": .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 5

'    .SetText 5, j, "Qty": .ColWidth(9) = 8: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    .Row = j: .Col = 6

'    .SetText 6, j, "Date": .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = True
'
'    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
'  End With
'
'ExitHere:
'  On Error GoTo 0
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

'Private Sub RowProcess()
'  On Error Resume Next
'  If GridCount >= pheight Then
'    With frmPrint.sprPrint
'      GridCount = 1
'      .MaxRows = .MaxRows + pheight + 1
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
'      .Row = .Row + 1: j = j + 1
'      .RowPageBreak = 1
'      iPageNo = iPageNo + 1
'      Header
'      j = j + 1: GridCount = GridCount + 1
'    End With
'  End If
'End Sub

Private Function sPrintcurrentSummary() As Boolean
  Dim Sql As String
  Dim i As Long
  Dim Cnt As Integer
  Dim sRow As Long
  Dim tamt As Double
  Dim tmp As String
  Dim subtot As Double
  Dim rsDetails As AceADODB.Recordset
  
  Try
  
  Set rsDetails = New AceADODB.Recordset
  
  Cnt = 0
  Sql = " select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount"
  Sql = Sql + " from SAL_Invoice_Hdr a "
  Sql = Sql + " Left outer join SAL_Invoice_Dtl d on a.InvoiceType=d.InvoiceType and a.InvoiceNo=d.InvoiceNo and a.InvoiceDate=d.InvoiceDate"
  Sql = Sql + " and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql + " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.InvoiceDate BETWEEN convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
  If cmbType <> "All" Then
    Sql = Sql + " and a.type = '" + Trim$(cmbType) + "'"
  End If
  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo=''"
  End If
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Customer"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql + "'" + Trim$(spCust.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and d.ItemCode in("
      sRow = 0
      For i = 1 To spItem.SelectionCount
        sRow = spItem.GetMultiSelItem(sRow)
        spItem.Row = sRow
spItem.Col = 1
        Sql = Sql + "'" + Trim$(spItem.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  If optRandom.Value = 1 Then
    Sql = Sql + " order by PartyName,a.InvoiceDate,a.InvoiceNo"
  Else
    Sql = Sql + " order by a.InvoiceDate,a.InvoiceNo"
  End If
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sPrintcurrentSummary = False
    GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  fOpenPort Me

  tamt = 0
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  gStrReportHeading = "Sales Register(Summary) between " & dpfromdt & " and " & dptodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  If optWRR Then
    sSendToPort CAlign(iPageWidth, "With RR")
  ElseIf optWORR Then
    sSendToPort CAlign(iPageWidth, "Without RR")
  End If
  sSendToPort ""
  sPageHeader
  i = 1
  Do While Not rsDetails.EOF
    If optRandom = True Then
      If tmp <> rsDetails("PartyName") Then
        If subtot > 0 Then
 sSendToPort LAlign(10, "Sub Total") & Space(14) & RAlign(14, Format(subtot, "0.00"))
End If
        subtot = 0
        If i > 1 Then
 sSendToPort String(iPageWidth, "-")
sSendToPort ""
End If
        sSendToPort LAlign(10, "Customer :") & Space(1) & LAlign(50, rsDetails("PartyName"))
        sSendToPort ""
      End If
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00"))
    Else
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00")) & Space(3) & LAlign(39, rsDetails("PartyName"))
    End If
    
    tamt = tamt + rsDetails("TotalAmount")
    subtot = subtot + rsDetails("TotalAmount")
    i = i + 1
    tmp = rsDetails("PartyName")
    rsDetails.MoveNext
  Loop
  If optRandom = True Then
    If rsDetails.EOF Then
      sSendToPort LAlign(10, "Sub Total") & Space(14) & RAlign(14, Format(subtot, "0.00"))
    End If
  End If
  sSendToPort ""
  sSendToPort LAlign(10, "Total") & Space(14) & RAlign(14, Format(tamt, "0.00"))
  sSendToPort ""
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  
  sPrintcurrentSummary = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
  Call StepClear(Me)
  Call StepEnable(Me)
  spCust.Visible = False
  spItem.Visible = False
  optAll.Value = True
  OptPAll.Value = True
  CmdStep(1).Enabled = False
  CmdStep(2).Enabled = False
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
'  Call Spread_PagePrint(spview, 5, 38, spview.MaxCols, 4, cdlLandscape)
  
End Sub

Private Sub cmdprint1_Click()
' On Error Resume Next
  
'  Call Spread_PagePrint(spviewsum, 5, 54, spviewsum.MaxCols, 4, cdlPortrait)

End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If frastep.UBound > Index Then
    If optRandom = True Then
      If spCust.SelectionCount >= 1 Then
        spCust.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
'  Else
    If OptPrdrandom = True Then
      If spItem.SelectionCount >= 1 Then
        spItem.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  End If
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If SumColrFlag <> "CR" Then
    If cmbReportType <> "Spread Format Report" Then
      If sPrintcurrentSummary = False Then
 GoTo ExitHere
End If
      Call FinaliseAndShowPrintPreview
    Else
      Erpno = Get_ERPandTSNo(Me.Tag, "SAL", mErp)
      TsNo = Get_ERPandTSNo(Me.Tag, "SAL", mTs)
      If sSpread_sPrintcurrentSummary = False Then
 GoTo ExitHere
End If
    End If
  Else
    If cmbReportType <> "Spread Format Report" Then
      If sPrintCurrent = False Then
 GoTo ExitHere
End If
      Call FinaliseAndShowPrintPreview
    Else
      Erpno = Get_ERPandTSNo(Me.Tag, "SAL", mErp)
      TsNo = Get_ERPandTSNo(Me.Tag, "SAL", mTs)
      If sSpread_sPrintcurrent = False Then
 GoTo ExitHere
End If
    End If
  End If

'  If SumColrFlag = "CR" Then
'    If sPrintcurrent = False Then: GoTo ExitHere
'  Else
'    If sPrintcurrentSummary = False Then: GoTo ExitHere
'  End If
'  If cmbReportType.ListIndex = 1 Then: Call FinaliseAndShowPrintPreview
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Public Sub sPageHeader()
  Try

'  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  
  If SumColrFlag = "CR" Then
    sSendToPort LAlign(10, "Inv.No") & Space(1) & LAlign(50, "Item Code") & Space(1) & RAlign(8, "Qty") & Space(1) & RAlign(9, "Rate")
    sSendToPort LAlign(10, "Inv.Date") & Space(1) & LAlign(50, "Item Description") & Space(10) & RAlign(9, "Ass.Rate")
  Else
    If optRandom = True Then
      sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount")
    Else
      sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount") & Space(3) & LAlign(39, "Customer Name")
    End If
  End If
  If gReportID = "PMI" Then
    sSendToPort Space(11) & LAlign(50, "Revision No")
  End If

  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim Sql As String
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim sRow As Long
  Dim tmp As String
  Dim tmp1 As String
  Dim iType As String
  Dim iNo As Integer
  Dim iDate As Date
  Dim INVType As String
  Dim tot As Double
  Dim Gtot As Double
  Dim Gtot1 As Double
  Dim etot As Double
  Dim ctot As Double
  Dim Cnt As Integer
  
  Dim rsDetails As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
 
  Try
  
  spview.Tag = 0
  Sql = " select a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql & " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount,"
  Sql = Sql & " a.ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode "
  Sql = Sql & " and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) ItemDescription,"
  Sql = Sql & " (select c.RevisionNo from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode "
  Sql = Sql & " and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) RevisionNo,"
  Sql = Sql & " a.InvoiceQty,a.Rate,a.AssessableRate, a.Type,"
  Sql = Sql & " (select c.Uomdesc from MATERIAL_MASTER c where c.ItemCode=a.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid "
  Sql = Sql & " and c.status = 1)Uom From SalesInvoice a "
  Sql = Sql & " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and InvoiceDate BETWEEN convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  If cmbType <> "All" Then
    Sql = Sql & " and a.type = '" & Trim$(cmbType) & "'"
  End If
  If optWRR Then
    Sql = Sql & " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql & " and a.RRNo=''"
  End If
  
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Customer"
GoTo ExitHere
End If
    Sql = Sql & " and CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql & "'" & Trim$(spCust.Text) & "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql & ")"
  End If
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.ItemCode in("
      sRow = 0
      For i = 1 To spItem.SelectionCount
        sRow = spItem.GetMultiSelItem(sRow)
        spItem.Row = sRow
spItem.Col = 1
        Sql = Sql + "'" + Trim$(spItem.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  Sql = Sql & " order by PartyName,InvoiceDate,InvoiceNo"
  
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sPrintCurrent = False
    GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  gStrReportHeading = "Sales Register(Columnar) between " & dpfromdt & " and " & dptodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  
  If optWRR Then
    sSendToPort CAlign(iPageWidth, "With RR")
  ElseIf optWORR Then
    sSendToPort CAlign(iPageWidth, "Without RR")
  End If
  
  sSendToPort ""
  sPageHeader
  Cnt = 1
tmp = ""
  Do While Not rsDetails.EOF
    If tmp <> (rsDetails("InvoiceType") & rsDetails("InvoiceNo")) Then
      If tmp <> "" Then
        Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
        Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim(iType) & "', " & Val(iNo) & ",'" & Format(iDate, "yyyy-MM-dd") & "', " & IIf(Trim$(INVType) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        Sql = Sql & " order by a.ExprSeq"
        If RSOpen(snap, Sql) = False Then GoTo ExitHere
        If Not snap.EOF Then
          tot = 0
          sSendToPort ""
          snap.MoveLast
snap.MoveFirst
          For j = 1 To snap.RecordCount
            sSendToPort Space(11) & LAlign(50, CheckNull(snap("ExprName"))) & Space(10) & RAlign(9, Format(snap("Resultvalue"), "0.00"))
            snap.MoveNext
          Next j
          sSendToPort ""
        End If
        Cnt = Cnt + 1
      End If
      sSendToPort ""
      If tmp1 <> rsDetails("PartyName") Then
        If optRandom.Value = 1 Then
          If i > 1 Then
 sSendToPort String(iPageWidth, "-")
End If
        End If
        sSendToPort ""
        sSendToPort LAlign(10, "Customer :") & Space(1) & LAlign(50, CheckNull(rsDetails("PartyName")))
        sSendToPort ""
      End If
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(50, rsDetails("ItemCode")) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(9, Format(rsDetails("Rate"), "0.00"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(1) & LAlign(50, CheckNull(rsDetails("ItemDescription"))) & Space(10) & RAlign(9, Format(rsDetails("AssessableRate"), "0.00"))
    Else
      sSendToPort Space(11) & LAlign(50, rsDetails("ItemCode")) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(9, Format(rsDetails("Rate"), "0.00"))
      sSendToPort Space(11) & LAlign(50, rsDetails("ItemDescription")) & Space(10) & RAlign(9, Format(rsDetails("AssessableRate"), "0.00"))
    End If
'    i = i + 1
    If gReportID = "PMI" And rsDetails("RevisionNo") <> "" Then
      sSendToPort Space(11) & LAlign(50, rsDetails("RevisionNo"))
    End If
    sSendToPort ""
    tmp = (rsDetails("InvoiceType") & rsDetails("InvoiceNo"))
    iType = rsDetails("InvoiceType")
    iNo = rsDetails("InvoiceNo")
    iDate = rsDetails("InvoiceDate")
    INVType = rsDetails("Type")
    tmp1 = rsDetails("PartyName")
    rsDetails.MoveNext
  Loop
  If rsDetails.EOF Then
    Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
    Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim(iType) & "', " & Val(iNo) & ",'" & Format(iDate, "yyyy-MM-dd") & "', " & IIf(Trim$(INVType) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " order by a.ExprSeq"
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
    If Not snap.EOF Then
      tot = 0
      sSendToPort ""
      snap.MoveLast
snap.MoveFirst
      For j = 1 To snap.RecordCount
        sSendToPort Space(11) & LAlign(50, CheckNull(snap("ExprName"))) & Space(10) & RAlign(9, Format(snap("Resultvalue"), "0.00"))
        snap.MoveNext
      Next j
    End If
  End If
bal:
  If rsDetails.EOF Then
    Sql = " Select a.ExprName, Sum(ResultValue) as ResultValue, a.GrandTotalFlag from "
    Sql = Sql & " SalesCalculationDetails a"
    Sql = Sql & " right Outer Join SalesInvoice b On b.InvoiceType = a.DocType "
    Sql = Sql & " and b.InvoiceNo = a.DocNo and b.InvoiceDate = a.DocDate "
    Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER c"
    Sql = Sql & " On c.PartyCode = b.CustomerCode and c.Branchid = b.Branchid "
    Sql = Sql & " and c.Compid = b.Compid where a.GrandTotalFlag <> 0"
    Sql = Sql & " and a.DocDate Between  convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111) "
    Sql = Sql & " and  convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111) "
    If optRandom.Value = True Then
      Sql = Sql & " and b.CustomerCode in("
        sRow = 0
        For i = 1 To spCust.SelectionCount
          sRow = spCust.GetMultiSelItem(sRow)
          spCust.Row = sRow
spCust.Col = 1
          Sql = Sql & "'" & Trim$(spCust.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
    End If
    If OptPrdrandom Then
      If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and b.ItemCode in("
        sRow = 0
        For i = 1 To spItem.SelectionCount
          sRow = spItem.GetMultiSelItem(sRow)
          spItem.Row = sRow
spItem.Col = 1
          Sql = Sql + "'" + Trim$(spItem.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)
      Sql = Sql + ")"
    End If
    If cmbType <> "All" Then
      Sql = Sql & " and b.type = '" & Trim$(cmbType) & "'"
    End If
    If optWRR Then
      Sql = Sql & " and b.RRNo <> ''"
    ElseIf optWORR Then
      Sql = Sql & " and b.RRNo = ''"
    End If
    Sql = Sql & " Group By a.ExprName, a.GrandTotalFlag, a.exprseq order by a.exprseq "
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
      
      
      If Not snap.EOF Then
        sSendToPort ""
        snap.MoveLast
snap.MoveFirst
        sSendToPort LAlign(40, "Cumulative Total")
        sSendToPort LAlign(40, "----------------")
        tot = 0
        For j = 1 To snap.RecordCount
          sSendToPort Space(11) & LAlign(50, CheckNull(snap("ExprName"))) & Space(10) & RAlign(9, Format(snap("ResultValue"), "0.00"))
          tot = tot + Format(snap("ResultValue"), "0.00")
          sSendToPort ""
          snap.MoveNext
        Next j
        sSendToPort Space(11) & LAlign(50, "Invoice Amount") & Space(10) & RAlign(9, Format(CStr(tot), "0.00"))
        sSendToPort ""
      End If
      sSendToPort ""
      rsDetails.Close
      Call EndofReport(iPageWidth, Val(Cnt))
  End If
  
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
  Set snap = Nothing
End Try
End Function

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try

  CurrDate Me
  cmbType.ListIndex = 0
  cmbReportType.ListIndex = 0
  'Frame2.Visible = False
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
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  CmdStep(1).Enabled = True
  spCust.Visible = False
End Sub

Private Sub OptPAll_Click()

  Try
  Me.MousePointer = vbHourglass

  CmdStep(2).Enabled = True
  spItem.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OptPrdrandom_Click()
  Dim Sql As String
  Dim sRow As Long
  Dim i As Long
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(2).Enabled = False

  Sql = " select distinct a.ItemCode [Item Code],(select b.ItemDescription from INV_MATERIAL_MASTER b"
  Sql = Sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.compid and b.status = 1) [Item Description] "
  Sql = Sql + " from SalesInvoice a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt.Value, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt.Value, "yyyy-MM-dd") + "',111) "
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
OptPAll = True
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql + "'" + Trim$(spCust.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
  Sql = Sql + " order by a.ItemCode"
'  adoList2.ConnectionString = db
'  adoList2.RecordSource = Sql
'  adoList2.Refresh
'  spItem.ClearSelection
'  spItem.MaxRows = spItem.DataRowCnt
'  If spItem.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    OptPAll.SetFocus
'    GoTo ExitHere
'  End If
'  spItem.Enabled = True
'  spItem.Visible = True
'  spItem.SetFocus
 Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    OptPAll.Value = True
    GoTo ExitHere
  Else
    spItem.Enabled = True
 spItem.Visible = True
    spItem.ClearSelection
    Call ADORS_Spread_DataFill(rs, spItem, True)
    
    spItem.MaxRows = spItem.DataRowCnt
    Call Spread_Set_Focus(spItem, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optRandom_Click()

  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(1).Enabled = False

  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql & " Where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql & " From SalesInvoice a "
  Sql = Sql & " Where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and InvoiceDate BETWEEN convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  
  If cmbType <> "All" Then
    Sql = Sql & " and a.type = '" & Trim$(cmbType) & "'"
  End If
  
  If optWRR Then
    Sql = Sql & " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql & " and a.RRNo=''"
  End If
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.ClearSelection
'  spCust.MaxRows = spCust.DataRowCnt
'  If spCust.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    OptAll.SetFocus
'    GoTo ExitHere
'  End If
'  spCust.Enabled = True
'  spCust.Visible = True
'  spCust.SetFocus
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    optAll.Value = True
    GoTo ExitHere
  Else
    spCust.Enabled = True
 spCust.Visible = True
    spCust.ClearSelection
    Call ADORS_Spread_DataFill(rs, spCust, True)
    
    spCust.MaxRows = spCust.DataRowCnt
    Call Spread_Set_Focus(spCust, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spCust_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spCust.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub spCust_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spCust.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub spItem_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spItem.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If spItem.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub Design_spview(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spview

End Sub

Private Sub Design_spviewsum(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spviewsum

End Sub

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust


spCust.Click +=  new EventHandler(spCust_Click)
spCust.KeyUp +=  new EventHandler(spCust_KeyUp)

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.Click +=  new EventHandler(spItem_Click)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spview("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterColumnarReport.frm", spview, "")

	Call Design_spviewsum("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterColumnarReport.frm", spviewsum, "")

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterColumnarReport.frm", spCust, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterColumnarReport.frm", spItem, "")

End Sub
