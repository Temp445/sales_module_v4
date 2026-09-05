VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptMonthWiseSales 
   Caption         =   "Monthwise Sales"
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
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   600
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton cmdexit1 
         Caption         =   "E&xit"
         Height          =   330
         Left            =   9225
         TabIndex        =   8
         Top             =   6960
         Width           =   1300
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "P&rint"
         Height          =   330
         Left            =   7935
         TabIndex        =   34
         Top             =   6960
         Width           =   1300
      End
Begin AceSpread spview
         Height          =   420
         Left            =   225
         TabIndex        =   9
         Top             =   225
         Visible         =   0   'False
         Width           =   540
         _Version        =   458752
         _ExtentX        =   953
         _ExtentY        =   741
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
         MaxCols         =   7
      End
   End
   Begin VB.Frame framain 
      Height          =   4530
      Left            =   3825
      TabIndex        =   10
      Top             =   495
      Width           =   3090
      Begin VB.Frame frastep 
         Caption         =   "Report Type"
         Height          =   930
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   3465
         Width           =   2370
         Begin VB.ComboBox cmbReportType 
            Height          =   330
            ItemData        =   "rptMonthWise.frx":44A77
            Left            =   90
            List            =   "rptMonthWise.frx":44A81
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   210
            Width           =   2205
         End
         Begin VB.CheckBox chkPrintdate 
            Caption         =   "Print.Date"
            Height          =   255
            Left            =   90
            TabIndex        =   32
            Top             =   615
            Width           =   1080
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
         Index           =   5
         Left            =   2610
         TabIndex        =   30
         Top             =   3720
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
         Index           =   4
         Left            =   2610
         TabIndex        =   12
         Top             =   2985
         Width           =   345
      End
      Begin VB.Frame frastep 
         Caption         =   "Month"
         Height          =   660
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   135
         Width           =   2355
         Begin VB.ComboBox CmbMonth 
            Height          =   330
            ItemData        =   "rptMonthWise.frx":44AAF
            Left            =   1440
            List            =   "rptMonthWise.frx":44AC5
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   210
            Width           =   825
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "mmm-yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   28
            Top             =   210
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "MMM-yyyy"
            Format          =   22282243
            CurrentDate     =   39314
         End
      End
      Begin VB.Frame frastep 
         Caption         =   "Total "
         Height          =   660
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   2790
         Width           =   2370
         Begin VB.CheckBox ChkTotal 
            Caption         =   "Itemwise Total Required"
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   255
            Width           =   2025
         End
      End
      Begin VB.Frame frastep 
         Caption         =   "Customer"
         Enabled         =   0   'False
         Height          =   660
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1470
         Width           =   2370
         Begin VB.OptionButton OptCustAll 
            Caption         =   "&All"
            Height          =   210
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Select to choose all Customers (Alt + A)"
            Top             =   285
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.OptionButton OptCustSelect 
            Caption         =   "&Selective"
            Height          =   210
            Left            =   1260
            TabIndex        =   22
            Top             =   300
            Width           =   960
         End
      End
      Begin VB.Frame frastep 
         Caption         =   "Item"
         Enabled         =   0   'False
         Height          =   660
         Index           =   3
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
         Top             =   2130
         Width           =   2370
         Begin VB.OptionButton optItemall 
            Caption         =   "&All"
            Height          =   210
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Select to choose all Customers (Alt + A)"
            Top             =   285
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.OptionButton optItemSelective 
            Caption         =   "&Selective"
            Height          =   210
            Left            =   1260
            TabIndex        =   19
            Top             =   285
            Width           =   960
         End
      End
      Begin VB.Frame frastep 
         Caption         =   "Report Based on"
         Enabled         =   0   'False
         Height          =   660
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   810
         Width           =   2370
         Begin VB.ComboBox CmbReport 
            Height          =   330
            ItemData        =   "rptMonthWise.frx":44ADB
            Left            =   120
            List            =   "rptMonthWise.frx":44AE5
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   210
            Width           =   2160
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
         Left            =   2610
         TabIndex        =   15
         Top             =   360
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
         Left            =   2610
         TabIndex        =   14
         Top             =   975
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
         Index           =   2
         Left            =   2610
         TabIndex        =   13
         Top             =   1680
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
         Index           =   3
         Left            =   2610
         TabIndex        =   11
         Top             =   2310
         Width           =   345
      End
   End
   Begin VB.Frame FraSprItem 
      Caption         =   "Item"
      Height          =   1560
      Left            =   3105
      TabIndex        =   4
      Top             =   5655
      Width           =   4635
Begin AceSpread SprItem
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   4380
         _Version        =   458752
         _ExtentX        =   7726
         _ExtentY        =   2143
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
         MaxRows         =   4
         OperationMode   =   4
         ScrollBars      =   2
      End
   End
   Begin VB.Frame fraSprCust 
      Caption         =   "Customer"
      Height          =   1560
      Left            =   6960
      TabIndex        =   2
      Top             =   1980
      Width           =   4635
Begin AceSpread SprCust
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Width           =   4365
         _Version        =   458752
         _ExtentX        =   7699
         _ExtentY        =   2143
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
         MaxRows         =   4
         OperationMode   =   4
         ScrollBars      =   2
         ScrollBarTrack  =   1
      End
   End
Begin AceSpread spprint
      Height          =   465
      Left            =   585
      TabIndex        =   29
      Top             =   1110
      Visible         =   0   'False
      Width           =   645
      _Version        =   458752
      _ExtentX        =   1138
      _ExtentY        =   820
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxCols         =   8
      MaxRows         =   54
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   420
      Left            =   5970
      TabIndex        =   1
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   420
      Left            =   4770
      TabIndex        =   6
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Height          =   420
      Left            =   3570
      TabIndex        =   0
      Top             =   5100
      Width           =   1215
   End
End
Attribute VB_Name = "rptMonthWiseSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum eCust
  mPartyCode = 1
  mPartyName = 2
End Enum

Public Enum eItem
  mItemCode1 = 1
  mItemDescription = 2
End Enum

Private Function ItemwiseCustomerwiseDisplay() As Boolean
  Dim sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim i As Integer
  Dim ToMonthyear As Date
  Dim toDate As Date
  Dim lcnt As Integer
  Dim CustomerCode As Integer
  Dim CustomerName As String
  Dim ItemCode As String
  Dim Qty As String
  Dim RecCnt As Integer
  Dim Cnt As Integer
  Dim rCnt As Integer
  Dim tmp As String
  Dim TotQty(5) As Double
  
  Try
  
  ItemwiseCustomerwiseDisplay = False
  
  If OptCustSelect = True Then
    If SprCust.SelectionCount < 1 Then
      sShowMessage "Please select atlease one item"
      GoTo ExitHere
    End If
  Else
    If optItemSelective = True Then
      If sprItem.SelectionCount < 1 Then
        sShowMessage "Please select atlease one item"
        GoTo ExitHere
      End If
    End If
  End If
  
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))

  sql = " Select distinct ItemCode " ', RevisionNo"
  sql = sql & " From (Select distinct Month(a.InvoiceDate) MthCode, b.ItemCode " ', c.RevisionNo "
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.InvoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Left Outer Join Material_Master c "
  sql = sql & " On b.ItemCode = c.ItemCode and b.BranchID = c.BranchID and b.CompID = c.CompID"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  sql = sql & ")x  Where Itemcode is not Null "
  If optItemSelective = True Then
    sql = sql & " and ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
  sql = sql & " Order by ItemCode"
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, sql) = False Then
    GoTo ExitHere
  End If
  If rsDetails.EOF Then
    MsgBox "No Records Found. "
    ItemwiseCustomerwiseDisplay = False
    GoTo ExitHere
  End If
  
  sql = " Select ItemCode, Itemdescription, CustomerCode, PartyName, RevisionNo, MthCode, Qty"
  sql = sql & " From (Select distinct b.ItemCode, d.ItemDescription, d.RevisionNo, a.CustomerCode, c.PartyName, Month(a.InvoiceDate) MthCode, "
  sql = sql & " Isnull(Sum(B.InvoiceQty), 0) As Qty"
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.InvoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Left Outer Join FAS_Party_Master c"
  sql = sql & " On c.PartyCode = a.CustomerCode and c.BranchId = b.BranchId and c.CompId = b.CompId"
  sql = sql & " Left Outer Join INV_Material_Master d"
  sql = sql & " On d.ItemCode = b.ItemCode and d.BranchId = c.BranchId and d.CompId = c.CompId"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  If optItemSelective = True Then
    sql = sql & " and b.ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
  If OptCustSelect = True Then
    sql = sql & " and a.CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  sql = sql & " Group by b.ItemCode, d.ItemDescription, d.RevisionNo, a.CustomerCode, c.PartyName, Month(a.InvoiceDate))x"
  sql = sql & " Where ItemCode is not Null and Qty > 0 "
  sql = sql & " Order by ItemCode, CustomerCode"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If rs.EOF Then
    MsgBox "No Records Found. "
    ItemwiseCustomerwiseDisplay = False
    GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 1 Then
    pheight = 75
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    
    sql = CAlign(iPageWidth, "Itemwise Customerwise Sales Report")
    gStrReportHeading = sql & Chr(13) & " For " & Left(MonthName(Month(dtpFrom)), 3) & " And " & Left(MonthName(Month(toDate)), 3)
    sSendToPort gStrReportHeading
    
    Call sPageHeader
    
    CustomerCode = 0
    ItemCode = ""
     
    TotQty(0) = 0
TotQty(1) = 0
TotQty(2) = 0
TotQty(3) = 0
TotQty(4) = 0
TotQty(5) = 0
    RecCnt = 0
    Do While Not rsDetails.EOF
      rs.Filter = adFilterNone
      rs.Filter = " ItemCode = '" & rsDetails("ItemCode") & "'"
      
      If Not rs.EOF Then
        If ItemCode <> rs("ItemCode") Then
          If ItemCode <> "" Then
            sSendToPort LAlign(37, Trim$(CustomerName)) & Space(1) & tmp
            If ChkTotal.Value = 1 Then
              sSendToPort
              sSendToPort LAlign(10, "Total Qty:") & Space(28) & IIf(TotQty(0) <> 0, RAlign(6, Val(TotQty(0))), "") & Space(1) & IIf(TotQty(1) <> 0, RAlign(6, Val(TotQty(1))), "") & Space(1) & IIf(TotQty(2) <> 0, RAlign(6, Val(TotQty(2))), "") & Space(1) & IIf(TotQty(3) <> 0, RAlign(6, Val(TotQty(3))), "") & Space(1) & IIf(TotQty(4) <> 0, RAlign(6, Val(TotQty(4))), "") & Space(1) & IIf(TotQty(5) <> 0, RAlign(6, Val(TotQty(5))), "")
            End If
            tmp = ""
            TotQty(0) = 0
TotQty(1) = 0
TotQty(2) = 0
TotQty(3) = 0
TotQty(4) = 0
TotQty(5) = 0
          End If
          sSendToPort
          sSendToPort LAlign(15, "ItemCode:") & LAlign(65, rs("ItemCode") & "-" & rs("Itemdescription"))
          If gReportID = "PMI" And rs("RevisionNo") <> "" Then
            sSendToPort LAlign(15, "Revision No :") & LAlign(65, rs("RevisionNo"))
          End If
          sSendToPort
          ItemCode = rs("ItemCode")
          RecCnt = RecCnt + 1
        End If
        Cnt = 1
       For rCnt = 1 To rs.RecordCount
          If CustomerCode <> rs("CustomerCode") Then
            If rCnt > 1 Then
              sSendToPort LAlign(37, Trim$(CustomerName)) & Space(1) & tmp
            End If
            Cnt = 1
            tmp = ""
            CustomerCode = rs("CustomerCode")
            CustomerName = rs("PartyName")
            tmp = tmp & RAlign(6, rs("Qty"))
            If ChkTotal.Value = 1 Then TotQty(Cnt - 1) = TotQty(Cnt - 1) + rs("Qty")
            Debug.Print rCnt - 1
            Debug.Print TotQty(rCnt - 1)
          Else
            tmp = tmp & IIf(tmp <> "", Space(1), "") & RAlign(6, rs("Qty"))
            If ChkTotal.Value = 1 Then TotQty(Cnt - 1) = TotQty(Cnt - 1) + rs("Qty")
            Debug.Print Cnt - 1
            Debug.Print TotQty(Cnt - 1)
          End If
          Cnt = Cnt + 1
          rs.MoveNext
        Next rCnt
      End If
      rsDetails.MoveNext
      If rsDetails.EOF Then
        sSendToPort LAlign(37, Trim$(CustomerName)) & Space(1) & tmp
        If ChkTotal.Value = 1 Then sSendToPort
sSendToPort LAlign(10, "Total Qty:") & Space(28) & IIf(TotQty(0) <> 0, RAlign(6, Val(TotQty(0))), "") & Space(1) & IIf(TotQty(1) <> 0, RAlign(6, Val(TotQty(1))), "") & Space(1) & IIf(TotQty(2) <> 0, RAlign(6, Val(TotQty(2))), "") & Space(1) & IIf(TotQty(3) <> 0, RAlign(6, Val(TotQty(3))), "") & Space(1) & IIf(TotQty(4) <> 0, RAlign(6, Val(TotQty(4))), "") & Space(1) & IIf(TotQty(5) <> 0, RAlign(6, Val(TotQty(5))), "")
      End If
    Loop
    Call EndofReport(iPageWidth, Val(RecCnt))
   Else
      spview.MaxRows = 500
      Call Spread_HideRow(spview, False, 6, spview.MaxRows)
      
      For i = 6 To spview.MaxRows
        spview.ClearRange 1, i, spview.MaxCols, i, True
        For Cnt = 1 To spview.MaxCols
          spview.Col = Cnt
spview.Row = i
spview.BackColor = RGB(255, 255, 255)
        Next Cnt
        spview.RemoveCellSpan 1, i
        spview.SetCellBorder 1, spview.Row, spview.MaxCols, spview.MaxRows, CellBorderIndexLeft, 0, CellBorderStyleBlank
      Next i
      spview.Col = 1
spview.Row = 4
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignCenter
      i = 6
      spview.SetText 1, 1, gCompName
      spview.SetText 1, 3, "Itemwise Customerwise Sales Report" & Space(2) & "(" & " For " & Left(MonthName(Month(dtpFrom)), 3) & " And " & Left(MonthName(Month(toDate)), 3) & " )"
      spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
      Call sPageHeader
      cmdView.Tag = 1
      Do While Not rsDetails.EOF
        rs.Filter = adFilterNone
        rs.Filter = " ItemCode = " & rsDetails("ItemCode")
        If Not rs.EOF Then
          If ItemCode <> rs("ItemCode") Then
            If ItemCode <> "" Then
              spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
              spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
              spview.SetText 1, i, CustomerName
  '            If gReportID = "PMI" And rs("RevisionNo") <> "" Then
  '              sSendToPort LAlign(20, rs("RevisionNo"))
  '            End If
              i = i + 1
            End If
            spview.Col = 1
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
            spview.AddCellSpan spview.Col, spview.Row, spview.MaxCols, 1
            spview.SetText spview.Col, spview.Row, "Item Code:  " & rs("ItemCode") & " - " & rs("Itemdescription")
            ItemCode = rs("ItemCode")
            spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
            i = i + 1
spprint.Tag = Val(spprint.Tag) + 1
          End If
          Cnt = 0
spview.Tag = 2
          Do While Not rs.EOF
            If CustomerName <> rs("Partyname") Then
              If Cnt <> 0 Then
                spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
                spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
                spview.SetText 1, i, CustomerName
                spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'                If gReportID = "PMI" And rs("RevisionNo") <> "" Then
'                  sSendToPort LAlign(20, rs("RevisionNo"))
'                End If
                i = i + 1
cmdView.Tag = Val(cmdView.Tag) + 1
              End If
              spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
              CustomerName = rs("Partyname")
              Cnt = Cnt + 1
spview.Tag = 2

              spview.Col = Val(spview.Tag)
spview.Row = i
spview.BackColor = RGB(100, 255, 255)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spview.SetText Val(spview.Tag), i, rs("Qty")
              spview.Tag = Val(spview.Tag) + 1
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i, CellBorderIndexLeft, 0, CellBorderStyleSolid
            Else
              spview.Col = Val(spview.Tag)
spview.Row = i
spview.BackColor = RGB(100, 255, 255)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spview.SetText Val(spview.Tag), i, rs("Qty")
              spview.Tag = Val(spview.Tag) + 1
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i, CellBorderIndexLeft, 0, CellBorderStyleSolid
            End If
            rs.MoveNext
          Loop
        End If
        rsDetails.MoveNext
      Loop
      If rsDetails.EOF Then
        spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
        spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
        spview.SetText 1, i, CustomerName
        spview.SetCellBorder 1, i + 1, spview.MaxCols, i + 1, CellBorderIndexTop, 0, CellBorderStyleSolid
        i = i + 1
        spview.Col = 1
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
spview.Font = "Arial"
spview.FontBold = True

        spview.SetText 1, i, "Grant Total : "
        For Cnt = 2 To 7
          spview.SetCellBorder Cnt, i, Cnt, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
          spview.Col = Cnt
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = True
          Qty = ""
          Qty = Val(Spread_Cell_Total(spview, Val(Cnt), 6, Val(Cnt), spview.MaxRows))
          If Val(Qty) > 0 Then
            spview.SetText Cnt, i, Val(Qty)
          End If
        Next Cnt
      End If
      rCnt = 54
CmdClear.Tag = 0
      If cmdprint = True Then
        For i = 6 To spview.MaxRows
          If i = rCnt Or i = spview.DataRowCnt Then
            If i = spview.DataRowCnt Then
              rCnt = Val(rCnt) - i
              spview.MaxRows = Val(spview.DataRowCnt + rCnt - 3)
              i = i + 1
              spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
              spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows, CellBorderIndexLeft, 0, CellBorderStyleBlank
              spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows, CellBorderIndexRight, 0, CellBorderStyleBlank
              spview.SetCellBorder 1, i + 1, spview.MaxCols, spview.MaxRows, CellBorderIndexTop, 0, CellBorderStyleBlank
              spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows, CellBorderIndexBottom, 0, CellBorderStyleBlank
              i = spview.MaxRows
            Else
              spview.InsertRows i, 1
            End If
            spview.AddCellSpan 1, i, 5, 1
            spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
            CmdClear.Tag = Val(CmdClear.Tag) + 1
            spview.Col = 1
spview.Row = i
spview.Font = "Arial"
spview.FontBold = True
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignCenter
            spview.SetText 1, i, gCompName
            spview.AddCellSpan 6, i, spview.MaxCols, 1
            spview.SetCellBorder 6, i, 6, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
            spview.Col = 6
spview.Row = i
spview.Font = "Arial"
spview.FontBold = True
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignCenter

            spview.SetText 6, i, "Page : " & Val(CmdClear.Tag)
            If i <> spview.MaxRows Then
              Call Spread_HideRow(spview, True, Val(rCnt) + 1, spview.DataRowCnt)
            End If
            spview.PrintOrientation = PrintOrientationPortrait
            spview.PrintPageOrder = PageOrderDownThenOver
            spview.PrintFirstPageNumber = 1
            spview.PrintPageStart = 1
            spview.PrintScalingMethod = PrintScalingMethodSmartPrint
            spview.BorderStyle = BorderStyleFixedSingle
            spview.PrintBorder = True
            spview.PrintCenterOnPageH = True
            spview.PrintCenterOnPageV = True
            spview.PrintMarginTop = 500
            spview.PrintMarginLeft = 50
            spview.PrintMarginRight = 1.25
            spview.Action = ActionPrint
            Call Spread_HideRow(spview, False, 6, spview.DataRowCnt)
            If i <> spview.DataRowCnt Then
              Call Spread_HideRow(spview, True, 6, Val(rCnt))
            End If
            rCnt = rCnt + 54
          End If
        Next i
      Else
        spview.MaxRows = spview.DataRowCnt
        Frame2.Height = 7455
        Frame2.Width = 10785
        spview.Height = 6630
        spview.Width = 10290
        Frame2.Visible = True
        spview.Visible = True
      End If
  End If
  
  ItemwiseCustomerwiseDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDetails = Nothing
End Try
End Function

Private Function CustomerwiseItemwiseDisplay() As Boolean
  Dim sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rsReturn As AceADODB.Recordset
  Dim ToMonthyear As Date
  Dim toDate As Date
  Dim i As Integer
  Dim CustomerCode As Integer
  Dim Item(1) As String
  Dim Cnt As Integer
  Dim lcnt As Integer
  Dim tmp As String
  Dim Qty As String

  Try
  
  CustomerwiseItemwiseDisplay = False
  
  If OptCustSelect = True Then
    If SprCust.SelectionCount < 1 Then
      sShowMessage "Please select atlease one item"
      GoTo ExitHere
    End If
  Else
    If optItemSelective = True Then
      If sprItem.SelectionCount < 1 Then
        sShowMessage "Please select atlease one item"
        GoTo ExitHere
      End If
    End If
  End If
  
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))

  sql = " Select distinct CustomerCode"
  sql = sql & " From (Select distinct Month(a.InvoiceDate) MthCode, a.CustomerCode, b.ItemCode "
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  sql = sql & ")x  "
  If OptCustSelect = True Then
    sql = sql & " Where CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  sql = sql & " Order by CustomerCode"
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, sql) = False Then
    GoTo ExitHere
  End If
  If rsDetails.EOF Then
    MsgBox "No Records Found. "
    CustomerwiseItemwiseDisplay = False
    GoTo ExitHere
  End If
  
  sql = " Select CustomerCode, PartyName, ItemCode, Itemdescription, RevisionNo, MthCode, Qty"
  sql = sql & " From (Select distinct a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription, Month(a.InvoiceDate) MthCode, "
  sql = sql & " Isnull(Sum(B.InvoiceQty), 0) As Qty, d.RevisionNo"
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Left Outer Join FAS_Party_Master c"
  sql = sql & " On c.PartyCode = a.CustomerCode and c.BranchId = b.BranchId and c.CompId = b.CompId"
  sql = sql & " Left Outer Join INV_Material_Master d"
  sql = sql & " On d.ItemCode = b.ItemCode and d.BranchId = c.BranchId and d.CompId = c.CompId"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  If OptCustSelect = True Then
    sql = sql & " and a.CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  If optItemSelective = True Then
    sql = sql & " and b.ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
  sql = sql & " Group by a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription, d.RevisionNo, Month(a.InvoiceDate))x"
  sql = sql & " Where Qty > 0 "
  sql = sql & " Order by CustomerCode, ItemCode, MthCode"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If rs.EOF Then
    MsgBox "No Records Found. "
    CustomerwiseItemwiseDisplay = False
    GoTo ExitHere
  End If
  
  sql = " Select ItemCode, Month(ReceiptEntryDate) as MthCode, Sum(ReceiptQty) Qty, VendorCode"
  sql = sql & " From ReceiptEntry"
  sql = sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and InwardType = 'R'"
  sql = sql & " and ReceiptEntryDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  If OptCustSelect = True Then
    sql = sql & " and VendorCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  If optItemSelective = True Then
    sql = sql & " and ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
  sql = sql & " Group by VendorCode, ItemCode, Month(ReceiptEntryDate)"
  sql = sql & " Order by Mthcode, VendorCode"
  Set rsReturn = New AceADODB.Recordset
  If RSOpen(rsReturn, sql) = False Then
    GoTo ExitHere
  End If
    
  If cmbReportType.ListIndex = 1 Then
    pheight = 75
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    
    sql = CAlign(iPageWidth, "Customerwise Itemwise Sales Report")
    gStrReportHeading = sql & Chr(13) & " For " & Left(MonthName(Month(dtpFrom)), 3) & " And " & Left(MonthName(Month(toDate)), 3)
    sSendToPort gStrReportHeading
    
    Call sPageHeader
    
    Item(0) = ""
    CustomerCode = 0
    
    Do While Not rsDetails.EOF
      rs.Filter = adFilterNone
      rs.Filter = " CustomerCode = " & rsDetails("CustomerCode")
      If Not rs.EOF Then
        If CustomerCode <> rs("CustomerCode") Then
          If CustomerCode <> 0 Then
            sSendToPort LAlign(20, Trim$(Item(0))) & Space(1) & tmp
            sSendToPort LAlign(50, Trim$(Item(1)))
            If gReportID = "PMI" And rs("RevisionNo") <> "" Then
              sSendToPort LAlign(20, rs("RevisionNo"))
            End If
            sSendToPort
          End If
          sSendToPort LAlign(15, "Customer Name:") & LAlign(50, rs("PartyName"))
          sSendToPort
          CustomerCode = rs("CustomerCode")
        End If
         
        Cnt = 0
        Do While Not rs.EOF
          If Item(0) <> rs("ItemCode") Then
            If Cnt <> 0 Then
              sSendToPort LAlign(20, Trim$(Item(0))) & Space(1) & tmp
              sSendToPort LAlign(50, Trim$(Item(1)))
              If gReportID = "PMI" And rs("RevisionNo") <> "" Then
                sSendToPort LAlign(20, rs("RevisionNo"))
              End If
              sSendToPort Space(21) & Qty
            End If
            tmp = ""
            Item(0) = rs("ItemCode")
            Item(1) = rs("Itemdescription")
            tmp = tmp & RAlign(7, rs("Qty"))
            Cnt = Cnt + 1
          Else
            tmp = tmp & Space(1) & RAlign(7, rs("Qty"))
          End If
          rs.MoveNext
          
  '        lCnt = 1
  '
  '        If Item(0) <> rs("ItemCode") Then
  '          rsReturn.Filter = adFilterNone
  '          rsReturn.Filter = " VendorCode = " & rsDetails("CustomerCode")
  '          If Not rsReturn.EOF Then
  '            rsReturn.Filter = adFilterNone
  '            rsReturn.Filter = " ItemCode = '" & Item(0) & "' and MthCode = " & rs("MthCode") & ""
  '            Do While Not rsReturn.EOF
  '              If lCnt = 1 Then
  '                Qty = Qty & RAlign(7, rsReturn("Qty"))
  '              Else
  '                Qty = Qty & Space(1) & RAlign(7, rsReturn("Qty"))
  '              End If
  '              rsReturn.MoveNext
  '              lCnt = lCnt + 1
  '  '            sSendToPort Space(21) & tmp
  '            Loop
  '          End If
  '        End If
        Loop
      End If
      
  '    Cnt = rsDetails.RecordCount
  '    If optItemSelective = True Then
  '      Cnt = sprItem.SelectionCount
  ''    Else
  ''      Cnt = rs.RecordCount
  '    End If
      
      rsDetails.MoveNext
      If rsDetails.EOF Then
        sSendToPort LAlign(20, Trim$(Item(0))) & Space(1) & tmp
        sSendToPort LAlign(50, Trim$(Item(1)))
        sSendToPort Space(21) & Qty
      End If
    Loop
     
    Call EndofReport(iPageWidth, Val(Cnt))
  Else
      spview.MaxRows = 500
      Call Spread_HideRow(spview, False, 6, spview.MaxRows)
      
      For i = 6 To spview.MaxRows
        spview.ClearRange 1, i, spview.MaxCols, i, True
        For Cnt = 1 To spview.MaxCols
          spview.Col = Cnt
spview.Row = i
spview.BackColor = RGB(255, 255, 255)
        Next Cnt
        spview.RemoveCellSpan 1, i
        spview.SetCellBorder 1, spview.Row, spview.MaxCols, spview.MaxRows, CellBorderIndexLeft, 0, CellBorderStyleBlank
      Next i
      spview.Col = 1
spview.Row = 4
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignCenter
      i = 6
      spview.SetText 1, 1, gCompName
      spview.SetText 1, 3, "Customerwise Itemwise Sales Report" & Space(2) & "(" & " For " & Left(MonthName(Month(dtpFrom)), 3) & " And " & Left(MonthName(Month(toDate)), 3) & " )"
      spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
      Call sPageHeader
      cmdView.Tag = 1
      Do While Not rsDetails.EOF
        rs.Filter = adFilterNone
        rs.Filter = " CustomerCode = " & rsDetails("CustomerCode")
        If Not rs.EOF Then
          If CustomerCode <> rs("CustomerCode") Then
            If CustomerCode <> 0 Then
              spview.SetCellBorder 1, i, 1, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
              spview.SetText 1, i, Trim$(Item(0))
              spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
              spview.SetText 1, i + 1, Trim$(Item(1))
  '            If gReportID = "PMI" And rs("RevisionNo") <> "" Then
  '              sSendToPort LAlign(20, rs("RevisionNo"))
  '            End If
              i = i + 2
            End If
            spview.Col = 1
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
            spview.AddCellSpan spview.Col, spview.Row, spview.MaxCols, 1
            spview.SetText spview.Col, spview.Row, "Customer Name:  " & rs("PartyName")
            CustomerCode = rs("CustomerCode")
            spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
            i = i + 1
spprint.Tag = Val(spprint.Tag) + 1
          End If
          Cnt = 0
spview.Tag = 2
          Do While Not rs.EOF
            If Item(0) <> rs("ItemCode") Then
              If Cnt <> 0 Then
                spview.SetCellBorder 1, i, 1, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
                spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
                spview.SetText 1, i, Trim$(Item(0))
                spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
                spview.SetText 1, i + 1, Trim$(Item(1))
'                If gReportID = "PMI" And rs("RevisionNo") <> "" Then
'                  sSendToPort LAlign(20, rs("RevisionNo"))
'                End If
                i = i + 2
cmdView.Tag = Val(cmdView.Tag) + 1
spview.Tag = 2
              End If
              tmp = ""
              spview.SetCellBorder 1, i, spview.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
              Item(0) = rs("ItemCode")
              Item(1) = rs("Itemdescription")
              Cnt = Cnt + 1
              spview.Col = Val(spview.Tag)
spview.Row = i
spview.BackColor = RGB(100, 255, 255)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spview.SetText Val(spview.Tag), i, rs("Qty")
              spview.Tag = Val(spview.Tag) + 1
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
            Else
              spview.Col = Val(spview.Tag)
spview.Row = i
spview.BackColor = RGB(100, 255, 255)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spview.SetText Val(spview.Tag), i, rs("Qty")
              spview.Tag = Val(spview.Tag) + 1
              spview.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
            End If
            rs.MoveNext
          Loop
        End If
        rsDetails.MoveNext
      Loop
      If rsDetails.EOF Then
      
        spview.SetCellBorder 1, i, 1, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
        spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
        spview.SetText 1, i, Trim$(Item(0))
        spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
        spview.SetText 1, i + 1, Trim$(Item(1))
        spview.SetCellBorder 1, i + 2, spview.MaxCols, i + 2, CellBorderIndexTop, 0, CellBorderStyleSolid
        i = i + 2
        spview.Col = 1
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignRight
spview.Font = "Arial"
spview.FontBold = True

        spview.SetText 1, i, "Grant Total : "
        For Cnt = 2 To 7
          spview.SetCellBorder Cnt, i, Cnt + 2, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
          spview.Col = Cnt
spview.Row = i
spview.BackColor = RGB(255, 255, 0)
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = True
          Qty = ""
          Qty = Val(Spread_Cell_Total(spview, Val(Cnt), 6, Val(Cnt), spview.MaxRows))
          If Val(Qty) > 0 Then
            spview.SetText Cnt, i, Val(Qty)
          End If
        Next Cnt
      End If
      spview.MaxRows = spview.DataRowCnt
      Frame2.Height = 7455
      Frame2.Width = 10785
      spview.Height = 6630
      spview.Width = 10290
      Frame2.Visible = True
      spview.Visible = True
  End If
  
  CustomerwiseItemwiseDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDetails = Nothing
  Set rsReturn = Nothing
End Try
End Function

Public Sub sPageHeader()
  Dim ToMonthyear As Date
  Dim toDate As Date
  Dim i As Integer
  Dim rCnt As Integer
  Dim var As String
  Dim tmp As Variant
  Dim Cnt As Integer
  
  Try
    
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))
  
  tmp = Month(dtpFrom)
  Cnt = tmp + (Val(CmbMonth.Text) - 1)
  i = Val(Cnt - tmp) + 1
  If cmbReportType.ListIndex = 1 Then
    For rCnt = tmp To Cnt
      If rCnt > 12 Then
        rCnt = rCnt - 12
        var = var & Space(1) & RAlign(IIf(CmbReport.ListIndex = 0, 7, 6), Left$(MonthName(rCnt), 3))
        i = i - 1
      Else
        var = var & Space(1) & RAlign(IIf(CmbReport.ListIndex = 0, 7, 6), Left$(MonthName(rCnt), 3))
        i = i - 1
      End If
      If i = 0 Then
 Exit For
End If
    Next rCnt
  
    sSendToPort String(iPageWidth, "-")
    If CmbReport.ListIndex = 0 Then
      sSendToPort LAlign(20, "ItemCode") & Space(1) & var
      sSendToPort LAlign(50, "ItemDescription")
      If gReportID = "PMI" Then
        sSendToPort LAlign(20, "Revision No")
      End If
    Else
      sSendToPort LAlign(37, "Customer Name") & Space(1) & LAlign(42, var)
    End If
    sSendToPort String(iPageWidth, "-")
  Else
    spview.Tag = 2
    For rCnt = tmp To Cnt
      If rCnt > 12 Then
        rCnt = rCnt - 12
        var = Left$(MonthName(rCnt), 3)
        If cmdprint = True And CmbReport.ListIndex = 0 Then
          spprint.SetText Val(spview.Tag), 6, var
        Else
          spview.SetText Val(spview.Tag), 5, var
        End If
        i = i - 1
spview.Tag = Val(spview.Tag) + 1
      Else
        var = Left$(MonthName(rCnt), 3)
        If cmdprint = True And CmbReport.ListIndex = 0 Then
          spprint.SetText Val(spview.Tag), 6, var
        Else
          spview.SetText Val(spview.Tag), 5, var
        End If
        i = i - 1
spview.Tag = Val(spview.Tag) + 1
      End If
      If i = 0 Then
 Exit For
End If
    Next rCnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrint() As Boolean
  Dim sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim ToMonthyear As Date
  Dim toDate As Date
  Dim i As Integer
  Dim CustomerCode As Integer
  Dim Item(1) As String
  Dim Cnt As Integer
  Dim Qty As String
  Dim gTotal(6) As String
  
  Try
  
  sPrint = False
  
  If OptCustSelect = True Then
    If SprCust.SelectionCount < 1 Then
      sShowMessage "Please select atlease one item"
      GoTo ExitHere
    End If
  Else
    If optItemSelective = True Then
      If sprItem.SelectionCount < 1 Then
        sShowMessage "Please select atlease one item"
        GoTo ExitHere
      End If
    End If
  End If
  
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))

  sql = " Select distinct CustomerCode"
  sql = sql & " From (Select distinct Month(a.InvoiceDate) MthCode, a.CustomerCode, b.ItemCode "
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  sql = sql & ")x  "
  If OptCustSelect = True Then
    sql = sql & " Where CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  sql = sql & " Order by CustomerCode"
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, sql) = False Then
    GoTo ExitHere
  End If
  If rsDetails.EOF Then
    MsgBox "No Records Found. "
    sPrint = False
    GoTo ExitHere
  End If
  
  sql = " Select CustomerCode, PartyName, ItemCode, Itemdescription, RevisionNo, MthCode, Qty"
  sql = sql & " From (Select distinct a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription, Month(a.InvoiceDate) MthCode, "
  sql = sql & " Isnull(Sum(B.InvoiceQty), 0) As Qty, d.RevisionNo"
  sql = sql & " From SAL_Invoice_hdr a"
  sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Left Outer Join FAS_Party_Master c"
  sql = sql & " On c.PartyCode = a.CustomerCode and c.BranchId = b.BranchId and c.CompId = b.CompId"
  sql = sql & " Left Outer Join INV_Material_Master d"
  sql = sql & " On d.ItemCode = b.ItemCode and d.BranchId = c.BranchId and d.CompId = c.CompId"
  sql = sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  sql = sql & " and a.InvoiceDate between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "'"
  If OptCustSelect = True Then
    sql = sql & " and a.CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  If optItemSelective = True Then
    sql = sql & " and b.ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
  sql = sql & " Group by a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription, d.RevisionNo, Month(a.InvoiceDate))x"
  sql = sql & " Where Qty > 0 "
  sql = sql & " Order by CustomerCode, ItemCode, MthCode"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If rs.EOF Then
    MsgBox "No Records Found. "
    sPrint = False
    GoTo ExitHere
  End If
    
    CmdClear.Tag = 0
    spprint.Tag = 0
    spprint.SetText 1, 5, "Printed Date : " & Format(ServerDate, "dd/MM/yyyy")
    i = 7
    spprint.SetText 1, 1, gCompName
    spprint.SetText 1, 3, "Customerwise Itemwise Sales Report" & Space(2) & "(" & " For " & Left(MonthName(Month(dtpFrom)), 3) & " And " & Left(MonthName(Month(toDate)), 3) & " )"
    spprint.SetCellBorder 1, i, spprint.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
    Call sPageHeader
    cmdView.Tag = 1
    Do While Not rsDetails.EOF
      rs.Filter = adFilterNone
      rs.Filter = " CustomerCode = " & rsDetails("CustomerCode")
      If Not rs.EOF Then
        If CustomerCode <> rs("CustomerCode") Then
          If CustomerCode <> 0 Then
            spprint.Col = 1
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
            spprint.SetText 1, i, 1
            spprint.SetCellBorder 2, i, 2, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
            spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
            spprint.SetText 2, i, Trim$(Item(0))
            spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
            spprint.SetText 2, i + 1, Trim$(Item(1))
'            If gReportID = "PMI" And rs("RevisionNo") <> "" Then
'              sSendToPort LAlign(20, rs("RevisionNo"))
'            End If
            i = i + 2
          End If
          If i < 51 Then
            spprint.Col = 1
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
            spprint.AddCellSpan spprint.Col, spprint.Row, spprint.MaxCols, 1
            spprint.SetText spprint.Col, spprint.Row, "Customer Name:  " & rs("PartyName")
            CustomerCode = rs("CustomerCode")
            spprint.SetCellBorder 1, i, spprint.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
            i = i + 1
spprint.Tag = spprint.Tag + 1
          Else
            gTotal(0) = Val(Spread_Cell_Total(spprint, 3, 7, 3, 53)) + Val(gTotal(0))
            gTotal(1) = Val(Spread_Cell_Total(spprint, 4, 7, 4, 53)) + Val(gTotal(1))
            gTotal(2) = Val(Spread_Cell_Total(spprint, 5, 7, 5, 53)) + Val(gTotal(2))
            gTotal(3) = Val(Spread_Cell_Total(spprint, 6, 7, 6, 53)) + Val(gTotal(3))
            gTotal(4) = Val(Spread_Cell_Total(spprint, 7, 7, 7, 53)) + Val(gTotal(4))
            gTotal(5) = Val(Spread_Cell_Total(spprint, 8, 7, 8, 53)) + Val(gTotal(5))
            GoTo pprint
          End If
        End If
        Cnt = 0
spview.Tag = 3
rec:
        Do While Not rs.EOF
          If Item(0) <> rs("ItemCode") Then
            If Cnt <> 0 Then
              spprint.Col = 1
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
              spprint.SetText 1, i, Val(cmdView.Tag)
              spprint.SetCellBorder 1, i, 2, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
              spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
              spprint.SetText 2, i, Trim$(Item(0))
              spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
              spprint.SetText 2, i + 1, Trim$(Item(1))
'              If gReportID = "PMI" And rs("RevisionNo") <> "" Then
'                sSendToPort LAlign(20, rs("RevisionNo"))
'              End If
              i = i + 2
cmdView.Tag = Val(cmdView.Tag) + 1
            End If

            spprint.SetCellBorder 1, i, spprint.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
            Item(0) = rs("ItemCode")
            Item(1) = rs("Itemdescription")
            Cnt = Cnt + 1
spview.Tag = 3

            spprint.Col = Val(spview.Tag)
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignRight
            spprint.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
            spprint.SetText Val(spview.Tag), i, rs("Qty")
            spview.Tag = Val(spview.Tag) + 1
            spprint.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
          Else
            spprint.Col = Val(spview.Tag)
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignRight
            spprint.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
            spprint.SetText Val(spview.Tag), i, rs("Qty")
            spview.Tag = Val(spview.Tag) + 1
            spprint.SetCellBorder Val(spview.Tag), i, Val(spview.Tag), i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
          End If
          cmdExit.Tag = 0
          If Val(i) >= 53 Then
            gTotal(0) = Val(Spread_Cell_Total(spprint, 3, 7, 3, 53)) + Val(gTotal(0))
            gTotal(1) = Val(Spread_Cell_Total(spprint, 4, 7, 4, 53)) + Val(gTotal(1))
            gTotal(2) = Val(Spread_Cell_Total(spprint, 5, 7, 5, 53)) + Val(gTotal(2))
            gTotal(3) = Val(Spread_Cell_Total(spprint, 6, 7, 6, 53)) + Val(gTotal(3))
            gTotal(4) = Val(Spread_Cell_Total(spprint, 7, 7, 7, 53)) + Val(gTotal(4))
            gTotal(5) = Val(Spread_Cell_Total(spprint, 8, 7, 8, 53)) + Val(gTotal(5))
            GoTo pprint
          End If
          rs.MoveNext
        Loop
      End If
      rsDetails.MoveNext
    Loop
    If rsDetails.EOF Then
      spprint.Col = 1
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
      spprint.SetText 1, i, Val(cmdView.Tag)
      spprint.SetCellBorder 2, i, 2, i + 1, CellBorderIndexLeft, 0, CellBorderStyleSolid
      spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
      spprint.SetText 2, i, Trim$(Item(0))
      spprint.Col = 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
      spprint.SetText 2, i + 1, Trim$(Item(1))
      spprint.SetCellBorder 1, i + 2, spprint.MaxCols, i + 2, CellBorderIndexTop, 0, CellBorderStyleSolid
      i = i + 2
      spprint.AddCellSpan 1, i, 2, 1
      spprint.Col = 1
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignRight
spprint.Font = "Arial"
spprint.FontBold = True

      spprint.SetText 1, i, "Grant Total : "
      For Cnt = 1 To 6
        spprint.SetCellBorder Cnt + 2, i, Cnt + 2, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
        spprint.Col = Cnt + 2
spprint.Row = i
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignLeft
spprint.Font = "Arial"
spprint.FontBold = True
        Qty = ""
        Qty = Val(Spread_Cell_Total(spprint, Cnt + 2, 7, Cnt + 2, spprint.MaxRows))
        spprint.SetText Cnt + 2, i, Val(Val(Qty) + Val(gTotal(Cnt - 1)))
      Next Cnt
      cmdExit.Tag = 1
i = i + 1
    End If
pprint:
    spprint.SetCellBorder 1, i, spprint.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
    CmdClear.Tag = Val(CmdClear.Tag) + 1
    spprint.Col = 1
spprint.Row = 54
spprint.Font = "Arial"
spprint.FontBold = True
spprint.TypeVAlign = TypeVAlignCenter
spprint.TypeHAlign = TypeHAlignCenter
    spprint.SetText 1, 54, gCompName
    spprint.SetText 7, 54, "Page " & IIf(Round(i) = 0, 1, Round(i)) & " of " & Val(CmdClear.Tag)
    spprint.PrintOrientation = PrintOrientationPortrait
    spprint.PrintPageOrder = PageOrderDownThenOver
    spprint.PrintFirstPageNumber = 1
    spprint.PrintPageStart = 1
    spprint.PrintScalingMethod = PrintScalingMethodSmartPrint
    spprint.BorderStyle = BorderStyleFixedSingle
    spprint.PrintBorder = True
    spprint.PrintCenterOnPageH = True
    spprint.PrintCenterOnPageV = True
    spprint.PrintMarginTop = 500
    spprint.PrintMarginLeft = 50
    spprint.PrintMarginRight = 1.25
    spprint.Action = ActionPrint
    spprint.ClearRange 1, 7, spprint.MaxCols, 53, True
    If cmdExit.Tag = 1 Then
      GoTo sPrint
    Else
      For i = 7 To 53
        spprint.RemoveCellSpan 1, i
      Next i
      spprint.SetCellBorder 1, 7, spprint.MaxCols, 53, CellBorderIndexLeft, 0, CellBorderStyleBlank
      spprint.SetCellBorder 1, 7, spprint.MaxCols, 53, CellBorderIndexRight, 0, CellBorderStyleBlank
      spprint.SetCellBorder 1, 7 + 1, spprint.MaxCols, 53, CellBorderIndexTop, 0, CellBorderStyleBlank
      spprint.SetCellBorder 1, 7, spprint.MaxCols, 53, CellBorderIndexBottom, 0, CellBorderStyleBlank
      spprint.ClearRange 1, 7, spprint.MaxCols, 53, True
      i = 7
Cnt = 0
      GoTo rec
    End If
sPrint:
  
  sPrint = True
  sShowMessage "Print Job Over."
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDetails = Nothing
End Try
End Function

Private Function sPrintcurrent() As Boolean
'  Dim Sql As String
'  Dim rs As ADODB.Recordset
'
'  Dim CurDate As String
'  Dim Cnt As Integer
'
'  Dim ToMonthyear As Date
'  Dim ToDate As Date

'  Dim Name As String
'  Dim iCode As String
'  Dim rCnt As Integer
'  Dim var As String
'  Dim tmp As Variant
'  Dim cval As String
'  Dim vCnt As Integer
  
  Try
  
  Select Case CmbReport.ListIndex
  Case 0
    If CustomerwiseItemwiseDisplay() = False Then
      GoTo ExitHere
    End If
  Case 1
    If ItemwiseCustomerwiseDisplay() = False Then
      GoTo ExitHere
    End If
  Case Else
    'No Data to Fill
  End Select

  
'  sPrintcurrent = False
'
'  If OptCustSelect = True Then
'    If SprCust.SelectionCount < 1 Then
'      sShowMessage "Please select atlease one item"
'      GoTo ExitHere
'    End If
'  Else
'    If optItemSelective = True Then
'      If sprItem.SelectionCount < 1 Then
'        sShowMessage "Please select atlease one item"
'        GoTo ExitHere
'      End If
'    End If
'  End If
'
'  CurDate = ServerDate
'
'  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text), DTPFrom.Value), "yyyy-MM")
'  ToDate = Format(ToMonth, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))
'
''  NoofDays = Format(ToMonth, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonth, "yyyy-MM-" & "01"))))
''  todate = NoofDays
''  tmp = Month(ToMonth)
''  cval = MonthName(tmp)
''  vCnt = tmp + Val(CmbMonth.Text)
'
'  Sql = " Select distinct Month(a.InvoiceDate) MthCode, a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription,"
'  Sum(B.InvoiceQty) As Qty
'  Sql = Sql & " From SAL_Invoice_hdr a"
'  Sql = Sql & " Left Outer Join SAL_Invoice_Dtl b"
'  Sql = Sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
'  Sql = Sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
'  Sql = Sql & " Left Outer Join FAS_Party_Master c"
'  Sql = Sql & " On c.PartyCode = a.CustomerCode and c.BranchId = b.BranchId and c.CompId = b.CompId"
'  Sql = Sql & " Left Outer Join INV_Material_Master d"
'  Sql = Sql & " On d.ItemCode = b.ItemCode and d.BranchId = c.BranchId and d.CompId = c.CompId"
'  Sql = Sql & " Where a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
'  Sql = Sql & " and a.InvoiceDate between '" & Format(DTPFrom, "yyyy-MM-01") & "' and '" & Format(ToDate, "yyyy-MM-dd") & "'"
'  If OptCustSelect = True Then
'    With SprCust
'      Sql = Sql & " and a.CustomerCode in "
'      Sql = Sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
'    End With
'  Else
'    If optItemSelective = True Then
'      With sprItem
'        Sql = Sql & " and b.ItemCode in "
'        Sql = Sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
'      End With
'    End If
'  End If
'  Sql = Sql & " Group by Month(a.InvoiceDate), a.CustomerCode, c.PartyName, b.ItemCode, d.ItemDescription"
'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then
'    GoTo ExitHere
'  End If
'  If rs.EOF Then
'    MsgBox "No Records Found. "
'    GoTo ExitHere
'  End If
'
'  pheight = 75
'  pPaper = vbPRPSFanfoldUS
'  Orient = 1
'  iPageWidth = 132
'  iLineNo = 1
'  fOpenPort Me
'
'  sSendToPort ""
'  sSendToPort CAlign(iPageWidth, gCompName)
'  sSendToPort ""
'
'  Sql = CAlign(iPageWidth, "Sales Report")
'  Sql = Sql & Chr(13) & CAlign(iPageWidth, "[" & Format(DTPFrom, "yyyy-MM-01") & "' and '" & Format(ToMonth, "yyyy-MM-dd") & "']")
'  gStrReportHeading = Sql
'
'  sSendToPort gStrReportHeading
'  sSendToPort CAlign(iPageWidth, "[" & Format(DTPFrom, "yyyy-MM-01") & "' and '" & Format(ToMonth, "yyyy-MM-dd") & "']")
'  Call sPageHeader
'
'  If CmbReport.ListIndex = 0 Then
'    Cnt = 1
'    Do While Not rs.EOF
'      If Name <> rs("Partyname") Then
'        If Cnt > 1 Then
'          sSendToPort ""
'          sSendToPort "Customer Name :" & LAlign(30, rs("Partyname"))
'        End If
'      End If
'      sSendToPort ""
'      sSendToPort LAlign(20, CheckNull(rs("ItemCode"))) & Space(2) & LAlign(82, CheckNull(rs("Qty")))
'      sSendToPort LAlign(30, CheckNull(rs("ItemDescription")))
'      Name = CheckNull(rs("Partyname"))
'      Cnt = Cnt + 1
'      rs.MoveNext
'    Loop
'  Else
'    If CmbReport.ListIndex = 1 Then
'      Cnt = 1
'      Do While Not rs.EOF
'        If iCode <> rs("ItemCode") Then
'          If Cnt > 1 Then
'            sSendToPort ""
'            sSendToPort "ItemCode :" & LAlign(20, rs("ItemCode"))
'            sSendToPort LAlign(30, rs("ItemDescription"))
'          End If
'        End If
'        sSendToPort LAlign(30, CheckNull(rs("PartyName"))) & Space(2) & LAlign(82, CheckNull(rs("Qty")))
'        iCode = CheckNull(rs("itemCode"))
'        Cnt = Cnt + 1
'        rs.MoveNext
'      Loop
'    End If
'   End If
'  Call EndofReport(iPageWidth, Cnt - 1) 'CStr(rs.RecordCount)

  sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub CmbReport_Click()
    
  Try
  Me.MousePointer = vbHourglass

  If CmbReport.ListIndex = 0 Then
    frastep(4).Enabled = False
    ChkTotal.Value = 0
'    FraStep(2).Move 4335, 3000
'    FraStep(3).Move 4335, 3660
'    fraSprCust.Move 7065, 2805
'    FraSprItem.Move 7065, 4365
  ElseIf CmbReport.ListIndex = 1 Then
    frastep(4).Enabled = True
'    FraStep(2).Move 4335, 3660
'    FraStep(3).Move 4335, 3000
''    fraSprCust.Move 7065, 4365
'    FraSprItem.Move 7065, 2805
  End If
  
  optCustAll.Value = True
  optItemAll.Value = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdClear_Click()
' On Error Resume Next
  CmdStep(2).Enabled = False
  CmdStep(3).Enabled = False
  fraSprCust.Visible = False
  FraSprItem.Visible = False
  optCustAll.Value = True
  optItemAll.Value = True
  Call StepClear(Me)
End Sub

Private Sub cmdExit1_Click()
' On Error Resume Next
  
  Frame2.Height = 885
  Frame2.Width = 1095
  spview.Height = 525
  spview.Width = 705
  Frame2.Visible = False
  Call CmdClear_Click
  
End Sub

Private Sub cmdprint_Click()
  
  
  Try
  Me.MousePointer = vbHourglass

  If CmbReport.ListIndex = 0 Then
    If sPrint = False Then
 GoTo ExitHere
End If
  Else
    If ItemwiseCustomerwiseDisplay = False Then
 GoTo ExitHere
End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If frastep.UBound > Index Then
    If OptCustSelect = True Then
      If SprCust.SelectionCount >= 1 Then
        fraSprCust.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  Else
    If optItemSelective = True Then
      If sprItem.SelectionCount >= 1 Then
        FraSprItem.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  End If
End Sub

Private Sub Cmdview_Click()

  Try
  Me.MousePointer = vbHourglass
  
  If sPrintcurrent() = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
  
  frastep(4).Enabled = False
  fraSprCust.Visible = False
  FraSprItem.Visible = False
  CmbMonth.ListIndex = 5
  CmbReport.ListIndex = 0
  cmbReportType.ListIndex = 0
  dtpFrom = Format(ServerDate, "mm-yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub OptCustAll_Click()
' On Error Resume Next
  CmdStep(2).Enabled = True
  fraSprCust.Visible = False
  Call Spread_Clear(SprCust)
  SprCust.MaxRows = 1
End Sub

Private Sub OptCustSelect_Click()
  Dim sql As String
  Dim rCnt As Integer
  Dim rs As AceADODB.Recordset
  Dim ToMonthyear As Date
  Dim toDate As Date

  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(2).Enabled = False
  
  If CmbReport.ListIndex = 1 Then
    If optItemSelective = True Then
      If sprItem.SelectionCount < 1 Then
        sShowMessage "Please Select atlease one item"
        GoTo ExitHere
      End If
    End If
  End If
  
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))
  
  sql = " Select distinct a.CustomerCode, c.PartyName "
  sql = sql & " From SAL_Invoice_hdr a "
  If optItemSelective.Value = True Then
    sql = sql & " Left Outer Join SAL_Invoice_Dtl b"
    sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
    sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  End If
  sql = sql & " Left Outer Join FAS_Party_Master c "
  sql = sql & " On c.PartyCode = a.CustomerCode and c.BranchId = a.BranchId and c.CompId = a.CompId "
  sql = sql & " Where a.InvoiceDate Between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "' and a.BranchId = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If optItemSelective.Value = True Then
    sql = sql & " and b.ItemCode in "
    sql = sql & Spread_Get_SelectionList(sprItem, eItem.mItemCode1)
  End If
 
  sql = sql & " Order by c.PartyName "
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
 
  If Not rs.EOF Then
      SprCust.ClearSelection
      SprCust.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        SprCust.SetText eCust.mPartyCode, rCnt, rs("CustomerCode")
        SprCust.SetText eCust.mPartyName, rCnt, rs("PartyName")
        rs.MoveNext
      Next rCnt
    fraSprCust.Enabled = True
    fraSprCust.Visible = True
  Else
    fraSprCust.Visible = False
    sShowMessage "No Records Found."
    optCustAll.Value = True
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optItemAll_Click()
' On Error Resume Next
  CmdStep(3).Enabled = True
  FraSprItem.Visible = False
  Call Spread_Clear(sprItem)
  sprItem.MaxRows = 1
End Sub

Private Sub optItemSelective_Click()
  Dim sql As String
  Dim rCnt As Integer
  Dim rs As AceADODB.Recordset
  Dim ToMonthyear As Date
  Dim toDate As Date
 
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(3).Enabled = False
  
  If CmbReport.ListIndex = 0 Then
    If OptCustSelect = True Then
       If SprCust.SelectionCount < 1 Then
        sShowMessage "Please Select Atlease One Customer"
        GoTo ExitHere
      End If
    End If
  End If
   
  ToMonthyear = Format(DateAdd("m", Val(CmbMonth.Text - 1), dtpFrom.Value), "yyyy-MM")
  toDate = Format(ToMonthyear, "yyyy-MM-") & Day(DateAdd("d", -1, DateAdd("m", 1, Format(ToMonthyear, "yyyy-MM-" & "01"))))
  
  sql = " Select distinct b.ItemCode, d.ItemDescription "
  sql = sql & " From SAL_Invoice_Dtl b "
  sql = sql & " Left Outer Join Sal_Invoice_Hdr a"
  sql = sql & " On b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo and b.invoiceDate = a.InvoiceDate"
  sql = sql & " and b.BranchId = a.BranchId and b.CompId = a.CompId"
  sql = sql & " Left Outer Join INV_Material_Master d "
  sql = sql & " On d.ItemCode = b.ItemCode and d.BranchId = b.BranchId and d.CompId = b.CompId "
  sql = sql & " Where b.InvoiceDate Between '" & Format(dtpFrom, "yyyy-MM-01") & "' and '" & Format(toDate, "yyyy-MM-dd") & "' and b.BranchId = '" & gBranchID & "' and b.Compid = '" & gCompID & "'"
  If OptCustSelect.Value = True Then
    sql = sql & " and a.CustomerCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCust.mPartyCode)
  End If
  sql = sql & " Order by b.ItemCode "
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
  
  If Not rs.EOF Then
      sprItem.ClearSelection
      sprItem.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprItem.SetText eItem.mItemCode1, rCnt, rs("ItemCode")
        sprItem.SetText eItem.mItemDescription, rCnt, rs("ItemDescription")
        rs.MoveNext
      Next rCnt
    FraSprItem.Enabled = True
    FraSprItem.Visible = True
  Else
    FraSprItem.Visible = False
    sShowMessage "No Records Found."
    optItemAll.Value = True
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SprCust_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If SprCust.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub sprItem_Click(ByVal Col As Long, ByVal Row As Long)
' On Error Resume Next
  If sprItem.SelectionCount >= 1 Then
    CmdStep(3).Enabled = True
  Else
    CmdStep(3).Enabled = False
  End If
End Sub

Private Sub Design_spview(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spview

End Sub

Private Sub Design_SprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprItem

End Sub

Private Sub Design_SprCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprCust


SprCust.Click +=  new EventHandler(SprCust_Click)

End Sub

Private Sub Design_spprint(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spprint

End Sub

Private Sub InitializeSpreads()

	Call Design_spview("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptMonthWise.frm", spview, "")

	Call Design_SprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptMonthWise.frm", SprItem, "")

	Call Design_SprCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptMonthWise.frm", SprCust, "")

	Call Design_spprint("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptMonthWise.frm", spprint, "")

End Sub
