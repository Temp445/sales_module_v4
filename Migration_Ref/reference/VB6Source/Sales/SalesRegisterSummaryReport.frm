VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesRegisterSummaryReport 
   Caption         =   "Sales Register Summary Report "
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
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   5835
      TabIndex        =   4
      Top             =   4785
      Width           =   1305
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4540
      TabIndex        =   27
      Top             =   4785
      Width           =   1300
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
      Left            =   6540
      TabIndex        =   26
      Top             =   2985
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
      Left            =   6525
      TabIndex        =   25
      Top             =   2385
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
      Left            =   6540
      TabIndex        =   24
      Top             =   1800
      Width           =   345
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
      Left            =   6540
      TabIndex        =   23
      Top             =   1110
      Width           =   345
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   4298
      TabIndex        =   21
      Top             =   4065
      Width           =   1785
      Begin VB.CheckBox chkCust 
         Caption         =   "Customer Wise"
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Products"
      Enabled         =   0   'False
      Height          =   570
      Index           =   2
      Left            =   3930
      TabIndex        =   17
      Top             =   2235
      Width           =   2520
      Begin VB.OptionButton OptPrdrandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1530
         TabIndex        =   19
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton OptPAll 
         Caption         =   "All"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   225
         Value           =   -1  'True
         Width           =   950
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Type of Invoice"
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   3930
      TabIndex        =   11
      Top             =   2790
      Width           =   2520
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "SalesRegisterSummaryReport.frx":0000
         Left            =   90
         List            =   "SalesRegisterSummaryReport.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   255
         Width           =   2340
      End
   End
   Begin VB.Frame FraStep 
      Height          =   765
      Index           =   0
      Left            =   3930
      TabIndex        =   5
      Top             =   900
      Width           =   2520
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101580801
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   7
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101580801
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1380
         TabIndex        =   9
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   570
      Index           =   1
      Left            =   3930
      TabIndex        =   0
      Top             =   1665
      Width           =   2520
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   255
         Left            =   1530
         TabIndex        =   2
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Value           =   -1  'True
         Width           =   570
      End
   End
Begin AceSpread spCust
      Height          =   1035
      Left            =   7125
      TabIndex        =   10
      Top             =   1770
      Visible         =   0   'False
      Width           =   4710
      _Version        =   458752
      _ExtentX        =   8308
      _ExtentY        =   1826
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
      Height          =   1290
      Left            =   7125
      TabIndex        =   20
      Top             =   2865
      Visible         =   0   'False
      Width           =   4710
      _Version        =   458752
      _ExtentX        =   8308
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
      SelectBlockOptions=   0
   End
   Begin VB.Frame Frame5 
      Caption         =   "Customer RR"
      Height          =   540
      Left            =   3840
      TabIndex        =   13
      Top             =   3525
      Width           =   2700
      Begin VB.OptionButton optWORR 
         Caption         =   "Without RR"
         Height          =   210
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optWRR 
         Caption         =   "with RR"
         Height          =   210
         Left            =   630
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton optRRAll 
         Caption         =   "All"
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3255
      TabIndex        =   3
      Top             =   4785
      Width           =   1300
   End
End
Attribute VB_Name = "rptSalesRegisterSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Public sprow As Long
Private Erpno As String
Private TsNo As String

Dim k As Integer
Dim j As Long
Public Sub Header()
  Try

  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
'  sSendToPort String(iPageWidth, "-")
  
  Call RptStandard_Header(5, j, Erpno, TsNo, True)
  GridCount = GridCount + 5
  Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
  Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
  spCAlign frmPrint.sprPrint, 1, j, 5, "S.No", True
  spCAlign frmPrint.sprPrint, 2, j, 10, "Invoice No", True
  spCAlign frmPrint.sprPrint, 3, j, 10, "Invoice Date", True
  spCAlign frmPrint.sprPrint, 4, j, 12, "Invoice Amount", True
  spCAlign frmPrint.sprPrint, 5, j, 10, "RR No.", True
  j = j + 1
GridCount = GridCount + 1
 
'  If optRandom.Value = 1 Then
'    sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount")
'  Else
'    sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount") & Space(3) & LAlign(39, "Customer Name")
'  End If
  
'  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Private Function sSpread_sPrintcurrentSummary() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim GrandTotal As Double
  Dim sRow As Long
  Dim subtot As Variant
  Dim L As Variant
  Dim i As Integer
  Dim Sql As String
  Dim pName As String
  Dim gvtot As Variant
  Dim edtot As Variant
  Dim cess2 As Variant
  Dim cess1 As Variant
  Dim cess3 As Variant
  Dim tmp As Variant
  Dim tot(6) As String
  Dim subtot1(100) As Variant
  Dim subtot2(100) As Variant
  Dim subtot3(100) As Variant
  Dim subtot4(100) As Variant
  Dim subtot5(100) As Variant
  Dim subtot11(1) As Variant
  Dim subtot12(1) As Variant
  Dim subtot13(1) As Variant
  Dim subtot14(1) As Variant
  Dim subtot15(1) As Variant
  Dim subtot31(1) As Variant
  Dim subtot32(1) As Variant
  Dim subtot33(1) As Variant
  Dim subtot34(1) As Variant
  Dim subtot35(1) As Variant
  Dim subtoted As Variant
  Dim subtotgv As Variant
  Dim subtotcess As Variant
  Dim subtotscess As Variant
  
  Try
  
  Set rsDetails = New AceADODB.Recordset
  
 ' Cnt = 0
  Sql = " select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount,a.rrno,a.rrdate"
  Sql = Sql + " from SAL_Invoice_Hdr a "
  Sql = Sql + " Left outer join SAL_Invoice_Dtl d on a.InvoiceType=d.InvoiceType and a.InvoiceNo=d.InvoiceNo and a.InvoiceDate=d.InvoiceDate"
  Sql = Sql + " and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql + " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.InvoiceDate BETWEEN convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
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
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If optRandom.Value = 1 Then
    Sql = Sql + " order by a.InvoiceNo,a.InvoiceDate"
  Else
    Sql = Sql + " order by a.InvoiceNo,a.InvoiceDate"
  End If
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_sPrintcurrentSummary = False
    GoTo ExitHere
  End If
  
  
      j = 1
GridCount = 1
k = 1
iPageNo = 1
pheight = 54
Orient = 1
L = 1
      gStrReportHeading = "Sales Summary  Report Between " & Format(dpfromdt, "dd/MM/yyyy") & " and " & Format(dpTodt, "dd/MM/yyyy") & ""
      Header
      
      'j = j + 1: GridCount = GridCount + 1: RowProcess1

      rsDetails.MoveLast
rsDetails.MoveFirst

      GrandTotal = 0
cmdView.Tag = 1
pName = ""
      
      Do While Not rsDetails.EOF
      
        If j > 500 Then
          frmPrint.sprPrint.MaxRows = j + 1
        End If
        
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
             

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, k
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 2, j, rsDetails("InvoiceType") & rsDetails("Invoiceno")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 3, j, rsDetails("invoicedate")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, j, rsDetails("rrno")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        
        
        '.Col = 7: .Row = j
        'subtot = subtot + Val(tot(0)) + Val(tot(4))
        'subtotgv = subtotgv + Val(tot(3))
        'subtoted = subtoted + Val(tot(2))
        'subtotcess = subtotcess + Val(tot(0))
        'subtotscess = subtotscess + Val(tot(4))
        '.SetText 7, j, Format(Val(tot(0)) + Val(tot(4)), "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .Font = "Arial": .FontBold = False
        'tot(0) = 0
        'tot(4) = 0
        'tot(2) = 0
        'tot(3) = 0
        'Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 8, j, 8, j
        GrandTotal = GrandTotal + rsDetails("Totalamount")
        pName = rsDetails("PartyName")
        rsDetails.MoveNext
        j = j + 1
GridCount = GridCount + 1
        
      If GridCount = 54 Then
        gvtot = gvtot + GrandTotal
'        edtot = edtot + subtoted
'        cess2 = cess2 + subtotcess
'        cess1 = cess1 + subtotscess
'        cess3 = cess1 + cess2
      subtot1(L) = Format(GrandTotal, "0.00")
'      subtot2(L) = Format(subtoted, "0.00")
'      subtot3(L) = Format(subtotcess, "0.00")
'      subtot4(L) = Format(subtotscess, "0.00")
'      subtot5(L) = Format(subtot, "0.00")
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 4, j, 4, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 4, j, 4, j
      frmPrint.sprPrint.SetText 4, j, Format(GrandTotal, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 4, j, 4, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 4, j, 4, j
'      .SetText 4, j, Format(subtoted, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 5, j, 5, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 5, j, 5, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 5, j, 5, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 5, j, 5, j
'      .SetText 5, j, Format(subtotcess, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 6, j, 6, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 6, j, 6, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 6, j, 6, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 6, j, 6, j
'      .SetText 6, j, Format(subtotscess, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 7, j, 7, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 7, j, 7, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 7, j, 7, j
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 7, j, 7, j
'      .SetText 7, j, Format(subtot, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
      L = L + 1
      'subtot = 0
      GrandTotal = 0
'      subtoted = 0
'      subtotcess = 0
'      subtotscess = 0
      
      End If
       k = k + 1
      RowProcess1
      cmdView.Tag = cmdView.Tag + 1
      Loop
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 8, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
      frmPrint.sprPrint.SetText 4, j, Format(GrandTotal, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
'      .SetText 4, j, Format(subtoted, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 5, j, 5, j
'      .SetText 5, j, Format(subtotcess, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 6, j, 6, j
'      .SetText 6, j, Format(subtotscess, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 7, j, 7, j
'      .SetText 7, j, Format(subtot, "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
'     Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 10, j, 10, j
     subtot11(1) = Format(GrandTotal, "0.00")
'     subtot12(1) = Format(subtoted, "0.00")
'     subtot13(1) = Format(subtotcess, "0.00")
'     subtot14(1) = Format(subtotscess, "0.00")
'     subtot15(1) = Format(subtot, "0.00")
        gvtot = gvtot + GrandTotal
'        edtot = edtot + subtoted
'        cess2 = cess2 + subtotcess
'        cess1 = cess1 + subtotscess
'        cess3 = cess1 + cess2
        subtot31(1) = Format(gvtot, "0.00")
'        subtot32(1) = Format(edtot, "0.00")
'        subtot33(1) = Format(cess2, "0.00")
'        subtot34(1) = Format(cess1, "0.00")
'        subtot35(1) = Format(cess3, "0.00")
         j = j + 2
    frmPrint.sprPrint.SetText 4, j + 1, "GRAND TOTAL"
frmPrint.sprPrint.FontBold = True
     j = j + 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, 3, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 3, j, 3, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, 3, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 3, j, 3, j
    frmPrint.sprPrint.SetText 3, j + 1, "PageNo" & Space(15) & "INVOICE.VALUE"
frmPrint.sprPrint.FontBold = True
   For i = 1 To L
   Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 3, j + i, frmPrint.sprPrint.MaxCols, j + i
    frmPrint.sprPrint.SetText 3, j + i + 1, "PageNo" & (i)
    frmPrint.sprPrint.SetText 4, j + i + 1, subtot1(i)
'    frmPrint.sprPrint.SetText 3, j + i + 1, subtot2(i)
'    frmPrint.sprPrint.SetText 4, j + i + 1, subtot3(i)
'    frmPrint.sprPrint.SetText 5, j + i + 1, subtot4(i)
'    frmPrint.sprPrint.SetText 6, j + i + 1, subtot5(i)
   Next i
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 3, j + i, frmPrint.sprPrint.MaxCols, j + i
    frmPrint.sprPrint.SetText 4, j + i, subtot11(1)
'    frmPrint.sprPrint.SetText 3, j + i, subtot12(1)
'    frmPrint.sprPrint.SetText 4, j + i, subtot13(1)
'    frmPrint.sprPrint.SetText 5, j + i, subtot14(1)
'    frmPrint.sprPrint.SetText 6, j + i, subtot15(1)
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 3, j + i + 1, frmPrint.sprPrint.MaxCols, j + i + 1
    frmPrint.sprPrint.SetText 3, j + i + 1, "Total"
    frmPrint.sprPrint.SetText 4, j + i + 1, subtot31(1)
'    frmPrint.sprPrint.SetText 3, j + i + 1, subtot32(1)
'    frmPrint.sprPrint.SetText 4, j + i + 1, subtot33(1)
'    frmPrint.sprPrint.SetText 5, j + i + 1, subtot34(1)
'    frmPrint.sprPrint.SetText 6, j + i + 1, subtot35(1)
    frmPrint.Show
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
  sSpread_sPrintcurrentSummary = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub RowProcess1()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
j = j + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Header
     ' j = j + 1: GridCount = GridCount + 1
  End If
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me
'  DoEvents
'  iLineNo = 1
'  iPageNo = 1

  If cmbType <> "Spread Format Report" Then
    If sPrintcurrentSummary = False Then
 GoTo ExitHere
End If
    Call FinaliseAndShowPrintPreview
  Else
    If sSpread_sPrintcurrentSummary = False Then
 GoTo ExitHere
End If
  End If

'  If sPrintcurrentSummary = False Then: GoTo ExitHere
'
'  Call FinaliseAndShowPrintPreview
  
'  sClosePort
'  Set frmOutput = New frmReportViewer
'  frmOutput.Caption = gStrReportHeading
'  frmOutput.txtTotalPages.Text = CStr(iPageNo)
'  frmOutput.wbrView.Navigate strFileName
'  frmOutput.Tag = strFileName
'  gStrReportHeading = ""
'  frmOutput.Show vbModal
'  DoEvents
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

  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  
  If optRandom.Value = 1 Then
    sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount")
  Else
    sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(14, "Invoice Amount") & Space(3) & LAlign(39, "RR No.")
  End If
  
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

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
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.TotalAmount,a.rrno,a.rrdate"
  Sql = Sql + " from SAL_Invoice_Hdr a "
  Sql = Sql + " Left outer join SAL_Invoice_Dtl d on a.InvoiceType=d.InvoiceType and a.InvoiceNo=d.InvoiceNo and a.InvoiceDate=d.InvoiceDate"
  Sql = Sql + " and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql + " where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.InvoiceDate BETWEEN convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
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
  gStrReportHeading = "Sales Register(Summary) between " & dpfromdt & " and " & dpTodt
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
    If optRandom.Value = 1 Then
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
      'sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00"))
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00")) & Space(9) & LAlign(10, CheckNull(rsDetails("rrno")))
    Else
      'sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00")) & Space(3) & LAlign(39, rsDetails("PartyName"))
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(14, Format(rsDetails("TotalAmount"), "0.00")) & Space(9) & LAlign(10, CheckNull(rsDetails("rrno")))
    End If
    
    tamt = tamt + rsDetails("TotalAmount")
    subtot = subtot + rsDetails("TotalAmount")
    i = i + 1
    tmp = rsDetails("PartyName")
    rsDetails.MoveNext
  Loop
  If optRandom.Value = 1 Then
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

Private Sub cmdClear_Click()
' On Error Resume Next
  
  'Clearing Form Controls
  Call ClearForm(Me)
  
  Call StepClear(Me)
  Call StepEnable(Me)
  
  spCust.Visible = False
  spItem.Visible = False
  
  optAll.Value = True
  OptPAll.Value = True
  optRRAll.Value = True
  
  cmbType.ListIndex = 0
  
'  CmdStep(1).Enabled = False
'  CmdStep(2).Enabled = False
  
  MDISales.Toolbar1.Buttons(1).Enabled = True
  MDISales.Toolbar1.Buttons(2).Enabled = True
  
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
    If OptPrdrandom = True Then
      If spItem.SelectionCount >= 1 Then
        spItem.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  End If
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try

  CurrDate Me
  cmbType.ListIndex = 0
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

Private Sub optAll_Click()
' On Error Resume Next
  CmdStep(1).Enabled = True
  spCust.Visible = False
  Frame3.Enabled = True
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
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim i As Long

  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(2).Enabled = False

  Sql = " select distinct a.ItemCode [Item Code],(select b.ItemDescription from INV_MATERIAL_MASTER b"
  Sql = Sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.compid and b.status = 1) [Item Description] "
  Sql = Sql + " from SalesInvoice a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt.Value, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt.Value, "yyyy-MM-dd") + "',111) "
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
'
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
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql + " from SalesInvoice a where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and InvoiceDate BETWEEN convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
  If cmbType <> "All" Then
    Sql = Sql + " and a.type = '" + Trim$(cmbType) + "'"
  End If
  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo=''"
  End If
'
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.ClearSelection
'  spCust.MaxRows = spCust.DataRowCnt
'  If spCust.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    optAll.SetFocus
'    GoTo ExitHere
'  End If

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
  
  Frame3.Enabled = False
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

Private Sub spItem_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spItem.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust


spCust.Click +=  new EventHandler(spCust_Click)

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.Click +=  new EventHandler(spItem_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterSummaryReport.frm", spCust, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesRegisterSummaryReport.frm", spItem, "")

End Sub
