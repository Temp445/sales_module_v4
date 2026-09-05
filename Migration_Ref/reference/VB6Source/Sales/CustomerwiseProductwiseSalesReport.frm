VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptCustomerwiseProductwiseSalesReport 
   Caption         =   "Customerwise Productwise Sales Report"
   ClientHeight    =   6120
   ClientLeft      =   1785
   ClientTop       =   1800
   ClientWidth     =   9255
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
   ScaleHeight     =   6120
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
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
      Left            =   6945
      TabIndex        =   10
      Top             =   3450
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
      Left            =   6945
      TabIndex        =   8
      Top             =   2775
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
      Left            =   6945
      TabIndex        =   5
      Top             =   2175
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
      Left            =   6945
      TabIndex        =   2
      Top             =   1485
      Width           =   345
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   780
      Index           =   3
      Left            =   4350
      TabIndex        =   24
      Top             =   3210
      Width           =   2520
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "CustomerwiseProductwiseSalesReport.frx":0000
         Left            =   90
         List            =   "CustomerwiseProductwiseSalesReport.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   345
         Width           =   2340
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Type of Invoice"
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   150
         Width           =   1110
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4350
      TabIndex        =   21
      Top             =   1275
      Width           =   2520
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103481345
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   0
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103219201
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   90
         TabIndex        =   23
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1275
         TabIndex        =   22
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Products"
      Enabled         =   0   'False
      Height          =   585
      Index           =   2
      Left            =   4350
      TabIndex        =   20
      Top             =   2625
      Width           =   2520
      Begin VB.OptionButton OptPrdrandom 
         Caption         =   "Selec&tive"
         Height          =   225
         Left            =   1515
         TabIndex        =   7
         Top             =   255
         Width           =   960
      End
      Begin VB.OptionButton OptPAll 
         Caption         =   "&All"
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Value           =   -1  'True
         Width           =   950
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Accounts"
      Enabled         =   0   'False
      Height          =   585
      Index           =   1
      Left            =   4350
      TabIndex        =   19
      Top             =   2040
      Width           =   2520
      Begin VB.OptionButton OptAccRandom 
         Caption         =   "&Selective"
         Height          =   225
         Left            =   1515
         TabIndex        =   4
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton OptAAll 
         Caption         =   "&All"
         Height          =   210
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   950
      End
   End
   Begin VB.Frame FraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4350
      TabIndex        =   18
      Top             =   3990
      Width           =   2520
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "CustomerwiseProductwiseSalesReport.frx":0088
         Left            =   90
         List            =   "CustomerwiseProductwiseSalesReport.frx":0092
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6255
      TabIndex        =   15
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   5040
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4950
      TabIndex        =   14
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   5040
      Width           =   1320
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3660
      TabIndex        =   13
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   5040
      Width           =   1300
   End
Begin AceSpread spCust
      Height          =   1035
      Left            =   7470
      TabIndex        =   16
      Top             =   2010
      Visible         =   0   'False
      Width           =   4215
      _Version        =   458752
      _ExtentX        =   7435
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
      Height          =   1275
      Left            =   7470
      TabIndex        =   17
      Top             =   3075
      Visible         =   0   'False
      Width           =   4215
      _Version        =   458752
      _ExtentX        =   7435
      _ExtentY        =   2249
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
End
Attribute VB_Name = "rptCustomerwiseProductwiseSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim k As Integer
Dim j As Long
Dim FDATE As Date
Dim TDATE As Date
Dim Head As Long

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Private Sub Header()
  Dim mstr As String
  Dim i As Integer
  Dim rCnt As Integer
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    frmPrint.sprPrint.MaxCols = 9
    
    If CustProdFlag = "C" Then
      mstr = "Customerwise Productwise Sales"
    ElseIf CustProdFlag = "P" Then
      mstr = "Productwise Customerwise Sales"
    End If
    mstr = mstr + " between " & dpfromdt & " and " & dptodt

    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
    Set rs = New Recordset
    If RSOpen(rs, Sql) = False Then GoTo ExitHere
    If rs.RecordCount > 0 Then
      frmPrint.sprPrint.AddCellSpan 1, j, 4, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If
    
    If iPageNo = 1 Then
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(1) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, mstr
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 6
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 27
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "PartNo"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 13
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "G.Value"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "G.Total"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv. Date"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 27
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Private Sub RowProcess()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
j = j + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Header
      j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Function PrintView() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim AcCode As String
  Dim AcCode1 As String
  Dim sRow As Long
  Dim Sql As String
  Dim i As Long
  Dim tqty As Double
  Dim tGV As Double
  Dim Gtot As Double
  Dim Gtot1 As Double
  Dim Gtot2 As Double
  Dim mstr As String
  
  Try
  PrintView = False
  
  If OptAccRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
  End If
  
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
  End If
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = "Select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName "
  Sql = Sql + " from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
  Sql = Sql + " and b.Compid=a.Compid) PartyName,"
  Sql = Sql + " g.Partno, a.ItemCode, c.ItemDescription, c.RevisionNo,"
  Sql = Sql + " a.InvoiceQty,a.Rate,(a.InvoiceQty*a.Rate) GV,a.ItemGroupKey, f.UomDesc, A.Totalamount from SalesInvoice a"
  Sql = Sql + " Left outer join INV_MATERIAL_MASTER c"
  Sql = Sql + " on c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1"
  Sql = Sql + " Left outer join MATERIAL_MASTER f"
  Sql = Sql + " on f.ItemCode=a.ItemCode and f.Branchid=a.Branchid and f.Compid=a.Compid "
  Sql = Sql + " Left outer join SAL_CUSTOMER_ITEMMASTER g on g.ItemCode=a.ItemCode and a.CustomerCode = g.CustomerCode"
  Sql = Sql + " and g.Branchid=a.Branchid and g.Compid=a.Compid and g.status = 1"
  Sql = Sql + " where a.InvoiceDate BETWEEN CONVERT(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111) and "
  Sql = Sql + " CONVERT(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
  If OptAccRandom Then
'    If spCust.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
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
'    If spItem.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
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
  Sql = Sql + " and a.Branchid ='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
  If cmbType.Text <> "All" Then
 Sql = Sql + " and a.Type = '" & Trim$(cmbType.Text) & "'"
End If
  If CustProdFlag = "C" Then
    Sql = Sql + " Order By PartyName,InvoiceDate,InvoiceType,InvoiceNo,a.ItemCode"
  Else
    Sql = Sql + " Order By a.ItemCode,InvoiceDate,InvoiceType,InvoiceNo"
  End If
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    Screen.MousePointer = vbNormal
    PrintView = False
    GoTo ExitHere
  End If
  
    If CustProdFlag = "C" Then
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
    Else
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
    End If
    
    Header
    
    j = j + 1
GridCount = GridCount + 1

    rsDetails.MoveLast
rsDetails.MoveFirst
  
    tqty = 0
tGV = 0
Gtot = 0
    
    Do While Not rsDetails.EOF
      
      If j > 500 Then
        frmPrint.sprPrint.MaxRows = j + 4
      End If
      If CustProdFlag = "C" Then
        AcCode1 = rsDetails("PartyName")
      Else
        AcCode1 = rsDetails("ItemCode") & rsDetails("ItemGroupKey")
      End If
      If AcCode <> AcCode1 Then
        If tGV > 0 Then
        
          frmPrint.sprPrint.AddCellSpan 1, j, 7, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
          
          frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(tGV, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
          
          frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(tqty, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
          
          tGV = 0
tqty = 0
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 2
GridCount = GridCount + 2
RowProcess
        End If
      End If
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      If GridCount = 53 Then
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("Invoicetype") & "-" & rsDetails("Invoiceno")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("PartyName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("Itemcode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("Partno")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("InvoiceQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("GV"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("Totalamount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

'      If gReportID = "PMI" And rsDetails("RevisionNo") <> "" Then
'        sSendToPort Space(11) & LAlign(45, rsDetails("RevisionNo"))
'      End If
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
  
      
      frmPrint.sprPrint.SetText 1, j, Format(rsDetails("InvoiceDate"), "dd/MM/yyyy")
      frmPrint.sprPrint.SetText 3, j, rsDetails("Itemdescription")

      If CustProdFlag = "P" Then
        AcCode = rsDetails("ItemCode") & rsDetails("ItemGroupKey")
      Else
        AcCode = rsDetails("PartyName")
      End If
      mstr = (rsDetails("InvoiceType") & rsDetails("InvoiceNo"))
      tGV = tGV + rsDetails("GV")
      tqty = tqty + rsDetails("Totalamount")
      Gtot = Gtot + rsDetails("GV")
      Gtot1 = Gtot1 + rsDetails("Totalamount")
      Gtot2 = Gtot2 + rsDetails("InvoiceQty")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      rsDetails.MoveNext
    Loop
    If rsDetails.EOF Then
      If tGV > 0 Then
        frmPrint.sprPrint.MaxRows = j + 4
        
        frmPrint.sprPrint.AddCellSpan 1, j, 7, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        
        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(tGV, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        
        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(tqty, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)

        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
        
        frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Grand Total : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        
        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(Gtot2, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        
        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(Gtot, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        
        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(Gtot1, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontSize = 8
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(259, 259, 0)
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
      tqty = 0
tGV = 0
    End If
    If CustProdFlag = "C" Then
      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColHidden = False
    Else
      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColHidden = True
    End If
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
  totalpages = iPageNo
  rsDetails.Close
  frmPrint.Show
  
  PrintView = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

 Private Function Sales() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim AcCode As String
  Dim sRow As Long
  Dim Sql As String
  Dim i As Long
  Dim tqty As Double
  Dim tqty1 As Double
  Dim tGV As Double
  Dim Gtot As Double
  Dim Gtot1 As Double
  Dim mstr As String
  
  Try
  Set rsDetails = New AceADODB.Recordset

  Sql = "Select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName "
  Sql = Sql + " from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
  Sql = Sql + " and b.Compid=a.Compid) PartyName,a.ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c"
  Sql = Sql + " where c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) ItemDescription,"
  Sql = Sql + " (select c.RevisionNo from INV_MATERIAL_MASTER c"
  Sql = Sql + " where c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) RevisionNo,"
  Sql = Sql + " a.InvoiceQty,a.Rate,(a.InvoiceQty*a.Rate) GV,a.ItemGroupKey,a.Totalamount from SalesInvoice a"
  Sql = Sql + " where a.InvoiceDate BETWEEN CONVERT(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
  Sql = Sql + " CONVERT(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
  If OptAccRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
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
    Sql = Sql & " and a.ItemCode in("
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
  Sql = Sql + " and a.Branchid ='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
  If cmbType.Text <> "All" Then
 Sql = Sql + " and a.Type = '" & Trim$(cmbType.Text) & "'"
End If
  If CustProdFlag = "C" Then
    Sql = Sql + " Order By PartyName,InvoiceDate,InvoiceType,InvoiceNo,ItemCode"
  Else
    Sql = Sql + " Order By ItemCode,InvoiceDate,InvoiceType,InvoiceNo"
  End If
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    Screen.MousePointer = vbNormal
    Sales = False
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
  If CustProdFlag = "C" Then
    mstr = "Customerwise Productwise Sales"
  ElseIf CustProdFlag = "P" Then
    mstr = "Productwise Customerwise Sales"
  End If
  mstr = mstr + " between " & dpfromdt & " and " & dptodt
  sSendToPort CAlign(iPageWidth, mstr)
  sSendToPort ""

  sPageHeader
  i = 1
  tqty = 0
tGV = 0
Gtot = 0
  Do While Not rsDetails.EOF
    If CustProdFlag = "C" Then
      If AcCode <> rsDetails("PartyName") Then
        If tGV > 0 Then
          sSendToPort Space(11) & LAlign(45, "Sub Total") & Space(10) & RAlign(11, Format(tGV, "0.00"))
          sSendToPort Space(11) & LAlign(45, "Inv Total") & Space(10) & RAlign(11, Format(tqty1, "0.00"))
          tGV = 0
tqty1 = 0
          sSendToPort ""
        End If
        If CustProdFlag = "C" Then
          If i > 1 Then
 sSendToPort String(iPageWidth, "-")
End If
        End If
        sSendToPort ""
        sSendToPort "Customer :" & rsDetails("PartyName")
        sSendToPort ""
      End If
      If mstr <> (rsDetails("InvoiceType") & rsDetails("InvoiceNo")) Then
        sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(45, Trim(rsDetails("ItemCode") & "[" & Trim(rsDetails("ItemGroupKey") & "]"))) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(11, Format(rsDetails("GV"), "0.00"))
        sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(1) & LAlign(45, CheckNull(rsDetails("ItemDescription"))) & Space(1) & RAlign(8, Format(rsDetails("Rate"), "0.00")) & Space(1) & RAlign(11, Format(rsDetails("Totalamount"), "0.00"))
      Else
        sSendToPort Space(11) & LAlign(45, Trim(rsDetails("ItemCode") & "[" & Trim(rsDetails("ItemGroupKey") & "]"))) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(11, Format(rsDetails("GV"), "0.00"))
        sSendToPort Space(11) & LAlign(45, rsDetails("ItemDescription")) & Space(1) & RAlign(8, Format(rsDetails("Rate"), "0.00")) & Space(1) & RAlign(11, Format(rsDetails("Totalamount"), "0.00"))
      End If
      
      If gReportID = "PMI" And rsDetails("RevisionNo") <> "" Then
        sSendToPort Space(11) & LAlign(45, rsDetails("RevisionNo"))
      End If
      
      AcCode = rsDetails("PartyName")
      mstr = (rsDetails("InvoiceType") & rsDetails("InvoiceNo"))
      tGV = tGV + rsDetails("GV")
      Gtot = Gtot + rsDetails("GV")
      Gtot1 = Gtot1 + rsDetails("Totalamount")
      tqty1 = tqty1 + rsDetails("Totalamount")
      sSendToPort ""
    ElseIf CustProdFlag = "P" Then
      If AcCode <> rsDetails("ItemCode") & rsDetails("ItemGroupKey") Then
        If tqty > 0 Then
          sSendToPort ""
          sSendToPort Space(11) & LAlign(45, "Sub Total") & Space(1) & RAlign(8, Format(tqty, "0.00")) & Space(1) & RAlign(11, Format(tGV, "0.00"))
          sSendToPort Space(11) & LAlign(45, "Inv Total") & Space(2) & RAlign(19, Format(tqty1, "0.00"))
          tqty = 0
tGV = 0
tqty1 = 0
          sSendToPort ""
        End If
        sSendToPort "Item Description :" & rsDetails("ItemDescription") & "( " & Trim(rsDetails("ItemCode")) & " )" & "[" & rsDetails("ItemGroupKey") & "]"
      End If
      If gReportID = "PMI" Then
        sSendToPort "Revision No : " & rsDetails("RevisionNo")
      End If
      sSendToPort ""
      sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(45, rsDetails("CustomerCode")) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(11, Format(rsDetails("GV"), "0.00"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(1) & LAlign(45, rsDetails("PartyName")) & Space(1) & RAlign(8, Format(rsDetails("Rate"), "0.00")) & RAlign(12, Format(rsDetails("Totalamount"), "0.00"))
      AcCode = rsDetails("ItemCode") & rsDetails("ItemGroupKey")
      tqty = tqty + rsDetails("InvoiceQty")
      tGV = tGV + rsDetails("GV")
      Gtot = Gtot + rsDetails("GV")
      Gtot1 = Gtot1 + rsDetails("Totalamount")
      tqty1 = tqty1 + rsDetails("Totalamount")
    End If
    i = i + 1
    rsDetails.MoveNext
  Loop
  If rsDetails.EOF Then
    If CustProdFlag = "C" Then
      If tGV > 0 Then
        sSendToPort Space(11) & LAlign(45, "Sub Total") & Space(10) & RAlign(11, Format(tGV, "0.00"))
        sSendToPort Space(11) & LAlign(45, "Grand Total(Sub)") & Space(10) & RAlign(11, Format(Gtot, "0.00"))
        sSendToPort Space(11) & LAlign(45, "Grand Total(Inv)") & Space(10) & RAlign(11, Format(Gtot1, "0.00"))
      End If
    Else
      If tGV > 0 Then
        sSendToPort ""
        sSendToPort Space(11) & LAlign(45, "Sub Total") & Space(1) & RAlign(8, Format(tqty, "0.00")) & Space(1) & RAlign(11, Format(tGV, "0.00"))
        sSendToPort Space(11) & LAlign(45, "Grand Total(Sub)") & Space(10) & RAlign(11, Format(Gtot, "0.00"))
        sSendToPort Space(11) & LAlign(45, "Grand Total(Inv)") & Space(10) & RAlign(11, Format(Gtot1, "0.00"))
      End If
    End If
    tqty = 0
tGV = 0
  End If
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  Sales = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sPageHeader()

  Try
  Me.MousePointer = vbHourglass

  Try
  
'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  If CustProdFlag = "C" Then
    sSendToPort LAlign(10, "Inv. No") & Space(1) & LAlign(45, "Item Code") & Space(1) & RAlign(8, "Quantity") & Space(1) & RAlign(11, "Goods Value")
    sSendToPort LAlign(10, "Inv. Date") & Space(1) & LAlign(45, "Item Description") & Space(1) & RAlign(8, "Rate") & Space(1) & RAlign(11, "Inv Value")
    If gReportID = "PMI" Then
      sSendToPort Space(11) & LAlign(45, "Revision No")
    End If
  Else
    sSendToPort LAlign(10, "Inv. Type") & Space(1) & LAlign(45, "Customer Code") & Space(1) & RAlign(8, "Quantity") & Space(1) & RAlign(11, "Goods Value")
    sSendToPort LAlign(10, "Inv. Date") & Space(1) & LAlign(45, "Customer Name") & Space(1) & RAlign(8, "Rate") & Space(1) & RAlign(11, "Inv Value")
  End If
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call StepClear(Me)
  Call StepEnable(Me)
  spItem.Visible = False
  spCust.Visible = False
  OptAAll.Value = True
  OptPAll.Value = True
  cmbType.ListIndex = 0
  CmdStep(1).Enabled = False
  CmdStep(2).Enabled = False
  cmbReportType.ListIndex = 0
  MDISales.Toolbar1.Buttons(1).Enabled = True
  MDISales.Toolbar1.Buttons(2).Enabled = True
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound > Index Then
    If OptAccRandom = True Then
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

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If DateValue(dptodt) < DateValue(dpfromdt) Then
    sShowMessage "To Date should be greater than From Date"
    dptodt.SetFocus
    GoTo ExitHere
  End If

  If cmbReportType.ListIndex = 0 Then
    If PrintView() = False Then GoTo ExitHere
  Else
    Set frmCurrent = Me
    DoEvents
    iLineNo = 1
    iPageNo = 1
    If Sales = False Then
 GoTo ExitHere
End If
    Call FinaliseAndShowPrintPreview
    sClosePort
    Set frmOutput = New frmReportViewer
    frmOutput.Caption = gStrReportHeading
    frmOutput.txtTotalPages.Text = CStr(iPageNo)
    frmOutput.wbrView.Navigate strFilename
    frmOutput.Tag = strFilename
    gStrReportHeading = ""
'    frmOutput.Show vbModal
    DoEvents
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
  End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim StDate As Variant
  Dim toDate As Variant
  Dim TmpDate As Variant
  Dim i As Long
  
  Try
  
  dpfromdt = Format(ServerDate, "dd-mm-yyyy")
  StDate = "01" + "-" + Mid$(dpfromdt, 4, 7)
  toDate = Mid$(dpfromdt, 4, 7)
  i = 31
  Do While True
    TmpDate = Trim$(str(i)) + "-" + toDate
    If IsDate(TmpDate) Then Exit Do
    i = i - 1
  Loop
  dptodt = TmpDate
  cmbType.ListIndex = 0
  If CustProdFlag = "C" Then
    rptCustomerwiseProductwiseSalesReport.Caption = "Customerwise Productwise Sales"
  ElseIf CustProdFlag = "P" Then   'Productwise Supplierwise
    rptCustomerwiseProductwiseSalesReport.Caption = "Productwise Customerwise Sales"
  End If
  cmbReportType.ListIndex = 0
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
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub OptAAll_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(1).Enabled = True
  spCust.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OptAccRandom_Click()
  
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(1).Enabled = False

  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b where"
  Sql = Sql + " b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql + " from SAL_INVOICE_HDR a where a.InvoiceDate between"
  Sql = Sql + " convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.ClearSelection
'  spCust.MaxRows = spCust.DataRowCnt
'  spCust.Enabled = True
'  spCust.Col = 1: spCust.ColHidden = True
'  spCust.Visible = True
'  spCust.SetFocus
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    'OptAll.Value = True
    GoTo ExitHere
  Else
    spCust.Col = 1
spCust.ColHidden = True

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

Private Sub OptAccRandom_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  If KeyAscii = 13 Then SendKeys Chr(9)
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
  If OptAccRandom Then
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
'  adoList1.ConnectionString = db
'  adoList1.RecordSource = Sql
'  adoList1.Refresh
'  spItem.ClearSelection
'  spItem.MaxRows = spItem.DataRowCnt
'  'spItem.Col = 1: spItem.ColHidden = True
'  spItem.Visible = True
'  spItem.Enabled = True
'  spItem.SetFocus
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
   ' OptAll.Value = True
    GoTo ExitHere
  Else
    spItem.Enabled = True
 spItem.Visible = True
    spItem.ClearSelection
    Call ADORS_Spread_DataFill(rs, spItem, True)
    
    spItem.MaxRows = spItem.DataRowCnt
    Call Spread_Set_Focus(spItem, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OptPrdrandom_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  If KeyAscii = 13 Then SendKeys Chr(9)
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

Private Sub spCust_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spCust.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub spCust_KeyDown(sender as object, e as KeyDownEventArgs)
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

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust


spCust.Click +=  new EventHandler(spCust_Click)
spCust.KeyDown +=  new EventHandler(spCust_KeyDown)

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.Click +=  new EventHandler(spItem_Click)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerwiseProductwiseSalesReport.frm", spCust, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerwiseProductwiseSalesReport.frm", spItem, "")

End Sub
