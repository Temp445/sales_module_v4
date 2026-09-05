VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesLedgerReport 
   Caption         =   "Sales Ledger Report"
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
   Icon            =   "SalesLedgerReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
      Index           =   2
      Left            =   7035
      TabIndex        =   7
      Top             =   2685
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
      Left            =   7035
      TabIndex        =   4
      Top             =   2010
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
      Left            =   7035
      TabIndex        =   2
      Top             =   1260
      Width           =   345
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   780
      Index           =   1
      Left            =   4340
      TabIndex        =   19
      Top             =   1755
      Width           =   2625
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "SalesLedgerReport.frx":000C
         Left            =   90
         List            =   "SalesLedgerReport.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   345
         Width           =   2445
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Type of Invoice"
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1110
      End
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   585
      Index           =   2
      Left            =   4340
      TabIndex        =   18
      Top             =   2535
      Width           =   2625
      Begin VB.OptionButton OptAAll 
         Caption         =   "&All"
         Height          =   210
         Left            =   195
         MaskColor       =   &H00400000&
         TabIndex        =   5
         Top             =   255
         Value           =   -1  'True
         Width           =   950
      End
      Begin VB.OptionButton OptAccRandom 
         Caption         =   "&Selective"
         Height          =   225
         Left            =   1545
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4340
      TabIndex        =   15
      Top             =   990
      Width           =   2625
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1305
         TabIndex        =   1
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   108003329
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   0
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   108003329
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1290
         TabIndex        =   17
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   16
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Frame FraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4348
      TabIndex        =   10
      Top             =   3120
      Width           =   2625
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "SalesLedgerReport.frx":0093
         Left            =   90
         List            =   "SalesLedgerReport.frx":009D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   2430
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6300
      TabIndex        =   13
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   4290
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5010
      TabIndex        =   12
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   4290
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3720
      TabIndex        =   11
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   4290
      Width           =   1300
   End
Begin AceSpread spList
      Height          =   1545
      Left            =   7425
      TabIndex        =   14
      Top             =   2655
      Width           =   4725
      _Version        =   458752
      _ExtentX        =   8334
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
End
Attribute VB_Name = "rptSalesLedgerReport"
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

Private Sub RowProcess()
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
      j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Sub Header()
  Dim mstr As String
  Dim i As Integer
  Dim rCnt As Integer
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    frmPrint.sprPrint.MaxCols = 8
    Orient = 2
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    frmPrint.sprPrint.Col = 6
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
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      
      frmPrint.sprPrint.Col = 6
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

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, gCompName
frmPrint.sprPrint.Col = 1
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sales Ledger Report between " & dpfromdt & " and " & dptodt
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code/Description"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Excise"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Cess"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Vat"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1

    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "No"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
'    .Col = 2: .Row = j

'    .SetText .Col, j, "Description": .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True
    
    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Duty"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Amount"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function Sales() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim GrandTotal As Double
  Dim sRow As Long
  Dim i As Integer
  Dim Sql As String
  Dim pName As String
  Dim tmp As Variant
  Dim tot(6) As String
  
  Try

  Set rsDetails = New AceADODB.Recordset
  
  If OptAccRandom Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atleast one item"
GoTo ExitHere
End If
  End If
  
  Sql = " Select distinct a.InvoiceType, a.Invoiceno, A.Invoicedate,a.ItemCode,c.ItemDescription, a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode"
  Sql = Sql + " and b.Branchid='" & gBranchID & "' and b.Compid='" & gCompID & "') PartyName,sum(a.TotalAmount) TotalAmount From SalesInvoice a  "
  Sql = Sql + " Left outer join INV_MATERIAL_MASTER c on c.ItemCode=a.ItemCode"
  Sql = Sql + " and c.Branchid=a.Branchid and c.Compid=a.Compid"
  Sql = Sql + " Where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
  If OptAccRandom Then
'    If spList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql + "'" + Trim$(spList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'" 'and a.Finyear = '" & gFinyear & "'"
  Sql = Sql + " group By a.InvoiceType, a.Invoiceno, A.Invoicedate,a.ItemCode,c.ItemDescription, a.CustomerCode order by PartyName"
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    Screen.MousePointer = vbNormal
    Sales = False
    GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 1 Then
    pheight = 75
    pPaper = vbPRPSA4
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
  
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "Sales Ledger Report between " & dpfromdt & " and " & dptodt)
    sSendToPort ""
    sPageHeader
    i = 1
    GrandTotal = 0
    Do While Not rsDetails.EOF
      sSendToPort LAlign(5, CStr(i)) & Space(1) & LAlign(45, rsDetails("PartyName")) & Space(1) & RAlign(12, Format(rsDetails("TotalAmount"), "00.00"))
      sSendToPort ""
      i = i + 1
      GrandTotal = GrandTotal + rsDetails("Totalamount")
      rsDetails.MoveNext
    Loop
    sSendToPort Space(6) & LAlign(45, "Total :") & Space(1) & RAlign(12, Format(GrandTotal, "00.00"))
    rsDetails.Close
  
    Call EndofReport(iPageWidth, i - 1)
  Else
  
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
      
      Header
      
      j = j + 1
GridCount = GridCount + 1
RowProcess

      rsDetails.MoveLast
rsDetails.MoveFirst

      GrandTotal = 0
cmdview.Tag = 1
pName = ""
      
      Do While Not rsDetails.EOF
      
        If j > 500 Then
          frmPrint.sprPrint.MaxRows = j + 1
        End If
        
        If Trim$(pName) <> rsDetails("PartyName") Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          If Val(cmdview.Tag) > 1 Then
            j = j + 1
GridCount = GridCount + 1
RowProcess
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          End If
          frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, rsDetails("PartyName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceType") & rsDetails("Invoiceno")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 2, j, rsDetails("ItemCode") & " - " & rsDetails("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        Sql = " SELECT distinct Exprname, Resultvalue FROM salesCalculationDetails "
        Sql = Sql & " where Doctype = '" & Trim$(rsDetails("InvoiceType")) & "' and Docno = " & rsDetails("Invoiceno") & " "
        Sql = Sql & " and Docdate = '" & Format$(rsDetails("Invoicedate"), "yyyy-MM-dd") & "'"
        Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
        Set rs = New AceADODB.Recordset
        If RSOpen(rs, Sql) = False Then GoTo ExitHere
        If rs.RecordCount > 0 Then
          If Trim$(pName) <> rsDetails("PartyName") Then
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          End If
          For i = 1 To rs.RecordCount
            If Trim$(rs("Exprname")) = "CESS" Or Trim$(rs("Exprname")) = "SHE CESS 1%" Then
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 5, j, 5, j
              frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
              frmPrint.sprPrint.GetText frmPrint.sprPrint.Col, j, tmp
              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(Val(tmp) + rs("Resultvalue"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
              tot(0) = Val(tot(0)) + Val(Format(Val(tmp) + rs("Resultvalue"), "00.00"))
            ElseIf Trim$(rs("Exprname")) = "EXCISE DUTY" Then
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 4, j, 4, j
              frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Resultvalue"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
              tot(2) = Val(tot(2)) + Val(rs("Resultvalue"))
            ElseIf Trim$(rs("Exprname")) = "GOODS Value" Then
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, 4, j
              frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Resultvalue"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
              tot(3) = Val(tot(3)) + Val(rs("Resultvalue"))
            ElseIf Trim$(rs("Exprname")) = "Sub Total" Then
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 7, j, 7, j
              frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Resultvalue"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
              tot(4) = Val(tot(4)) + Val(rs("Resultvalue"))
            ElseIf Trim$(rs("Exprname")) = "VAT" Then
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 6, j, 6, j
              frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Resultvalue"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
              tot(5) = Val(tot(5)) + Val(rs("Resultvalue"))
            End If
            rs.MoveNext
          Next i
        End If
        
        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Format(rsDetails("TotalAmount"), "00.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = False
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 8, j, 8, j
        GrandTotal = GrandTotal + rsDetails("Totalamount")
        pName = rsDetails("PartyName")
        rsDetails.MoveNext
        j = j + 1
GridCount = GridCount + 1
RowProcess
cmdview.Tag = cmdview.Tag + 1
      Loop
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 8, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Total : "
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, frmPrint.sprPrint.Col, j, frmPrint.sprPrint.Col, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, frmPrint.sprPrint.MaxCols, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Col = frmPrint.sprPrint.MaxCols
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.MaxCols, j, Format(GrandTotal, "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, i + 2, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 3, j, Format(Val(tot(3)), "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(Val(tot(2)), "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(Val(tot(0)), "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(Val(tot(5)), "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(Val(tot(4)), "00.00")
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
      frmPrint.Show
    totalpages = iPageNo
  End If
  
  Sales = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub sPageHeader()

  Try

  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(5, "Sl No.") & Space(1) & LAlign(45, "Customer Name") & Space(1) & LAlign(12, "Total Amount")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
  OptAAll.Value = True
  CmdStep(2).Enabled = False
  Call StepClear(Me)
  Call StepEnable(Me)
  spList.Visible = False
  MDISales.Toolbar1.Buttons(1).Enabled = True
  MDISales.Toolbar1.Buttons(2).Enabled = True
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdview.Enabled = True
  End If
  spList.Enabled = False
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If DateValue(dptodt) < DateValue(dpfromdt) Then
    sShowMessage "To Date should be greater than From Date"
    dptodt.SetFocus
    GoTo ExitHere
  End If
  
'  Set frmCurrent = Me
'  DoEvents
'  iLineNo = 1
'  iPageNo = 1

  If Sales = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
  
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

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub dpfromdt_Change()
  Try
  Me.MousePointer = vbHourglass

  If CDate(dpfromdt.Value) > CDate(dptodt.Value) Then
    sShowMessage "From date can't be greater than to date"
    dpfromdt.SetFocus
    dpfromdt.Value = Format(dptodt, "dd/mm/yyyy")
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dptodt_Change()
  Try
  Me.MousePointer = vbHourglass

  If CDate(dptodt.Value) < CDate(dpfromdt.Value) Then
    sShowMessage "To date can't be Less than from date"
    dptodt.SetFocus
    dptodt.Value = Format(dpfromdt, "dd/mm/yyyy")
    GoTo ExitHere
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
  cmbType.ListIndex = 0
  cmbReportType.ListIndex = 0
  spList.Visible = False
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

  CmdStep(2).Enabled = True
  Call Spread_Clear(spList)
  spList.Visible = False
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
  
  CmdStep(2).Enabled = False

  If OptAAll Then
 GoTo ExitHere
End If
  Sql = " Select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode"
  Sql = Sql + " and b.Branchid=a.BranchID and b.Compid=a.CompID) [Customer Name] From "
  Sql = Sql + " SAL_INVOICE_HDR a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " order by a.CustomerCode"
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.ClearSelection
'  spList.Enabled = True
'  spList.MaxRows = spList.DataRowCnt
'  If spList.DataRowCnt = 0 Then
'    sShowMessage "No Records Found."
'    OptAAll.SetFocus
'    GoTo ExitHere
'  End If
'  spList.Enabled = True
'  spList.Visible = True
'  spList.SetFocus
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    OptAAll.Value = True
    GoTo ExitHere
  Else
    spList.Enabled = True
 spList.Visible = True
    spList.ClearSelection
    Call ADORS_Spread_DataFill(rs, spList, True)
    
    spList.MaxRows = spList.DataRowCnt
    Call Spread_Set_Focus(spList, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OptAccRandom_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyAscii = 13 Then
    SendKeys Chr(9)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList


spList.Click +=  new EventHandler(spList_Click)
spList.KeyUp +=  new EventHandler(spList_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesLedgerReport.frm", spList, "")

End Sub
