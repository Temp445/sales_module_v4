VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptPendingSalesOrderReport 
   Caption         =   "Pending Sales Order Report"
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
      Left            =   6960
      TabIndex        =   8
      Top             =   2205
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
      Left            =   6960
      TabIndex        =   5
      Top             =   1635
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
      Left            =   6960
      TabIndex        =   2
      Top             =   975
      Width           =   345
   End
   Begin VB.Frame frastep 
      Caption         =   "Cust / Item"
      Enabled         =   0   'False
      Height          =   570
      Index           =   2
      Left            =   4335
      TabIndex        =   24
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   2085
      Width           =   2535
      Begin VB.OptionButton Optselective 
         Caption         =   "&Selective"
         Height          =   210
         Left            =   1515
         TabIndex        =   7
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton optall 
         Caption         =   "&All"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Select to choose all Customers (Alt + A)"
         Top             =   240
         Value           =   -1  'True
         Width           =   480
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4335
      TabIndex        =   21
      Top             =   735
      Width           =   2535
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219873281
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
         Format          =   219873281
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   23
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1260
         TabIndex        =   22
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Report"
      Enabled         =   0   'False
      Height          =   585
      Index           =   1
      Left            =   4335
      TabIndex        =   20
      Top             =   1500
      Width           =   2535
      Begin VB.OptionButton optcustwise 
         Caption         =   "&Customer"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   1350
      End
      Begin VB.OptionButton optitemwise 
         Caption         =   "&Item"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame fraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4335
      TabIndex        =   19
      Top             =   2655
      Width           =   2535
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "PendingSalesOrderReport.frx":0000
         Left            =   90
         List            =   "PendingSalesOrderReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2385
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6315
      TabIndex        =   13
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   3825
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5025
      TabIndex        =   12
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   3825
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3735
      TabIndex        =   11
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   3825
      Width           =   1300
   End
   Begin VB.Frame fraAsOn 
      Caption         =   "Rpt Based On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9510
      TabIndex        =   16
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   135
      Visible         =   0   'False
      Width           =   1395
      Begin MSComCtl2.DTPicker dtpAsOn 
         Height          =   315
         Left            =   105
         TabIndex        =   17
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   219873281
         CurrentDate     =   39346
      End
   End
   Begin VB.Frame fraSchedule 
      Height          =   585
      Left            =   9510
      TabIndex        =   14
      Top             =   705
      Visible         =   0   'False
      Width           =   1770
      Begin VB.CheckBox chkSchedule 
         Caption         =   "Sch.date wise"
         Height          =   210
         Left            =   210
         TabIndex        =   15
         Top             =   240
         Width           =   1365
      End
   End
Begin AceSpread spList
      Height          =   1530
      Left            =   7365
      TabIndex        =   18
      Top             =   1605
      Width           =   4830
      _Version        =   458752
      _ExtentX        =   8520
      _ExtentY        =   2699
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
Attribute VB_Name = "rptPendingSalesOrderReport"
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
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    
    If optAll = True Then
      frmPrint.Tag = "ALL"
    Else
     frmPrint.Tag = "SELECTIVE"
    End If
    
    If Pen_Enquiry = eFormList.mSalOrder Then
      mstr = "Sales"
      frmPrint.sprPrint.MaxCols = 9
    Else
      mstr = "Job Work"
      frmPrint.sprPrint.MaxCols = 10
    End If
    
    frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    
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
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
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
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, gCompName
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "List Of Pending " & mstr & "Orders ( " & frmPrint.Tag & " )" & " Between ( " & dpfromdt & " to " & dptodt & " )"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer PO No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    If Pen_Enquiry = eFormList.mSalOrder Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid From"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Bal Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    Else
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentItem Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "RecdUom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid From"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Bal Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    End If
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer PO Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    If Pen_Enquiry = eFormList.mSalOrder Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid To"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Else
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item GroupKey"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentUom"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid To"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    End If
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Private Function sPrintDetails() As Boolean
  Dim ord As String
  Dim sRow As Long
  Dim genrec As New AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim CurDate As Date
    
  Try
    
  Set genrec = New AceADODB.Recordset
  CurDate = ServerDate
  If Pen_Enquiry = eFormList.mSalOrder Then
    Sql = " Select Distinct x.OrderType,x.OrderNo,x.OrderDate,x.CustomerCode,x.CustomerName, x.ReferenceNo, x.ReferenceDate, x.ItemCode,'' as ItemGroupKey,"
    Sql = Sql & " x.ItemDescription, x.UOMDesc, x.UOMDecimal,x.OrderQty,x.DCQty,x.InvoiceQty,x.Rate, x.Validfrom, x.Validto, x.Partno "
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode,(select top 1 c.PartyName "
    Sql = Sql & " From FAS_PARTY_MASTER c where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "'"
    Sql = Sql & " and Compid ='" & gCompID & "') CustomerName, a.ReferenceNo, a.ReferenceDate, "
    Sql = Sql & " a.ItemCode, e.ItemDescription,  e.UOMDesc, e.UOMDecimal, a.OrderQty,"
    Sql = Sql & " isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType "
    Sql = Sql & " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
    Sql = Sql & " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty,  "
    Sql = Sql & " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.ItemCode "
    Sql = Sql & " and b.OrderType = a.OrderType And b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate and b.Branchid=a.Branchid  "
    Sql = Sql & " and b.Compid=a.Compid and InvoiceDate <= '" & Format(dtpAsOn, "yyyy-MM-dd") & "'),0) InvoiceQty,"
    Sql = Sql & " a.rate, a.Validfrom, a.Validto, a.Partno from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
    Sql = Sql & " left outer join Material_Master e "
    Sql = Sql & " on e.ItemCode=a.ItemCode and e.Branchid ='" & gBranchID & "' and e.Compid ='" & gCompID & "'"
    Sql = Sql & " where a.OrderingType='Specific' and a.Status = 'A'and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql & " and a.OrderDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
    Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)) x where (x.OrderQty-(DCQty+InvoiceQty)) >0"
    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and x.CustomerCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql + " and x.ItemCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    End If
  Else
    Sql = " Select Distinct x.OrderType,x.OrderNo,x.OrderDate,x.CustomerCode,x.CustomerName, x.ReferenceNo,"
    Sql = Sql & " x.ReferenceDate,x.ReceivedItemCode, x.SentItemCode, x.ItemGroupKey,x.RecdDesc,x.SentDesc, x.RecdUom,"
    Sql = Sql & " X.SentUom , X.UOMDecimal, X.OrderQty, X.DCQty, X.InvoiceQty, X.Rate, X.Validfrom, X.Validto, X.PartNo"
    Sql = Sql & " from (Select  a.OrderType,a.OrderNo,a.OrderDate,a.CustomerCode,(select top 1 c.PartyName"
    Sql = Sql & " from  FAS_PARTY_MASTER c where c.PartyCode=a.CustomerCode and c.Branchid=a.Branchid and c.Compid=a.Compid ) CustomerName,"
    Sql = Sql & " a.ReferenceNo, a.ReferenceDate, a.ItemGroupKey,a.ReceivedItemCode, a.SentItemCode,"
    Sql = Sql & " e.ItemDescription as SentDesc, g.ItemDescription as RecdDesc,e.UOMDesc as SentUom, g.UomDesc as RecdUom,"
    Sql = Sql & " e.UOMDecimal, a.OrderQty, isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.SentItemCode"
    Sql = Sql & " and d.Ordertype=a.OrderType and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate"
    Sql = Sql & " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty, isnull((select sum(b.InvoiceQty)"
    Sql = Sql & " from SAL_INVOICE_DTL b where b.ItemCode=a.SentItemCode  and b.OrderType = a.OrderType And b.OrderNo=a.OrderNo"
    Sql = Sql & " and b.OrderDate=a.OrderDate and b.Branchid=a.Branchid  and b.Compid=a.Compid and InvoiceDate <=  '" & Format(dtpAsOn, "yyyy-MM-dd") & "'),0) InvoiceQty"
    Sql = Sql & " , a.rate, a.Validfrom, a.Validto, f.Partno"
    Sql = Sql & " from dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
    Sql = Sql & " Left outer join Material_Master e  on e.ItemCode=a.SentItemCode and e.Branchid=a.Branchid  and e.Compid=a.Compid"
    Sql = Sql & " Left outer join Material_Master g  on g.ItemCode=a.ReceivedItemCode and g.Branchid=a.Branchid  and g.Compid=a.Compid"
    Sql = Sql & " Left outer join SAL_Customer_ItemMaster f on f.CustomerCode = a.CustomerCode and f.ItemCode=a.SentItemCode and f.Branchid=a.Branchid and f.Compid=a.Compid"
    Sql = Sql & " where a.OrderingType='Specific' and a.Status = 'A' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql & " and a.OrderDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111) and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)) x"
    Sql = Sql & " where (x.OrderQty-(DCQty+InvoiceQty)) >0 "

    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and x.CustomerCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql + " and x.SentItemCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    End If
  End If
  Sql = Sql + " order by x.OrderType,x.OrderNo,x.OrderDate"
  If RSOpen(genrec, Sql) = False Then GoTo ExitHere
  If genrec.RecordCount = 0 Then
    sPrintDetails = False
    sShowMessage "No Pending Orders"
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
    If optAll Then
      gStrReportHeading = "List of Pending Orders (ALL)"
    Else
      gStrReportHeading = "List of Pending Orders (Selective)"
    End If
    gStrLoginUserName = "between " & dpfromdt & " to " & dptodt
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort CAlign(iPageWidth, gStrLoginUserName)
    sSendToPort ""
    sPageHeader
    i = 1
    Do While Not genrec.EOF
      If ord <> genrec("CustomerCode") Then
        If i > 1 Then
 sSendToPort String(iPageWidth, "-")
sSendToPort ""
End If
        sSendToPort "Customer : " & genrec("CustomerName")
      End If
      If Sql <> (genrec("OrderType") & genrec("OrderNo")) Then
        sSendToPort ""
        sSendToPort LAlign(10, genrec("OrderType") & genrec("OrderNo")) & Space(1) & LAlign(10, genrec("ReferenceNo")) & Space(1) & LAlign(39, genrec("ItemCode") & IIf(Pen_Enquiry = eFormList.mJWOrder, "[" & genrec("ItemGroupKey") & "]", "")) & Space(1) & RAlign(9, genrec("OrderQty")) & Space(1) & RAlign(8, genrec("OrderQty") - (genrec("DCQty") + genrec("InvoiceQty")))
        sSendToPort LAlign(10, genrec("OrderDate")) & Space(1) & LAlign(10, genrec("ReferenceDate")) & Space(1) & LAlign(39, genrec("ItemDescription")) & Space(1) & RAlign(9, genrec("UOMDesc"))
        i = i + 1
      Else
        sSendToPort Space(22) & LAlign(39, genrec("ItemCode") & IIf(Pen_Enquiry = eFormList.mJWOrder, "[" & genrec("ItemGroupKey") & "]", "")) & Space(1) & RAlign(9, genrec("OrderQty")) & Space(1) & RAlign(8, genrec("OrderQty") - (genrec("DCQty") + genrec("InvoiceQty")))
        sSendToPort Space(22) & LAlign(39, genrec("ItemDescription")) & Space(1) & RAlign(9, genrec("UOMDesc"))
      End If
  '    i = i + 1
      ord = genrec("CustomerCode")
      Sql = (genrec("OrderType") & genrec("OrderNo"))
      genrec.MoveNext
    Loop
    genrec.Close
    EndofReport iPageWidth, i - 1
  Else 'spread report
    frmPrint.Show
  
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
    
    Header
    
    j = j + 1
GridCount = GridCount + 1

    genrec.MoveLast
genrec.MoveFirst
    
    If Pen_Enquiry = eFormList.mJWOrder Then
      For k = 1 To genrec.RecordCount
          If Sql <> (genrec("OrderType") & genrec("OrderNo")) Then
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderType") & " - " & genrec("OrderNo")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("CustomerName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("ReferenceNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("ReceivedItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("SentItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("Partno")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("RecdUom")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("Validfrom")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderQty")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderQty") - (genrec("DCQty") + genrec("InvoiceQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
         
          If Sql <> (genrec("OrderType") & genrec("OrderNo")) Then
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderDate")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("ReferenceDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("ItemGroupKey"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("RecdDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("SentDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("SentUom")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("ValidTo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          ord = genrec("CustomerCode")
          Sql = (genrec("OrderType") & genrec("OrderNo"))
          genrec.MoveNext      Next k
    Else
      For k = 1 To genrec.RecordCount
          If Sql <> (genrec("OrderType") & genrec("OrderNo")) Then
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderType") & " - " & genrec("OrderNo")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("CustomerName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("ReferenceNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("ItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("Partno")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("UOMDesc")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("Validfrom")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderQty")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderQty") - (genrec("DCQty") + genrec("InvoiceQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
         
          If Sql <> (genrec("OrderType") & genrec("OrderNo")) Then
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("OrderDate")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("ReferenceDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, genrec("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(genrec("ValidTo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          ord = genrec("CustomerCode")
          Sql = (genrec("OrderType") & genrec("OrderNo"))
          genrec.MoveNext
      Next k
    End If
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    totalpages = iPageNo
  End If
  
  sPrintDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
End Function

Public Sub sPageHeader()
  
  Try

  sSendToPort String(iPageWidth, "-")
  If gReportID <> "PMI" Then
    sSendToPort LAlign(10, "Order No") & Space(1) & LAlign(10, "Ref. No") & Space(1) & LAlign(39, "Item Code") & Space(1) & RAlign(9, "Order Qty") & Space(1) & RAlign(8, "Bal. Qty")
    sSendToPort LAlign(10, "Order Date") & Space(1) & LAlign(10, "Ref. Date") & Space(1) & LAlign(39, "Item Description") & Space(1) & RAlign(9, "UOM")
  Else
    sSendToPort LAlign(10, "Order No") & Space(1) & LAlign(10, "Ref. No") & Space(1) & LAlign(38, "Item Code") & Space(1) & RAlign(10, "Sch Qty") & Space(1) & RAlign(8, "Bal. Qty")
    sSendToPort LAlign(10, "Order Date") & Space(1) & LAlign(10, "Ref. Date") & Space(1) & LAlign(38, "Item Description") & Space(1) & RAlign(10, "Sch. Date")
    sSendToPort Space(22) & LAlign(38, "Revision No")
  End If
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintDetails1() As Boolean
  Dim ord As String
  Dim sRow As Long
  Dim genrec As New AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim CurDate As Date
  Dim balqty As Double
  Dim Det As String
  Dim DisQty As Double
    
  Try
     
  CurDate = ServerDate
  
  Set genrec = New AceADODB.Recordset
  
  If Pen_Enquiry = eFormList.mSalOrder Then
    Sql = " select Distinct x.OrderType,x.OrderNo,x.OrderDate,x.CustomerCode,x.CustomerName,x.ReferenceNo,"
    Sql = Sql & " x.ReferenceDate, IsNull(x.RevisionNo, '')RevisionNo, x.ItemCode,'' as ItemGroupKey,"
    Sql = Sql & " x.ItemDescription, x.UOMDesc, x.UOMDecimal,x.OrderQty, x.ScheduleDate, x.ScheduleQty,"
    Sql = Sql & " x.DCQty,x.InvoiceQty,x.Rate,"
    Sql = Sql & " Case When(X.CumSchQty - (X.DCQty + X.InvoiceQty)) < ScheduleQty"
    Sql = Sql & " Then (x.CumSchQty-(x.DCQty+x.InvoiceQty)) Else ScheduleQty End as PendingQty"
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, e.RevisionNo,"
    Sql = Sql & " (select top 1 c.PartyName From FAS_PARTY_MASTER c where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "' and Compid ='" & gCompID & "') CustomerName,"
    Sql = Sql & " a.ReferenceNo, a.ReferenceDate, a.ItemCode, e.ItemDescription, e.UOMDesc, e.UOMDecimal, a.OrderQty,"
    Sql = Sql & " f.ScheduleDate, f.ScheduleQty,"
    Sql = Sql & " isNull((Select Sum(g.ScheduleQty) from SAL_Order_Schedule g Where f.OrderType = g.OrderType and f.OrderNo=g.OrderNo and f.OrderDate=g.OrderDate"
    Sql = Sql & " and f.ItemCode=g.ItemCode  and f.Branchid=g.Branchid and f.Compid=g.Compid and g.ScheduleDate<=f.ScheduleDate), 0) as CumSchQty,"
    Sql = Sql & " isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType  and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate"
    Sql = Sql & " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty,"
    Sql = Sql & " isnull((select sum(b.InvoiceQty) from SALESINVOICE b where b.ItemCode=a.ItemCode  and b.OrderType = a.OrderType And b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate"
    Sql = Sql & " and a.CustomerCode = b.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid and InvoiceDate <= '" & Format(dtpAsOn, "yyyy-MM-dd") & "'),0) InvoiceQty, a.rate"
    Sql = Sql & " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " Left outer join Material_Master e"
    Sql = Sql & " On e.ItemCode=a.ItemCode and e.Branchid ='" & gBranchID & "' and e.Compid ='" & gCompID & "'"
    Sql = Sql & " Left outer join SAL_Order_Schedule f"
    Sql = Sql & " on f.OrderType = a.OrderType and f.OrderNo=a.OrderNo and f.OrderDate=a.OrderDate"
    Sql = Sql & " and f.ItemCode=a.ItemCode  and f.Branchid=a.Branchid and f.Compid=a.Compid"
    Sql = Sql & " where a.OrderingType='Specific' and a.Status = 'A' and a.ItemStatus <> 'C'"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    Sql = Sql & " and a.OrderDate between '" & Format(dpfromdt, "yyyy-MM-dd") & "' and '" & Format(dptodt, "yyyy-MM-dd") & "'"
    Sql = Sql & " ) x Where (x.CumSchQty-(x.DCQty+x.InvoiceQty)) > 0 "
    
'    Sql = " select Distinct x.OrderType,x.OrderNo,x.OrderDate,x.CustomerCode,x.CustomerName, x.ReferenceNo, x.ReferenceDate, IsNull(x.RevisionNo, '')RevisionNo, x.ItemCode,'' as ItemGroupKey,"
'    Sql = Sql & " x.ItemDescription, x.UOMDesc, x.UOMDecimal,x.OrderQty, x.ScheduleDate, x.ScheduleQty,x.DCQty,x.InvoiceQty,x.Rate "
'    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, e.RevisionNo, "
'    Sql = Sql & " (select top 1 c.PartyName From FAS_PARTY_MASTER c where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "'"
'    Sql = Sql & " and Compid ='" & gCompID & "') CustomerName, a.ReferenceNo, a.ReferenceDate, "
'    Sql = Sql & " a.ItemCode, e.ItemDescription, e.UOMDesc, e.UOMDecimal, a.OrderQty, f.ScheduleDate, f.ScheduleQty,"
'    Sql = Sql & " isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType "
'    Sql = Sql & " and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty,  "
'    Sql = Sql & " isnull((select sum(b.InvoiceQty) from SALESINVOICE b where b.ItemCode=a.ItemCode "
'    Sql = Sql & " and b.OrderType = a.OrderType And b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate and a.CustomerCode = b.CustomerCode and b.Branchid=a.Branchid"
'    Sql = Sql & " and b.Compid=a.Compid and InvoiceDate <= '" & Format(dtpAsOn, "yyyy-MM-dd") & "'),0) InvoiceQty,"
'    Sql = Sql & " a.rate from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
'    Sql = Sql & " Left outer join Material_Master e "
'    Sql = Sql & " On e.ItemCode=a.ItemCode and e.Branchid ='" & gBranchID & "' and e.Compid ='" & gCompID & "'"
'    Sql = Sql & " Left outer join SAL_Order_Schedule f on f.OrderType = a.OrderType "
'    Sql = Sql & " and f.OrderNo=a.OrderNo and f.OrderDate=a.OrderDate and f.ItemCode=a.ItemCode "
'    Sql = Sql & " and f.Branchid=a.Branchid and f.Compid=a.Compid"
'    Sql = Sql & " where a.OrderingType='Specific' and a.Status = 'A' and a.ItemStatus <> 'C' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
'    Sql = Sql & " and a.OrderDate between '" & Format(dpfromdt, "yyyy-MM-dd") & "' and '" & Format(dptodt, "yyyy-MM-dd") & "'"
'    Sql = Sql & " ) x where (x.OrderQty-(DCQty+InvoiceQty)) >0"

    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and x.CustomerCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql + " and x.ItemCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    End If
  Else
    Sql = " select Distinct x.OrderType,x.OrderNo,x.OrderDate,x.CustomerCode,x.CustomerName, x.ReferenceNo, x.ReferenceDate, IsNull(x.RevisionNo, '')RevisionNo, x.SentItemCode ItemCode,"
    Sql = Sql & " x.ItemGroupKey,x.ItemDescription, x.UOMDesc, x.UOMDecimal,x.OrderQty,x.DCQty,x.InvoiceQty,x.Rate From "
    Sql = Sql & " (Select a.OrderType,a.OrderNo,a.OrderDate,a.CustomerCode,(select top 1 c.PartyName from "
    Sql = Sql & " From (Select FAS_PARTY_MASTER c where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "'"
    Sql = Sql & " and Compid ='" & gCompID & "') CustomerName, a.ReferenceNo, a.ReferenceDate, "
    Sql = Sql & " a.ItemGroupKey,a.SentItemCode, e.ItemDescription,e.RevisionNo, e.UOMDesc, e.UOMDecimal, a.OrderQty,"
    Sql = Sql & " isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.SentItemCode and d.Ordertype=a.OrderType and d.OrderNo=a.OrderNo "
    Sql = Sql & " and d.OrderDate=a.OrderDate and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty,"
    Sql = Sql & " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.SentItemCode "
    Sql = Sql & " and b.OrderType = a.OrderType And b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate  and b.Branchid=a.Branchid"
    Sql = Sql & " and b.Compid=a.Compid and InvoiceDate <= '" & Format(dtpAsOn, "yyyy-MM-dd") & "'),0) InvoiceQty, a.rate "
    Sql = Sql & " From dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
    Sql = Sql & " left outer join Material_Master e "
    Sql = Sql & " on e.ItemCode=a.SentItemCode and e.Branchid ='" & gBranchID & "' and e.Compid ='" & gCompID & "'"
    Sql = Sql & " where a.OrderingType='Specific' and a.Status = 'A' and a.ItemStatus <> 'C' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql & " and a.OrderDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
    Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)) x where (x.OrderQty-(DCQty+InvoiceQty)) >0"
    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and x.CustomerCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql + " and x.SentItemCode in("
        sRow = 0
        For i = 1 To spList.SelectionCount
          sRow = spList.GetMultiSelItem(sRow)
          spList.Row = sRow
spList.Col = 1
          Sql = Sql + "'" + Trim$(spList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)
      Sql = Sql + ")"
    End If
  End If
  If chkSchedule.Value = 1 Then
    Sql = Sql & " Order by ScheduleDate"
  Else
    Sql = Sql & " Order by x.OrderType,x.OrderNo,x.OrderDate"
  End If
  If RSOpen(genrec, Sql) = False Then GoTo ExitHere
  If genrec.RecordCount = 0 Then
    sPrintDetails1 = False
    sShowMessage "No Pending Orders"
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
  
  If optAll Then
    gStrReportHeading = "List of Pending Orders (ALL)"
  Else
    gStrReportHeading = "List of Pending Orders (Selective)"
  End If
  
  gStrLoginUserName = "between " & dpfromdt & " to " & dptodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort CAlign(iPageWidth, gStrLoginUserName)
  sSendToPort ""
  sPageHeader
  
  i = 1
  Do While Not genrec.EOF
    If ord <> genrec("CustomerCode") Then
      If i > 1 Then
 sSendToPort String(iPageWidth, "-")
sSendToPort ""
End If
      sSendToPort "Customer : " & genrec("CustomerName")
    End If
    If Det <> (genrec("OrderType") & genrec("OrderNo")) & genrec("ItemCode") Then
      balqty = (genrec("DCQty") + genrec("InvoiceQty"))
    End If
    DisQty = 0
    If genrec("ScheduleQty") > balqty Then
      DisQty = genrec("ScheduleQty") - balqty
      balqty = 0
    Else
      DisQty = genrec("ScheduleQty")
      balqty = balqty - genrec("ScheduleQty")
    End If
    
    If Det <> (genrec("OrderType") & genrec("OrderNo")) & genrec("ItemCode") Then
      sSendToPort ""
      sSendToPort LAlign(10, genrec("OrderType") & genrec("OrderNo")) & Space(1) & LAlign(10, genrec("ReferenceNo")) & Space(1) & LAlign(38, genrec("ItemCode") & IIf(Pen_Enquiry = eFormList.mJWOrder, "[" & genrec("ItemGroupKey") & "]", "")) & Space(1) & RAlign(10, genrec("ScheduleQty")) & Space(1) & RAlign(8, CStr(DisQty))
      sSendToPort LAlign(10, genrec("OrderDate")) & Space(1) & LAlign(10, genrec("ReferenceDate")) & Space(1) & LAlign(38, genrec("ItemDescription")) & Space(1) & RAlign(10, genrec("ScheduleDate"))
      If genrec("RevisionNo") <> "" Then
        sSendToPort Space(22) & LAlign(38, genrec("RevisionNo"))
      End If
      i = i + 1
    Else
      sSendToPort ""
      sSendToPort Space(60) & Space(1) & RAlign(10, genrec("ScheduleQty")) & Space(1) & RAlign(8, CStr(DisQty))
      sSendToPort Space(60) & Space(1) & RAlign(10, genrec("ScheduleDate"))
    End If
'    i = i + 1
    ord = genrec("CustomerCode")
    Sql = (genrec("OrderType") & genrec("OrderNo"))
    Det = (genrec("OrderType") & genrec("OrderNo") & genrec("ItemCode"))
    genrec.MoveNext
  Loop
  genrec.Close
  EndofReport iPageWidth, i - 1
   
  sPrintDetails1 = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call StepClear(Me)
  Call StepEnable(Me)
  spList.Visible = False
  optAll.Value = True
  CmdStep(1).Enabled = False
  CmdStep(2).Enabled = False
  cmbReportType.ListIndex = 0
  chkPrintDate.Value = 1
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
 
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me

  If dptodt < dpfromdt Then
    sShowMessage "To Date cannot be less than From Date "
    dptodt.SetFocus
    GoTo ExitHere
  End If
  
'  iLineNo = 1
'  iPageNo = 1
  
  If gReportID <> "PMI" Then
    If sPrintDetails = False Then
 GoTo ExitHere
End If
    If cmbReportType.ListIndex = 1 Then
      Call FinaliseAndShowPrintPreview
    End If
  Else
    If sPrintDetails1 = False Then
 GoTo ExitHere
End If
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

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If FraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    CmdView.Enabled = True
  End If
  spList.Enabled = False
End Sub

Private Sub dptodt_Change()
' On Error Resume Next
  
  dtpAsOn.MinDate = dptodt.Value
  dtpAsOn.MaxDate = ServerDate
End Sub

Private Sub dptodt_Click()
' On Error Resume Next
  
  dtpAsOn.MinDate = dptodt.Value
  dtpAsOn.MaxDate = ServerDate
End Sub

Private Sub dtpAsOn_Change()
  
  Try
  Me.MousePointer = vbHourglass

'  dtpAsOn.MinDate = dptodt.Value
'  dtpAsOn.MaxDate = ServerDate
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpAsOn_Click()

  Try
  Me.MousePointer = vbHourglass

'  dtpAsOn.MinDate = dptodt.Value
'  dtpAsOn.MaxDate = ServerDate
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
  
  dpfromdt = Format(ServerDate, "dd/mm/yyyy")
  dptodt = Format(dpfromdt, "dd/mm/yyyy")
  
  
  dpfromdt.MaxDate = Format(dpfromdt, "dd/mm/yyyy")
  dptodt.MaxDate = Format(dpfromdt, "dd/mm/yyyy")
  
  If gReportID = "PMI" Then
    fraAsOn.Visible = True
    dtpAsOn.Value = Format(dptodt, "dd/mm/yyyy")
  End If

  If Pen_Enquiry = eFormList.mSalOrder Then
     rptPendingSalesOrderReport.Caption = "Sales Order Pending Report"
     If gReportID = "PMI" Then
       fraSchedule.Visible = True
     End If
  Else
    rptPendingSalesOrderReport.Caption = "JobWork Order Pending Report"
  End If
  
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
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  CmdStep(2).Enabled = True
  spList.Visible = False
End Sub

Private Sub optcustwise_Click()
' On Error Resume Next
  OptSelective_Click
End Sub

Private Sub optitemwise_Click()
' On Error Resume Next
  OptSelective_Click
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim CurDate As Date
  Dim rs As AceADODB.Recordset

  Try
  If optAll Then
 GoTo ExitHere
End If
  
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  CmdStep(2).Enabled = False
  
  If Pen_Enquiry = eFormList.mSalOrder Then
    If optcustwise Then
      Sql = " select distinct CustomerCode [Customer Code],CustomerName [Customer Name] from"
      Sql = Sql + " ( Select distinct CustomerCode,(select top 1 c.PartyName from FAS_PARTY_MASTER c "
      Sql = Sql + " where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "' and Compid ='" & gCompID & "') CustomerName,"
      Sql = Sql + " a.OrderQty,isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType"
      Sql = Sql + " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
      Sql = Sql + " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty, "
      Sql = Sql + " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.ItemCode and b.Ordertype=a.OrderType"
      Sql = Sql + " and b.OrderNo=a.OrderNo and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate "
      Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid),0) InvoiceQty"
      Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + " where a.OrderingType ='Specific' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise Then
      Sql = " select distinct ItemCode [Item Code],ItemDescription [Item Description] from"
      Sql = Sql + " ( Select distinct ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c "
      Sql = Sql + " where c.ItemCode=a.ItemCode and c.Branchid ='" & gBranchID & "' and c.Compid ='" & gCompID & "') ItemDescription,"
      Sql = Sql + " a.OrderQty,isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType"
      Sql = Sql + " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
      Sql = Sql + " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty, "
      Sql = Sql + " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.ItemCode and b.Ordertype=a.OrderType"
      Sql = Sql + " and b.OrderNo=a.OrderNo and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate "
      Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid),0) InvoiceQty"
      Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + " where a.OrderingType ='Specific' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    End If
  Else
    If optcustwise Then
      Sql = " select distinct CustomerCode [Customer Code],CustomerName [Customer Name] from"
      Sql = Sql + " ( Select distinct CustomerCode,(select top 1 c.PartyName from FAS_PARTY_MASTER c "
      Sql = Sql + " where c.PartyCode=a.CustomerCode and a.Branchid ='" & gBranchID & "' and Compid ='" & gCompID & "') CustomerName,"
      Sql = Sql + " a.OrderQty,isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.SentItemCode and d.Ordertype=a.OrderType"
      Sql = Sql + " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
      Sql = Sql + " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty, "
      Sql = Sql + " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.SentItemCode and b.Ordertype=a.OrderType"
      Sql = Sql + " and b.OrderNo=a.OrderNo and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate "
      Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid),0) InvoiceQty"
      Sql = Sql + " from dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + " where a.OrderingType ='Specific' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise Then
      Sql = " select distinct SentItemCode [Item Code],ItemDescription [Item Description] from"
      Sql = Sql + " ( Select distinct SentItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c "
      Sql = Sql + " where c.ItemCode=a.SentItemCode and c.Branchid ='" & gBranchID & "' and c.Compid ='" & gCompID & "') ItemDescription,"
      Sql = Sql + " a.OrderQty,isnull((select sum(d.DCQty) from SAL_DC_DTL d where d.ItemCode=a.SentItemCode and d.Ordertype=a.OrderType"
      Sql = Sql + " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
      Sql = Sql + " and d.Branchid=a.Branchid and d.Compid=a.Compid),0) DCQty, "
      Sql = Sql + " isnull((select sum(b.InvoiceQty) from SAL_INVOICE_DTL b where b.ItemCode=a.SentItemCode and b.Ordertype=a.OrderType"
      Sql = Sql + " and b.OrderNo=a.OrderNo and b.OrderNo=a.OrderNo and b.OrderDate=a.OrderDate "
      Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid),0) InvoiceQty"
      Sql = Sql + " from dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + " where a.OrderingType ='Specific' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    End If
  End If
  Sql = Sql & " and a.OrderDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)) x"
  Sql = Sql + " where (x.OrderQty-(DCQty+InvoiceQty)) >0"

'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.ClearSelection
'  spList.MaxRows = spList.DataRowCnt
'  If spList.DataRowCnt = 0 Then
'    sShowMessage "No Records Exist."
'    OptAll.SetFocus
'    GoTo ExitHere
'  End If
'  If optcustwise Then
'    spList.Col = 1: spList.ColHidden = True
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
    optAll.Value = True
    GoTo ExitHere
  Else
    If optcustwise Then
    spList.Col = 1
spList.ColHidden = True
    End If
    spList.Enabled = True
spList.Visible = True
    spList.ClearSelection
    Call ADORS_Spread_DataFill(rs, spList, True)
    
    spList.MaxRows = spList.DataRowCnt
    Call Spread_Set_Focus(spList, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
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

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\PendingSalesOrderReport.frm", spList, "")

End Sub
