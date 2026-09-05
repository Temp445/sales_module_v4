VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesReceiptRegisterReport 
   Caption         =   "Sales Receipt Register Report"
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
      Index           =   3
      Left            =   6555
      TabIndex        =   10
      Top             =   3465
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
      Left            =   6555
      TabIndex        =   8
      Top             =   2790
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
      Left            =   6555
      TabIndex        =   5
      Top             =   2145
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
      Left            =   6555
      TabIndex        =   2
      Top             =   1485
      Width           =   345
   End
   Begin VB.Frame Frame6 
      Caption         =   "RR"
      Height          =   540
      Left            =   5340
      TabIndex        =   17
      Top             =   3960
      Width           =   2700
      Begin VB.OptionButton optRRAll 
         Caption         =   "All"
         Height          =   210
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optWRR 
         Caption         =   "with RR"
         Height          =   210
         Left            =   630
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
      Begin VB.OptionButton optWORR 
         Caption         =   "Without RR"
         Height          =   210
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Type of Invoice"
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   3975
      TabIndex        =   32
      Top             =   3225
      Width           =   2520
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "SalesReceiptRegisterReport.frx":0000
         Left            =   90
         List            =   "SalesReceiptRegisterReport.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   285
         Width           =   2340
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Product"
      Enabled         =   0   'False
      Height          =   585
      Index           =   2
      Left            =   3975
      TabIndex        =   31
      Top             =   2625
      Width           =   2520
      Begin VB.OptionButton optIRandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1395
         TabIndex        =   7
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton optIAll 
         Caption         =   "All"
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   540
      Left            =   2760
      TabIndex        =   30
      Top             =   3960
      Width           =   2520
      Begin VB.OptionButton optTpending 
         Caption         =   "Pending"
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optTAll 
         Caption         =   "All"
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optTclose 
         Caption         =   "Closed"
         Height          =   240
         Left            =   690
         TabIndex        =   12
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame frastep 
      Height          =   705
      Index           =   0
      Left            =   3975
      TabIndex        =   27
      Top             =   1290
      Width           =   2520
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   109314049
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   0
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   109314049
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1260
         TabIndex        =   29
         Top             =   135
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   28
         Top             =   135
         Width           =   735
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   585
      Index           =   1
      Left            =   3975
      TabIndex        =   26
      Top             =   2010
      Width           =   2520
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1380
         TabIndex        =   4
         Top             =   255
         Width           =   960
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame FraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   3975
      TabIndex        =   25
      Top             =   4500
      Width           =   2520
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "SalesReceiptRegisterReport.frx":0087
         Left            =   90
         List            =   "SalesReceiptRegisterReport.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   210
         Width           =   2325
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   5760
      TabIndex        =   22
      Top             =   5535
      Width           =   1300
   End
Begin AceSpread spItem
      Height          =   1260
      Left            =   6990
      TabIndex        =   23
      Top             =   2685
      Visible         =   0   'False
      Width           =   3540
      _Version        =   458752
      _ExtentX        =   6244
      _ExtentY        =   2222
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
Begin AceSpread spList
      Height          =   1275
      Left            =   6990
      TabIndex        =   24
      Top             =   1380
      Visible         =   0   'False
      Width           =   3540
      _Version        =   458752
      _ExtentX        =   6244
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
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4470
      TabIndex        =   21
      Top             =   5535
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3180
      TabIndex        =   20
      Top             =   5535
      Width           =   1300
   End
End
Attribute VB_Name = "rptSalesReceiptRegisterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim j As Long

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Private Sub Header()
  Dim mstr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    
    frmPrint.sprPrint.MaxCols = 9

    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    
    frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(15) & "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
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
      
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(15) & "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If

    If iPageNo = 1 Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText 1, j, gCompName
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
    frmPrint.sprPrint.SetText 1, j, "Sales Receipt Register between " & dpfromdt & " and " & dptodt
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

    frmPrint.sprPrint.SetText 1, j, "Invoice No"
frmPrint.sprPrint.ColWidth(1) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText 2, j, "Invoice Date"
frmPrint.sprPrint.ColWidth(2) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText 3, j, "RR No"
frmPrint.sprPrint.ColWidth(3) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText 4, j, "Inv. Qty"
frmPrint.sprPrint.ColWidth(4) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText 5, j, "Rejection/"
frmPrint.sprPrint.ColWidth(5) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText 6, j, "Approved"
frmPrint.sprPrint.ColWidth(6) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText 7, j, "Invoice"
frmPrint.sprPrint.ColWidth(7) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

    frmPrint.sprPrint.SetText 8, j, "Rejection/"
frmPrint.sprPrint.ColWidth(8) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

    frmPrint.sprPrint.SetText 9, j, "SIDBI"
frmPrint.sprPrint.ColWidth(9) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
    

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText 5, j, "Short Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText 6, j, "Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText 7, j, "Amount"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

    frmPrint.sprPrint.SetText 8, j, "Debit Amt."
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

    frmPrint.sprPrint.SetText 9, j, "Amount"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    
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

  If sPrintCurrent = False Then
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

Public Sub sPageHeader()
  
  Try
  
  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  
  sSendToPort LAlign(10, "Inv. No") & Space(1) & LAlign(10, "Inv. Date") & Space(1) & LAlign(10, "RR No") & Space(1) & RAlign(8, "Inv. Qty") & Space(1) & RAlign(11, "Rejection/") & Space(1) & RAlign(8, "Approved") & Space(1) & RAlign(10, "Invoice") & Space(1) & RAlign(10, "Rejection/") & Space(1) & RAlign(10, "SIDBI")
  sSendToPort Space(42) & RAlign(11, "Short Qty") & Space(1) & RAlign(8, "Quantity") & Space(1) & RAlign(10, "Amount") & Space(1) & RAlign(10, "Debit Amt.") & Space(1) & RAlign(10, "Amount")
  
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  
End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim tmp As String
  Dim sRow As Long
  Dim i As Integer
  Dim InvQty As Double
  Dim RecQty As Double
  Dim AccQty As Double
  Dim Qty(3) As String
  Dim tot As Double
  Dim Gtot As Double

  Try
  
  If optRandom Then
    If spList.SelectionCount < 1 Then
      sShowMessage "Please select atleast one Customer"
      GoTo ExitHere
    End If
  End If
  
  If optIRandom Then
    If spItem.SelectionCount < 1 Then
      sShowMessage "Please select atleast one item"
      GoTo ExitHere
    End If
  End If
  
  Set rsDetails = New AceADODB.Recordset       'rsDetails Recordset Initialized
  
  Sql = " select distinct a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,"
  Sql = Sql + " a.ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode"
  Sql = Sql + " and c.Branchid=a.Branchid and c.Compid=a.Compid)ItemDescription,"
  Sql = Sql + " (select c1.PartNo from SAL_Customer_ItemMaster c1 where c1.CustomerCode = a.CustomerCode "
  Sql = Sql + " and c1.ItemCode=a.ItemCode and c1.Branchid=a.Branchid and c1.Compid=a.Compid) PartNo,"
  Sql = Sql + " a.InvoiceQty,a.ReceivedQty,a.AcceptedQty,Round(((a.InvoiceQty-a.ReceivedQty)/a.InvoiceQty) * TotalAmount,2) as DebitAmount,"
  Sql = Sql + " a.RRNo,a.RRDate, a.TotalAmount from SalesInvoice a"
  Sql = Sql + " where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
  Sql = Sql + " convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If optRandom Then
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
  If optIRandom Then
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
  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo=''"
  End If
  If cmbType <> "All" Then
    Sql = Sql + " and a.Type = '" + Trim$(cmbType) + "'"
  End If
  If optTclose.Value = True Then Sql = Sql & " and a.Invoiceqty=a.AcceptedQty "
  If optTpending.Value = True Then Sql = Sql & " and a.Invoiceqty<>a.AcceptedQty"
  Sql = Sql + " order by PartyName, a.InvoiceDate, a.InvoiceType, a.InvoiceNo"
  If RSOpen(rsDetails, Sql) = False Then
    GoTo ExitHere
  End If
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sPrintCurrent = False
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
   gStrReportHeading = "Sales Register between " & dpfromdt & " and " & dptodt
   sSendToPort CAlign(iPageWidth, gStrReportHeading)
   sSendToPort ""
   
   sPageHeader
   
   i = 1
InvQty = 0
RecQty = 0
AccQty = 0
   Do While Not rsDetails.EOF
     If tmp <> rsDetails("PartyName") Then
       If tmp <> "" Then
         sSendToPort Space(11) & LAlign(10, "Total") & Space(1) & Space(30) & RAlign(8, CStr(InvQty)) & Space(1) & RAlign(8, CStr(RecQty))
         sSendToPort Space(61) & RAlign(8, CStr(AccQty))
         InvQty = 0
RecQty = 0
AccQty = 0
       End If
       If i > 1 Then
 sSendToPort String(iPageWidth, "-")
sSendToPort ""
End If
       sSendToPort LAlign(15, "Customer Name:") & Space(1) & LAlign(50, rsDetails("PartyName"))
       sSendToPort ""
     End If
     
     sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(10, rsDetails("InvoiceDate")) & Space(1) & LAlign(10, IIf(IsNull(rsDetails("RRNo")), "", rsDetails("RRNo"))) & Space(1) & RAlign(8, rsDetails("InvoiceQty")) & Space(1) & RAlign(11, (rsDetails("InvoiceQty") - rsDetails("ReceivedQty"))) & Space(1) & RAlign(8, rsDetails("AcceptedQty")) & Space(1) & RAlign(10, rsDetails("TotalAmount")) & Space(1) & RAlign(10, rsDetails("DebitAmount")) & Space(1) & RAlign(10, (rsDetails("TotalAmount") - rsDetails("DebitAmount")))

     InvQty = InvQty + rsDetails("InvoiceQty")
     RecQty = RecQty + rsDetails("ReceivedQty")
     AccQty = AccQty + rsDetails("AcceptedQty")
     
     i = i + 1
     Sql = CheckNull(rsDetails("PartNo"))
     tmp = rsDetails("PartyName")
     rsDetails.MoveNext
     sSendToPort ""
   Loop
   
   If rsDetails.EOF And InvQty > 0 Then
     sSendToPort Space(11) & LAlign(10, "Total") & Space(1) & Space(30) & RAlign(8, CStr(InvQty)) & Space(1) & RAlign(8, CStr(RecQty))
     sSendToPort Space(61) & RAlign(8, CStr(AccQty))
     sSendToPort ""
     InvQty = 0
RecQty = 0
AccQty = 0
   End If
   rsDetails.Close
   Call EndofReport(iPageWidth, i - 1)
  Else
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
      
      Header
      
      j = j + 1
GridCount = GridCount + 1

      rsDetails.MoveLast
rsDetails.MoveFirst
      
      InvQty = 0
RecQty = 0
AccQty = 0
      Do While Not rsDetails.EOF
        If tmp <> rsDetails("PartyName") Then
          If tmp <> "" Then
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
            
            frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText 1, j, "Total : "
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
            
            frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText 9, j, Val(tot)
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
            
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            tot = 0
            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "Customer Name : " & rsDetails("PartyName")
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
        
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceType") & rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, CheckNull(rsDetails("InvoiceDate"))
        
        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 3, j, IIf(IsNull(rsDetails("RRNo")), "", rsDetails("RRNo"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, rsDetails("InvoiceQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, frmPrint.sprPrint.Row, (rsDetails("InvoiceQty") - rsDetails("ReceivedQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 6, j, rsDetails("AcceptedQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 7, j, rsDetails("TotalAmount")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, rsDetails("DebitAmount")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 9, j, (rsDetails("TotalAmount") - rsDetails("DebitAmount"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

       
        tot = Val(tot) + (rsDetails("TotalAmount") - rsDetails("DebitAmount"))
        Gtot = Val(Gtot) + (rsDetails("TotalAmount") - rsDetails("DebitAmount"))
        j = j + 1
GridCount = GridCount + 1
RowProcess
        tmp = rsDetails("PartyName")
        rsDetails.MoveNext
      Loop
      
      If rsDetails.EOF And tot > 0 Then
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
            
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "Sub Total : "
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 9, j, Val(tot)
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "Grant Total : "
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 9, j, Val(Gtot)
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
      
      Sql = "select a.RefNo, a.RefDate, a.Amount from fas_billdetails a where a.vrroute='02'"
      If optRandom Then
        Sql = Sql & " and a.lrcode in("
        With spList
          sRow = 0
          For i = 1 To frmPrint.sprPrint.SelectionCount
            sRow = frmPrint.sprPrint.GetMultiSelItem(sRow)
            frmPrint.sprPrint.Row = sRow
frmPrint.sprPrint.Col = 1
            Sql = Sql + "'" + Trim$(frmPrint.sprPrint.Text) + "',"
          Next i
          Sql = Left$(Sql, Len(Sql) - 1)
          sRow = 0
          For i = 1 To spList.SelectionCount
            sRow = spList.GetMultiSelItem(sRow)
            spList.Row = sRow
spList.Col = 1
            Sql = Sql + "'" + Trim$(spList.Text) + "',"
          Next i
          Sql = Left$(Sql, Len(Sql) - 1)
        Sql = Sql + ")"
      Else
        Sql = Sql & "and a.lrcode= 11 "
      End If
      Sql = Sql + " and a.refdate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
      Sql = Sql + " convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) "
      Sql = Sql + " and a.compcode='" & gCompID + gBranchID & "'"
      If RSOpen(rsDetails, Sql) = False Then
 GoTo ExitHere
End If
    
      If Not rsDetails.EOF Then
        .AddCellSpan 1, j, 2, 1
        .Col = 1
.Row = j

        .SetText 1, j, "Debit Amount Details: "
.BackColor = RGB(255, 255, 0)
.TypeVAlign = TypeVAlignCenter
.TypeHAlign = TypeHAlignRight
.Font = "Arial"
.FontBold = True

        For i = 1 To rsDetails.RecordCount
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
              
          .AddCellSpan 1, j, 2, 1
          .Col = 1
.Row = j

          .SetText 1, j, "Sub Total : "
.BackColor = RGB(255, 255, 0)
.TypeVAlign = TypeVAlignCenter
.TypeHAlign = TypeHAlignRight
.Font = "Arial"
.FontBold = True
          
          .Col = 9
.Row = j

          .SetText 9, j, Val(tot)
.BackColor = RGB(255, 255, 0)
.TypeVAlign = TypeVAlignCenter
.TypeHAlign = TypeHAlignRight
.Font = "Arial"
.FontBold = True
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, .MaxCols, j
          
          j = j + 1
GridCount = GridCount + 1
RowProcess
          
          .AddCellSpan 1, j, 2, 1
          .Col = 1
.Row = j

          .SetText 1, j, "Grant Total : "
.BackColor = RGB(255, 255, 0)
.TypeVAlign = TypeVAlignCenter
.TypeHAlign = TypeHAlignRight
.Font = "Arial"
.FontBold = True
          
          .Col = 9
.Row = j

          .SetText 9, j, Val(Gtot)
.BackColor = RGB(255, 255, 0)
.TypeVAlign = TypeVAlignCenter
.TypeHAlign = TypeHAlignRight
.Font = "Arial"
.FontBold = True
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, .MaxCols, j
          
          rsDetails.MoveNext
        Next i
      End If
      
      
      .MaxRows = .DataRowCnt + 1
    End With
    Gtot = 0
tot = 0
    totalpages = iPageNo
    frmPrint.Show
  End If

  
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  Call StepClear(Me)
  Call StepEnable(Me)
  spList.Visible = False
  spItem.Visible = False
  OptAll.Value = True
  optIAll.Value = True
  chkPrintDate.Value = 1
  CmdStep(1).Enabled = False
  CmdStep(2).Enabled = False
'  MDISales.Toolbar1.Buttons(1).Enabled = True
'  MDISales.Toolbar1.Buttons(2).Enabled = True
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound > Index Then
    If optRandom = True Then
      If spList.SelectionCount >= 1 Then
        spList.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
    If optIRandom = True Then
      If spItem.SelectionCount >= 1 Then
        spItem.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  End If
End Sub

Private Sub dpfromdt_Change()
' On Error Resume Next
  optRandom_Click
End Sub

Private Sub dptodt_Change()
' On Error Resume Next
  optRandom_Click
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try
  
  CurrDate Me
  cmbType.ListIndex = 0
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
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  CmdStep(1).Enabled = True
  spList.Visible = False
End Sub

Private Sub optIAll_Click()
' On Error Resume Next
  CmdStep(2).Enabled = True
  spItem.Visible = False
End Sub

Private Sub optIRandom_Click()
  Dim Sql As String
  Dim sRow As Long
  Dim i As Long
  Dim rs As AceADODB.Recordset
   
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(2).Enabled = False

  Sql = " select distinct a.ItemCode [Item Code],(select b.ItemDescription from INV_MATERIAL_MASTER b "
  Sql = Sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Item Description]" 'and b.status = 1
  Sql = Sql + " from SalesInvoice a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt.Value, "yyyy-MM-dd") + "',111)  "
  Sql = Sql + " and convert(datetime,'" + Format(dptodt.Value, "yyyy-MM-dd") + "',111) "
  If optRandom Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
optIAll = True
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
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
  If cmbType <> "All" Then
    Sql = Sql + " and a.Type = '" + Trim$(cmbType) + "'"
  End If

  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo=''"
  End If
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' "
'  adoList1.ConnectionString = db
'  adoList1.RecordSource = Sql
'  adoList1.Refresh
'  spItem.ClearSelection
'  spItem.MaxRows = spItem.DataRowCnt
'  If spItem.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    optIAll.SetFocus
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
    optIAll.Value = True
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

Private Sub optRandom_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(1).Enabled = False

  If OptAll Then
 GoTo ExitHere
End If
  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid ='" & gCompID & "') [Customer Name]"
  Sql = Sql + " from SAL_INVOICE_HDR a where a.InvoiceDate BETWEEN CONVERT(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111) and a.Branchid='" & gBranchID & "' "
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  If optWRR Then
    Sql = Sql + " and a.RRNo <> ''"
  ElseIf optWORR Then
    Sql = Sql + " and a.RRNo is null"
  End If
  If cmbType <> "All" Then
    Sql = Sql + " and a.Type = '" + Trim$(cmbType) + "'"
  End If
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.ClearSelection
'  spList.MaxRows = spList.DataRowCnt
'  If spList.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    OptAll.SetFocus
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
    OptAll.Value = True
    GoTo ExitHere
  Else
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

Private Sub spItem_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spItem.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spItem_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spItem.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub spList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.Click +=  new EventHandler(spItem_Click)
spItem.KeyUp +=  new EventHandler(spItem_KeyUp)

End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList


spList.Click +=  new EventHandler(spList_Click)
spList.KeyUp +=  new EventHandler(spList_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesReceiptRegisterReport.frm", spItem, "")

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesReceiptRegisterReport.frm", spList, "")

End Sub
