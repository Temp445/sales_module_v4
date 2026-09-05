VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptPendingSalesQuotationReport 
   Caption         =   "Pending Sales Quotation Report"
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
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   570
      Index           =   2
      Left            =   4350
      TabIndex        =   11
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   2235
      Width           =   2625
      Begin VB.OptionButton Optselective 
         Caption         =   "&Selective"
         Height          =   255
         Left            =   1530
         TabIndex        =   10
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton optall 
         Caption         =   "&All"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Select to choose all Customers (Alt + A)"
         Top             =   210
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   795
      Index           =   1
      Left            =   4350
      TabIndex        =   21
      Top             =   1440
      Width           =   2625
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1335
         TabIndex        =   7
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108134401
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   105
         TabIndex        =   5
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108134401
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1320
         TabIndex        =   6
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Frame frastep 
      Height          =   615
      Index           =   0
      Left            =   4350
      TabIndex        =   0
      Top             =   825
      Width           =   2625
      Begin VB.OptionButton optitemwise 
         Caption         =   "&Itemwise"
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
         Left            =   1545
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optcustwise 
         Caption         =   "&Customerwise"
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
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1350
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
      Left            =   7125
      TabIndex        =   3
      Top             =   1005
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
      Left            =   7125
      TabIndex        =   8
      Top             =   1695
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
      Left            =   7125
      TabIndex        =   12
      Top             =   2385
      Width           =   345
   End
   Begin VB.Frame frastep 
      Caption         =   "Report Type"
      Height          =   930
      Index           =   3
      Left            =   4350
      TabIndex        =   15
      Top             =   2835
      Width           =   2625
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print.Date"
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   615
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "PendingSalesQuotationReport.frx":0000
         Left            =   90
         List            =   "PendingSalesQuotationReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   210
         Width           =   2430
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
      Left            =   7125
      TabIndex        =   16
      Top             =   3090
      Width           =   345
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6450
      TabIndex        =   19
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   3990
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5155
      TabIndex        =   18
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   3990
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3870
      TabIndex        =   17
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   3990
      Width           =   1300
   End
Begin AceSpread spList
      Height          =   1545
      Left            =   7590
      TabIndex        =   20
      Top             =   2355
      Width           =   4590
      _Version        =   458752
      _ExtentX        =   8096
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
Attribute VB_Name = "rptPendingSalesQuotationReport"
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
    
    If Pen_Enquiry = eFormList.mSalQuotation Then
      mstr = "Sales"
    Else
      mstr = "Job Work"
    End If
    If Pen_Enquiry = eFormList.mJWQuotation Then
      frmPrint.sprPrint.MaxCols = 10
    Else
      frmPrint.sprPrint.MaxCols = 9
    End If
    
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
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
    frmPrint.sprPrint.SetText 1, j, "List Of Pending " & mstr & " Quotation ( " & frmPrint.Tag & " )" & " Between ( " & dpfromdt & " to " & dptodt & " )"
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Quot. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 27
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    If Pen_Enquiry = eFormList.mJWQuotation Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      k = 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Else
      k = 0
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 28
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    End If
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid From"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Quot. Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText 1, j, "Quot. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText 2, j, "Item Group Key"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    If Pen_Enquiry = eFormList.mJWQuotation Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      k = 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Else
      k = 0
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    End If

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Valid To"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9 + k

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Public Sub sPageHeader()
  
  Try

  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(14, "Quotation No") & Space(1) & LAlign(50, "Item Code") & Space(1) & RAlign(8, "Quantity")
  sSendToPort LAlign(14, "Quotation Date") & Space(1) & LAlign(50, "Item Description") & Space(1) & RAlign(8, "Rate")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  
End Try
End Sub

Private Function sPrintDetails() As Boolean
  Dim sRow As Long
  Dim quo As String
  Dim quoT As String
  Dim genrec As AceADODB.Recordset
  Dim Sql As String
  Dim i As Integer
  
  Try
  
  Set genrec = New AceADODB.Recordset
  If Pen_Enquiry = eFormList.mSalQuotation Then
    
    Sql = " Select A.QuotationType, a.QuotationNo, a.QuotationDate, a.CustomerCode, a.CustomerName,"
    Sql = Sql & " a.ValidFrom, a.ValidTo, a.TotalAmount, a.ItemCode, a.ItemDescription, a.UOM, a.Qty, a.Rate,"
    Sql = Sql & " a.AssessableRate, b.PartNo from SalesQuotation a"
    Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster b on b.ItemCode = a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " where a.OrderType='' and a.Branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
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
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and a.ItemCode in("
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
    Sql = " Select distinct a.QuotationType,a.QuotationNo,a.SentItemCode ItemCode,a.ItemGroupKey,"
    Sql = Sql & " SentDesc ItemDescription, a.ReceivedItemcode,a.ReceivedDesc, a.QuotationDate, a.ValidFrom,"
    Sql = Sql & " a.ValidTo, a.rate, a.AssRate AssessableRate, a.Sendqty Qty, a.TotalAmount, a.CustomerCode,"
    Sql = Sql & " a.CustomerName, b.PartNo, (select c.UomDesc"
    Sql = Sql & " from MATERIAL_MASTER c where c.ItemCode=a.SentItemCode and c.Branchid=a.Branchid"
    Sql = Sql & " and c.Compid=a.Compid) Uom from JobWorkQuotation a"
    Sql = Sql & " Left Outer Join Partmast b on b.StageCode = a.SentItemcode and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " Where a.OrderType='' and Revisionno = (Select max(Revisionno) from JobWorkQuotation d where d.SentItemCode=a.SentItemCode and d.Branchid=a.Branchid and d.Compid=a.Compid)"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
    If optcustwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
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
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
    ElseIf optitemwise And OptSelective Then
      If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
      Sql = Sql & " and a.SentItemCode in("
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
  
  Sql = Sql & " and a.QuotationDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " order by a.QuotationDate"
  If RSOpen(genrec, Sql) = False Then GoTo ExitHere
  If genrec.RecordCount = 0 Then
    sShowMessage "No Pending Quotation"
    sPrintDetails = False
    GoTo ExitHere
  End If
  If cmbReportType.ListIndex = 1 Then
    pheight = 75
    pPaper = vbPRPSA4
    Orient = 1
    iLineNo = 1
    iPageWidth = 80
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    
    sSendToPort ""
    If optAll Then
      gStrReportHeading = "List of Pending Quotations (ALL)"
    Else
      gStrReportHeading = "List of Pending Quotations (Selective)"
    End If
    gStrLoginUserName = "(" & dpfromdt & " to " & dptodt & ")"
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort CAlign(iPageWidth, gStrLoginUserName)
    sSendToPort ""
    sPageHeader '''Printing the Header Details
    i = 1
    Do While Not genrec.EOF
      If optitemwise Then
        If quo <> genrec("ItemCode") Then
          sSendToPort ""
          sSendToPort "Item : " & genrec("ItemDescription") & "(" & Trim(genrec("ItemCode")) & ")"
          sSendToPort ""
          sSendToPort LAlign(14, genrec("QuotationType") & genrec("QuotationNo")) & Space(1) & LAlign(50, genrec("CustomerName")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort LAlign(14, genrec("QuotationDate")) & Space(52) & RAlign(8, Format(genrec("Rate"), "0.00"))
        Else
          sSendToPort Space(15) & LAlign(50, genrec("CustomerName")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort Space(66) & RAlign(8, Format(genrec("Rate"), "0.00"))
        End If
        quo = genrec("ItemCode")
      Else
        If quo <> genrec("CustomerCode") Then
          sSendToPort ""
          sSendToPort "Customer : " & genrec("CustomerName")
          sSendToPort ""
          sSendToPort LAlign(14, genrec("QuotationType") & genrec("QuotationNo")) & Space(1) & LAlign(50, genrec("ItemCode")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort LAlign(14, genrec("QuotationDate")) & Space(1) & LAlign(50, genrec("ItemDescription")) & Space(1) & RAlign(8, Format(genrec("Rate"), "0.00"))
        Else
          sSendToPort Space(15) & LAlign(50, genrec("ItemCode")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort Space(15) & LAlign(50, genrec("ItemDescription")) & Space(1) & RAlign(8, Format(genrec("Rate"), "0.00"))
        End If
        quo = genrec("CustomerCode")
      End If
      sSendToPort ""
      i = i + 1
      genrec.MoveNext
    Loop
    genrec.Close
    EndofReport iPageWidth, i - 1
  Else
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
      
      For i = 1 To genrec.RecordCount
        If optitemwise Then
          If Trim$(Sql) <> genrec("ItemCode") Then
            frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.SetText 1, j, "Item Code : " & genrec("ItemCode") & " - " & genrec("ItemDescription")
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
        Else
          If Trim$(quo) <> genrec("CustomerCode") Then
            frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.SetText 1, j, "Customer Name : " & genrec("CustomerName")
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
        End If
        If Trim$(quoT) <> (genrec("QuotationType") & genrec("Quotationno")) Then
          frmPrint.sprPrint.SetText 1, j, genrec("QuotationType") & genrec("Quotationno")
        End If
        frmPrint.sprPrint.SetText 2, j, genrec("CustomerName")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If Pen_Enquiry = eFormList.mJWQuotation Then
          If optcustwise = True Then
            frmPrint.sprPrint.SetText 3, j, genrec("ReceivedItemCode") & "  [" & genrec("ItemGroupKey") & "]"
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          Else
            frmPrint.sprPrint.SetText 3, j, genrec("ReceivedItemCode")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
          k = 1
        Else
          k = 0
        End If
        frmPrint.sprPrint.SetText 3 + k, j, genrec("ItemCode")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 4 + k, j, genrec("PartNo")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 5 + k, j, genrec("UOM")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 6 + k, j, genrec("ValidFrom")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 7 + k, j, Format(genrec("Rate"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.SetText 8 + k, j, Format(genrec("Qty"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.SetText 9 + k, j, Format((genrec("Rate") * genrec("Qty")), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        If Trim$(quoT) <> (genrec("QuotationType") & genrec("Quotationno")) Then
          frmPrint.sprPrint.SetText 1, j, genrec("QuotationDate")
        End If
        If Pen_Enquiry = eFormList.mJWQuotation Then
          frmPrint.sprPrint.SetText 2, j, genrec("ItemGroupKey")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.SetText 3, j, genrec("ReceivedDesc")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.SetText 3 + k, j, genrec("ItemDescription")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 6 + k, j, genrec("ValidTo")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.SetText 7 + k, j, Format(genrec("AssessableRate"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + k
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        quo = genrec("CustomerCode")
        Sql = genrec("ItemCode")
        quoT = genrec("QuotationType") & genrec("Quotationno")
        genrec.MoveNext
      Next i
      If optitemwise Then
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColHidden = True
      Else
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColHidden = True
      End If
      totalpages = iPageNo
  End If
  sPrintDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
 End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  CmdStep(2).Enabled = False
  optAll.Value = True
  Call StepClear(Me)
  Call StepEnable(Me)
  cmbReportType.ListIndex = 0
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
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

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me

  If dpfromdt > dptodt Then
    sShowMessage "From Date cannot be greater than To Date "
    dpfromdt.SetFocus
    GoTo ExitHere
  End If

'  iLineNo = 1
'  iPageNo = 1

  If sPrintDetails = False Then GoTo ExitHere
  
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

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  dpfromdt = Format(CurDate, "dd/MM/yyyy")
  dptodt = Format(CurDate, "dd/MM/yyyy")
  
  If Pen_Enquiry = eFormList.mSalQuotation Then
    rptPendingSalesQuotationReport.Caption = "Sales Quotation Pending Report"
  Else
    rptPendingSalesQuotationReport.Caption = "JobWork Quotation Pending Report"
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
  spList.Visible = False
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
  If optAll Then
 GoTo ExitHere
End If
  
  Me.MousePointer = vbHourglass
  
  CmdStep(2).Enabled = False
  
  Spread_Clear spList
  If Pen_Enquiry = eFormList.mSalQuotation Then
    If optcustwise = True Then
      Sql = "Select Distinct CustomerCode [CustomerCode], CustomerName [Customer Name] from SalesQuotation a"
      Sql = Sql & " where a.OrderType='' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise = True Then
      Sql = "select Distinct ItemCode [Item Code],ItemDescription [Item Description]  from SalesQuotation a where"
      Sql = Sql & " a.OrderType='' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    End If
  Else
    If optcustwise = True Then
      Sql = "Select Distinct CustomerCode [CustomerCode], CustomerName [Customer Name] from JobWorkQuotation a"
      Sql = Sql & " where a.OrderType='' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise = True Then
      Sql = "select Distinct SentItemCode [Item Code],(Select Distinct SentDescription from SAl_JobWorkEnquiry_Dtl where SentItemCode=a.SentItemCode and Branchid=a.Branchid and Compid=a.Compid) [Item Description]  from JobWorkQuotation a where"
      Sql = Sql & " a.OrderType='' and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    End If
  End If
  
  Sql = Sql & " and a.QuotationDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)"
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

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\PendingSalesQuotationReport.frm", spList, "")

End Sub
