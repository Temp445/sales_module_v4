VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptPendingSalesEnquiryReport 
   Caption         =   "Pending Sales Enquiry Report"
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
   Begin VB.Frame frmmain 
      Height          =   3240
      Left            =   4095
      TabIndex        =   19
      Top             =   810
      Width           =   3435
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
         Left            =   2850
         TabIndex        =   11
         Top             =   2460
         Width           =   345
      End
      Begin VB.Frame frastep 
         Caption         =   "Report Type"
         Height          =   930
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   2175
         Width           =   2625
         Begin VB.CheckBox chkPrintdate 
            Caption         =   "Print.Date"
            Height          =   255
            Left            =   90
            TabIndex        =   10
            Top             =   615
            Width           =   1080
         End
         Begin VB.ComboBox cmbReportType 
            Height          =   330
            ItemData        =   "PendingSalesEnquiryReport.frx":0000
            Left            =   90
            List            =   "PendingSalesEnquiryReport.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   210
            Width           =   2430
         End
      End
      Begin VB.Frame frastep 
         Caption         =   "Cust / Item"
         Enabled         =   0   'False
         Height          =   570
         Index           =   2
         Left            =   150
         TabIndex        =   17
         ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
         Top             =   1560
         Width           =   2625
         Begin VB.OptionButton Optselective 
            Caption         =   "&Selective"
            Height          =   255
            Left            =   1530
            TabIndex        =   7
            Top             =   210
            Width           =   975
         End
         Begin VB.OptionButton optall 
            Caption         =   "&All"
            Height          =   255
            Left            =   120
            TabIndex        =   6
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
         Left            =   150
         TabIndex        =   16
         Top             =   750
         Width           =   2625
         Begin MSComCtl2.DTPicker dptodt 
            Height          =   315
            Left            =   1335
            TabIndex        =   4
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
            Left            =   105
            TabIndex        =   3
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
            TabIndex        =   21
            Top             =   150
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Date"
            Height          =   210
            Left            =   1320
            TabIndex        =   20
            Top             =   150
            Width           =   555
         End
      End
      Begin VB.Frame frastep 
         Height          =   615
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   135
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
            TabIndex        =   1
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
            TabIndex        =   0
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
         Left            =   2850
         TabIndex        =   2
         Top             =   315
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
         Left            =   2850
         TabIndex        =   5
         Top             =   990
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
         Left            =   2850
         TabIndex        =   8
         Top             =   1695
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6465
      TabIndex        =   14
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   4140
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5175
      TabIndex        =   13
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   4140
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3885
      TabIndex        =   12
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   4140
      Width           =   1300
   End
Begin AceSpread spList
      Height          =   1545
      Left            =   7560
      TabIndex        =   22
      Top             =   2490
      Width           =   4335
      _Version        =   458752
      _ExtentX        =   7646
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
Attribute VB_Name = "rptPendingSalesEnquiryReport"
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
  Dim i As Integer
  Dim rCnt As Integer
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
  
    frmPrint.sprPrint.MaxCols = 9
    
    If OptAll = True Then
      frmPrint.sprPrint.Tag = "ALL"
    Else
      frmPrint.sprPrint.Tag = "SELECTIVE"
    End If
    
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    If FrmFlg = "SalEnquiry" Then
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.Tag = "Sales"
    Else
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.Tag = "Job Work Enquiry"
    End If
    
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
      
      If FrmFlg = "SalEnquiry" Then
        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
      Else
        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
      End If
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
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
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

    frmPrint.sprPrint.SetText 1, j, "List of Pending Enquiry ( " & frmPrint.sprPrint.Tag & " )" & " Between ( " & dpfromdt & " to " & dptodt & " )"
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. No"
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 27
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Target"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Cust.Type"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
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
  If optcustwise Then
    sSendToPort LAlign(12, "Enquiry No") & Space(1) & LAlign(50, "Item Code") & Space(1) & RAlign(8, "Quantity")
    sSendToPort LAlign(12, "Enquiry Date") & Space(1) & LAlign(50, "Item Description") & Space(1) & RAlign(8, "Rate")
  Else
    sSendToPort LAlign(12, "Enquiry No") & Space(1) & LAlign(50, "Customer") & Space(1) & RAlign(8, "Quantity")
    sSendToPort LAlign(12, "Enquiry Date") & Space(52) & RAlign(8, "Rate")
  End If
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintDetails() As Boolean
  Dim enq As String
  Dim genrec As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim enqt As String
  Dim i As Integer

  Try
   
  Set genrec = New AceADODB.Recordset
  
  If FrmFlg = "SalEnquiry" Then
    Sql = " select Distinct EnquiryType,EnquiryNo,EnquiryDate,ReferenceNo,ReferenceDate,a.CustomerCode,"
    Sql = Sql & " CustomerType,CustomerName, a.ItemType Type,a.ItemCode,ItemDescription,UOM,Qty,a.Rate,a.AssessableRate,"
    Sql = Sql & " TargetDate,Specification, b.PartNo from SalesEnquiry a Left Outer Join SAL_Customer_ItemMaster b"
    Sql = Sql & " on b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " where a.QuotationType=''"
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
    Sql = "Select Distinct a.EnquiryType,a.EnquiryNo,a.ItemGroupKey,a.SentItemCode ItemCode,"
    Sql = Sql & " a.SentDescription ItemDescription,a.EnquiryDate,a.Rate,a.SendQty Qty, a.CustomerCode,"
    'CustomerType passed in ''
    Sql = Sql & " CustomerName,'' as Type, c.UomDesc Uom,a.AssRate AssessableRate,ReferenceNo,"
    Sql = Sql & " ReferenceDate, b.PartNo from jobworkEnquiry a Left Outer Join SAL_Customer_ItemMaster b"
    Sql = Sql & " on b.ItemCode=a.SentItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid"
    Sql = Sql & " Left outer join Material_Master c  on c.ItemCode=a.SentItemCode and c.Branchid =a.Branchid"
    Sql = Sql & " and c.Compid =a.Compid where a.QuotationType=''"
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
  Sql = Sql & " and a.EnquiryDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
  Sql = Sql & " order by a.EnquiryDate"
  If RSOpen(genrec, Sql) = False Then GoTo ExitHere
  If genrec.RecordCount = 0 Then
    sPrintDetails = False
    sShowMessage "No Pending Enquiries"
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
    If OptAll Then
      gStrReportHeading = "List of Pending Enquiry (All)"
    Else
      gStrReportHeading = "List of Pending Enquiry (Selective)"
    End If
    gStrLoginUserName = "between " & dpfromdt & " to " & dptodt
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader 'Printing the Header Details
    i = 1
    Do While Not genrec.EOF
      If optitemwise Then
        If enq <> genrec("ItemCode") Then
          sSendToPort ""
          sSendToPort "Item : " & genrec("ItemDescription") & "(" & Trim(genrec("ItemCode")) & ")" & IIf(FrmFlg = "JWEnquiry", "[" & genrec("ItemGroupKey") & "]", "")
          sSendToPort ""
          sSendToPort LAlign(12, genrec("EnquiryType") & genrec("EnquiryNo")) & Space(1) & LAlign(50, genrec("CustomerName")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort LAlign(12, genrec("EnquiryDate")) & Space(52) & RAlign(8, Format(genrec("Rate"), "0.00"))
        Else
          sSendToPort Space(13) & LAlign(50, genrec("CustomerName")) & Space(1) & RAlign(8, genrec("qty"))
          sSendToPort Space(64) & RAlign(8, Format(genrec("Rate"), "0.00"))
        End If
        enq = genrec("ItemCode")
      Else
        If enq <> genrec("CustomerCode") Then
          sSendToPort ""
          sSendToPort "Customer : " & genrec("CustomerName")
          sSendToPort ""
          sSendToPort LAlign(12, genrec("EnquiryType") & genrec("EnquiryNo")) & Space(1) & LAlign(50, genrec("ItemCode")) & Space(1) & RAlign(8, genrec("qty"))  '& IIf(Pen_Enquiry = eFormList.mJWEnquiry, "[" & genrec("ItemGroupKey") & "]", ""))
          sSendToPort LAlign(12, genrec("EnquiryDate")) & Space(1) & LAlign(50, genrec("ItemDescription")) & Space(1) & RAlign(8, Format(genrec("Rate"), "0.00"))
        Else
          sSendToPort Space(13) & LAlign(50, genrec("ItemCode") & Space(1) & RAlign(8, genrec("qty"))) '& IIf(FrmFlg = "JWEnquiry", "[" & genrec("ItemGroupKey") & "]", ""))
          sSendToPort Space(13) & LAlign(50, genrec("ItemDescription")) & Space(1) & RAlign(8, Format(genrec("Rate"), "0.00"))
        End If
        enq = genrec("CustomerCode")
      End If
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

    genrec.MoveLast
genrec.MoveFirst
      Call Spread_HideRow(frmPrint.sprPrint, False, 6, frmPrint.sprPrint.MaxRows)
      Call Spread_HideCol(frmPrint.sprPrint, False, 1, frmPrint.sprPrint.MaxCols)
      
      For i = 1 To genrec.RecordCount
        If optitemwise Then
          If Trim$(Sql) <> genrec("ItemCode") Then
            If i <> 1 Then
 j = j + 1
GridCount = GridCount + 1
End If
            frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText 1, frmPrint.sprPrint.Row, "Item Code: " & genrec("ItemCode") & " - " & genrec("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
j = j + 1
          End If
        Else
          If Trim$(enq) <> genrec("CustomerCode") Then
            If i <> 1 Then
 j = j + 1
GridCount = GridCount + 1
RowProcess
End If
            frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
            If FrmFlg = "SalEnquiry" Then
              frmPrint.sprPrint.SetText 1, frmPrint.sprPrint.Row, "Customer Name: " & genrec("CustomerName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
            Else
              frmPrint.sprPrint.SetText 1, frmPrint.sprPrint.Row, "Customer Name: " & genrec("CustomerName") & "[" & genrec("ItemGroupKey") & "]"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
            End If
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
j = j + 1
          End If
        End If
        
        If Trim$(enqt) <> (genrec("EnquiryType") & genrec("EnquiryNo")) Then
          frmPrint.sprPrint.SetText 1, j, genrec("EnquiryType") & " - " & genrec("EnquiryNo")
        End If
        
        frmPrint.sprPrint.SetText 2, j, genrec("CustomerName")
        frmPrint.sprPrint.SetText 3, j, genrec("ReferenceNo")
        If FrmFlg = "SalEnquiry" Then
          frmPrint.sprPrint.SetText 9, j, genrec("TargetDate")
        Else
          frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.ColHidden = True
        End If
        frmPrint.sprPrint.SetText 4, j, genrec("Partno")
        frmPrint.sprPrint.SetText 5, j, genrec("ItemCode")
        frmPrint.sprPrint.SetText 6, j, genrec("UOM")
        frmPrint.sprPrint.SetText 7, j, Format(genrec("Rate"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.SetText 8, j, Format(genrec("qty"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        If Trim$(enqt) <> (genrec("EnquiryType") & genrec("EnquiryNo")) Then
          frmPrint.sprPrint.SetText 1, j, genrec("EnquiryDate")
        End If
        If FrmFlg <> "SalEnquiry" Then
 frmPrint.sprPrint.SetText 2, j, genrec("ItemGroupKey")
End If
        frmPrint.sprPrint.SetText 3, j, genrec("ReferenceDate")
        frmPrint.sprPrint.SetText 4, j, IIf(genrec("Type") = "E", "Existing", "New")
        frmPrint.sprPrint.SetText 5, j, genrec("ItemDescription")
        frmPrint.sprPrint.SetText 7, j, Format(genrec("AssessableRate"), "00.00")
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        enq = genrec("CustomerCode")
        Sql = genrec("ItemCode")
        enqt = genrec("EnquiryType") & genrec("EnquiryNo")
        genrec.MoveNext
      Next i
      If optitemwise Then
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColHidden = True
      Else
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColHidden = True
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

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call Form_Clear(Me)
  chkPrintDate.Value = 1
  cmbReportType.ListIndex = 0
  Call StepClear(Me)
  Call StepEnable(Me)
  OptAll.Value = True
  CmdStep(2).Enabled = False
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

  If sPrintDetails = False Then
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

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdview.Enabled = True
  End If
  spList.Enabled = False
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  Try
  
  CurDate = Format(ServerDate, "dd/mm/yyyy")
  dpfromdt = Format(CurDate, "dd/mm/yyyy")
  dptodt = Format(CurDate, "dd/mm/yyyy")
  
  If FrmFlg = "JWEnquiry" Then
    rptPendingSalesEnquiryReport.Caption = "JobWork Enquiry Pending Report"
  Else
    rptPendingSalesEnquiryReport.Caption = "Sales Enquiry Pending Report"
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
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  spList.Visible = False
  CmdStep(2).Enabled = True
End Sub

Private Sub optcustwise_Click()
' On Error Resume Next
  OptSelective_Click
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass

  CmdStep(2).Enabled = False

  If OptAll Then
 GoTo ExitHere
End If
  If Pen_Enquiry = eFormList.mSalEnquiry Then
    If optcustwise Then
      Sql = "Select Distinct CustomerCode [Customer Code],CustomerName [Customer Name]"
      Sql = Sql + " from SalesEnquiry a where a.QuotationType ='' and a.Branchid ='" & gBranchID & "'"
      Sql = Sql + " and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise Then
      Sql = "select Distinct a.ItemCode Item,a.ItemDescription from SalesEnquiry a where"
      Sql = Sql & " a.QuotationType='' and a.Branchid ='" & gBranchID & "'"
      Sql = Sql + " and a.Compid = '" & gCompID & "'"
    End If
  Else
    If optcustwise Then
      Sql = "Select Distinct CustomerCode [Customer Code],CustomerName [Customer Name]"
      Sql = Sql + " from jobworkEnquiry a where a.QuotationType ='' and a.Branchid ='" & gBranchID & "'"
      Sql = Sql + " and a.Compid = '" & gCompID & "'"
    ElseIf optitemwise Then
      Sql = "select Distinct a.SentItemCode Item,a.SentDescription"
      Sql = Sql & " from jobworkEnquiry a where a.QuotationType='' and a.Branchid ='" & gBranchID & "'"
      Sql = Sql + " and a.Compid = '" & gCompID & "'"
    End If
  End If
  Sql = Sql & " and a.EnquiryDate between convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and convert(datetime, '" & Format(dptodt, "yyyy-MM-dd") & "',111)"
  
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.ClearSelection
'  If spList.DataRowCnt = 0 Then
'    sShowMessage "No Records Found"
'    optAll.SetFocus
'    GoTo ExitHere
'  End If
'  If optcustwise Then
'    spList.Col = 1: spList.ColHidden = True
'  End If
'  spList.MaxRows = spList.DataRowCnt
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
    'To Hide Customer Code
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
  End If
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

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\PendingSalesEnquiryReport.frm", spList, "")

End Sub
