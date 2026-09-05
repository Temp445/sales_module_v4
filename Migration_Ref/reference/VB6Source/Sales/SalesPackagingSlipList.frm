VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptSalesPackagingSlipList 
   Caption         =   "Sales Packaging Slip List"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00404040&
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
Begin AceSpread fpSpread1
      Height          =   1620
      Left            =   540
      TabIndex        =   23
      Top             =   14355
      Visible         =   0   'False
      Width           =   2190
      _Version        =   458752
      _ExtentX        =   3863
      _ExtentY        =   2858
      _StockProps     =   64
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
      MaxCols         =   2
      MaxRows         =   6
   End
   Begin VB.CommandButton cmdStep 
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
      Left            =   11325
      TabIndex        =   8
      Top             =   4245
      Width           =   345
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8070
      Left            =   4410
      TabIndex        =   14
      Top             =   3210
      Width           =   13380
      Begin VB.Frame Frame2 
         Caption         =   "Print Label Starts From"
         Height          =   705
         Left            =   2955
         TabIndex        =   21
         Top             =   4350
         Width           =   1905
         Begin VB.ComboBox cmbLabelNo 
            Enabled         =   0   'False
            Height          =   330
            ItemData        =   "SalesPackagingSlipList.frx":0338
            Left            =   105
            List            =   "SalesPackagingSlipList.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   240
            Width           =   1680
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   400
         Left            =   4635
         TabIndex        =   13
         Top             =   6105
         Width           =   1300
      End
      Begin VB.Frame ReportType 
         Caption         =   "Report Type"
         Height          =   930
         Left            =   2705
         TabIndex        =   19
         Top             =   5085
         Width           =   2550
         Begin VB.CheckBox chkPrintdate 
            Caption         =   "Print Date"
            Height          =   255
            Left            =   90
            TabIndex        =   10
            Top             =   615
            Width           =   1080
         End
         Begin VB.ComboBox cmbReportType 
            Enabled         =   0   'False
            Height          =   330
            ItemData        =   "SalesPackagingSlipList.frx":038B
            Left            =   90
            List            =   "SalesPackagingSlipList.frx":0395
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   210
            Width           =   2370
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   3345
         TabIndex        =   12
         Top             =   6105
         Width           =   1300
      End
      Begin VB.Frame FraStep 
         Caption         =   "Select"
         Height          =   1440
         Index           =   0
         Left            =   1130
         TabIndex        =   15
         Top             =   120
         Width           =   5700
         Begin VB.Frame fraInvoiceNo 
            Height          =   585
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   2850
            Begin VB.TextBox txtInvoiceTo 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1560
               TabIndex        =   5
               Top             =   195
               Width           =   1185
            End
            Begin VB.TextBox txtInvoiceFrom 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   90
               TabIndex        =   4
               Top             =   195
               Width           =   1185
            End
            Begin VB.Label lblTo 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "to"
               Height          =   210
               Index           =   1
               Left            =   1350
               TabIndex        =   18
               Top             =   210
               Width           =   165
            End
         End
         Begin VB.Frame fraInvoiceDate 
            Height          =   585
            Left            =   1680
            TabIndex        =   16
            Top             =   135
            Visible         =   0   'False
            Width           =   2850
            Begin MSComCtl2.DTPicker dtpInvoiceTo 
               Height          =   330
               Left            =   1560
               TabIndex        =   2
               Top             =   180
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   582
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
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   135790595
               CurrentDate     =   38257
            End
            Begin MSComCtl2.DTPicker dtpInvoiceFrom 
               Height          =   330
               Left            =   90
               TabIndex        =   1
               Top             =   180
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   582
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
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   135790595
               CurrentDate     =   38257
            End
            Begin VB.Label lblTo 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "to"
               Height          =   210
               Index           =   0
               Left            =   1350
               TabIndex        =   17
               Top             =   240
               Width           =   165
            End
         End
         Begin VB.CommandButton cmdProceed 
            Caption         =   "&Proceed"
            Height          =   390
            Left            =   4620
            TabIndex        =   7
            Top             =   870
            Width           =   960
         End
         Begin VB.CheckBox chkInvoiceNo 
            Caption         =   "By Invoice No"
            Height          =   225
            Left            =   120
            TabIndex        =   3
            Top             =   930
            Width           =   1365
         End
         Begin VB.CheckBox chkInvoiceDate 
            Caption         =   "By Invoice Date"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   390
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&Proceed"
         Height          =   400
         Left            =   2055
         TabIndex        =   11
         Top             =   6105
         Width           =   1300
      End
Begin AceSpread sprInvoice
         Height          =   2670
         Left            =   285
         TabIndex        =   20
         Top             =   1650
         Width           =   7830
         _Version        =   458752
         _ExtentX        =   13811
         _ExtentY        =   4710
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
         MaxCols         =   6
         OperationMode   =   3
         ScrollBars      =   2
         SelectBlockOptions=   0
      End
   End
End
Attribute VB_Name = "rptSalesPackagingSlipList"
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
End Enum

Private i As Integer
Private retVal As Variant
Dim k As Integer
Dim j As Long
Public Sub sprint()
  'No Change
End Sub
Public Sub sView()
  'No Change
End Sub

Private Sub RowProcess1()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.SprPrint.MaxRows = frmPrint.SprPrint.MaxRows + pheight + 1
      frmPrint.SprPrint.Row = frmPrint.SprPrint.Row + 1
j = j + 1
      frmPrint.SprPrint.RowPageBreak = 1
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
  
    frmPrint.SprPrint.MaxCols = 8
    frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(1) = 15:

    frmPrint.SprPrint.Col = 2
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(2) = 20

    frmPrint.SprPrint.Col = 3
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(3) = 15

    frmPrint.SprPrint.Col = 4
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(4) = 20

    frmPrint.SprPrint.Col = 5
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(5) = 15

    frmPrint.SprPrint.Col = 6
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(6) = 20

    frmPrint.SprPrint.Col = 7
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(7) = 15

    frmPrint.SprPrint.Col = 8
frmPrint.SprPrint.Row = j
frmPrint.SprPrint.ColWidth(8) = 20
'
'      SprInvoice.ColWidth(1) = 47.88
'      SprInvoice.ColWidth(2) = 47.88
'      SprInvoice.RowHeight(1) = 119
'      SprInvoice.RowHeight(2) = 119
    
    frmPrint.SprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.Row = j

    frmPrint.SprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
    
    frmPrint.SprPrint.Col = 7
frmPrint.SprPrint.Row = j
    frmPrint.SprPrint.AddCellSpan frmPrint.SprPrint.Col, j, frmPrint.SprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmPrint.SprPrint.SetText frmPrint.SprPrint.Col, j, Space(20) & "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
    Set rs = New Recordset
    If RSOpen(rs, Sql) = False Then GoTo ExitHere
    If rs.RecordCount > 0 Then
      frmPrint.SprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.Row = j

      frmPrint.SprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
      
      frmPrint.SprPrint.Col = 7
frmPrint.SprPrint.Row = j
      frmPrint.SprPrint.AddCellSpan frmPrint.SprPrint.Col, j, frmPrint.SprPrint.MaxCols, 1
      frmPrint.SprPrint.SetText frmPrint.SprPrint.Col, j, Space(20) & "TS Ref. No   : " & rs("TsNo")
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If
    
    If iPageNo = 1 Then
      frmPrint.SprPrint.AddCellSpan 1, j, frmPrint.SprPrint.MaxCols, 1
      frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.Row = j

      frmPrint.SprPrint.SetText 1, j, gCompName
frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True
frmPrint.SprPrint.FontSize = 12
frmPrint.SprPrint.FontUnderline = True
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.SprPrint.AddCellSpan 1, j, frmPrint.SprPrint.MaxCols, 1
    frmPrint.SprPrint.Col = 1
frmPrint.SprPrint.Row = j

    frmPrint.SprPrint.SetText 1, j, "PACKAGING SLIP LIST"
frmPrint.SprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True
frmPrint.SprPrint.FontSize = 10
frmPrint.SprPrint.FontUnderline = True

    j = j + 1
GridCount = GridCount + 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sPageHeader()
  
  Try
  
  If CmdView Then
    sSendToPort String(iPageWidth, "-")
  Else
    Printer.Print Space(1) & String$(35, "-") & Space(1) & String$(35, "-") & Space(1) & String$(35, "-")
k = k + 1
RowProcess
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Function Dumrow(ByVal rownos As Integer)
  Dim rowcnt As Integer
  
' On Error Resume Next
  
  For rowcnt = 1 To rownos
    Printer.Print
  Next rowcnt
  
End Function

Private Function sPrintCurrent()
  Dim snap As AceADODB.Recordset
  Dim Sql As String
     
  Try
 
  If SprInvoice.SelectionCount < 1 Then
    sShowMessage "Please select atlease one item"
    GoTo ExitHere
  End If

    For i = 1 To SprInvoice.SelectionCount
      SprInvoice.GetText mInvNo, SprInvoice.ActiveRow, retVal
      Sql = " select a.InvoiceNo, a.InvoiceDate, a.CustomerCode, (select b.PartyName from FAS_Party_Master b"
      Sql = Sql + " where b.PartyCode = a.CustomerCode and b.Branchid = a.Branchid and b.Compid = a.Compid) CustomerName,"
      Sql = Sql + " a.NoofPackets, a.ItemCode, (select c.ItemDescription from INV_Material_Master c"
      Sql = Sql + " where c.Itemcode = a.Itemcode and c.Branchid = a.Branchid and c.Compid = a.Compid) ItemDescription,"
      Sql = Sql + " (Select  Top 1 PartNO From PartMast Where StageCode = a.ItemCode and BranchID = a.BranchID and"
      Sql = Sql + " CompID = a.COmpID)PartNo, a.InvoiceQty from SalesInvoice a where InvoiceNo ='" & retVal & " '"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      Sql = Sql + " order by a.InvoiceNo, a.InvoiceDate"
    Next i
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
     
  k = 1
  iPageNo = 1
  iPageWidth = 80
  pheight = 75
  Printer.PaperSize = vbPRPSFanfoldUS
  Printer.FontName = "Draft 12cpi"
  If Printer.FontName <> "Draft 12cpi" Then
    Printer.FontName = "Roman 12cpi"
    If Printer.FontName <> "Roman 12cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If
   
  Do While Not snap.EOF
    Dumrow 2
    k = k + 3
RowProcess
    For i = 1 To snap.RecordCount
      sPageHeader
      Printer.Print Space(1); LAlign(35, "Customer Name:" & snap("CustomerName") & "") & Space(1) & LAlign(35, "Customer Name:" & snap("CustomerName") & "") & Space(1) & LAlign(35, "Customer Name:" & snap("CustomerName") & "")
k = k + 1
RowProcess

      Printer.Print Space(1); LAlign(35, "Part Name    :" & snap("ItemDescription") & "") & Space(1) & LAlign(35, "Part Name    :" & snap("ItemDescription") & "") & Space(1) & LAlign(35, "Part Name    :" & snap("ItemDescription") & "")
k = k + 1
RowProcess

      Printer.Print Space(1); LAlign(35, "Part NO      :" & snap("PartNo") & "") & Space(1) & LAlign(35, "Part NO      :" & snap("PartNo") & "") & Space(1) & LAlign(35, "Part NO      :" & snap("PartNo") & "")
k = k + 1
RowProcess

      Printer.Print Space(1); LAlign(35, "Qty          :") & Space(1) & LAlign(35, "Qty          :") & Space(1) & LAlign(35, "Qty          :")
k = k + 1
RowProcess

      Printer.Print Space(1); LAlign(35, "Invoice No   :" & snap("InvoiceNo") & "") & Space(1) & LAlign(35, "Invoice No   :" & snap("InvoiceNo") & "") & Space(1) & LAlign(35, "Invoice No   :" & snap("InvoiceNo") & "")
k = k + 1
RowProcess

      Printer.Print Space(1); LAlign(35, "Invoice Qty  :" & snap("InvoiceQty") & "") & Space(1) & LAlign(35, "Invoice Qty  :" & snap("InvoiceQty") & "") & Space(1) & LAlign(35, "Invoice Qty  :" & snap("InvoiceQty") & "")
k = k + 1
RowProcess
      If Not snap.RecordCount Then
      End If
      snap.MoveNext
    Next i
  Loop
  Printer.Print Space(1) & String$(35, "-") & Space(1) & String$(35, "-") & Space(1) & String$(35, "-")
k = k + 1
RowProcess

  Dumrow 3
k = k + 3
RowProcess
  If snap.EOF Then

End If
  snap.Close
  Printer.EndDoc
  sShowMessage "Print Over"
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set snap = Nothing
End Try
End Function

Private Function sPrintView() As Boolean
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim cCol As Integer
  
  Try
  
  If SprInvoice.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
  
    For i = 1 To SprInvoice.SelectionCount
      SprInvoice.GetText mInvNo, SprInvoice.ActiveRow, retVal
      Sql = " select a.InvoiceNo, a.InvoiceDate, a.CustomerCode, (select b.PartyName from FAS_Party_Master b"
      Sql = Sql + " where b.PartyCode = a.CustomerCode and b.Branchid = a.Branchid and b.Compid = a.Compid) CustomerName,"
      Sql = Sql + " a.NoofPackets, a.ItemCode, (select c.ItemDescription from INV_Material_Master c"
      Sql = Sql + " where c.Itemcode = a.Itemcode and c.Branchid = a.Branchid and c.Compid = a.Compid) ItemDescription,"
      Sql = Sql + " (Select  Top 1 PartNO From PartMast Where StageCode = a.ItemCode and BranchID = a.BranchID and"
      Sql = Sql + " CompID = a.COmpID)PartNo, a.InvoiceQty from SalesInvoice a where a.InvoiceNo ='" & retVal & " '"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      Sql = Sql + " order by a.InvoiceNo, a.InvoiceDate"
    Next i  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
  
  k = 1
  If cmbReportType.ListIndex = 1 Then
    iPageNo = 1
    'iPageWidth = 132
    pheight = 75
    'Printer.PaperSize = vbPRPSFanfoldStdGerman
    'pheight = 75
    'pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    'fOpenPort Me
    Printer.FontName = "Draft 8cpi"
    If Printer.FontName <> "Draft 8cpi" Then
      Printer.FontName = "Roman 8cpi"
      If Printer.FontName <> "Roman 8cpi" Then
        Printer.FontName = "Courier New"
        Printer.FontSize = 11
      End If
    End If
    fOpenPort Me
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "PACKAGING SLIP LIST")
    sSendToPort ""
    For i = 1 To snap.RecordCount
      sPageHeader
      sSendToPort LAlign(32, "Customer Name:" & snap("CustomerName") & "") & Space(1) & LAlign(32, "Customer Name:" & snap("CustomerName") & "") & Space(1) & LAlign(32, "Customer Name:" & snap("CustomerName") & "") & Space(1) & LAlign(32, "Customer Name:" & snap("CustomerName") & "")
      sSendToPort LAlign(32, "Part Name    :" & snap("ItemDescription") & "") & Space(1) & LAlign(32, "Part Name    :" & snap("ItemDescription") & "") & Space(1) & LAlign(32, "Part Name    :" & snap("ItemDescription") & "") & Space(1) & LAlign(32, "Part Name    :" & snap("ItemDescription") & "")
      sSendToPort LAlign(32, "Part NO      :" & snap("PartNo") & "") & Space(1) & LAlign(32, "Part NO      :" & snap("PartNo") & "") & Space(1) & LAlign(32, "Part NO      :" & snap("PartNo") & "") & Space(1) & LAlign(32, "Part NO      :" & snap("PartNo") & "")
      sSendToPort LAlign(32, "Qty          :") & Space(1) & LAlign(32, "Qty          :") & Space(1) & LAlign(32, "Qty          :") & Space(1) & LAlign(32, "Qty          :")
      sSendToPort LAlign(32, "Invoice No   :" & snap("InvoiceNo") & "") & Space(1) & LAlign(32, "Invoice No   :" & snap("InvoiceNo") & "") & Space(1) & LAlign(32, "Invoice No   :" & snap("InvoiceNo") & "") & Space(1) & LAlign(32, "Invoice No   :" & snap("InvoiceNo") & "")
      sSendToPort LAlign(32, "Invoice Qty  :" & snap("InvoiceQty") & "") & Space(1) & LAlign(32, "Invoice Qty  :" & snap("InvoiceQty") & "") & Space(1) & LAlign(32, "Invoice Qty  :" & snap("InvoiceQty") & "") & Space(1) & LAlign(32, "Invoice Qty  :" & snap("InvoiceQty") & "")
      snap.MoveNext
    Next i
    If snap.EOF Then
 sSendToPort ""
End If
    snap.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
      frmPrint.Show
      
      Header
      frmPrint.SprPrint.Tag = j + 1
      For i = 1 To snap.RecordCount
        j = 2
cCol = 1
        For k = 1 To 4
          Spread_Set_CellBorder frmPrint.SprPrint, CellBorderIndexOutline, 1, Val(frmPrint.SprPrint.Tag), frmPrint.SprPrint.MaxCols, Val(frmPrint.SprPrint.Tag)
          frmPrint.SprPrint.SetText cCol, Val(frmPrint.SprPrint.Tag), "Customer Name"
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = cCol
frmPrint.SprPrint.TypeHAlign = TypeHAlignRight
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True

          frmPrint.SprPrint.SetText j, Val(frmPrint.SprPrint.Tag), snap("CustomerName")
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = j
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 1
GridCount = GridCount + 1
RowProcess

          Spread_Set_CellBorder frmPrint.SprPrint, CellBorderIndexOutline, 1, Val(frmPrint.SprPrint.Tag), frmPrint.SprPrint.MaxCols, Val(frmPrint.SprPrint.Tag)
          frmPrint.SprPrint.SetText cCol, Val(frmPrint.SprPrint.Tag), "Part Name"
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = cCol
frmPrint.SprPrint.TypeHAlign = TypeHAlignRight
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True

          frmPrint.SprPrint.SetText j, Val(frmPrint.SprPrint.Tag), snap("ItemDescription")
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = j
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 1
GridCount = GridCount + 1
RowProcess
          
          Spread_Set_CellBorder frmPrint.SprPrint, CellBorderIndexOutline, 1, Val(frmPrint.SprPrint.Tag), frmPrint.SprPrint.MaxCols, Val(frmPrint.SprPrint.Tag)
          frmPrint.SprPrint.SetText cCol, Val(frmPrint.SprPrint.Tag), "Part No"
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = cCol
frmPrint.SprPrint.TypeHAlign = TypeHAlignRight
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True

          frmPrint.SprPrint.SetText j, Val(frmPrint.SprPrint.Tag), snap("PartNo")
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = j
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 1
GridCount = GridCount + 1
RowProcess
          
          Spread_Set_CellBorder frmPrint.SprPrint, CellBorderIndexOutline, 1, Val(frmPrint.SprPrint.Tag), frmPrint.SprPrint.MaxCols, Val(frmPrint.SprPrint.Tag)
          frmPrint.SprPrint.SetText cCol, Val(frmPrint.SprPrint.Tag), "Invoice No"
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = cCol
frmPrint.SprPrint.TypeHAlign = TypeHAlignRight
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True

          frmPrint.SprPrint.SetText j, Val(frmPrint.SprPrint.Tag), snap("InvoiceQty")
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = j
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 1
GridCount = GridCount + 1
RowProcess
          
          Spread_Set_CellBorder frmPrint.SprPrint, CellBorderIndexOutline, 1, Val(frmPrint.SprPrint.Tag), frmPrint.SprPrint.MaxCols, Val(frmPrint.SprPrint.Tag)
          frmPrint.SprPrint.SetText cCol, Val(frmPrint.SprPrint.Tag), "Invoice Qty"
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = cCol
frmPrint.SprPrint.TypeHAlign = TypeHAlignRight
frmPrint.SprPrint.Font = "Arial"
frmPrint.SprPrint.FontBold = True

          frmPrint.SprPrint.SetText j, Val(frmPrint.SprPrint.Tag), snap("InvoiceNo")
frmPrint.SprPrint.Row = Val(frmPrint.SprPrint.Tag)
frmPrint.SprPrint.Col = j
frmPrint.SprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 1
GridCount = GridCount + 1
RowProcess

          j = j + 2
cCol = cCol + 2
frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) - 5
GridCount = GridCount + 1
RowProcess
        Next k
        frmPrint.SprPrint.Tag = Val(frmPrint.SprPrint.Tag) + 6
'        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, Val(.Tag), .MaxCols, Val(.Tag)

        snap.MoveNext
      Next i
      frmPrint.SprPrint.MaxRows = frmPrint.SprPrint.DataRowCnt + 1
      frmPrint.Show
  End If
  
  sPrintView = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function sPrintPackingList() As Boolean
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  Dim cCol As Integer
  
  Dim curX As Integer
  Dim curY As Integer
  
  Dim lblLeftX As Integer
  Dim lblTopY As Integer
  
  Dim tmp As Integer
  Dim tmp2 As Integer
  Dim txtToPrint As String
  
  Dim LabelWidth As Integer
  Dim LabelHeight As Integer
  
  Dim LabelPerPageCount As Integer
  Dim LabelPerRow As Integer
  
  Dim LabelStartNo As Integer
  Dim LabelNo As Integer
  Dim OffX As Integer
  Dim OffY As Integer
  Dim HorGap As Integer
  Dim VerGap As Integer
  Dim fTxt As String
  Dim k As Integer
  Dim s As String
  Dim tXPos As Single
  Dim tYPos  As Single

  Try
  
  fTxt = ""
  
  Sql = " select CompanyName, City, Pincode,(Address1+', '+Address2) as Add1, City+'-'+Pincode as Add2, Convert(Varchar, Phone) as Add3 "
  Sql = Sql & "from company_master where compid = '" & gCompID & "'"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If SprInvoice.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If

    For i = 1 To SprInvoice.SelectionCount
      Sql = " Execute SAL_Rpt_SalesPackagingSlipList_Gen"
      
      SprInvoice.GetText eInv.mInvNo, SprInvoice.ActiveRow, retVal
      Sql = Sql & " @InvNo=" & Val(retVal) & ","
      
      Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', "
      Sql = Sql & " @FinYear='" & gFinyear & "'"
    Next i
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  Else
    'Continue
  End If
  
  LabelWidth = 99
  LabelHeight = 43
  
  LabelPerPageCount = 12
  LabelPerRow = 2
  
  HorGap = 1    'Horizontal Gap between two Labels
  VerGap = 4    'Vertical Gap between two Labels
  
  OffX = -2  'Horizontal Offset by Printer
  OffY = -5  'Vertical Offset by Printer
  
  LabelStartNo = cmbLabelNo

  Printer.ScaleMode = 6
  Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
 
'  For k = 1 To snap.RecordCount
'  For i = 1 To snap.RecordCount
'    'TopLeft
'    curX = IIf(k Mod 2 = 1, 4, LabelWidth + HorGap)
'    curY = IIf(k <= 2, 9, (LabelNo \ 2) * (LabelHeight + VerGap))
    
    LabelNo = LabelStartNo - 1
     
    Do While Not snap.EOF
      
      LabelNo = LabelNo + 1
      
      If LabelNo > LabelPerPageCount Then
        LabelNo = 1
        Printer.NewPage
      Else
        'Continue
      End If
      
      'TopLeft
      lblLeftX = IIf(LabelNo Mod LabelPerRow = 1, 4, 4 + LabelWidth + HorGap)
      lblTopY = IIf(LabelNo <= LabelPerRow, 9, 9 + ((LabelNo - 1) \ LabelPerRow) * (LabelHeight + VerGap))

      Printer.FontName = "Arial"
      
      Printer.FontSize = 11
      Printer.FontBold = True
      
      'Set Print With ForeColor: Company Name and Address For Using Navy Blue
      Printer.ForeColor = RGB(32, 42, 68)     'Navy Blue
      
      'Company Name
      curY = 2
      txtToPrint = UCase$(rs("CompanyName"))
      tmp = Printer.TextWidth(txtToPrint)
      
      If tmp >= (LabelWidth - 12) Then      '12mm allocated for left band printing "PACKING SLIP"
        curX = 0
      Else
        curX = 12 + ((LabelWidth - 12 - tmp) \ 2)
      End If
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      'Company Address
      curY = curY + 5         '5mm line spacing for font Arial 11
      txtToPrint = UCase$(rs("City") & "-" & rs("Pincode"))
      tmp = Printer.TextWidth(txtToPrint)
      If tmp >= (LabelWidth - 12) Then
        curX = 0
      Else
        curX = 12 + ((LabelWidth - 12 - tmp) \ 2)
      End If
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
            
      Printer.FontSize = 10
      Printer.FontBold = False
      'Customer Name
      curY = curY + 6
      curX = 13
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Line (0, -1)-(LabelWidth - 12, -1), vbBlack
      
      'Set Print With ForeColor: Other(s) For Using Military Green
      Printer.ForeColor = RGB(75, 83, 32)     'Military Green
      
      txtToPrint = "Customer Name"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = Left(snap("CustomerName"), 25)
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Item Description
      curY = curY + 5
      curX = 13
      txtToPrint = "Description"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = Left(snap("ItemDescription"), 25)
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Part No
      curY = curY + 5
      curX = 13
      txtToPrint = "Part No"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("PartNo")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Invoice Qty
      curY = curY + 5
      curX = 13
      txtToPrint = "Quantity"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("PacketQty")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Invoice No & Date
      curY = curY + 5
      curX = 13
      txtToPrint = "Invoice No"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("InvoiceNo") & Space(12) & "Date : " & Format(snap("InvoiceDate"), "dd.MM.yyyy")

      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      'No. of Packets
      curY = curY + 5
      curX = 13
      txtToPrint = "No.of Pkts"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("NoofPackets") & " " & Trim$("X") & " " & snap("PacketQty")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      Printer.FontBold = True
      
      'PACKING **-------------------------
      fTxt = "PACKING"
      
      curX = 2
      curY = 4    'Per character 5mm. Total for 7 character 35mm. So label width 43 - 35 = 8. Split as top 4 and bottom 4
      
      For tmp = 1 To Len(Trim$(fTxt))
        txtToPrint = Mid$(Trim$(fTxt), tmp, 1)
        Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
        Printer.Print txtToPrint

'        Printer.Line (0, -1)-(LabelWidth - 82, -1), vbYellow
      
        curY = curY + 5
      Next tmp
      '-----------------------------------
      
      'SLIP **----------------------------
      fTxt = "SLIP"
      
      curX = 7
      curY = 11    'Per character 5mm. Total for 4 character 20mm. So label width 43 - 20 = 23. Split as top 11 and bottom 12
      
      For tmp = 1 To Len(Trim$(fTxt))
        txtToPrint = Mid$(Trim$(fTxt), tmp, 1)
        
'        Line (lblLeftX + curX+ OffX, lblTopY + curY+ OffY)-(lblLeftX + curX + TextWidth(txtToPrint) + OffX, lblTopY + curY + TextHeight(txtToPrint) + OffY), vbYellow, BF
        
        Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
        Printer.Print txtToPrint
  
        curY = curY + 5
      Next tmp
      '-----------------------------------
      
      snap.MoveNext
    Loop
    
  Printer.EndDoc

'      frmPrint.Show
'    End With
'  End If
  
  sPrintPackingList = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function sPrintPackingList_Old() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  Dim cCol As Integer
  
  Dim curX As Integer
  Dim curY As Integer
  
  Dim lblLeftX As Integer
  Dim lblTopY As Integer
  
  Dim tmp As Integer
  Dim tmp2 As Integer
  Dim txtToPrint As String
  
  Dim LabelWidth As Integer
  Dim LabelHeight As Integer
  
  Dim LabelPerPageCount As Integer
  Dim LabelPerRow As Integer
  
  Dim LabelStartNo As Integer
  Dim LabelNo As Integer
  Dim OffX As Integer
  Dim OffY As Integer
  Dim HorGap As Integer
  Dim VerGap As Integer
  
  Dim k As Integer
  Dim s As String
  Dim tXPos As Single
  Dim tYPos  As Single

  Try
  
  Sql = " select CompanyName, City, Pincode,(Address1+', '+Address2) as Add1, City+'-'+Pincode as Add2, Convert(Varchar, Phone) as Add3 "
  Sql = Sql & "from company_master where compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If SprInvoice.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If

    For i = 1 To SprInvoice.SelectionCount
      SprInvoice.GetText mInvNo, SprInvoice.ActiveRow, retVal
      Sql = " Execute SAL_Rpt_SalesPackagingSlipList_Gen"
      SprInvoice.GetText eInv.mInvNo, SprInvoice.ActiveRow, retVal
      Sql = Sql & " @InvNo=" & Val(retVal) & ","
      Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', "
      Sql = Sql & " @FinYear='" & gFinyear & "'"
    Next i
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
  
  LabelWidth = 99
  LabelHeight = 43
  
  LabelPerPageCount = 12
  LabelPerRow = 2
  
  HorGap = 1    'Horizontal Gap between two Labels
  VerGap = 4    'Vertical Gap between two Labels
  
  OffX = -2  'Horizontal Offset by Printer
  OffY = -5  'Vertical Offset by Printer
  
  LabelStartNo = cmbLabelNo

  Printer.ScaleMode = 6
  Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
  
'  For k = 1 To snap.RecordCount
'  For i = 1 To snap.RecordCount
'    'TopLeft
'    curX = IIf(k Mod 2 = 1, 4, LabelWidth + HorGap)
'    curY = IIf(k <= 2, 9, (LabelNo \ 2) * (LabelHeight + VerGap))
    
    LabelNo = LabelStartNo - 1
     
    Do While Not snap.EOF
      
      LabelNo = LabelNo + 1
      
      If LabelNo > LabelPerPageCount Then
        LabelNo = 1
        Printer.NewPage
      End If
      
      'TopLeft
      lblLeftX = IIf(LabelNo Mod LabelPerRow = 1, 4, 4 + LabelWidth + HorGap)
      lblTopY = IIf(LabelNo <= LabelPerRow, 9, 9 + ((LabelNo - 1) \ LabelPerRow) * (LabelHeight + VerGap))
      
'      Printer.Line (1, 1)-(lblLeftX + LabelWidth - 1, lblTopY + LabelHeight - 1), , B
      
      Printer.FontName = "Arial"
      
      Printer.FontSize = 11
      Printer.FontBold = True

      'Company Name
      curY = 2
      txtToPrint = UCase$(rs("CompanyName"))
      tmp = Printer.TextWidth(txtToPrint)
      If tmp >= (LabelWidth - 12) Then      '12mm allocated for left band printing "PACKING SLIP"
        curX = 0
      Else
        curX = 12 + ((LabelWidth - 12 - tmp) \ 2)
      End If
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      'Company Address
      curY = curY + 5         '5mm line spacing for font Arial 11
      txtToPrint = UCase$(rs("City") & "-" & rs("Pincode"))
      tmp = Printer.TextWidth(txtToPrint)
      If tmp >= (LabelWidth - 12) Then
        curX = 0
      Else
        curX = 12 + ((LabelWidth - 12 - tmp) \ 2)
      End If
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
            
      Printer.FontSize = 10
      Printer.FontBold = False
      'Customer Name
      curY = curY + 6
      curX = 13
      
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Line (0, -1)-(LabelWidth - 12, -1), vbBlack
      
      txtToPrint = "Customer Name"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = Left(snap("CustomerName"), 25)
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Item Description
      curY = curY + 5
      curX = 13
      txtToPrint = "Description"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = Left(snap("ItemDescription"), 25)
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Part No
      curY = curY + 5
      curX = 13
      txtToPrint = "Part No"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("PartNo")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Invoice Qty
      curY = curY + 5
      curX = 13
      txtToPrint = "Quantity"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("PacketQty")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      'Invoice No & Date
      curY = curY + 5
      curX = 13
      txtToPrint = "Invoice No"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("InvoiceNo") & Space(12) & "Date : " & Format(snap("InvoiceDate"), "dd.MM.yyyy")

      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint

      
      'No. of Packets
      curY = curY + 5
      curX = 13
      txtToPrint = "No.of Pkts"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 40
      txtToPrint = ":"
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      curX = 42
      txtToPrint = snap("NoofPackets") & " " & Trim$("X") & " " & snap("PacketQty")
      Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
      Printer.Print txtToPrint
      
      Printer.FontBold = True
      
      'PACKING
      curX = 2
      curY = 4    'Per character 5mm. Total for 7 character 35mm. So label width 43 - 35 = 8. Split as top 4 and bottom 4
      For tmp = 1 To 7
        txtToPrint = Mid$("PACKING", tmp, 1)
        Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
        Printer.Print txtToPrint
      
        curY = curY + 5
      Next tmp
      
      'SLIP
      curX = 7
      curY = 11    'Per character 5mm. Total for 4 character 20mm. So label width 43 - 20 = 23. Split as top 11 and bottom 12
      For tmp = 1 To 4
        txtToPrint = Mid$("SLIP", tmp, 1)
        Printer.CurrentX = lblLeftX + curX + OffX
Printer.CurrentY = lblTopY + curY + OffY
        Printer.Print txtToPrint
      
        curY = curY + 5
      Next tmp
      
      snap.MoveNext
    Loop
    
  Printer.EndDoc

'      frmPrint.Show
'    End With
'  End If
  
  sPrintPackingList_Old = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function sPrintPack() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  Dim cCol As Integer
  
  Dim curX As Integer
  Dim curY As Integer
  
  Dim LabelWidth As Integer
  Dim LabelHeight As Integer
  
  Dim LabelNo As Integer
  Dim OffX As Integer
  Dim OffY As Integer
  Dim HorGap As Integer
  Dim VerGap As Integer
  
  Dim k As Integer
  Dim s As String
  Dim tXPos As Single
  Dim tYPos  As Single

  Try
  
  Sql = " select CompanyName, City, Pincode,(Address1+', '+Address2) as Add1, City+'-'+Pincode as Add2, Convert(Varchar, Phone) as Add3 "
  Sql = Sql & "from company_master where compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If SprInvoice.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If

    For i = 1 To SprInvoice.SelectionCount
      SprInvoice.GetText mInvNo, SprInvoice.ActiveRow, retVal
      Sql = " Execute SAL_Rpt_SalesPackagingSlipList_Gen"
      SprInvoice.GetText eInv.mInvNo, SprInvoice.ActiveRow, retVal
      Sql = Sql & " @InvNo=" & Val(retVal) & ","
      Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', "
      Sql = Sql & " @FinYear='" & gFinyear & "'"
    Next i
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
  
  LabelWidth = 99
  LabelHeight = 43
  
  HorGap = 2    'Horizontal Gap between two Labels
  VerGap = 5    'Vertical Gap between two Labels
  
  OffX = -5  'Horizontal Offset by Printer
  OffY = -5  'Vertical Offset by Printer

  Printer.ScaleMode = 6
  s = ""
  Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
  Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
  
  
'  For k = 1 To snap.RecordCount
'  For i = 1 To snap.RecordCount
'    'TopLeft
'    curX = IIf(k Mod 2 = 1, 4, LabelWidth + HorGap)
'    curY = IIf(k <= 2, 9, (LabelNo \ 2) * (LabelHeight + VerGap))
    
    For LabelNo = 1 To 12
      'TopLeft
      curX = IIf(LabelNo Mod 2 = 1, 4, LabelWidth + HorGap)
      curY = IIf(LabelNo <= 2, 9, (LabelNo \ 2) * (LabelHeight + VerGap))
    
      curX = 13 + LabelWidth - 3
      Printer.CurrentX = 13 + OffX
Printer.CurrentY = curY + OffY
Printer.Print "P"
      
      curX = 21 + LabelWidth - 3
      Printer.CurrentX = 21 + OffX
Printer.CurrentY = curY + OffY
Printer.Print "Customer"
      
      curX = 51 + LabelWidth - 3
      Printer.CurrentX = 51 + OffX
Printer.CurrentY = curY + OffY
Printer.Print ":"
      
      curX = 58 + LabelWidth - 3
      Printer.CurrentX = 58 + OffX
Printer.CurrentY = curY + OffY
Printer.Print snap("CustomerName")
      
      LabelNo = LabelNo + 1
      curX = 13 + LabelWidth - 3
      Printer.CurrentX = 13 + OffX
Printer.CurrentY = curY + OffY
Printer.Print "A"
'
'      curY = curY + LabelHeight - 3
'      Printer.CurrentX = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print "BR"
'
'      curX = curX - LabelWidth + 3
'      Printer.CurrentX = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print "BL"
      
    Next LabelNo
    
'    Printer.FontSize = 12
'    Printer.CurrentX = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(13); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "              :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(51); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(24); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(57); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(105); "G"
'
'    curY = curY + LabelHeight - 3
'    Printer.FontSize = 12
'    Printer.CurrentY = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(14); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "               :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(52); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(25); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(58); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(106); "G"
'
'    curY = curY + LabelHeight - 3
'    Printer.FontSize = 12
'    Printer.CurrentY = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    curY = curY + LabelHeight - 3
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(14); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "               :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(52); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(25); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(58); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(106); "G"
'
'    Printer.FontSize = 12
'    Printer.CurrentX = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    curX = curX - LabelWidth - 3
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(14); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "               :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(52); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(25); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(58); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(106); "G"
'
'    curY = curY + LabelHeight - 3
'    Printer.FontSize = 12
'    Printer.CurrentY = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(14); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "               :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(52); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(25); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(58); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(106); "G"
'
'    curY = curY + LabelHeight - 3
'    Printer.FontSize = 12
'    Printer.CurrentY = curX + OffX: Printer.CurrentY = curY + OffY: Printer.Print Space(15); rs("CompanyName") & LAlign(32); rs("CompanyName") & Chr(10) & Space(15) & rs("City") & "-" & rs("Pincode") & LAlign(34) & rs("City") & "-" & rs("Pincode")
'    Printer.FontSize = 10
'    Printer.Print "P" & Space(6) & "Customer Name"; "      :      " & snap("CustomerName") & LAlign(14); "P" & Space(6); "Customer Name"; "        :      " & snap("CustomerName")

'    Printer.Print "A" & Space(2) & "S" & Space(3) & "Description"; "               :      " & snap("ItemDescription") & LAlign(53); "A" & Space(2) & "S" & Space(3); "Description"; "                :      " & snap("ItemDescription")

'    Printer.Print "C" & Space(2) & "L" & Space(3) & "Part No"; "                    :      " & snap("PartNo") & LAlign(40); "C" & Space(2) & "L" & Space(3); "Part No"; "                     :      " & snap("PartNo")

'    Printer.Print "K" & Space(2) & "I" & Space(4) & "Quantity"; "                   :       " & snap("InvoiceQty") & LAlign(52); "K" & Space(2) & "I" & Space(4); "Quantity"; "                    :       " & snap("InvoiceQty")

'    Printer.Print "I" & Space(2) & " P" & Space(3) & "Invoice No"; "               :       " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate"); LAlign(25); "I" & Space(2) & " P" & Space(3); "Invoice No" & "                 :   " & snap("InvoiceNo") & "  " & "Date" & ":" & snap("InvoiceDate") & "             :       " & snap("InvoiceNo")

'    Printer.Print "N" & Space(6) & "No.of Pkts"; "                :       " & snap("NoofPackets") & LAlign(58); "N" & Space(6); "No.of Pkts"; "                  :       " & snap("NoofPackets")
'    Printer.Print "G" & LAlign(106); "G"
'
    
    
'  Next k
  
  Printer.EndDoc

      frmPrint.Show
'    End With
'  End If
  
  sPrintPack = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function sPrint2() As Boolean
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim cCol As Integer
  
  Dim curX As Integer
  Dim curY As Integer
  
  Dim LabelWidth As Integer
  Dim LabelHeight As Integer
  
  Dim LabelNo As Integer
  Dim OffX As Integer
  Dim OffY As Integer
  Dim HorGap As Integer
  Dim VerGap As Integer
  
  LabelWidth = 99
  LabelHeight = 43
  
  HorGap = 2    'Horizontal Gap between two Labels
  VerGap = 5    'Vertical Gap between two Labels
  
  OffX = -5  'Horizontal Offset by Printer
  OffY = -5  'Vertical Offset by Printer

  Printer.ScaleMode = 6
'  Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
'  Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
  
  If cmbReportType.ListIndex = 1 Then
  
  For LabelNo = 1 To 12
    'TopLeft
    curX = IIf(LabelNo Mod 2 = 1, 4, LabelWidth + HorGap)
    curY = IIf(LabelNo <= 2, 9, (LabelNo \ 2) * (LabelHeight + VerGap))
  
    Printer.CurrentX = curX + OffX
Printer.CurrentY = curY + OffY
Printer.Print "1"
    
    curX = curX + LabelWidth - 3
    Printer.CurrentX = curX + OffX
Printer.CurrentY = curY + OffY
Printer.Print "2"
    
    curY = curY + LabelHeight - 3
    Printer.CurrentX = curX + OffX
Printer.CurrentY = curY + OffY
Printer.Print "3"
    
    curX = curX - LabelWidth + 3
    Printer.CurrentX = curX + OffX
Printer.CurrentY = curY + OffY
Printer.Print "4"
    
  Next LabelNo
  
  Printer.EndDoc
'     sSendToPort LAlign(30, 1 & Space(1) & LAlign(50, 0 & Space(25) & LAlign(50, 2)))
'
'     sSendToPort LAlign(30, 2 & Space(1) & LAlign(50, Space(25) & LAlign(50, 5)))
'
'     sSendToPort LAlign(30, 3 & Space(1) & LAlign(50, Space(25) & LAlign(50, 6)))
'
'     sSendToPort LAlign(30, 8 & Space(1) & LAlign(50, Space(25) & LAlign(50, 7)))
  
      
   
  Else

          
'      .UnitType = UnitTypeTwips
      
'      .MaxCols = 2
'      .MaxRows = 2
'
'      .ColWidth(1) = 45: .ColWidth(2) = 45

'      .RowHeight(1) = 125: .RowHeight(2) = 125
'      frmPrint.Show
      
     ' Header

'     .Tag = Val(.Tag) + 1: GridCount = GridCount + 1: RowProcess

'     .Tag = Val(.Tag) + 1: GridCount = GridCount + 1: RowProcess

'     .Tag = Val(.Tag) + 1: GridCount = GridCount + 1: RowProcess
'
      Call spSetCellBorder(SprInvoice, 1, 1, SprInvoice.MaxCols, SprInvoice.MaxRows, mBorderOutline)
'       .AddCellSpan 1, Val(.Tag), 6, 6
      SprInvoice.SetText 1, 1, "1""000000000000000000000000000000000023467907654323456787654156565656565656565656565656565656"
      SprInvoice.SetText 1, 2, "1""000000000000000000000000000000000023467907654323455656565656565656"
      
      
      SprInvoice.SetText 1, 2, "2""35687698698696967567533217567808974683564757787683785678566888888688"
    
      Call spSetCellBorder(SprInvoice, 1, 1, SprInvoice.MaxCols, SprInvoice.MaxRows, mBorderRight + mBorderBottom)
      
      SprInvoice.SetText 2, 1, "3""3755555555555555555555555555555555555557799999999999999978889999"
      SprInvoice.SetText 2, 2, "4""6877672469706457689769099976756845648970000000000005677747445745"
      
      SprInvoice.Action = ActionPrint
     frmPrint.Show
  End If
  
  
  sPrint2 = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Public Sub RowProcess()
' On Error Resume Next
  
  If k > pheight Then
    Printer.NewPage
    If gReportID = "KUS" Then
      Dumrow 11
      k = 11
    Else
      Dumrow 5
      k = 5
    End If
    iPageNo = iPageNo + 1
    sPageHeader
  End If
End Sub

'
Private Function Validate() As Boolean
  Dim FromNo As Long        'Invoice No Lower Range
  Dim ToNo As Long          'Invoice No Upper Range
  
  Try
  
  Validate = False

  If chkInvoiceDate.Value = 1 Then
    If ValidateDateFromTo(dtpInvoiceFrom.Value, dtpInvoiceTo.Value, dtpInvoiceFrom.MinDate, dtpInvoiceTo.MaxDate) = False Then
      GoTo ExitHere
    End If
  ElseIf chkInvoiceNo.Value = 1 Then
    If txtInvoiceFrom.Text = "" Then
      MsgBox "Pls Type the InvoiceNo"
      GoTo ExitHere
    End If
    If ValidateNumberFromTo(txtInvoiceFrom, txtInvoiceTo, Val(txtInvoiceFrom.Tag), Val(txtInvoiceTo.Tag)) = False Then
      GoTo ExitHere
    End If
  Else
    MsgBox "Pls select atleast one"
    GoTo ExitHere
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
  End If
 
 Validate = True
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

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

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

Private Sub cmdClear_Click()

' On Error Resume Next
  
  Form_Clear Me
  chkPrintDate.Value = 1
  cmbReportType.ListIndex = 0
  cmbLabelNo.ListIndex = 0
  StepClear Me
  StepEnable Me
  
  fraStep(0).Enabled = True
  chkInvoiceDate.Value = 0
  dtpInvoiceFrom.Value = dtpInvoiceFrom.MinDate
  dtpInvoiceTo.Value = dtpInvoiceTo.MaxDate

  chkInvoiceNo.Value = 0
  txtInvoiceFrom.Text = ""
  txtInvoiceTo.Text = ""
  
  SprInvoice.MaxRows = 1
  SprInvoice.ClearRange 1, 1, -1, -1, True
  SprInvoice.Enabled = False
  CmdView.Enabled = False
  
'  MDISales.Toolbar1.Buttons(1).Enabled = True
'  MDISales.Toolbar1.Buttons(2).Enabled = True
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdprint_Click()
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
  
'
      Call spSetCellBorder(SprInvoice, 1, 1, SprInvoice.MaxCols, SprInvoice.MaxRows, mBorderOutline)
'       .AddCellSpan 1, Val(.Tag), 6, 6
      SprInvoice.SetText 1, 1, "1""000000000000000000000000000000000023467907654323456787654156565656565656565656565656565656"
      SprInvoice.SetText 1, 2, "1""000000000000000000000000000000000023467907654323455656565656565656"
      
      
      SprInvoice.SetText 1, 2, "2""35687698698696967567533217567808974683564757787683785678566888888688"
    
      Call spSetCellBorder(SprInvoice, 1, 1, SprInvoice.MaxCols, SprInvoice.MaxRows, mBorderRight + mBorderBottom)
      
      SprInvoice.SetText 2, 1, "3""3755555555555555555555555555555555555557799999999999999978889999"
      SprInvoice.SetText 2, 2, "4""6877672469706457689769099976756845648970000000000005677747445745"
      
      SprInvoice.Action = ActionPrint
  

  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

End Try
End Sub

Private Sub cmdProceed_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Try
  
  If Validate = False Then
    GoTo ExitHere
  End If
  
  Sql = "Select a.InvoiceType [Inv. Type], a.InvoiceNo [Inv. No], a.InvoiceDate [Inv. Date], b.PartyName [Customer], a.CustomerCode, y.OrderFlag "
  Sql = Sql & " from SAL_INVOICE_HDR a Left Outer Join FAS_PARTY_MASTER b On a.CustomerCode = b.PartyCode and a.CompId=b.Compid"
  Sql = Sql & " left Outer join Sal_PackingList y On a.InvoiceType=y.OrderType"
  Sql = Sql & " and a.InvoiceNo=y.OrderNo and a.InvoiceDate=y.OrderDate"
  Sql = Sql & " and a.Branchid=y.Branchid and a.Compid=y.Compid"
  Sql = Sql & " where a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
  
  If chkInvoiceNo.Value = 1 Then
    Sql = Sql & " and a.InvoiceNo between " & Val(Trim$(txtInvoiceFrom.Text)) & " and " & Val(Trim$(txtInvoiceTo.Text))
  End If
  
  If chkInvoiceDate.Value = 1 Then
    Sql = Sql & " and a.InvoiceDate between '" & Format(dtpInvoiceFrom, "yyyy-MM-dd") & "' and '" & Format(dtpInvoiceTo, "yyyy-MM-dd") & "'"
  End If
  
  Sql = Sql & " Order by a.InvoiceNo "

'  adoInvoice.ConnectionString = db.ConnectionString: adoInvoice.CommandType = adCmdText
'  adoInvoice.RecordSource = Sql
'  sprInvoice.MaxRows = 0
'  sprInvoice.DataSource = adoInvoice
'  adoInvoice.Refresh
'  sprInvoice.MaxRows = sprInvoice.DataRowCnt
'
'  sprInvoice.Enabled = True
'  cmdview.Enabled = True
'
'  If cmdProceed Then
'  Call sprInvoice_Click(1, 1)
'  End If
'  fraStep(0).Enabled = False
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  Else
    SprInvoice.Enabled = True
 SprInvoice.Visible = True
    SprInvoice.ClearSelection
    Call ADORS_Spread_DataFill(rs, SprInvoice, True)
    CmdView.Enabled = True
    SprInvoice.MaxRows = SprInvoice.DataRowCnt
    Call Spread_Set_Focus(SprInvoice, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Call StepProceedNext(Me, Index)
  
  If fraStep.UBound = Index Then
    cmdStep(Index).Enabled = False
    CmdView.Enabled = True
    cmbLabelNo.Enabled = True
  Else
    'Continue
  End If
End Sub

Private Sub cmdView_Click()
 
  Try
  Me.MousePointer = vbHourglass

'  If cmbReportType.ListIndex = 0 Then
''    If sPrint2 = False Then: GoTo ExitHere

'    If sPrintView = False Then: GoTo ExitHere
'  Else
'    Set frmCurrent = Me
'    DoEvents
'    iLineNo = 1
'    iPageNo = 1
'    If sPrintView = False Then: GoTo ExitHere

''    If sPrint2 = False Then: GoTo ExitHere
'    sClosePort
'    Set frmOutput = New frmReportViewer
'    frmOutput.Caption = gStrReportHeading
'    frmOutput.txtTotalPages.Text = CStr(iPageNo)
'    frmOutput.wbrView.Navigate strFilename
'    frmOutput.Tag = strFilename
'    gStrReportHeading = ""
'    frmOutput.Show vbModal
'    DoEvents
'    Screen.MousePointer = vbNormal
'  End If
  
  If sPrintPackingList = False Then
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
  
  Call FindFillMinMaxValue(anByNone)
  CmdView.Enabled = False
  cmbReportType.ListIndex = 0
  cmbLabelNo.ListIndex = 0
  SprInvoice.UnitType = UnitTypeTwips
  
'  If Orient = 1 Then
    SprInvoice.PrintMarginLeft = 0
    SprInvoice.PrintMarginTop = 0
'  End If
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

Private Sub sprInvoice_Click(sender as object, e as ClickEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  Spread_BkColour SprInvoice, vbWhite, 1, 1, 4, SprInvoice.DataRowCnt
  SprInvoice.Col = 4
SprInvoice.Row = SprInvoice.ActiveRow
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
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

Private Sub Design_fpSpread1(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpSpread1

End Sub

Private Sub Design_sprInvoice(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprInvoice


sprInvoice.Click +=  new EventHandler(sprInvoice_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_fpSpread1("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesPackagingSlipList.frm", fpSpread1, "")

	Call Design_sprInvoice("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesPackagingSlipList.frm", sprInvoice, "")

End Sub
