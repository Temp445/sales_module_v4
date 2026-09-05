VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptPurchaseValueReport 
   Caption         =   "Purchase Value Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6690
      TabIndex        =   15
      Top             =   3825
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5400
      TabIndex        =   16
      Top             =   3825
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4110
      TabIndex        =   17
      Top             =   3825
      Width           =   1300
   End
   Begin VB.Frame fratype 
      Height          =   945
      Index           =   3
      Left            =   4485
      TabIndex        =   11
      Top             =   2730
      Width           =   3135
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptPurchaseValueReport.frx":0000
         Left            =   1035
         List            =   "rptPurchaseValueReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   210
         Width           =   1995
      End
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   585
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   14
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.Frame FraStep 
      Height          =   765
      Index           =   0
      Left            =   4740
      TabIndex        =   5
      Top             =   1395
      Width           =   2625
      Begin MSComCtl2.DTPicker dtptodt 
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
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
         Format          =   20185089
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dtpfromdt 
         Height          =   315
         Left            =   60
         TabIndex        =   7
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
         Format          =   20185089
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1290
         TabIndex        =   9
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Select Item"
      Enabled         =   0   'False
      Height          =   525
      Index           =   1
      Left            =   4995
      TabIndex        =   2
      Top             =   2205
      Width           =   2115
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   195
         Left            =   1035
         TabIndex        =   4
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   570
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
      Left            =   7470
      TabIndex        =   1
      Top             =   1635
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
      Left            =   7245
      TabIndex        =   0
      Top             =   2295
      Width           =   345
   End
Begin AceSpread spVendor
      Height          =   1305
      Left            =   7695
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   4050
      _Version        =   458752
      _ExtentX        =   7144
      _ExtentY        =   2302
      _StockProps     =   64
      DAutoSizeCols   =   1
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
End
Attribute VB_Name = "rptPurchaseValueReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private sprow As Long
Private sRow(4) As String
Private sShe As Integer

Private Erpno As String
Private TsNo As String
Public Sub Header()
  
  Dim Spr As AceSpread
  
  Try
  
  Set Spr = frmprint.sprPrint
  
  Spr.Sheet = Spr.Sheet
  Call spSetColWidth(Spr, 1, 15, 2, 30, 3, 8, 4, 8, 5, 10)
  
  Call RptStandard_Header(5, sprow, Erpno, TsNo, IIf(chkPrintDate, True, False))
  GridCount = GridCount + 6
  
  Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
  Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Item Code", "Description", "Quantity", "Rate", "Goods Value")
  
  Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)

  sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow, Val(Spr.Sheet)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
  Dim i As Integer
  Dim sql As String
  Dim Spr As AceSpread
  Dim genrec As AceADODB.Recordset
  Dim TotQty(4) As String
  Dim ItemNature As String
  Dim Total As Double
  Dim ShtCount As Long
  
  Try
  sSpread_sPrintcurrent = False
  
  Sql = " Select distinct a.doctype,a.docno ,A.DocDate, b.Itemcode,(Select c.ItemDescription"
  Sql = Sql & " from INV_MATERIAL_MASTER c where c.ItemCode=b.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid"
  Sql = Sql & " and c.status = 1) ItemDescription, sum(b.Qty) as Qty, b.Rate, sum(InvoiceAmount ) as InvoiceAmount,"
  Sql = Sql & " (Select c.ItemNature from MATERIAL_MASTER c where c.ItemCode=b.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid"
  Sql = Sql & " and c.status = 1) Nature from PUR_PurchaseInvoice_Hdr a"
  Sql = Sql & " left outer join PUR_PurchaseInvoice_Dtl b"
  Sql = Sql & " on a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate and a.Branchid = b.Branchid and a.Compid = b.Compid"
  Sql = Sql & " AND A.Finyear = b.Finyear Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  Sql = Sql & " and a.DocDate BETWEEN '" & Format(dtpfromdt, "yyyy-MM-dd") & "' and '" & Format(dtptodt, "yyyy-MM-dd") & "'"
  If optRandom = True Then
      Sql = Sql & "  and b.Itemcode in "
      Sql = Sql & Spread_Get_SelectionList(spVendor, 1)  End If
  Sql = Sql & " group by a.doctype,a.docno,A.DocDate,b.Itemcode, b.rate,a.Branchid, a.Compid,InvoiceAmount"
  Sql = Sql & " ORDER BY  Nature,a.doctype,a.docno,B.ITEMCODE"
  
  If RSOpen(genrec, Sql) = False Then Exit Function
  If genrec.RecordCount = 0 Then
    sSpread_sPrintcurrent = False
    sShowMessage "No Record's Found"
    GoTo ExitHere
  End If
  
  Set Spr = frmprint.sprPrint
  sprow = 1
GridCount = 1
  Spr.SheetCount = 1
  ShtCount = 1
  i = 1
  
    sRow(0) = sprow
sRow(1) = sprow
sRow(2) = sprow
sRow(3) = sprow
sRow(4) = sprow
    
    Do While Not genrec.EOF
     ' If Trim$(CheckNull(genrec("Nature"))) = "A" Or Trim$(CheckNull(genrec("Nature"))) = "L" Then
      If ItemNature <> Trim$(genrec("Nature")) And sprow <> 1 Then
        Call spSetText(Spr, 1, sprow, "Total", mAlignRightCenter, True, 4)
        Call spSetText(Spr, 5, sprow, Format(Total, "0.00"), mAlignRightCenter, True, 4)
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        Total = 0
        
        ShtCount = Val(Spr.SheetCount) + 1
        Spr.SheetCount = ShtCount
        sprow = 1
GridCount = 1
      End If
    
      If Trim$(CheckNull(genrec("Nature"))) = "L" Then 'Or Trim$(CheckNull(genrec("Nature"))) = "L"
        Spr.sheet = ShtCount
        gStrReportHeading = "Purchase Value Report - Consumables - " & Format(dtpfromdt, "MMM-yy")
        
        If Val(sRow(0)) = 1 Then
          iPageNo = 1
pheight = 54
Orient = 2
          sShe = Spr.Sheet
          Spr.SheetName = "Consumables"
          Call Header
          sRow(0) = sprow
        End If
        Spr.AddCellSpan 1, sprow, 1, 1
        Call spSetTextColArray(Spr, 1, Val(sRow(0)), 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, Val(sRow(0)), 1, mAlignRightCenter, False, genrec("Qty"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, Val(sRow(0)), Spr.MaxCols, Val(sRow(0)), mBorderAllSides)
        sRow(0) = Val(sRow(0)) + 1
GridCount = Val(sRow(0))
sprow = Val(sRow(0))
RowProcess Me, sprow
        TotQty(0) = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty(0))
        sRow(0) = sprow
        
      ElseIf Trim$(CheckNull(genrec("Nature"))) = "I" Then
        Spr.sheet = ShtCount
        gStrReportHeading = "Purchase Value Report - Instruments - " & Format(dtpfromdt, "MMM-yy")
        
        If Val(sRow(1)) = 1 Then
          iPageNo = 1
pheight = 54
Orient = 2
          sprow = sRow(1)
          sShe = Spr.Sheet
          Spr.SheetName = "Instruments"
          Call Header
          sRow(1) = sprow
        End If
        Call spSetTextColArray(Spr, 1, Val(sRow(1)), 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, Val(sRow(1)), 1, mAlignRightCenter, False, genrec("Qty"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, Val(sRow(1)), Spr.MaxCols, Val(sRow(1)), mBorderAllSides)
        sRow(1) = Val(sRow(1)) + 1
GridCount = Val(sRow(1))
sprow = Val(sRow(1))
RowProcess Me, sprow
        sRow(1) = sprow
        TotQty(1) = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty(1))

      ElseIf Trim$(CheckNull(genrec("Nature"))) = "R" Then
        Spr.sheet = ShtCount
        gStrReportHeading = "Purchase Value Report - Raw Material - " & Format(dtpfromdt, "MMM-yy")
        
        If Val(sRow(2)) = 1 Then
          iPageNo = 1
pheight = 54
Orient = 2
          sprow = sRow(2)
          sShe = Spr.Sheet
          Spr.SheetName = "Raw Material"
          Call Header
          sRow(2) = sprow
        End If
        Call spSetTextColArray(Spr, 1, Val(sRow(2)), 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, Val(sRow(2)), 1, mAlignRightCenter, False, genrec("Qty"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, Val(sRow(2)), Spr.MaxCols, Val(sRow(2)), mBorderAllSides)
        sRow(2) = Val(sRow(2)) + 1
GridCount = Val(sRow(2))
sprow = Val(sRow(2))
RowProcess Me, sprow
        sRow(2) = sprow
        TotQty(2) = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty(2))

      ElseIf Trim$(CheckNull(genrec("Nature"))) = "T" Then
        Spr.sheet = ShtCount
        
        If Val(sRow(3)) = 1 Then
          iPageNo = 1
pheight = 54
Orient = 2
          sprow = sRow(3)
          sShe = Spr.Sheet
          Spr.SheetName = "Tools"
          Call Header
          sRow(3) = sprow
        End If
        Call spSetTextColArray(Spr, 1, Val(sRow(3)), 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, Val(sRow(3)), 1, mAlignRightCenter, False, genrec("Qty"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, Val(sRow(3)), Spr.MaxCols, Val(sRow(3)), mBorderAllSides)
        sRow(3) = Val(sRow(3)) + 1
GridCount = Val(sRow(3))
sprow = Val(sRow(3))
RowProcess Me, sprow
        sRow(3) = sprow
        TotQty(3) = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty(3))
        
      ElseIf Trim$(CheckNull(genrec("Nature"))) = "A" Then
        Spr.sheet = ShtCount
        gStrReportHeading = "Purchase Value Report - Spares - " & Format(dtpfromdt, "MMM-yy")

        If Val(sRow(4)) = 1 Then
          iPageNo = 1
pheight = 54
Orient = 2
          sShe = Spr.sheet
          Spr.SheetName = "Spares"
          Call Header
          sRow(4) = sprow
        End If
        Spr.AddCellSpan 1, sprow, 1, 1
        Call spSetTextColArray(Spr, 1, Val(sRow(4)), 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, Val(sRow(4)), 1, mAlignRightCenter, False, genrec("Qty"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, Val(sRow(4)), Spr.MaxCols, Val(sRow(4)), mBorderAllSides)
        sRow(4) = Val(sRow(4)) + 1
GridCount = Val(sRow(4))
sprow = Val(sRow(4))
RowProcess Me, sprow
        TotQty(0) = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty(0))
        sRow(4) = sprow

      End If
      
      ItemNature = Trim$(CheckNull(genrec("Nature")))
      Total = Format(Val(genrec("Qty") * genrec("Rate")), "0.00") + Val(Total)
      
      genrec.MoveNext
    Loop
    
    If genrec.EOF = True Then
      Call spSetText(Spr, 1, sprow, "Total", mAlignRightCenter, True, 4)
      Call spSetText(Spr, 5, sprow, Format(Total, "0.00"), mAlignRightCenter, True, 4)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      Total = 0
    End If
    
'    sSendToPort String(iPageWidth, "-")
'    sSendToPort Space(55) & RAlign(10, "Total") & Space(1) & RAlign(12, Format(TotQty, "0.00"))
  
    Spr.MaxRows = Spr.DataRowCnt + 2
    Spr.MaxCols = Spr.DataColCnt
    totalpages = iPageNo
    For i = 1 To Spr.SheetCount
      Spr.sheet = i
      Spr.ActiveSheet = i
      
      Spr.Row = 1
Spr.Row2 = Spr.MaxRows
Spr.Col = 1
Spr.Col2 = Spr.MaxCols
      Spr.BlockMode = True
      Spr.Lock = True
      Spr.BlockMode = False
      Spr.GridShowHoriz = False
Spr.GridShowVert = False

      Spr.ColHeadersShow = False
Spr.RowHeadersShow = False
      Spr.MaxRows = Spr.DataRowCnt + 2
    Next i
    
    Spr.ActiveSheet = 1
    
  frmprint.Show
  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
End Function

Public Sub sView()
  'No Change
End Sub

Public Sub sPrint()
  'No Change
End Sub

Public Sub sPageHeader()
  
  Try
  
  sSendToPort String(iPageWidth, "-")
  
  sSendToPort LAlign(12, "Item Code") & Space(1) & LAlign(30, "Description") & Space(1) & RAlign(10, "Quantity") & Space(1) & RAlign(10, "Rate") & Space(1) & RAlign(12, "Goods Value")
  
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Private Function SPrint_Current() As Boolean
  Dim genrec As AceADODB.Recordset
  Dim sql As String
  Dim sRow As Long
  Dim i As Long
  Dim TotQty As Double
  
  Try
  
  SPrint_Current = False
  
  Set genrec = New AceADODB.Recordset
  
  Sql = " Select distinct a.doctype,a.docno ,A.DocDate, b.Itemcode,(Select c.ItemDescription"
  Sql = Sql & " from INV_MATERIAL_MASTER c where c.ItemCode=b.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid"
  Sql = Sql & " and c.status = 1) ItemDescription, sum(b.Qty) as Qty, b.Rate, sum(InvoiceAmount ) as InvoiceAmount,"
  Sql = Sql & " (Select c.ItemNature from MATERIAL_MASTER c where c.ItemCode=b.ItemCode  and c.Branchid=a.Branchid and c.Compid=a.Compid"
  Sql = Sql & " and c.status = 1) Nature from PUR_PurchaseInvoice_Hdr a"
  Sql = Sql & " left outer join PUR_PurchaseInvoice_Dtl b"
  Sql = Sql & " on a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate and a.Branchid = b.Branchid and a.Compid = b.Compid"
  Sql = Sql & " AND A.Finyear = b.Finyear Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  Sql = Sql & " and a.DocDate BETWEEN '" & Format(dtpfromdt, "yyyy-MM-dd") & "' and '" & Format(dtptodt, "yyyy-MM-dd") & "'"
  Sql = Sql & " group by a.doctype,a.docno,A.DocDate,b.Itemcode, b.rate,a.Branchid, a.Compid,InvoiceAmount"
  Sql = Sql & " ORDER BY  Nature,a.doctype,a.docno,B.ITEMCODE"
  
  If RSOpen(genrec, Sql) = False Then Exit Function
  If genrec.RecordCount = 0 Then
    SPrint_Current = False
    sShowMessage "No Record's Found"
    GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  fOpenPort Me

  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  
  sSendToPort ""
  gStrReportHeading = "Purchase Value Report between " & Format(dtpfromdt, "dd-MM-yyyy") & " and " & Format(dtptodt, "dd-MM-yyyy")
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
  sPageHeader 'Printing the Header Details
  i = 1
  Do While Not genrec.EOF
    sSendToPort LAlign(12, genrec("ItemCode")) & Space(1) & LAlign(30, CheckNull(genrec("ItemDescription"))) & Space(1) & RAlign(10, genrec("Qty")) & Space(1) & RAlign(10, Format(genrec("Rate"), "0.00")) & Space(1) & RAlign(12, Format((genrec("Qty") * genrec("Rate")), "0.00"))
    i = i + 1
    TotQty = Val(genrec("Qty") * genrec("Rate")) + Val(TotQty)
    genrec.MoveNext
  Loop
  sSendToPort String(iPageWidth, "-")
  sSendToPort Space(55) & RAlign(10, "Total") & Space(1) & RAlign(12, Format(TotQty, "0.00"))
  genrec.Close
  Call EndofReport(iPageWidth, Val(i - 1))
  
  SPrint_Current = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
End Function

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdClear_Click()
' On Error Resume Next
  
  Form_Clear Me
  Call StepClear(Me)
  StepEnable Me
  OptAll.Value = True
  cmbType.ListIndex = 0
  chkPrintDate.Value = 1
  chkPrintDate.Value = 1
  dtpfromdt = Format(ServerDate, "dd/MM/yyyy")
  dtptodt = Format(ServerDate, "dd/MM/yyyy")
  spVendor.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If optRandom.Value = True Then
    If spVendor.SelectionCount = 0 Then
      sShowMessage "Select atleast one Customer"
      Exit Sub
    End If
  End If

  Call StepProceedNext(Me, Index)
  cmdview.Enabled = True
  
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass
'
'  ERPNo = Get_ERPandTSNo(Me.Tag, "PUR", mErp)
'  TSNo = Get_ERPandTSNo(Me.Tag, "PUR", mTs)
  If cmbType.ListIndex = 1 Then
    If SPrint_Current = False Then GoTo ExitHere
    
    Call FinaliseAndShowPrintPreview
  Else
    Erpno = Get_ERPandTSNo("mnuBaseMasterReport", "SAL", mErp)
    TsNo = Get_ERPandTSNo("mnuBaseMasterReport", "SAL", mTs)
    If sSpread_sPrintcurrent = False Then
 GoTo ExitHere
End If
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
  
  cmdview.Enabled = False
'  Call sView
  cmbType.ListIndex = 0
  dtpfromdt = Format(ServerDate, "dd/MM/yyyy")
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
    RptFlg = ""
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  
  spVendor.Visible = False
  
End Sub

Private Sub optRandom_Click()
  Dim rs As AceADODB.Recordset
  Dim i As Integer
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Select distinct a.Itemcode, b.ItemDescription from PUR_PurchaseInvoice_dtl a"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER b"
  Sql = Sql & " On b.ItemCode=a.ItemCode  and b.Branchid=a.Branchid and b.Compid=a.Compid and b.status = 1"
  Sql = Sql & " Where DocDate BETWEEN '" & Format(dtpfromdt, "yyyy-MM-dd") & "' and '" & Format(dtptodt, "yyyy-MM-dd") & "'"
  Sql = Sql & " Order By a.Itemcode"
  Set rs = New AceADODB.Recordset
  
  If RSOpen(rs, Sql) = False Then Exit Sub
  
  If rs.RecordCount = 0 Then
    sShowMessage "No Record's Found"
    spVendor.Visible = False
    GoTo ExitHere
  Else
    Call Spread_Clear(spVendor)
    spVendor.ClearSelection
    spVendor.MaxRows = rs.RecordCount
    For i = 1 To rs.RecordCount
      spVendor.SetText 1, i, rs("ItemCode")
      spVendor.SetText 2, i, rs("ItemDescription")
      rs.MoveNext
    Next i
    spVendor.Visible = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spVendor(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spVendor

End Sub

Private Sub InitializeSpreads()

	Call Design_spVendor("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptPurchaseValueReport.frm", spVendor, "")

End Sub
