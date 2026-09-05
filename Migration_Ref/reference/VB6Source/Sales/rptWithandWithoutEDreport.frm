VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptWithandWithoutEDreport 
   Caption         =   "With and Without ED Report"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4245
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5145
      TabIndex        =   15
      Top             =   4245
      Width           =   1305
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   405
      Left            =   3870
      TabIndex        =   14
      ToolTipText     =   "Click to OK"
      Top             =   4245
      Width           =   1290
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
      Index           =   0
      Left            =   7200
      TabIndex        =   12
      Top             =   2640
      Width           =   405
   End
   Begin VB.Frame frastep 
      Caption         =   "Report Type"
      Height          =   930
      Index           =   1
      Left            =   4470
      TabIndex        =   9
      Top             =   3225
      Width           =   2670
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "rptWithandWithoutEDreport.frx":0000
         Left            =   90
         List            =   "rptWithandWithoutEDreport.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2370
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print.Date"
         Height          =   255
         Left            =   90
         TabIndex        =   10
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
      Index           =   1
      Left            =   7200
      TabIndex        =   8
      Top             =   3540
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   4440
      TabIndex        =   5
      Top             =   6435
      Visible         =   0   'False
      Width           =   2730
      Begin VB.CommandButton cmdExit1 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to Exit from the Current Screen"
         Top             =   165
         Width           =   1300
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   405
         Left            =   75
         TabIndex        =   7
         ToolTipText     =   "Click to OK"
         Top             =   165
         Width           =   1290
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "With and Without ED"
      Height          =   855
      Index           =   0
      Left            =   4470
      TabIndex        =   0
      Top             =   2355
      Width           =   2670
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   1365
         TabIndex        =   1
         Top             =   435
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   39268
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   435
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20185089
         CurrentDate     =   39268
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
         Height          =   180
         Left            =   1365
         TabIndex        =   4
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   195
         Width           =   1140
      End
   End
End
Attribute VB_Name = "rptWithandWithoutEDreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Dim j As Long
Dim k As Integer
Private Sub RowProcess()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmprint.sprPrint.MaxRows = frmprint.sprPrint.MaxRows + pheight + 1
      frmprint.sprPrint.Row = frmprint.sprPrint.Row + 1
j = j + 1
      frmprint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Header
      j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Sub Header()
  Dim mstr As String

  Try
  
    frmprint.sprPrint.MaxCols = 7

    frmprint.sprPrint.AddCellSpan 1, j, 2, 1
    frmprint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    
    frmprint.sprPrint.AddCellSpan 6, j, frmprint.sprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmprint.sprPrint.SetText 6, j, "Printed Date: " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmprint.sprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
    Spread_ERPnoPrint frmprint.sprPrint, 1, j
    j = j + 1
GridCount = GridCount + 1

    If iPageNo = 1 Then
      frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1
      frmprint.sprPrint.AddCellSpan 1, j, frmprint.sprPrint.MaxCols, 1
      frmprint.sprPrint.SetText 1, j, gCompName
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.FontSize = 12
frmprint.sprPrint.FontUnderline = True
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmprint.sprPrint.AddCellSpan 1, j, frmprint.sprPrint.MaxCols, 1
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1

    frmprint.sprPrint.SetText 1, j, " Details Reg.With and Without ED Details "
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.FontSize = 10
frmprint.sprPrint.FontUnderline = True
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1

    frmprint.sprPrint.SetText 1, j, " WITH ED "
frmprint.sprPrint.TypeHAlign = TypeHAlignLeft
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.FontSize = 10

    j = j + 2
GridCount = GridCount + 2

    Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexLeft, 1, j, frmprint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexTop, 1, j, frmprint.sprPrint.MaxCols, j

    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1

    frmprint.sprPrint.SetText 1, j, "S.No"
frmprint.sprPrint.ColWidth(1) = 5
frmprint.sprPrint.TypeHAlign = TypeHAlignLeft
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
    
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 2

    frmprint.sprPrint.SetText 2, j, "Tariff No"
frmprint.sprPrint.ColWidth(2) = 15
frmprint.sprPrint.TypeHAlign = TypeHAlignLeft
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
  
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 3

    frmprint.sprPrint.SetText 3, j, "Quantity"
frmprint.sprPrint.ColWidth(3) = 10
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
    
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 4

    frmprint.sprPrint.SetText 4, j, "Goods"
frmprint.sprPrint.ColWidth(4) = 10
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True

    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 5

    frmprint.sprPrint.SetText 5, j, "Exicise"
frmprint.sprPrint.ColWidth(5) = 10
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True

    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 6

    frmprint.sprPrint.SetText 6, j, "Cess"
frmprint.sprPrint.ColWidth(6) = 10
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True

    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 7

    frmprint.sprPrint.SetText 7, j, "Sur Cess"
frmprint.sprPrint.ColWidth(7) = 10
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True

    Spread_BkColour frmprint.sprPrint, &H8000000F, 1, j, frmprint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexLeft, 1, j, frmprint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexBottom, 1, j, frmprint.sprPrint.MaxCols, j

    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 4

    frmprint.sprPrint.SetText 4, j, "Value"
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
    
    frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 5

    frmprint.sprPrint.SetText 5, j, "Duty"
frmprint.sprPrint.TypeHAlign = TypeHAlignRight
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
    
    Spread_BkColour frmprint.sprPrint, &H8000000F, 1, j, frmprint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

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

Private Sub CmdClear_Click()
' On Error Resume Next
  
  Call StepClear(Me)
  MDISales.Toolbar1.Buttons(1).Enabled = True
  MDISales.Toolbar1.Buttons(2).Enabled = True
  Call StepEnable(Me)
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdExit1_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Frame1.Visible = False
  dtpFrom.Value = Format(ServerDate, "dd/MM/yyyy")
  dtpTo.Value = Format(ServerDate, "dd/MM/yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  cmbReportType.ListIndex = 0
End Sub

Private Sub CmdView_Click()
  Dim snap As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim Sql As String
  Dim i As Integer
  Dim tmp As Variant
  
  Try
  
  If cmdView = True Then
    If ValidateDateFromTo(dtpFrom.Value, dtpTo.Value, dtpFrom.MinDate, dtpTo.MaxDate) = False Then
      Call cmdExit1_Click
      GoTo ExitHere
    End If
  End If
  
'With ED
  Sql = " Select x.TariffNo, sum(invoiceqty) Qty, sum(gv) GV, sum(ED) ED, sum(cess) cess,sum(scess)SHE"
  Sql = Sql & " from (select a.invoiceNo, b.tariffno, a.invoiceqty, (a.invoiceqty * a.Rate) GV,"
  Sql = Sql & " isNull((select c.ResultValue from salesCalculationDetails c"
  Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
  Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql & " and c.ExprName in ('CENVAT CLAIMS','EXCISE DUTY  8%')),0) ED, isNull((select c.ResultValue from salesCalculationDetails c"
  Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
  Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql & " and c.ExprName = 'CESS 2%'),0) Cess,isNull((select c.ResultValue from salesCalculationDetails c"
  Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
  Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql & " and c.ExprName = 'S&H EDU.CESS 1%'),0) sCess from SalesInvoice a left outer join"
  Sql = Sql & " inv_material_master b on a.ItemCode = b.ItemCode and a.branchid = b.branchid"
  Sql = Sql & " and a.compid = b.compid where a.branchid ='" & gBranchID & "' and a.compid = '" & gCompID & "'"
  Sql = Sql & " and a.finyear = '" & gFinyear & "' and a.CustomerCode <> 28"
  Sql = Sql & " and a.invoiceDate between '" & Format(dtpFrom, "yyyy-MM-dd") & "' and '" & Format(dtpTo, "yyyy-MM-dd") & "' ) x where ed > 0 group by x.TariffNo"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If snap.RecordCount = 0 Then
    MsgBox "No records exist"
    GoTo ExitHere
  End If
  
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 75
Orient = 1
    
    Header
    
    j = j + 1
GridCount = GridCount + 1
  
    If Not snap.EOF Then
      For i = 1 To snap.RecordCount
        Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexOutline, 1, j, frmprint.sprPrint.MaxCols, j

        frmprint.sprPrint.SetText 1, j, Val(i) & "."
        frmprint.sprPrint.SetText 2, j, snap("Tariffno")
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 3, j, snap("Qty")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 4, j, snap("Gv")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 5, j, snap("Ed")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 6, j, snap("Cess")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 7, j, snap("She")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        snap.MoveNext
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Next i
      j = j + 1
GridCount = GridCount + 1
RowProcess
    End If
    
'Without ED
    Sql = " select y.TariffNo, sum(invoiceqty) Qty, sum(gv) GV, sum(ED) ED, sum(cess) cess,sum(scess)SHE from ("
    Sql = Sql & " select * from ( select a.invoiceNo, b.tariffno, a.invoiceqty, (a.invoiceqty * a.Rate) GV,"
    Sql = Sql & " isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName in ('CENVAT CLAIMS','EXCISE DUTY')),0) ED, isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName = 'CESS'),0) Cess,isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName = 'SHE CESS 1%'),0) sCess from SalesInvoice a left outer join inv_material_master b"
    Sql = Sql & " on a.ItemCode = b.ItemCode and a.branchid = b.branchid and a.compid = b.compid"
    Sql = Sql & " where a.branchid ='" & gBranchID & "' and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
    Sql = Sql & " and a.CustomerCode <> 28 and a.invoiceDate between '" & Format(dtpFrom, "yyyy-MM-dd") & "' and '" & Format(dtpTo, "yyyy-MM-dd") & "') y  where ed = 0 ) y"
    Sql = Sql & " group by y.tariffno order by y.tariffno, y.ED"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
    If Not snap.EOF Then
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmprint.sprPrint.AddCellSpan 1, j, 2, 1
      frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1

      frmprint.sprPrint.SetText 1, j, " WITHOUT ED "
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.FontSize = 10

      j = j + 2
GridCount = GridCount + 2
RowProcess
      Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexOutline, 1, j, frmprint.sprPrint.MaxCols, j
      
      For i = 1 To 7
        frmprint.sprPrint.GetText i, 9, tmp
        frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = i

        frmprint.sprPrint.SetText i, j, tmp
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
      Next i
      j = j + 1
      For i = 1 To snap.RecordCount
        Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexOutline, 1, j, frmprint.sprPrint.MaxCols, j
'        .SetCellBorder 1, j + 1, .MaxCols, j + 1, CellBorderIndexTop, 0, CellBorderStyleSolid
        
        frmprint.sprPrint.SetText 1, j, Val(i) & "."
        frmprint.sprPrint.SetText 2, j, snap("Tariffno")
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 3, j, snap("Qty")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 4, j, snap("Gv")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 5, j, snap("Ed")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 6, j, snap("Cess")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 7, j, snap("She")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        snap.MoveNext
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Next i
        j = j + 1
GridCount = GridCount + 1
RowProcess
    End If
    
'EOU
    Sql = " select y.TariffNo, y.CustomerCode, sum(invoiceqty) Qty, sum(gv) GV, sum(ED) ED, sum(cess) cess,sum(scess)SHE from ("
    Sql = Sql & " select * from ( select a.invoiceNo, b.tariffno, m.CustomerCode, a.invoiceqty, (a.invoiceqty * a.Rate) GV,"
    Sql = Sql & " isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.InvoiceType = c.DocType And A.InvoiceNo = c.DocNo And A.InvoiceDate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName in ('CENVAT CLAIMS','EXCISE DUTY')),0) ED, isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.InvoiceType = c.DocType And A.InvoiceNo = c.DocNo And A.InvoiceDate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName = 'CESS'),0) Cess,isNull((select c.ResultValue from salesCalculationDetails c"
    Sql = Sql & " Where A.InvoiceType = c.DocType And A.InvoiceNo = c.DocNo And A.InvoiceDate = c.DocDate"
    Sql = Sql & " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
    Sql = Sql & " and c.ExprName = 'SHE CESS 1%'),0) sCess from sal_invoice_dtl a Left Outer Join sal_Invoice_Hdr m"
    Sql = Sql & " On m.InvoiceType = a.InvoiceType and m.InvoiceNo = a.InvoiceNo and m.InvoiceDate = a.InvoiceDate"
    Sql = Sql & " and m.branchid = a.branchid and m.compid = a.compid and m.finyear = a.finyear"
    Sql = Sql & " left outer join inv_material_master b on a.ItemCode = b.ItemCode and a.branchid = b.branchid and a.compid = b.compid"
    Sql = Sql & " where a.branchid ='" & gBranchID & "' and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
    Sql = Sql & " and a.invoiceDate between '" & Format(dtpFrom, "yyyy-MM-dd") & "' and '" & Format(dtpTo, "yyyy-MM-dd") & "' and m.CustomerCode = 28"
    Sql = Sql & " ) x  where ed = 0 ) y group by y.tariffno, y.CustomerCode order by y.tariffno, y.ED"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
    If Not snap.EOF Then
      j = j + 1
GridCount = GridCount + 1
RowProcess
      frmprint.sprPrint.AddCellSpan 1, j, 2, 1
      frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = 1

      frmprint.sprPrint.SetText 1, j, " EOU "
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.FontSize = 10

      j = j + 2
GridCount = GridCount + 2
RowProcess
      
      Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexOutline, 1, j, frmprint.sprPrint.MaxCols, j
      
      For i = 1 To 7
        frmprint.sprPrint.GetText i, 9, tmp
        frmprint.sprPrint.Row = j
frmprint.sprPrint.Col = i

        frmprint.sprPrint.SetText i, j, tmp
frmprint.sprPrint.Font = "Arial"
frmprint.sprPrint.FontBold = True
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
      Next i
      j = j + 1
GridCount = GridCount + 1
RowProcess
      For i = 1 To snap.RecordCount
         Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexOutline, 1, j, frmprint.sprPrint.MaxCols, j
        
        frmprint.sprPrint.SetText 1, j, Val(i) & "."
        frmprint.sprPrint.SetText 2, j, snap("Tariffno")
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 3, j, snap("Qty")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 4, j, snap("Gv")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 5, j, snap("Ed")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 6, j, snap("Cess")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        frmprint.sprPrint.Col = 1
frmprint.sprPrint.Row = j

        frmprint.sprPrint.SetText 7, j, snap("She")
frmprint.sprPrint.TypeVAlign = TypeVAlignCenter
frmprint.sprPrint.TypeHAlign = TypeHAlignCenter
        snap.MoveNext
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Next i
    End If
    frmprint.sprPrint.MaxRows = frmprint.sprPrint.DataRowCnt + 1
  
  totalpages = iPageNo
  frmprint.Show
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs1 = Nothing
  Set snap = Nothing
End Try
End Sub

Private Sub Form_Load()
  
  Try
  
  dtpFrom.Value = Format(ServerDate, "dd/MM/yyyy")
  dtpTo.Value = Format(ServerDate, "dd/MM/yyyy")
  CmdStep(0).Enabled = True
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

