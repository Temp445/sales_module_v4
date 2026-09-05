VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptOrderDocument 
   Caption         =   "Sales Order Document"
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
   Begin VB.CommandButton cmdStep 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   8100
      TabIndex        =   13
      Top             =   2565
      Width           =   345
   End
   Begin VB.Frame frastep 
      Height          =   780
      Index           =   1
      Left            =   3975
      TabIndex        =   21
      Top             =   2310
      Width           =   4095
      Begin VB.ComboBox cmbRevNo 
         Height          =   330
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   345
         Width           =   840
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRevNo 
         AutoSize        =   -1  'True
         Caption         =   "Rev. No."
         Height          =   210
         Left            =   3180
         TabIndex        =   25
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         Height          =   210
         Left            =   2040
         TabIndex        =   24
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblEnqNo 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
         Height          =   210
         Left            =   990
         TabIndex        =   23
         Top             =   180
         Width           =   660
      End
      Begin VB.Label lblEnqType 
         AutoSize        =   -1  'True
         Caption         =   "Order Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   22
         Top             =   180
         Width           =   825
      End
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
      Left            =   6975
      TabIndex        =   4
      Top             =   1800
      Width           =   345
   End
   Begin VB.Frame frastep 
      Caption         =   "Status"
      Height          =   675
      Index           =   0
      Left            =   4830
      TabIndex        =   31
      Top             =   1620
      Width           =   2055
      Begin VB.ComboBox cmbAuthorise 
         Height          =   330
         ItemData        =   "rptOrderDocument.frx":0000
         Left            =   120
         List            =   "rptOrderDocument.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6480
      TabIndex        =   18
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5190
      TabIndex        =   17
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   3900
      TabIndex        =   16
      Top             =   4095
      Width           =   1300
   End
   Begin VB.Frame fraReportType 
      Height          =   945
      Left            =   4260
      TabIndex        =   19
      Top             =   3090
      Width           =   3135
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   585
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptOrderDocument.frx":0058
         Left            =   1035
         List            =   "rptOrderDocument.frx":0062
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   255
         Width           =   885
      End
   End
Begin AceSpread fpsList
      Height          =   1470
      Left            =   7860
      TabIndex        =   7
      Top             =   2580
      Visible         =   0   'False
      Width           =   3540
      _Version        =   458752
      _ExtentX        =   6244
      _ExtentY        =   2593
      _StockProps     =   64
      DAutoSizeCols   =   1
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
      OperationMode   =   4
      RowHeaderDisplay=   0
      ScrollBars      =   2
   End
   Begin VB.Frame FraSteps 
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
      Height          =   1320
      Left            =   4080
      TabIndex        =   26
      Top             =   1860
      Width           =   3870
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
         Index           =   2
         Left            =   3420
         TabIndex        =   2
         Top             =   135
         Width           =   345
      End
      Begin VB.Frame frastep 
         Caption         =   "Date"
         Height          =   630
         Index           =   2
         Left            =   105
         TabIndex        =   27
         Top             =   -30
         Width           =   3270
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   315
            Left            =   1995
            TabIndex        =   1
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109314049
            CurrentDate     =   37273
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   315
            Left            =   510
            TabIndex        =   0
            Top             =   210
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109314049
            CurrentDate     =   37273
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   210
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   210
            Left            =   1785
            TabIndex        =   28
            Top             =   255
            Width           =   180
         End
      End
      Begin VB.CommandButton cmdStep 
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
         Left            =   3420
         TabIndex        =   8
         Top             =   780
         Width           =   345
      End
      Begin VB.Frame frastep 
         Caption         =   "Customer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   105
         TabIndex        =   30
         Top             =   630
         Width           =   3270
         Begin VB.OptionButton optCustAll 
            Caption         =   "All"
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
            Left            =   390
            TabIndex        =   5
            Top             =   255
            Value           =   -1  'True
            Width           =   570
         End
         Begin VB.OptionButton optCustSelective 
            Caption         =   "Selective"
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
            Left            =   1980
            TabIndex        =   6
            Top             =   255
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "rptOrderDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim mRevNo As String
Dim mDate As String
Dim mFromDate As Date
Dim mToDate As Date
Dim mRevPrev As String
Dim mRevCurr As String
Dim j As Long
Dim k As Long

Private Function sPrintcurent_JobworkView() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim s As Integer
  Dim GenSpec As String
  Dim PartNo As String
  Dim CurDate As Date
  
  Try
  
  sPrintcurent_JobworkView = False
  
  CurDate = ServerDate
  Set snap = New AceADODB.Recordset
  Set rs = New AceADODB.Recordset
  Set rs1 = New AceADODB.Recordset
  Set rsDetails = New AceADODB.Recordset
  
  If cmbType.ListIndex = 1 Then
    Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, c.PartyName,"
    Sql = Sql & " a.ReferenceNo, a.ReferenceDate, a.ItemGroupKey,a.ReceivedItemCode,a.SentItemCode,"
    Sql = Sql & " d.ItemDescription as RecDesc, d1.ItemDescription as SentDesc, d2.PartNo, a.RevisionNo, b.sDesc as Recuom, b.sDesc as Sentuom, a.OrderQty,"
    Sql = Sql & " isNull(d.ItemSpecification,'') ItemSpecification,a.ItemGroupKey,PackingDesc,InsuranceDesc,DeliveryDesc,PaymentDesc,FreightDesc,ModeDesc,FormDesc,FormulaName,"
    Sql = Sql & " PostingAcName from dbo.fn_SAL_JobworkOrder ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
    Sql = Sql & " left outer join PartMast d2 on a.SentItemCode = d2.ItemCode"
    Sql = Sql & " and a.Branchid = d2.Branchid and a.Compid = d2.Compid left outer join INV_Material_Master d"
    Sql = Sql & " on a.ReceivedItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
    Sql = Sql & " left outer join INV_Material_Master d1 on a.SentItemCode = d1.ItemCode and a.Branchid = d1.Branchid and a.Compid = d1.Compid"
    Sql = Sql & " left outer join Inv_UOM_Master b on d.UOM = b.Code and d.Branchid = b.Branchid"
    Sql = Sql & " and d.Compid = b.Compid left outer join Inv_UOM_Master b1 on d1.UOM = b1.Code and d1.Branchid = b1.Branchid"
    Sql = Sql & " and d1.Compid = b1.Compid left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
    Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid"
    Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
    Sql = Sql & " On f.OrderType = a.OrderType and f.OrderNo = a.OrderNo and f.OrderDate = a.OrderDate"
    Sql = Sql & " and f.BranchID = a.BranchID and f.CompID = a.CompID"
    Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
    Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If snap.EOF Then
      sShowMessage "No Record(s) Found"
      sPrintcurent_JobworkView = False
      GoTo ExitHere
    End If
    
    pheight = 71
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "Sales Order")
    sSendToPort ""
    sSendToPort LAlign(17, "SO No.          :") & Space(1) & LAlign(20, snap("OrderType") & snap("OrderNo")) & Space(12) & LAlign(18, "SO Date          : ") & Space(1) & LAlign(10, snap("OrderDate"))
    sSendToPort LAlign(17, "Customer PO No. :") & Space(1) & LAlign(20, snap("ReferenceNo")) & Space(12) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, snap("ReferenceDate"))
    sSendToPort LAlign(17, "Revision Level  :") & Space(1) & LAlign(20, snap("RevisionLevel"))
    sSendToPort LAlign(17, "Customer Name   :") & Space(1) & LAlign(50, snap("PartyName"))
    sPageHeader
    
    i = 1
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      GenSpec = Trim$(snap("GeneralDetails"))
      For i = 1 To snap.RecordCount
        'ProcessFlow
        If Trim$(gReportID) = "PMI" Then
          Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
          Sql = Sql + " left outer join SAL_Order_Dtl b "
          Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
          Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
          If RSOpen(rs1, Sql) = False Then Exit Function
          If rs1.EOF Then
            Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
          End If
          'RawMaterials
          Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
          Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
          Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
          If RSOpen(rsDetails, Sql) = False Then Exit Function
        End If
                
        sSendToPort LAlign(20, snap("PartNo")) & Space(1) & LAlign(40, snap("ItemDescription")) & Space(1) & LAlign(3, snap("RevisionNo")) & Space(1) & LAlign(3, CheckNull(snap("sDesc"))) & Space(1) & RAlign(10, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty")))
        If Trim$(snap("ItemCode")) <> "" Then
 sSendToPort LAlign(20, snap("ItemCode"))
End If
'        If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'          WordSplit Trim$(snap("ItemSpecification")), 40
'          sSendToPort Space(21) & LAlign(40, WSplit(1))
'          For s = 2 To UBound(WSplit)
'            sSendToPort Space(21) & LAlign(40, WSplit(s))
'          Next s
'        End If
        sSendToPort ""
        snap.MoveNext
      Next i
    End If
    'Delivery Details
    Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
    Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
    If RSOpen(rs, Sql) = False Then Exit Function
    If Not rs.EOF Then
      PartNo = ""
      sSendToPort String(iPageWidth, "-")
      sSendToPort LAlign(17, "Delivery Details :") & Space(1) & LAlign(30, "Part No") & Space(1) & LAlign(10, "Sch. Date") & Space(1) & RAlign(10, "Sch.Qty")
      rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        If PartNo <> rs("PartNo") Then
          If i > 1 Then
 sSendToPort ""
End If
          sSendToPort Space(18) & LAlign(30, rs("PartNo")) & Space(1) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        Else
          sSendToPort Space(49) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        End If
        PartNo = rs("PartNo")
        rs.MoveNext
      Next i
    End If
    sSendToPort String(iPageWidth, "-")
    sSendToPort ""
                         'Process Flow Print
    If Trim$(gReportID) = "PMI" Then
      If Not rs1.EOF Then
        sSendToPort LAlign(20, "Operation(s):")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(5, ("S.No")) & Space(2) & LAlign(50, "Description")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rs1.RecordCount
          sSendToPort LAlign(5, rs1("ProcessSequence")) & Space(2) & LAlign(50, rs1("ProcessFlow"))
          rs1.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
        sSendToPort ""
      End If
                        'Raw Material Print
      If Not rsDetails.EOF Then
        sSendToPort LAlign(25, "Rawmaterial Details:")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(20, "Rawmaterial Code") & Space(1) & LAlign(40, "Description") & Space(1) & LAlign(5, "UOM") & Space(1) & RAlign(12, "Norms")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rsDetails.RecordCount
          sSendToPort LAlign(20, rsDetails("ChildCode")) & Space(1) & LAlign(40, rsDetails("ItemDescription")) & Space(1) & LAlign(5, rsDetails("PurchaseUOMDesc")) & Space(1) & RAlign(12, rsDetails("Qty"))
          rsDetails.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
      End If
    rs1.Close
    rsDetails.Close
    End If
    
    If Trim$(GenSpec) <> "" Then
      WordSplit Trim$(GenSpec), iPageWidth
      sSendToPort "Note :"
      sSendToPort LAlign(iPageWidth, WSplit(1))
      For s = 2 To UBound(WSplit)
        sSendToPort LAlign(iPageWidth, WSplit(s))
      Next s
      sSendToPort String(iPageWidth, "-")
    End If
                         'Order Issued
    Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
    Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
    Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
    Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      For i = 1 To snap.RecordCount
        If i = 1 Then
          sSendToPort LAlign(16, "Order Issued To:") & Space(1) & LAlign(50, snap("Description"))
        Else
          sSendToPort Space(17) & LAlign(50, snap("Description"))
        End If
        snap.MoveNext
      Next i
    End If
  
  Else

      frmPrint.Show
      
      Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, c.PartyName,"
      Sql = Sql & " a.ReferenceNo, a.ReferenceDate, a.ItemGroupKey,a.ReceivedItemCode,a.SentItemCode,"
      Sql = Sql & " d.ItemDescription as RecDesc, d1.ItemDescription as SentDesc, d2.PartNo, a.RevisionNo, b.sDesc as Recuom, b.sDesc as Sentuom, a.OrderQty,"
      Sql = Sql & " isNull(d.ItemSpecification,'') ItemSpecification,a.ItemGroupKey,PackingDesc,InsuranceDesc,DeliveryDesc,PaymentDesc,FreightDesc,ModeDesc,FormDesc,FormulaName,"
      Sql = Sql & " PostingAcName from dbo.fn_SAL_JobworkOrder ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
      Sql = Sql & " left outer join PartMast d2 on a.SentItemCode = d2.ItemCode"
      Sql = Sql & " and a.Branchid = d2.Branchid and a.Compid = d2.Compid left outer join INV_Material_Master d"
      Sql = Sql & " on a.ReceivedItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
      Sql = Sql & " left outer join INV_Material_Master d1 on a.SentItemCode = d1.ItemCode and a.Branchid = d1.Branchid and a.Compid = d1.Compid"
      Sql = Sql & " left outer join Inv_UOM_Master b on d.UOM = b.Code and d.Branchid = b.Branchid"
      Sql = Sql & " and d.Compid = b.Compid left outer join Inv_UOM_Master b1 on d1.UOM = b1.Code and d1.Branchid = b1.Branchid"
      Sql = Sql & " and d1.Compid = b1.Compid left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
      Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid"
      Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
      Sql = Sql & " On f.OrderType = a.OrderType and f.OrderNo = a.OrderNo and f.OrderDate = a.OrderDate"
      Sql = Sql & " and f.BranchID = a.BranchID and f.CompID = a.CompID"
      Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
      Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If snap.EOF Then
        sShowMessage "No Record(s) Found"
        sPrintcurent_JobworkView = False
        GoTo ExitHere
      End If
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
  
      frmPrint.sprPrint.Col = 5
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
        If RptFlg <> "View" And RptFlg <> "Print" And RptFlg <> "Obsolute" Then
          frmPrint.sprPrint.AddCellSpan 1, j, 4, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
      End If
      
      If iPageNo = 1 Then
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

        j = j + 2
GridCount = GridCount + 2
      End If
          
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText 1, j, "Job Work Order"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

      j = j + 2
GridCount = GridCount + 2
      
      frmPrint.sprPrint.MaxCols = 7
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(1) = 8

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(2) = 25

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(3) = 25

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(4) = 15

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(5) = 5

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.ColWidth(6) = 8

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.ColWidth(7) = 10
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "JO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderType") & snap("OrderNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "JO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer Name"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PartyName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Packing & Forwarding"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PackingDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Insurance"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("InsuranceDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DeliveryDesc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("DeliveryDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PaymentDetails"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PaymentDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Freight"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FreightDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Mode"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ModeDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Form"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "FormulaName"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormulaName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PostingAc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PostingAcName")
      
      j = j + 2
GridCount = GridCount + 2
RowProcess
                                                
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd. Item Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sent Item code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rev."
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd.Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "GroupKey"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd. Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sent Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sen.Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess

      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
'        GenSpec = Trim$(snap("GeneralDetails"))
        For i = 1 To snap.RecordCount
          'ProcessFlow
          If Trim$(gReportID) = "PMI" Then
            Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
            Sql = Sql + " left outer join SAL_JobworkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b "
            Sql = Sql + " on a.ItemCode= b.RecdItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
            If rs1.EOF Then
              Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
              Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
              If RSOpen(rs1, Sql) = False Then Exit Function
            End If
            'RawMaterials
            Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
            Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
            Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
            If RSOpen(rsDetails, Sql) = False Then Exit Function
          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemGroupKey")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ReceivedItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("SentItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RevisionNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("RecUom"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          j = j + 1
GridCount = GridCount + 1
RowProcess
          
          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RecDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("SentDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("SentUom"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

'          j = j + 1: GridCount = GridCount + 1: RowProcess
'          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
            
'          If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'            WordSplit Trim$(snap("ItemSpecification")), 40
'            .Col = 1: .Row = j

'            .SetText .Col, .Row, WSplit(1): .TypeHAlign = TypeHAlignLeft

'            j = j + 1: GridCount = GridCount + 1: RowProcess
'
'            For s = 2 To UBound(WSplit)
'              .Col = 1: .Row = j

'              .SetText .Col, .Row, WSplit(s): .TypeHAlign = TypeHAlignLeft

'              j = j + 1: GridCount = GridCount + 1: RowProcess
'            Next s
'          End If
          j = j + 1
GridCount = GridCount + 1
RowProcess
          snap.MoveNext
        Next i
      End If
      'Delivery Details
      Sql = "select distinct a.SentItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from Sal_JobWorkOrder_Schedule a"
      Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
      Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
      If RSOpen(rs, Sql) = False Then Exit Function
      If Not rs.EOF Then
        PartNo = ""
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Delivery Details :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        j = j + 1
GridCount = GridCount + 1
RowProcess

        rs.MoveLast
rs.MoveFirst
        For i = 1 To rs.RecordCount
          If PartNo <> rs("PartNo") Then
            If i > 1 Then
 j = j + 1
GridCount = GridCount + 1
RowProcess
End If
            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          PartNo = rs("PartNo")
          rs.MoveNext
        Next i
      End If
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
                           'Process Flow Print
      If Trim$(gReportID) = "PMI" Then
        If Not rs1.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Operation(s):"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "S.No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        
          For i = 1 To rs1.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessSequence")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessFlow")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rs1.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
                          'Raw Material Print
        If Not rsDetails.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Details:"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Norms"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          For i = 1 To rsDetails.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ChildCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("PurchaseUOMDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rsDetails.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
      rs1.Close
      rsDetails.Close
      End If
      
      If Trim$(GenSpec) <> "" Then
        WordSplit Trim$(GenSpec), iPageWidth
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Note :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(1)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        For s = 2 To UBound(WSplit)
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(s)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        Next s
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
                           'Order Issued
      Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
      Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
      Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
      Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        For i = 1 To snap.RecordCount
          If i = 1 Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Issued To :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          snap.MoveNext
        Next i
      End If
      j = j + 3
GridCount = GridCount + 3
RowProcess
      frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 5, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "System Administrator"
  frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      j = j + 3
GridCount = GridCount + 3
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
  End If
  
  sPrintcurent_JobworkView = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set rs = Nothing
  Set rs1 = Nothing
  Set rsDetails = Nothing
End Try
End Function

Private Function sPrintcurent_JobworkObsolete() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim s As Integer
  Dim GenSpec As String
  Dim PartNo As String
  Dim CurDate As Date
  
  Try
  
  sPrintcurent_JobworkObsolete = False
  
  CurDate = ServerDate
  Set snap = New AceADODB.Recordset
  Set rs = New AceADODB.Recordset
  Set rs1 = New AceADODB.Recordset
  Set rsDetails = New AceADODB.Recordset
  
  If cmbType.ListIndex = 1 Then

    Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a1.CustomerCode, c.PartyName,"
    Sql = Sql & " a1.ReferenceNo, a1.ReferenceDate, a.ItemGroupKey,a.ReceivedItemCode,a.SentItemCode,"
    Sql = Sql & " d.ItemDescription as RecDesc, d1.ItemDescription as SentDesc, d2.PartNo, a.RevisionNo, b.sDesc as Recuom, b.sDesc as Sentuom, a.OrderQty,"
    Sql = Sql & " isNull(d.ItemSpecification,'') ItemSpecification,a.ItemGroupKey,PackingDesc,InsuranceDesc,DeliveryDesc,PaymentDesc,FreightDesc,ModeDesc,FormDesc,FormulaName,"
    Sql = Sql & " PostingAcName from SAL_JobworkOrder_Dtl a"
    Sql = Sql & " left outer join SAL_JobworkOrder_Hdr a1 on a.OrderType = a1.OrderType"
    Sql = Sql & " and a.OrderNo = a1.OrderNo and a.OrderDate = a1.OrderDate and a.Branchid = a1.Branchid"
    Sql = Sql & " and a.Compid = a1.Compid left outer join PartMast d2 on a.SentItemCode = d2.ItemCode"
    Sql = Sql & " and a.Branchid = d2.Branchid and a.Compid = d2.Compid left outer join INV_Material_Master d"
    Sql = Sql & " on a.ReceivedItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
    Sql = Sql & " left outer join INV_Material_Master d1 on a.SentItemCode = d1.ItemCode and a.Branchid = d1.Branchid and a.Compid = d1.Compid"
    Sql = Sql & " left outer join Inv_UOM_Master b on d.UOM = b.Code and d.Branchid = b.Branchid"
    Sql = Sql & " and d.Compid = b.Compid left outer join Inv_UOM_Master b1 on d1.UOM = b1.Code and d1.Branchid = b1.Branchid"
    Sql = Sql & " and d1.Compid = b1.Compid left outer join FAS_Party_Master c on a1.CustomerCode = c.PartyCode"
    Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid"
    Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
    Sql = Sql & " On f.OrderType = a1.OrderType and f.OrderNo = a1.OrderNo and f.OrderDate = a1.OrderDate"
    Sql = Sql & " and f.BranchID = a1.BranchID and f.CompID = a1.CompID"
    Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Revisionno = " & cmbRevNo & " and a.Branchid = '" & gBranchID & "' "
    Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If snap.EOF Then
      sShowMessage "No Record(s) Found"
      sPrintcurent_JobworkObsolete = False
      GoTo ExitHere
    End If
    
    pheight = 71
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "Sales Order")
    sSendToPort ""
    sSendToPort LAlign(17, "SO No.          :") & Space(1) & LAlign(20, snap("OrderType") & snap("OrderNo")) & Space(12) & LAlign(18, "SO Date          : ") & Space(1) & LAlign(10, snap("OrderDate"))
    sSendToPort LAlign(17, "Customer PO No. :") & Space(1) & LAlign(20, snap("ReferenceNo")) & Space(12) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, snap("ReferenceDate"))
    sSendToPort LAlign(17, "Revision Level  :") & Space(1) & LAlign(20, snap("RevisionLevel"))
    sSendToPort LAlign(17, "Customer Name   :") & Space(1) & LAlign(50, snap("PartyName"))
    sPageHeader
    
    i = 1
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      GenSpec = Trim$(snap("GeneralDetails"))
      For i = 1 To snap.RecordCount
        'ProcessFlow
        If Trim$(gReportID) = "PMI" Then
          Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
          Sql = Sql + " left outer join SAL_Order_Dtl b "
          Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
          Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
          If RSOpen(rs1, Sql) = False Then Exit Function
          If rs1.EOF Then
            Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
          End If
          'RawMaterials
          Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
          Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
          Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
          If RSOpen(rsDetails, Sql) = False Then Exit Function
        End If
                
        sSendToPort LAlign(20, snap("PartNo")) & Space(1) & LAlign(40, snap("ItemDescription")) & Space(1) & LAlign(3, snap("RevisionNo")) & Space(1) & LAlign(3, CheckNull(snap("sDesc"))) & Space(1) & RAlign(10, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty")))
        If Trim$(snap("ItemCode")) <> "" Then
 sSendToPort LAlign(20, snap("ItemCode"))
End If
'        If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'          WordSplit Trim$(snap("ItemSpecification")), 40
'          sSendToPort Space(21) & LAlign(40, WSplit(1))
'          For s = 2 To UBound(WSplit)
'            sSendToPort Space(21) & LAlign(40, WSplit(s))
'          Next s
'        End If
        sSendToPort ""
        snap.MoveNext
      Next i
    End If
    'Delivery Details
    Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
    Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
    If RSOpen(rs, Sql) = False Then Exit Function
    If Not rs.EOF Then
      PartNo = ""
      sSendToPort String(iPageWidth, "-")
      sSendToPort LAlign(17, "Delivery Details :") & Space(1) & LAlign(30, "Part No") & Space(1) & LAlign(10, "Sch. Date") & Space(1) & RAlign(10, "Sch.Qty")
      rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        If PartNo <> rs("PartNo") Then
          If i > 1 Then
 sSendToPort ""
End If
          sSendToPort Space(18) & LAlign(30, rs("PartNo")) & Space(1) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        Else
          sSendToPort Space(49) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        End If
        PartNo = rs("PartNo")
        rs.MoveNext
      Next i
    End If
    sSendToPort String(iPageWidth, "-")
    sSendToPort ""
                         'Process Flow Print
    If Trim$(gReportID) = "PMI" Then
      If Not rs1.EOF Then
        sSendToPort LAlign(20, "Operation(s):")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(5, ("S.No")) & Space(2) & LAlign(50, "Description")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rs1.RecordCount
          sSendToPort LAlign(5, rs1("ProcessSequence")) & Space(2) & LAlign(50, rs1("ProcessFlow"))
          rs1.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
        sSendToPort ""
      End If
                        'Raw Material Print
      If Not rsDetails.EOF Then
        sSendToPort LAlign(25, "Rawmaterial Details:")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(20, "Rawmaterial Code") & Space(1) & LAlign(40, "Description") & Space(1) & LAlign(5, "UOM") & Space(1) & RAlign(12, "Norms")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rsDetails.RecordCount
          sSendToPort LAlign(20, rsDetails("ChildCode")) & Space(1) & LAlign(40, rsDetails("ItemDescription")) & Space(1) & LAlign(5, rsDetails("PurchaseUOMDesc")) & Space(1) & RAlign(12, rsDetails("Qty"))
          rsDetails.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
      End If
    rs1.Close
    rsDetails.Close
    End If
    
    If Trim$(GenSpec) <> "" Then
      WordSplit Trim$(GenSpec), iPageWidth
      sSendToPort "Note :"
      sSendToPort LAlign(iPageWidth, WSplit(1))
      For s = 2 To UBound(WSplit)
        sSendToPort LAlign(iPageWidth, WSplit(s))
      Next s
      sSendToPort String(iPageWidth, "-")
    End If
                         'Order Issued
    Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
    Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
    Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
    Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      For i = 1 To snap.RecordCount
        If i = 1 Then
          sSendToPort LAlign(16, "Order Issued To:") & Space(1) & LAlign(50, snap("Description"))
        Else
          sSendToPort Space(17) & LAlign(50, snap("Description"))
        End If
        snap.MoveNext
      Next i
    End If
  
  Else

      frmPrint.Show
      
      Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a1.CustomerCode, c.PartyName,"
      Sql = Sql & " a1.ReferenceNo, a1.ReferenceDate, a.ItemGroupKey,a.ReceivedItemCode,a.SentItemCode,"
      Sql = Sql & " d.ItemDescription as RecDesc, d1.ItemDescription as SentDesc, d2.PartNo, a.RevisionNo, b.sDesc as Recuom, b.sDesc as Sentuom, a.OrderQty,"
      Sql = Sql & " isNull(d.ItemSpecification,'') ItemSpecification,a.ItemGroupKey,PackingDesc,InsuranceDesc,DeliveryDesc,PaymentDesc,FreightDesc,ModeDesc,FormDesc,FormulaName,"
      Sql = Sql & " PostingAcName from SAL_JobworkOrder_Dtl a"
      Sql = Sql & " left outer join SAL_JobworkOrder_Hdr a1 on a.OrderType = a1.OrderType"
      Sql = Sql & " and a.OrderNo = a1.OrderNo and a.OrderDate = a1.OrderDate and a.Branchid = a1.Branchid"
      Sql = Sql & " and a.Compid = a1.Compid left outer join PartMast d2 on a.SentItemCode = d2.ItemCode"
      Sql = Sql & " and a.Branchid = d2.Branchid and a.Compid = d2.Compid left outer join INV_Material_Master d"
      Sql = Sql & " on a.ReceivedItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
      Sql = Sql & " left outer join INV_Material_Master d1 on a.SentItemCode = d1.ItemCode and a.Branchid = d1.Branchid and a.Compid = d1.Compid"
      Sql = Sql & " left outer join Inv_UOM_Master b on d.UOM = b.Code and d.Branchid = b.Branchid"
      Sql = Sql & " and d.Compid = b.Compid left outer join Inv_UOM_Master b1 on d1.UOM = b1.Code and d1.Branchid = b1.Branchid"
      Sql = Sql & " and d1.Compid = b1.Compid left outer join FAS_Party_Master c on a1.CustomerCode = c.PartyCode"
      Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid"
      Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
      Sql = Sql & " On f.OrderType = a1.OrderType and f.OrderNo = a1.OrderNo and f.OrderDate = a1.OrderDate"
      Sql = Sql & " and f.BranchID = a1.BranchID and f.CompID = a1.CompID"
      Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Revisionno = " & cmbRevNo & " and a.Branchid = '" & gBranchID & "' "
      Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If snap.EOF Then
        sShowMessage "No Record(s) Found"
        sPrintcurent_JobworkObsolete = False
        GoTo ExitHere
      End If
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
  
      frmPrint.sprPrint.Col = 5
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
        If RptFlg <> "View" And RptFlg <> "Print" And RptFlg <> "Obsolute" Then
          frmPrint.sprPrint.AddCellSpan 1, j, 4, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
      End If
      
      If iPageNo = 1 Then
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

        j = j + 2
GridCount = GridCount + 2
      End If
          
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText 1, j, "Job Work Order Obsolete"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

      j = j + 2
GridCount = GridCount + 2
      
      frmPrint.sprPrint.MaxCols = 7
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(1) = 8

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(2) = 25

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(3) = 25

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(4) = 15

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(5) = 5

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.ColWidth(6) = 8

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.ColWidth(7) = 10
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "JO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderType") & snap("OrderNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "JO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer Name"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PartyName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Packing & Forwarding"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PackingDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Insurance"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("InsuranceDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DeliveryDesc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("DeliveryDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PaymentDetails"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PaymentDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Freight"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FreightDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Mode"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ModeDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Form"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "FormulaName"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormulaName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, 2, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PostingAc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PostingAcName")
      
      j = j + 2
GridCount = GridCount + 2
RowProcess
                                                
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd. Item Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sent Item code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rev."
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd.Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "GroupKey"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Recd. Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sent Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sen.Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
'        GenSpec = Trim$(snap("GeneralDetails"))
        For i = 1 To snap.RecordCount
          'ProcessFlow
          If Trim$(gReportID) = "PMI" Then
            Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
            Sql = Sql + " left outer join SAL_JobworkOrder_Dtl b " '('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b "
            Sql = Sql + " on a.ItemCode= b.RecdItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
            If rs1.EOF Then
              Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
              Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
              If RSOpen(rs1, Sql) = False Then Exit Function
            End If
            'RawMaterials
            Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
            Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
            Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
            If RSOpen(rsDetails, Sql) = False Then Exit Function
          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemGroupKey")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ReceivedItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("SentItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RevisionNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("RecUom"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          j = j + 1
GridCount = GridCount + 1
RowProcess
          
          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RecDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("SentDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("SentUom"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'          j = j + 1: GridCount = GridCount + 1: RowProcess
'          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
            
'          If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'            WordSplit Trim$(snap("ItemSpecification")), 40
'            .Col = 1: .Row = j

'            .SetText .Col, .Row, WSplit(1): .TypeHAlign = TypeHAlignLeft

'            j = j + 1: GridCount = GridCount + 1: RowProcess
'
'            For s = 2 To UBound(WSplit)
'              .Col = 1: .Row = j

'              .SetText .Col, .Row, WSplit(s): .TypeHAlign = TypeHAlignLeft

'              j = j + 1: GridCount = GridCount + 1: RowProcess
'            Next s
'          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          snap.MoveNext
        Next i
      End If
      'Delivery Details
      Sql = "select distinct a.SentItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from Sal_JobWorkOrder_Schedule a"
      Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
      Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
      If RSOpen(rs, Sql) = False Then Exit Function
      If Not rs.EOF Then
        PartNo = ""
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Delivery Details :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        frmPrint.sprPrint.AddCellSpan 5, j, 2, 1
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        j = j + 1
GridCount = GridCount + 1
RowProcess

        rs.MoveLast
rs.MoveFirst
        For i = 1 To rs.RecordCount
          If PartNo <> rs("PartNo") Then
            If i > 1 Then
 j = j + 1
GridCount = GridCount + 1
RowProcess
End If
            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
            frmPrint.sprPrint.AddCellSpan 5, j, 2, 1
            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
            frmPrint.sprPrint.AddCellSpan 5, j, 2, 1
            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          PartNo = rs("PartNo")
          rs.MoveNext
        Next i
      End If
      j = j + 1
GridCount = GridCount + 1
RowProcess
                           'Process Flow Print
      If Trim$(gReportID) = "PMI" Then
        If Not rs1.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Operation(s):"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "S.No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        
          For i = 1 To rs1.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessSequence")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessFlow")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rs1.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
                          'Raw Material Print
        If Not rsDetails.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Details:"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Norms"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          For i = 1 To rsDetails.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ChildCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("PurchaseUOMDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rsDetails.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
      rs1.Close
      rsDetails.Close
      End If
      
      If Trim$(GenSpec) <> "" Then
        WordSplit Trim$(GenSpec), iPageWidth
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Note :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(1)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        For s = 2 To UBound(WSplit)
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(s)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        Next s
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
                           'Order Issued
      Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
      Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
      Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
      Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        For i = 1 To snap.RecordCount
          If i = 1 Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Issued To :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          snap.MoveNext
        Next i
      End If
      j = j + 3
GridCount = GridCount + 3
RowProcess
      frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 5, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "System Administrator"
  frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
      Screen.MousePointer = 1
      totalpages = iPageNo
  End If
  
  sPrintcurent_JobworkObsolete = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set rs = Nothing
  Set rs1 = Nothing
  Set rsDetails = Nothing
End Try
End Function

Public Sub sAudit()
    
 Try

  fraStep(0).Visible = False
  CmdStep(0).Visible = False
  
  fraStep(1).Visible = False
  CmdStep(1).Visible = False
  fraStep(2).Visible = True
  CmdStep(2).Visible = True
  fraStep(3).Visible = True
  CmdStep(3).Visible = True
  
  fraStep(2).Enabled = True
  CmdStep(2).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Public Sub sObsolute()
    
     
  Try

  fraStep(1).Left = 3815
  fraStep(1).Width = 4100
  CmdStep(1).Left = 7950

  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  fraStep(3).Visible = False
  CmdStep(3).Visible = False
  fraStep(0).Visible = False
  CmdStep(0).Visible = False

  lblRevNo.Enabled = True
  cmbRevNo.Enabled = True
  fraStep(1).Enabled = True
  CmdStep(1).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurent_SalesObsolete() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim s As Integer
  Dim GenSpec As String
  Dim PartNo As String
  Dim CurDate As Date
  
  Try
  
  sPrintcurent_SalesObsolete = False
  
  CurDate = ServerDate
  Set snap = New AceADODB.Recordset
  Set rs = New AceADODB.Recordset
  Set rs1 = New AceADODB.Recordset
  Set rsDetails = New AceADODB.Recordset
  
  If cmbType.ListIndex = 1 Then
    Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a1.CustomerCode, c.PartyName, a1.ReferenceNo,"
    Sql = Sql & " a1.ReferenceDate,a1.RevisionLevel, a.ItemCode, a.ItemDescription, a.PartNo, a.RevisionNo,"
    Sql = Sql & " b.sDesc, a.OrderQty, isNull(d.ItemSpecification,'') ItemSpecification, a1.GeneralDetails,"
    Sql = Sql & " PackingDesc, InsuranceDesc, DeliveryDesc, PaymentDesc, FreightDesc, ModeDesc, FormDesc, FormulaName, PostingAcName "
    Sql = Sql & " from SAL_Order_Dtl a left outer join SAL_Order_hdr a1 on a.OrderType = a1.OrderType"
    Sql = Sql & " and a.OrderNo = a1.OrderNo and a.OrderDate = a1.OrderDate and a.Branchid = a1.Branchid"
    Sql = Sql & " and a.Compid = a1.Compid left outer join Inv_UOM_Master b on a.UOM = b.Code and a.Branchid = b.Branchid"
    Sql = Sql & " and a.Compid = b.Compid left outer join FAS_Party_Master c on a1.CustomerCode = c.PartyCode"
    Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join INV_Material_Master d"
    Sql = Sql & " on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
    Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
    Sql = Sql & " On f.OrderType = a1.OrderType and f.OrderNo = a1.OrderNo and f.OrderDate = a1.OrderDate"
    Sql = Sql & " and f.BranchID = a1.BranchID and f.CompID = a1.CompID"
    Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
    Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    Sql = Sql & " and a.Revisionno = " & cmbRevNo & " and a.Branchid = '00'  and a.Compid = '01' and a.Finyear = '04-2007 03-2008'"
  
    If RSOpen(snap, Sql) = False Then Exit Function
    If snap.EOF Then
      sShowMessage "No Record(s) Found"
      sPrintcurent_SalesObsolete = False
      GoTo ExitHere
    End If
    
    pheight = 71
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "Sales Order")
    sSendToPort ""
    sSendToPort LAlign(17, "SO No.          :") & Space(1) & LAlign(20, snap("OrderType") & snap("OrderNo")) & Space(12) & LAlign(18, "SO Date          : ") & Space(1) & LAlign(10, snap("OrderDate"))
    sSendToPort LAlign(17, "Customer PO No. :") & Space(1) & LAlign(20, snap("ReferenceNo")) & Space(12) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, snap("ReferenceDate"))
    sSendToPort LAlign(17, "Revision Level  :") & Space(1) & LAlign(20, snap("RevisionLevel"))
    sSendToPort LAlign(17, "Customer Name   :") & Space(1) & LAlign(50, snap("PartyName"))
    sPageHeader
    
    i = 1
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      GenSpec = Trim$(snap("GeneralDetails"))
      For i = 1 To snap.RecordCount
        'ProcessFlow
        If Trim$(gReportID) = "PMI" Then
          Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
          Sql = Sql + " left outer join SAL_Order_Dtl b "
          Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
          Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
          If RSOpen(rs1, Sql) = False Then Exit Function
          If rs1.EOF Then
            Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
          End If
          'RawMaterials
          Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
          Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
          Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
          If RSOpen(rsDetails, Sql) = False Then Exit Function
        End If
                
        sSendToPort LAlign(20, snap("PartNo")) & Space(1) & LAlign(40, snap("ItemDescription")) & Space(1) & LAlign(3, snap("RevisionNo")) & Space(1) & LAlign(3, CheckNull(snap("sDesc"))) & Space(1) & RAlign(10, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty")))
        If Trim$(snap("ItemCode")) <> "" Then
 sSendToPort LAlign(20, snap("ItemCode"))
End If
'        If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'          WordSplit Trim$(snap("ItemSpecification")), 40
'          sSendToPort Space(21) & LAlign(40, WSplit(1))
'          For s = 2 To UBound(WSplit)
'            sSendToPort Space(21) & LAlign(40, WSplit(s))
'          Next s
'        End If
        sSendToPort ""
        snap.MoveNext
      Next i
    End If
    'Delivery Details
    Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
    Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
    If RSOpen(rs, Sql) = False Then Exit Function
    If Not rs.EOF Then
      PartNo = ""
      sSendToPort String(iPageWidth, "-")
      sSendToPort LAlign(17, "Delivery Details :") & Space(1) & LAlign(30, "Part No") & Space(1) & LAlign(10, "Sch. Date") & Space(1) & RAlign(10, "Sch.Qty")
      rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        If PartNo <> rs("PartNo") Then
          If i > 1 Then
 sSendToPort ""
End If
          sSendToPort Space(18) & LAlign(30, rs("PartNo")) & Space(1) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        Else
          sSendToPort Space(49) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        End If
        PartNo = rs("PartNo")
        rs.MoveNext
      Next i
    End If
    sSendToPort String(iPageWidth, "-")
    sSendToPort ""
                         'Process Flow Print
    If Trim$(gReportID) = "PMI" Then
      If Not rs1.EOF Then
        sSendToPort LAlign(20, "Operation(s):")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(5, ("S.No")) & Space(2) & LAlign(50, "Description")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rs1.RecordCount
          sSendToPort LAlign(5, rs1("ProcessSequence")) & Space(2) & LAlign(50, rs1("ProcessFlow"))
          rs1.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
        sSendToPort ""
      End If
                        'Raw Material Print
      If Not rsDetails.EOF Then
        sSendToPort LAlign(25, "Rawmaterial Details:")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(20, "Rawmaterial Code") & Space(1) & LAlign(40, "Description") & Space(1) & LAlign(5, "UOM") & Space(1) & RAlign(12, "Norms")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rsDetails.RecordCount
          sSendToPort LAlign(20, rsDetails("ChildCode")) & Space(1) & LAlign(40, rsDetails("ItemDescription")) & Space(1) & LAlign(5, rsDetails("PurchaseUOMDesc")) & Space(1) & RAlign(12, rsDetails("Qty"))
          rsDetails.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
      End If
    rs1.Close
    rsDetails.Close
    End If
    
    If Trim$(GenSpec) <> "" Then
      WordSplit Trim$(GenSpec), iPageWidth
      sSendToPort "Note :"
      sSendToPort LAlign(iPageWidth, WSplit(1))
      For s = 2 To UBound(WSplit)
        sSendToPort LAlign(iPageWidth, WSplit(s))
      Next s
      sSendToPort String(iPageWidth, "-")
    End If
                         'Order Issued
    Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
    Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
    Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
    Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      For i = 1 To snap.RecordCount
        If i = 1 Then
          sSendToPort LAlign(16, "Order Issued To:") & Space(1) & LAlign(50, snap("Description"))
        Else
          sSendToPort Space(17) & LAlign(50, snap("Description"))
        End If
        snap.MoveNext
      Next i
    End If
  
  Else

      frmPrint.Show
      
      Sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate, a1.CustomerCode, c.PartyName, a1.ReferenceNo,"
      Sql = Sql & " a1.ReferenceDate,a1.RevisionLevel, a.ItemCode, a.ItemDescription, a.PartNo, a.RevisionNo,"
      Sql = Sql & " b.sDesc, a.OrderQty, isNull(d.ItemSpecification,'') ItemSpecification, a1.GeneralDetails,"
      Sql = Sql & " PackingDesc, InsuranceDesc, DeliveryDesc, PaymentDesc, FreightDesc, ModeDesc, FormDesc, FormulaName, PostingAcName "
      Sql = Sql & " from SAL_Order_Dtl a left outer join SAL_Order_hdr a1 on a.OrderType = a1.OrderType"
      Sql = Sql & " and a.OrderNo = a1.OrderNo and a.OrderDate = a1.OrderDate and a.Branchid = a1.Branchid"
      Sql = Sql & " and a.Compid = a1.Compid left outer join FAS_Party_Master c on a1.CustomerCode = c.PartyCode"
      Sql = Sql & " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join INV_Material_Master d"
      Sql = Sql & " on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
      Sql = Sql & " left outer join Inv_UOM_Master b on d.UOM = b.Code and d.Branchid = b.Branchid and d.Compid = b.Compid"
      Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
      Sql = Sql & " On f.OrderType = a1.OrderType and f.OrderNo = a1.OrderNo and f.OrderDate = a1.OrderDate"
      Sql = Sql & " and f.BranchID = a1.BranchID and f.CompID = a1.CompID"
      Sql = Sql & " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
      Sql = Sql & " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      Sql = Sql & " and a.Revisionno = " & cmbRevNo & " and a.Branchid = '00'  and a.Compid = '01' and a.Finyear = '04-2007 03-2008'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If snap.EOF Then
        sShowMessage "No Record(s) Found"
        sPrintcurent_SalesObsolete = False
        GoTo ExitHere
      End If
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
  
      frmPrint.sprPrint.Col = 4
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
        If RptFlg <> "View" And RptFlg <> "Print" And RptFlg <> "Obsolute" Then
          frmPrint.sprPrint.AddCellSpan 1, j, 4, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
      End If
      
      If iPageNo = 1 Then
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

        j = j + 2
GridCount = GridCount + 2
      End If
          
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText 1, j, "Sales Order Obsolete"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

      j = j + 2
GridCount = GridCount + 2
      
      frmPrint.sprPrint.MaxCols = 5
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(1) = 20

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(2) = 36

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(3) = 10

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(4) = 6

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(5) = 10
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "SO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderType") & snap("OrderNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "SO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Level"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("RevisionLevel")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer Name"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PartyName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Packing & Forwarding"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PackingDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Insurance"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("InsuranceDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DeliveryDesc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("DeliveryDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PaymentDetails"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PaymentDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Freight"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FreightDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Mode"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ModeDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Form"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "FormulaName"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormulaName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PostingAc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PostingAcName")
      j = j + 2
GridCount = GridCount + 2
RowProcess
                                                
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
Cont:
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rev. No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      s = s + 1
      If s > 1 Then
 GoTo Cont1
End If
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        GenSpec = Trim$(snap("GeneralDetails"))
        For i = 1 To snap.RecordCount
          'ProcessFlow
          If Trim$(gReportID) = "PMI" Then
            Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
            Sql = Sql + " left outer join SAL_Order_Dtl b " '('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b "
            Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
            If rs1.EOF Then
              Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
              Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
              If RSOpen(rs1, Sql) = False Then Exit Function
            End If
            'RawMaterials
            Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
            Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
            Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
            If RSOpen(rsDetails, Sql) = False Then Exit Function
          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
Cont1:
          k = 0
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RevisionNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("sDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'          j = j + 1: GridCount = GridCount + 1: RowProcess
'          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, .MaxCols, j
            
'          If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'            WordSplit Trim$(snap("ItemSpecification")), 40
'            .Col = 1: .Row = j

'            .SetText .Col, .Row, WSplit(1): .TypeHAlign = TypeHAlignLeft

'            j = j + 1: GridCount = GridCount + 1: RowProcess
'
'            For s = 2 To UBound(WSplit)
'              .Col = 1: .Row = j

'              .SetText .Col, .Row, WSplit(s): .TypeHAlign = TypeHAlignLeft

'              j = j + 1: GridCount = GridCount + 1: RowProcess
'            Next s
'          End If
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          If pheight > 54 And k = 1 And s = 1 Then
 GoTo Cont
End If
          snap.MoveNext
        Next i
      End If
      'Delivery Details
      Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
      Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
      Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
      If RSOpen(rs, Sql) = False Then Exit Function
      If Not rs.EOF Then
        PartNo = ""
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Delivery Details :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 2, j, frmPrint.sprPrint.MaxCols, j

        rs.MoveLast
rs.MoveFirst
        For i = 1 To rs.RecordCount
          If PartNo <> rs("PartNo") Then
            If i > 1 Then
 j = j + 1
GridCount = GridCount + 1
RowProcess
End If
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          PartNo = rs("PartNo")
          rs.MoveNext
        Next i
      End If
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
                           'Process Flow Print
      If Trim$(gReportID) = "PMI" Then
        If Not rs1.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Operation(s):"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "S.No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        
          For i = 1 To rs1.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessSequence")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessFlow")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rs1.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
                          'Raw Material Print
        If Not rsDetails.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Details:"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Norms"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          For i = 1 To rsDetails.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ChildCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("PurchaseUOMDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rsDetails.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
      rs1.Close
      rsDetails.Close
      End If
      
      If Trim$(GenSpec) <> "" Then
        WordSplit Trim$(GenSpec), iPageWidth
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Note :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(1)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        For s = 2 To UBound(WSplit)
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(s)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        Next s
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
                           'Order Issued
      Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
      Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
      Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
      Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        For i = 1 To snap.RecordCount
          If i = 1 Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Issued To :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          snap.MoveNext
        Next i
      End If
      j = j + 3
GridCount = GridCount + 3
RowProcess

      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 3, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "System Administrator"
  frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      j = j + 1
GridCount = GridCount + 1
RowProcess
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  End If
  
  sPrintcurent_SalesObsolete = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set rs = Nothing
  Set rs1 = Nothing
  Set rsDetails = Nothing
End Try
End Function

Public Sub sRevisionSummary()
    
  Try

  fraStep(1).Left = 3815
  fraStep(1).Width = 4100
  CmdStep(1).Left = 7950

  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  fraStep(3).Visible = False
  CmdStep(3).Visible = False
  fraStep(0).Visible = False
  CmdStep(0).Visible = False

  fraStep(1).Enabled = True
  CmdStep(1).Enabled = True

  lblRevNo.Enabled = False
  cmbRevNo.Enabled = False
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sListing()
    
  Try

  fraStep(0).Visible = False
  CmdStep(0).Visible = False
  
  fraStep(1).Visible = False
  CmdStep(1).Visible = False
  fraStep(2).Visible = True
  CmdStep(2).Visible = True
  fraStep(3).Visible = True
  CmdStep(3).Visible = True
  
  fraStep(2).Enabled = True
  CmdStep(2).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sPrint()
  Try

  fraStep(1).Left = 4240
  fraStep(1).Width = 3160
  CmdStep(1).Left = 7470
  
  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  fraStep(3).Visible = False
  CmdStep(3).Visible = False
  
  fraStep(0).Visible = True
  CmdStep(0).Visible = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  fraStep(1).Left = 4240
  fraStep(1).Width = 3160
  CmdStep(1).Left = 7470
  
  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  fraStep(3).Visible = False
  CmdStep(3).Visible = False
  
  fraStep(0).Visible = True
  CmdStep(0).Visible = True
  frmPrint.cmdPPrint.Enabled = False
  frmPrint.cmdExcel.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
'Section 1 : ( Main Control(s) ) :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try
  CurrDate Me
  
  Call sView
  cmbType.ListIndex = 0
  cmbAuthorise.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub txtType_DblClick()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim mRevNo As String
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass
  cmbRevNo.Clear
  
  Sql = "Select Distinct OrderType, OrderNo, OrderDate"

  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobworkOrder_Hdr"
End If

  Sql = Sql & " Where Branchid='" & gBranchID & "'"
  Sql = Sql & " and Compid = '" & gCompID & "' and finyear = '" & gFinyear & "'"
  Sql = Sql & " Group by OrderType, OrderNo, OrderDate"
  Sql = Sql & " Order by OrderDate desc, OrderNo desc"

  If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of Order No's", "Order Type", "Order No", "Order Date", "Rev. No.", "Rev. Date") = True Then
  
    txttype.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)

    If Trim$(txttype.Text) = "" Then
      sShowMessage "Record(s) Not Found"
      GoTo ExitHere
     Else
        'Skip
    End If

    If Trim$(RptFlg) = "Obsolute" Or Trim$(RptFlg) = "RevisionSummary" Then
      Sql = "Select Max(CreatedDate) As CreatedDate, Max(RevisionNo) As RevNo"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Dtl n"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobworkOrder_Dtl n"
End If

      Sql = Sql & " Where n.OrderType='" & Trim$(txttype.Text) & "'"
      Sql = Sql & " and n.OrderNo='" & Trim$(txtNo.Text) & "'"
      Sql = Sql & " and n.OrderDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
      Sql = Sql & " and n.CreatedDate=( Select Max(m.CreatedDate) As CreatedDate"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Dtl m"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobworkOrder_Dtl m"
End If

      Sql = Sql & " Where m.OrderType='" & Trim$(txttype.Text) & "'"
      Sql = Sql & " and m.OrderNo='" & Trim$(txtNo.Text) & "'"
      Sql = Sql & " and m.OrderDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
      Sql = Sql & " and m.Branchid='" & gBranchID & "'"
      Sql = Sql & " and m.Compid = '" & gCompID & "' and m.finyear = '" & gFinyear & "'"
      Sql = Sql & " Group by m.OrderType, m.OrderNo, m.OrderDate"
      Sql = Sql & " )"
      Sql = Sql & " and n.Branchid='" & gBranchID & "'"
      Sql = Sql & " and n.Compid = '" & gCompID & "' and n.finyear = '" & gFinyear & "'"

      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

      If rs.RecordCount <= 0 Then
        sShowMessage "Obsolete does not exist"
        Call ClearControls
        GoTo ExitHere
      End If
      
      mRevNo = Trim$(rs("RevNo"))
      If Val(rs("RevNo")) <= 0 Or Trim$(mRevNo) = "" Then
        If Trim$(RptFlg) = "Obsolute" Then
          sShowMessage "Obsolete Document did not exist"
          Call ClearControls
          GoTo ExitHere
        End If
      End If
    
      If Val(mRevNo) = 0 Then
          cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(mRevNo)) / 1))) & Val(Val(mRevNo))
      ElseIf Val(mRevNo) >= 1 Then
        If Trim$(DocFlg) = "Sales" Or Trim$(DocFlg) = "Job Work" Then
          For i = Val(mRevNo) To 0 Step -1
            cmbRevNo.AddItem Space(5 - Len(Trim$(Int(i) / 1))) & Val(i)
          Next i
        ElseIf Trim$(RptFlg) = "Obsolute" Then
          For i = Val(mRevNo) - 1 To 0 Step -1
            cmbRevNo.AddItem Space(5 - Len(Trim$(Int(i) / 1))) & Val(i)
          Next i
        Else
          'Skip
        End If
      Else
        'Skip
      End If
      If cmbRevNo.ListCount >= 1 Then
 cmbRevNo.ListIndex = 0
End If
      
      rs.Close
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub OptCustAll_Click()
' On Error Resume Next
  cmbType.ListIndex = 0
  fpsList.Visible = False
End Sub

Private Sub optCustSelective_Click()
  Dim Sql As String
  Dim CurrDate As Date
  Dim rs As AceADODB.Recordset
 
  Try
  Me.MousePointer = vbHourglass
  
  CurrDate = ServerDate
  
  'Listing & Audit
  
  Sql = " Select Distinct IsNull(b.PartyName,'--') As CustomerName,"
  Sql = Sql & " a.CustomerCode"
  
  If DocFlg = "Sales" Then
    Sql = Sql & " From SAL_CUSTOMER_ITEMMASTER a"
  Else
    Sql = Sql & " From SAL_JobWork_ITEMMASTER a"
  End If
  
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On a.CustomerCode=b.PartyCode and a.BranchID=b.Branchid"
  Sql = Sql & " and a.CompID=b.Compid"
  Sql = Sql & " Where a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  
  If DocFlg = "Sales" Then
    Sql = Sql & " and a.Status=1 and a.CustomerCode in (Select CustomerCode From SAL_Order_Hdr "
    Sql = Sql & " Where CustomerCode = a.CustomerCode and BranchID = a.BranchID and CompID = a.CompID)"
  Else
    Sql = Sql & " and a.Isactive=1 and a.CustomerCode in (Select CustomerCode From SAL_JobworkOrder_Hdr"
    Sql = Sql & " Where CustomerCode = a.CustomerCode And BranchID = a.BranchID And CompID = a.CompID)"
  End If
    
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  fpsList.MaxRows = fpsList.DataRowCnt
'
'  If fpsList.DataRowCnt <= 0 Then
'    sShowMessage "Record(s) Not Found"
'    optCustAll.Value = True
'    GoTo ExitHere
'  Else
'    Call Spread_HideCol(fpsList, True, 2)
'    CmdStep(3).Enabled = False
'    fpsList.Enabled = True
'    fpsList.Visible = True
'    fpsList.SetFocus
'  End If
'
'  fpsList.ClearSelection
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

If rs.EOF Then
  sShowMessage "Record(s) Not Found"
  optCustAll.Value = True
  GoTo ExitHere
Else
 Call Spread_HideCol(fpsList, True, 2)
'    CmdStep(3).Enabled = False
 fpsList.Enabled = True
fpsList.Visible = True
fpsList.ClearSelection
  Call ADORS_Spread_DataFill(rs, fpsList, True)
  
  fpsList.MaxRows = fpsList.DataRowCnt
  Call Spread_Set_Focus(fpsList, 1, 1)
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmbRevNo_Click()

  Try
  mRevPrev = ""
mRevCurr = ""
  
  If Trim$(cmbRevNo.Text) = "" Then
'    lblRevNo(1).Caption = ""
  ElseIf Val(cmbRevNo.List(cmbRevNo.ListIndex)) = (cmbRevNo.ListCount - 1) Then
    If Val(cmbRevNo.List(cmbRevNo.ListIndex)) = 0 Then
      mRevCurr = "0"
    Else
      If cmbRevNo.ListCount >= 2 Then
        If Val(cmbRevNo.List(cmbRevNo.ListIndex)) <> 0 Then
'          If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex) - 1) & "-" & (Val(cmbRevNo.List(cmbRevNo.ListIndex)))
          mRevPrev = Val(cmbRevNo.List(cmbRevNo.ListIndex) - 1)
          mRevCurr = Val(cmbRevNo.List(cmbRevNo.ListIndex))
        Else
'          If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex)) & " Only"
          mRevCurr = Val(cmbRevNo.List(cmbRevNo.ListIndex))
        End If
      ElseIf cmbRevNo.ListCount = 1 Then
'        If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex)) & " Only"
        mRevCurr = Val(cmbRevNo.List(cmbRevNo.ListIndex))
      Else
        sShowMessage "does not exist Revision Summary"
      End If
    End If
  Else
'    If optSelective Then
'      If cmbRevNo.ListCount >= 2 Then
'        If Val(cmbRevNo.List(cmbRevNo.ListIndex)) <> 0 Then
'          If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex) - 1) & "-" & (Val(cmbRevNo.List(cmbRevNo.ListIndex)))
'          lblrevno.Tag = Val(cmbRevNo.List(cmbRevNo.ListIndex) - 1)
'          lblRevNo(1).Tag = Val(cmbRevNo.List(cmbRevNo.ListIndex))
'        Else
'          If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex)) & " Only"
'          lblRevNo(1).Tag = Val(cmbRevNo.List(cmbRevNo.ListIndex))
'        End If
'      ElseIf cmbRevNo.ListCount = 1 Then
'        If optSelective Then: lblRevNo(1).Caption = Val(cmbRevNo.List(cmbRevNo.ListIndex)) & " Only"
'        lblRevNo(1).Tag = Val(cmbRevNo.List(cmbRevNo.ListIndex))
'      Else
'        If optSelective Then: lblRevNo(1).Caption = "[None]"
'      End If
'    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub fpsList_KeyUp(sender as object, e as KeyUpEventArgs)
  
' On Error Resume Next
  
  If fpsList.SelectionCount >= 1 Then
    CmdStep(3).Enabled = True
  Else
    CmdStep(3).Enabled = False
  End If
End Sub

Private Sub fpsList_Click(sender as object, e as ClickEventArgs)

' On Error Resume Next
  
  If fpsList.SelectionCount >= 1 Then
    CmdStep(3).Enabled = True
  Else
    CmdStep(3).Enabled = False
  End If
  

'  Dim i As Integer
'  Dim retVal As Variant
'
'  On Error Resume Next
'  If Row = 0 And Col > 0 Then
'    fpsList.Tag = Col
'    With fpsList
'      .Row = 0: .Row2 = 0
'
'      .BlockMode = True
'      For i = 1 To .MaxCols
'        .Col = i: .Col2 = i: .GetText i, 0, retVal
'        .CellType = CellTypeButton
'        .TypeButtonText = Trim$(retVal)
'        .TypeButtonTextColor = RGB(0, 0, 0)
'      Next i
'
'      .Col = Col: .Col2 = Col
'      .BlockMode = False
'
'    End With
'  End If
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If optCustSelective.Value = True Then
    If fpsList.SelectionCount = 0 Then
      sShowMessage "Select atleast one Customer"
      Exit Sub
    End If
  End If
  
  If Index = 1 Then
    If RptFlg = "RevisionSummary" And Trim$(txttype) = "" Then
      sShowMessage "Select the Enquiry Type"
      Exit Sub
    End If
  ElseIf Index = 2 Then
    If RptFlg = "Listing" Then
      If dtpTodate < dtpFromDate Then
        sShowMessage "To date should be greater than From Date"
        Exit Sub
      End If
      
      If optCustSelective = True Then
        If fpsList.SelectionCount = 0 Then
          sShowMessage "Select atleast one Customer"
          Exit Sub
        End If
      End If
    End If
  ElseIf Index = 3 Then
    If dtpTodate < dtpFromDate Then
        sShowMessage "To date should be greater than From Date"
        Exit Sub
      End If
      
    If optCustSelective = True Then
      If fpsList.SelectionCount = 0 Then
        sShowMessage "Select atleast one Customer"
        Exit Sub
      End If
    End If
  End If
  
  Call StepProceedNext(Me, Index)
  fpsList.Enabled = False
  
  
'  If FraStep(0).Visible = True And cmdStep(0).Enabled = True Then
'    Call StepProceedNext(Me, Index)
'    Index = 2: Call StepProceedNext(Me, Index)
'  ElseIf FraStep(1).Visible = True And cmdStep(1).Enabled = True Then
'    If txtType = "" Then: sShowMessage "Select the Order Type": Exit Sub
'    Call StepProceedNext(Me, Index)
'    Index = 3: Call StepProceedNext(Me, Index)
'  ElseIf FraStep(2).Visible = True And cmdStep(2).Enabled = True Then
'    Call StepProceedNext(Me, Index)
'    Index = 2: Call StepProceedNext(Me, Index)
'  ElseIf FraStep(3).Visible = True And cmdStep(3).Enabled = True Then
'    Call StepProceedNext(Me, Index)
'    Index = 3: Call StepProceedNext(Me, Index)
'  Else
'    Call StepProceedNext(Me, Index)
'    cmdView.Enabled = True
'  End If
  
End Sub

Private Sub cmdview_Click()

  Try
  Me.MousePointer = vbHourglass

  If cmbType.ListIndex = 0 Then
  
    If Trim$(RptFlg) = "View" Then
      If Trim$(DocFlg) = "Sales" Then
        If sPrintcurrent_SalesView = False Then
 GoTo ExitHere
End If
      Else
        If sPrintcurent_JobworkView = False Then
 GoTo ExitHere
End If
      End If
    ElseIf Trim$(RptFlg) = "Print" Then
      If Trim$(DocFlg) = "Sales" Then
        If sPrintcurrent_SalesView = False Then
 GoTo ExitHere
End If
      Else
        If sPrintcurent_JobworkView = False Then
 GoTo ExitHere
End If
      End If
    ElseIf Trim$(RptFlg) = "Listing" Then
      If sPrintcurrent_Listing = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "RevisionSummary" Then
        If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Obsolute" Then
      If Trim$(DocFlg) = "Sales" Then
        If sPrintcurent_SalesObsolete = False Then
 GoTo ExitHere
End If
      Else
        If sPrintcurent_JobworkObsolete = False Then
 GoTo ExitHere
End If
      End If
    ElseIf Trim$(RptFlg) = "Audit" Then
      If sPrintcurrent_Audit = False Then
 GoTo ExitHere
End If
    Else
      'Skip
    End If
    
  ElseIf cmbType.ListIndex = 1 Then
  
    Set frmCurrent = Me
    DoEvents
    iLineNo = 1
    iPageNo = 1
    
    If Trim$(RptFlg) = "Listing" Then
      If sPrintcurrent_Listing = False Then
 GoTo ExitHere
End If
'    ElseIf Trim$(RptFlg) = "View" Then
'
'    ElseIf Trim$(RptFlg) = "View" Then  'Revision Details
'
'    ElseIf Trim$(RptFlg) = "View" Then  'Audit
    
    ElseIf Trim$(RptFlg) = "View" Or Trim$(RptFlg) = "Print" Then     'View
      If sPrintcurrent_SalesView = False Then
 GoTo ExitHere
End If
    End If
    
    sClosePort
    Set frmOutput = New frmReportViewer
    frmOutput.Caption = gStrReportHeading
    frmOutput.txtTotalPages.Text = CStr(iPageNo)
    frmOutput.wbrView.Navigate strFilename
    frmOutput.Tag = strFilename
    gStrReportHeading = ""
    frmOutput.Show vbModal
    DoEvents
    Screen.MousePointer = vbNormal
    Set frmOutput = Nothing
        
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub CmdClear_Click()
' On Error Resume Next
  
  Call ClearControls
  Call StepClear(Me)
  If RptFlg = "View" Then
    Call sView
  ElseIf RptFlg = "Print" Then
    Call sPrint
  ElseIf RptFlg = "Listing" Then
    Call sListing
  ElseIf RptFlg = "Print" Then
    Call sAudit
  ElseIf RptFlg = "RevisionSummary" Then
    Call sRevisionSummary
  Else
    Call sObsolute
  End If
  Call sButtonsAction
  Call StepEnable(Me)
  cmbType.ListIndex = 0
  chkPrintDate.Value = 1
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

'Section : Function/Procedure/Subroutine :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub ClearControls()

  Try
    
  fpsList.Visible = False
  optCustAll.Value = True
  CmdStep(1).Enabled = False
  cmbRevNo.Clear
  
  fraStep(0).Enabled = True
  CmdStep(0).Enabled = True
  fraStep(1).Enabled = True
  CmdStep(1).Enabled = True
  
  cmbType.ListIndex = 0
  chkPrintDate.Value = 0
  txttype.Text = ""
txttype.Tag = ""
txtNo.Text = ""
txtDate.Text = ""
mRevNo = ""
mDate = ""
  
  cmdview.Enabled = False
  If txttype.Visible = True Then
 txttype.SetFocus
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sPageHeader()
  Try
  
  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  If mDocSeq(1) = 1 Then
  
    sSendToPort String(iPageWidth, "-")
    sSendToPort LAlign(10, "Order No.") & Space(1) & LAlign(30, "Customer") & Space(1) & LAlign(20, "Part No.") & Space(1) & LAlign(4, "UOM") & Space(1) & LAlign(25, "Item Code") & Space(1) & LAlign(12, "UOM") & Space(1) & RAlign(12, "Rate") & Space(1) & RAlign(12, "Goods Value")
    sSendToPort LAlign(10, "Order Date") & Space(32) & LAlign(20, "Part Description") & Space(6) & LAlign(25, "Item Description") & Space(1) & LAlign(12, "Qty") & Space(1) & RAlign(12, "Ass.Rate") & Space(1) & RAlign(12, "Ass. Value")
    sSendToPort String(iPageWidth, "-")
    
  ElseIf mDocSeq(1) = 4 Then
  
    sSendToPort Space(2) & String(iPageWidth, "-")
    sSendToPort Space(2) & LAlign(20, "Item Code") & Space(1) & LAlign(33, "Item Description") & Space(1) & LAlign(3, "UOM") & Space(1) & RAlign(10, "Rate") & Space(1) & RAlign(10, "Value")
    sSendToPort Space(56) & Space(1) & LAlign(3, "Qty") & Space(1) & RAlign(10, "Rate") & Space(1) & RAlign(10, "Ass.Value")
    sSendToPort Space(2) & String(iPageWidth, "-")

  ElseIf mDocSeq(1) = 5 Then
  
    Printer.Print Space(2) & String(iPageWidth, "-")
    Printer.Print Space(2) & LAlign(20, "Item Code") & Space(1) & LAlign(33, "Item Description") & Space(1) & LAlign(3, "UOM") & Space(1) & RAlign(10, "Qty") & Space(1) & RAlign(10, "Value")
    Printer.Print Space(2) & Space(1) & LAlign(3, "Qty") & Space(1) & RAlign(10, "Rate") & Space(1) & RAlign(10, "Ass.Value")
    Printer.Print Space(2) & String(iPageWidth, "-")
    
  Else
    'Skip
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrent_Listing() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim SRow As Long
  Dim Gtot(2) As Double
  Dim tmp As String
  Dim CurrDate As Date
  Dim i As Integer
  Dim k As Integer
  Dim mEnqType As String
  Dim mCustNm As String
  Dim mAmendNo As String
    

  Try
  sPrintcurrent_Listing = False
  mAmendNo = ""
  
  CurrDate = ServerDate
  
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
  End If

  Sql = "Select a.OrderType, a.OrderNo, a.OrderDate,"
  Sql = Sql & " IsNull(m.OrderQty,'0') As Qty, m.Rate As Rate,"
  Sql = Sql & " cast((IsNull(m.OrderQty,'0')*m.Rate) As Decimal(15,2)) As GoodsValue,"
  Sql = Sql & " Case IsNull(a.OrderSource,'') When 'D' Then 'Direct'"
  Sql = Sql & " When 'Q' Then 'Quotation' When 'W' Then 'Without ItemCode'"
  Sql = Sql & " End As OrderSource, a.CustomerCode, a.ReferenceNo As RefNo, a.ReferenceDate As RefDate,"
  Sql = Sql & " a.OrderingType, a.CustomerCode, IsNull(e.PartyName,'') As CustomerName,"
  Sql = Sql & " IsNull(a.Remarks,'') As Remarks,"
  Sql = Sql & " Case IsNull(a.Status,'') When 'N' Then 'New' When 'M' Then 'Amendment'"
  Sql = Sql & " When 'A' Then 'Authorised' When 'H' Then 'Hold'"
  Sql = Sql & " End As OrdStatus, a.CreatedDate,"
  Sql = Sql & " IsNull(u.UserName,'') UserName, IsNull(v.GroupName,'') As GroupName,"

  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " IsNull(m.PartNo,'') As PartNo,"
    Sql = Sql & " IsNull(g.ItemDescription,'') As PartDescription,"
    Sql = Sql & " IsNull(h.SDesc,'') As PartUOM,"
    Sql = Sql & " IsNull(m.ItemCode,'') As ItemCode,"
    Sql = Sql & " IsNull(m.ItemDescription,'') As ItemDescription,"
    Sql = Sql & " IsNull(m.UOM,'') As UOM, m.AssessableRate As AssRate"
    Sql = Sql & " From SAL_Order_Hdr a"
    Sql = Sql & " Full Outer Join SAL_Order_Dtl m"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " IsNull(f.PartNo,'') As PartNo, IsNull(g.ItemDescription,'') As PartDescription,"
    Sql = Sql & " IsNull(m.ReceivedItemCode,'') As RecdItemCode, "
    Sql = Sql & " IsNull(k.ItemDescription,'') As RecdItemDesc,"
    Sql = Sql & " isnull(m.SentItemCode,'')As SentItemCode, IsNull(g.ItemDescription,'') As SentItemDesc,"
    Sql = Sql & " IsNull(i.SDesc,'') As RecdUOM,"
    Sql = Sql & " IsNull(h.SDesc,'') As SentUOM, m.RatePer, m.AssRate As AssRate, m.ItemGroupKey"
    Sql = Sql & " From SAL_JobWorkOrder_Hdr a"
    Sql = Sql & " Full Outer Join SAL_JobWorkOrder_Dtl m"
  End If
  Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
  Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
  Sql = Sql & " and a.Finyear=m.Finyear"

  Sql = Sql & " Left Outer Join FAS_Party_Master e"
  Sql = Sql & " On a.CustomerCode=e.PartyCode and a.Branchid=e.Branchid"
  Sql = Sql & " and a.Compid=e.Compid"
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join PartMast f"
End If
  If Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " On m.SentItemCode=f.StageCode and a.Branchid=f.BranchID and a.Compid=f.CompID"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER g On m.SentItemCode=g.ItemCode and a.Branchid=g.Branchid and a.Compid=g.Compid  and g.status=1"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER k On m.ReceivedItemCode=k.ItemCode and a.Branchid=k.Branchid and a.Compid=k.Compid and k.status=1"
    Sql = Sql & " Left Outer Join INV_UOM_Master h On g.UOM=h.Code and a.Branchid=h.Branchid and a.Compid=h.CompId and h.Status=1"
    Sql = Sql & " Left Outer Join INV_UOM_Master i On k.UOM=i.Code and a.Branchid=i.Branchid and a.Compid=i.CompId and i.Status=1"
  End If
  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER g On m.ItemCode=g.ItemCode"
    Sql = Sql & " and a.Branchid=g.Branchid and a.Compid=g.Compid and g.status=1"
    Sql = Sql & " Left Outer Join INV_UOM_Master h"
    Sql = Sql & " On g.UOM=h.Code and a.Branchid=h.Branchid and a.Compid=h.CompId"
    Sql = Sql & " and h.Status=1"
  End If
  Sql = Sql & " Left Outer Join dbo.fn_SEC_Users  ('" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "')  u"
  Sql = Sql & " On a.OrderAuthorizationBy=u.UserID and a.Branchid=u.BranchID"
  Sql = Sql & " and a.Compid=u.CompId and u.ActiveStatus=1"
  Sql = Sql & " Left Outer Join Sec_UserGroups v"
  Sql = Sql & " On u.UserGroupID=v.UserGroupID and a.Branchid=v.Branchid"
  Sql = Sql & " and a.Compid=v.CompID"

  Sql = Sql & " Where a.OrderDate Between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTodate.Value, "yyyy-MM-dd") & "'"

  If optCustSelective Then
'    If fpsList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": Exit Function
    
      Sql = Sql & " and a.CustomerCode in "
      Sql = Sql & Spread_Get_SelectionList(fpsList, 2)
    
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = IIf(mDocSeq(3) = 3, 2, 1)
'        Sql = Sql + "'" + Trim$(.Text) + "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)'    Sql = Sql + ")"
  End If

  Sql = Sql & " and a.CreatedDate=(Select Max(m.CreatedDate) As CreatedDate"

  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr m"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr m"
End If
  Sql = Sql & " Where OrderType = A.OrderType And OrderNo = A.OrderNo"
  Sql = Sql & " and OrderDate=a.OrderDate and Branchid=a.Branchid and Compid=a.Compid and Finyear=a.Finyear)"

  Sql = Sql & " and a.Branchid='" & gBranchID & "'"
  Sql = Sql & " and a.Compid ='" & gCompID & "' and a.Finyear='" & gFinyear & "'"

  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Order by a.CustomerCode, m.ItemCode"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Order by a.CustomerCode, m.SentItemCode"
End If

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If rs.RecordCount <= 0 Then
    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If
   
  If cmbType.ListIndex = 0 Then       'Spread
  
    frmPrint.Show
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
    mCustNm = ""
    
    Header
    
    j = j + 1
GridCount = GridCount + 1
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
                        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("OrderType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrderType")) & Trim$(rs("OrderNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(mCustNm) <> Trim$(rs("CustomerName")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("CustomerName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(rs("Qty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("GoodsValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("RefNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrdStatus"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else  'Job Work
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("GoodsValue"), "0.00")
frmPrint.sprPrint.ColWidth(10) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("RefNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 11

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrdStatus"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("OrderType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("OrderDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemGroupKey"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RatePer"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 11

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        mEnqType = Trim$(rs("OrderType"))
        mCustNm = Trim$(rs("CustomerName"))
        rs.MoveNext
      Next i
    
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    Screen.MousePointer = 1
    totalpages = iPageNo
     
  ElseIf cmbType.ListIndex = 1 Then   'File
  
    pheight = 71
    pPaper = vbPRPSFanfoldUS
    Orient = 1
    iPageWidth = 132
    iLineNo = 1
    fOpenPort Me
  
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    gStrReportHeading = "SALES Order (LISTING) BETWEEN " & mFromDate & " AND " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("OrderNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("OrderDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("Qty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
      Gtot(1) = Gtot(1) + Val(rs("GoodsValue"))
      Gtot(2) = Gtot(2) + Val(rs("AssValue"))
      i = i + 1
      sSendToPort ""
      rs.MoveNext
    Loop
    
    sSendToPort Space(10) & "Goods Value Grand Total      : " & Gtot(1) & Space(10) & "Assessable Value Grand Total : " & Gtot(2)

'    sSendToPort Space(10) & "Assessable Value Grand Total : " & GTot(2)
    Call EndofReport(iPageWidth, i - 1)
  End If
  
  rs.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function sPrintcurrent_RevisionSummary() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim mRev(2) As String
  Dim mEnqType As String
  Dim mCustNm As String
  Dim Ln As Integer
  Dim i As Integer
  Dim k As Integer

  Try
  sPrintcurrent_RevisionSummary = False

  Call RecordInsertion  'Process

  Sql = "Select Distinct a.OrderType, a.OrderNo, a.OrderDate,"
  Sql = Sql & " IsNull(b.PartNo,'') As PartNo,"
   Sql = Sql & " IsNull(a.RevisionNo,'') As RevisionNo,"
  Sql = Sql & " IsNull(a.CreatedDate,'') As RevisionDate,"
  Sql = Sql & " IsNull(a.CustomerCode,'') As CustomerCode,"
  Sql = Sql & " IsNull(c.PartyName,'') As CustomerName,"
  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " a.ItemCode, IsNull(a.ItemDescription,'') As ItemDescription,"
    Sql = Sql & " IsNull(d.SDesc,'') As UOM, a.OrderQty As Qty, a.Rate,"
    Sql = Sql & " IsNull(a.AssessableRate,'') As AssRate, a.RecStatus,"
    Sql = Sql & " "
  Else  'Jobwork
    Sql = Sql & " a.ReceivedItemCode As RecdItemCode,"
    Sql = Sql & " IsNull(r.ItemDescription,'') As RecdItemDesc,"
    Sql = Sql & " IsNull(j.SDesc,'') As RecdUOM, a.AssRate,"
    Sql = Sql & " a.SentItemCode,  IsNull(s.ItemDescription,'') As SentItemDesc,"
    Sql = Sql & " IsNull(d.SDesc,'') As SentUOM, a.OrderQty As Qty, a.RatePer, a.Rate,a.ItemGroupKey,"
  End If
  Sql = Sql & " IsNull(a.Remarks,'') As Remarks,"
  Sql = Sql & " IsNull(a.ReferenceNo,'') As RefNo,"
  Sql = Sql & " IsNull(a.ReferenceDate,'') As RefDate, IsNull(RecStatus,'') As RecStatus"
  Sql = Sql & " From SAL_SalseDocument_Tmp a"     'a.TargetDate,
  Sql = Sql & " Left Outer Join PartMast b"

  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " On a.ItemCode=b.StageCode "
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " On a.SentItemCode=b.StageCode "
End If

  Sql = Sql & " and a.Branchid=b.BranchID and a.Compid=b.CompID"

  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join INV_UOM_MASTER d"
End If
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " On a.UOM=d.Code and a.Branchid=d.Branchid and a.Compid=d.Compid"
End If
  
  Sql = Sql & " Left Outer Join Material_Master r"
  
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " On a.ItemCode=r.ItemCode"
End If
  If Trim$(DocFlg) = "Job work" Then
 Sql = Sql & " On a.ReceivedItemCode=r.ItemCode"
End If
  
  Sql = Sql & " and a.Branchid=r.BranchID and a.Compid=r.CompID"
  Sql = Sql & " Left Outer Join Material_Master s"
  
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " On a.ItemCode=s.ItemCode"
End If
  If Trim$(DocFlg) = "Job work" Then
 Sql = Sql & " On a.SentItemCode=s.ItemCode"
End If
  
  Sql = Sql & " and a.Branchid=s.BranchID and a.Compid=s.CompID"
  
   If Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " Left Outer Join Material_Master i"
    Sql = Sql & " On a.ReceivedItemCode=i.ItemCode and a.Branchid=i.Branchid"
    Sql = Sql & " and a.Compid=i.Compid"
    Sql = Sql & " Left Outer Join INV_UOM_MASTER j"
    Sql = Sql & " On i.UOM=j.Code and a.Branchid=j.Branchid and a.Compid=j.Compid"
    Sql = Sql & " Left Outer Join Material_Master k"
    Sql = Sql & " On a.SentItemCode=k.ItemCode and a.Branchid=k.Branchid"
    Sql = Sql & " and a.Compid=k.Compid"
    Sql = Sql & " Left Outer Join INV_UOM_MASTER d"
    Sql = Sql & " On k.UOM=d.Code and a.Branchid=d.Branchid and a.Compid=d.Compid"
  End If
  
  Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "'"
  Sql = Sql & " and a.OrderNo='" & Trim$(txtNo) & "'"

  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"

  Sql = Sql & " Order by x.RevisionNo, x.RevisionDate"

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If rs.RecordCount <= 0 Then
    'Temporary Table Droping
    Sql = "Drop Table SAL_SalseDocument_Tmp"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If

  If cmbType.ListIndex = 0 Then       'Spread
  
    frmPrint.Show
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
    mCustNm = ""
    
    Header
    j = j + 1
GridCount = GridCount + 1
RowProcess
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
                        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("OrderType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText 1, j, Trim$(rs("OrderType")) & Trim$(rs("OrderNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(mCustNm) <> Trim$(rs("CustomerName")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("CustomerName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("UOM"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("RefNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecStatus"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else  'Job Work
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("RefNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecStatus"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("OrderType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("OrderDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemGroupKey"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RatePer"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        mEnqType = Trim$(rs("OrderType"))
        mCustNm = Trim$(rs("CustomerName"))
        rs.MoveNext
      Next i
    
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    Screen.MousePointer = 1
    totalpages = iPageNo
     
  ElseIf cmbType.ListIndex = 1 Then   'File
    'Skip
  End If

'Temporary Table Droping
  Sql = "Drop Table SAL_SalseDocument_Tmp"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub RecordInsertion()
  Dim Sql As String
  Dim mQry As String
  Dim i As Integer
  Dim j As Integer

  Try

'Create Table - SAL_SalseDocument_Tmp
  Sql = "Create Table SAL_SalseDocument_Tmp(OrderType varchar(3) Not Null, OrderNo int Not Null,"
  Sql = Sql & " OrderDate datetime Not Null, OrderSource char(1) Not Null, CustomerCode int Not Null,"
  Sql = Sql & " ReferenceNo varchar(20) Not Null, ReferenceDate datetime Not Null,"
  Sql = Sql & " ValidFrom datetime Not Null, ValidTo datetime Not Null,"
  Sql = Sql & " OrderingType varchar(10) Not Null, Remarks varchar(50) Not Null,"
  Sql = Sql & " FormCode int Not Null, PackingCode int Not Null, DeliveryCode int Not Null,"
  Sql = Sql & " PaymentCode int Not Null, InsuranceCode int Not Null, ModeCode int Not Null,"
  Sql = Sql & " FreightCode int Not Null, FormulaCode int Not Null, TotalAmount float Not Null,"
  Sql = Sql & " PostingAc int Not Null, Status char(1) Not Null,"
  Sql = Sql & " OrderAuthorization char(1) Not Null, OrderAuthorizationDate datetime,"
  Sql = Sql & " OrderAuthorizationBy varchar(10) Not Null, PartNo varchar(20) Not Null,"
  Sql = Sql & " ScheduleType char(1) Not Null, OrderQty float Not Null,"
  Sql = Sql & " RevisionNo int Not Null, CreatedDate datetime Not Null,"
  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " RevisionLevel varchar(10) Not Null, GeneralDetails varchar(3000) Not Null,"
    Sql = Sql & " ItemCode varchar(20) Not Null, ItemDescription varchar(50) Not Null,"
    Sql = Sql & " UOM varchar(3) Not Null, Rate Money Not Null,"
    Sql = Sql & " AssessableRate float Not Null,"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " ItemGroupKey int Not Null, ReceivedItemCode varchar(20) Not Null,"
    Sql = Sql & " SentItemCode varchar(20) Not Null, RatePer float Not Null,"
    Sql = Sql & " Rate Money Not Null, AssRate money Not Null,"
  End If
  Sql = Sql & " Userid varchar(10) Not Null, Branchid varchar(3) Not Null,"
  Sql = Sql & " Compid varchar(3) Not Null, Finyear varchar(20) Not Null,"
  Sql = Sql & " RecStatus varchar(20) Not Null)"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

'For Select Fields
  mQry = "a.OrderType, a.OrderNo, a.OrderDate, a.OrderSource, a.CustomerCode,"
  mQry = mQry & " IsNull(a.ReferenceNo,'') As RefNo, a.ReferenceDate As RefDate, a.ValidFrom,"
  mQry = mQry & " a.ValidTo, IsNull(a.OrderingType,'') As OrderingType,"
  mQry = mQry & " IsNull(a.Remarks,'') As Remarks, IsNull(a.FormCode,'') As FormCode,"
  mQry = mQry & " IsNull(a.PackingCode,'') As PackingCode, IsNull(a.DeliveryCode,'') As DeliveryCode,"
  mQry = mQry & " IsNull(a.PaymentCode,'') As PaymentCode, IsNull(a.InsuranceCode,'') As InsuranceCode,"
  mQry = mQry & " IsNull(a.ModeCode,'') As ModeCode, IsNull(a.FreightCode,'') As FreightCode,"
  mQry = mQry & " IsNull(a.FormulaCode,'') As FormulaCode, IsNull(a.TotalAmount,'') As TotalAmount,"
  mQry = mQry & " IsNull(a.PostingAc,'') As PostingAc, IsNull(a.Status,'') As Status,"
  mQry = mQry & " IsNull(a.OrderAuthorization,'') As OrdAuth,"
  mQry = mQry & " a.OrderAuthorizationDate As OrdAuthDate,"
  mQry = mQry & " IsNull(a.OrderAuthorizationBy,'') As OrdAuthBy,"
  If Trim$(DocFlg) = "Sales" Then
 mQry = mQry & " IsNull(m.PartNo,'') As PartNo,"
End If
  If Trim$(DocFlg) = "Job Work" Then
 mQry = mQry & " '' As PartNo,"
End If
  mQry = mQry & " IsNull(m.ScheduleType,'') As ScheduleType, m.OrderQty,"
  mQry = mQry & " IsNull(m.RevisionNo,'') As RevNo, m.CreatedDate As RevDate,"

  If Trim$(DocFlg) = "Sales" Then
    mQry = mQry & " IsNull(a.RevisionLevel,'') As RevisionLevel,"
    mQry = mQry & " IsNull(a.GeneralDetails,'') As GeneralDetails,"
    mQry = mQry & " m.ItemCode, IsNull(m.ItemDescription,'') As ItemDescription,"
    mQry = mQry & " m.UOM, m.Rate, m.AssessableRate As AssRate,"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    mQry = mQry & " IsNull(m.ItemGroupKey,'') As ItemGroupKey,"
    mQry = mQry & " m.ReceivedItemCode As RecdItemCode,"
    mQry = mQry & " m.SentItemCode, m.RatePer, m.Rate, m.AssRate,"
  End If
  mQry = mQry & " a.Userid, a.Branchid, a.Compid, a.Finyear"

'Insert - From SAL_SalesPlan To SAL_SalesPlan_Tmp
  i = 0
  If i = 0 Then
  'All - INITIAL
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select " & Trim$(mQry) & ", 'INITIAL STATE' As RecStatus"

    If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr a"
End If
    If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Full Outer Join (Select * From SAL_Order_Dtl n"
End If
    If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr a"
End If
    If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Full Outer Join (Select * From SAL_JobWorkOrder_Dtl n"
End If

    Sql = Sql & " Where n.OrderType='" & Trim$(txttype) & "' and n.OrderNo='" & Trim$(txtNo) & "'"
    Sql = Sql & " and n.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and n.Revisionno=0"

'    If optAll Then: Sql = Sql & " and n.Revisionno=0"

'    If optSelective Then: Sql = Sql & " and n.Revisionno=" & Val(lblRevNo.Tag) & ""

    Sql = Sql & " ) m"
    Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo and a.OrderDate=m.OrderDate"
    Sql = Sql & " and a.Branchid=m.Branchid and a.Compid=m.Compid and a.Finyear=m.Finyear"

    Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
    Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and a.CreatedDate=m.CreatedDate"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " Order by a.CustomerCode, m.CreatedDate"   ', m.ItemCode"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If

  j = 0
'  If optAll Then

    For i = 0 To cmbRevNo.ListCount - 1
      j = i + 1

    'All - NEW
      Sql = "Insert Into SAL_SalseDocument_Tmp"
      Sql = Sql & " Select " & Trim$(mQry) & ", 'NEW' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr a"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr a"
End If
      Sql = Sql & " Full Outer Join ("
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Select * From SAL_Order_Dtl n"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl n"
End If

      Sql = Sql & " Where n.OrderType='" & Trim$(txttype) & "' and n.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and n.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and n.RevisionNo In(" & j & ")"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and n.ItemCode Not In("
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and n.SentItemCode Not In("
End If
      
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & "  Select o.ItemCode From ( Select * From SAL_Order_Dtl o"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & "  Select o.SentItemCode From ( Select * From SAL_JobWorkOrder_Dtl o"
End If

      Sql = Sql & " Where o.OrderType='" & Trim$(txttype) & "' and o.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and o.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and o.RevisionNo In(" & i & ")"
      Sql = Sql & " and o.Branchid='" & gBranchID & "' and o.Compid ='" & gCompID & "'"
'      Sql = Sql & " and o.Finyear='" & gFinyear & "'"
      Sql = Sql & " ) o"

      Sql = Sql & " Inner Join ("
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Select * From SAL_Order_Dtl p"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl p"
End If

      Sql = Sql & " Where p.OrderType='" & Trim$(txttype) & "' and p.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and p.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and p.RevisionNo In(" & j & ")"
      Sql = Sql & " and p.Branchid='" & gBranchID & "' and p.Compid ='" & gCompID & "'"
'      Sql = Sql & " and p.Finyear='" & gFinyear & "'"
      Sql = Sql & " ) p"
      Sql = Sql & " On o.OrderType=p.OrderType and o.OrderNo=p.OrderNo and o.OrderDate=p.OrderDate"
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and o.ItemCode=p.ItemCode"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and o.SentItemCode=p.SentItemCode"
End If
      Sql = Sql & " and o.Branchid=p.Branchid and o.Compid=p.Compid"
      Sql = Sql & " and o.Finyear=p.Finyear"
      Sql = Sql & " )"
      Sql = Sql & " ) m"
      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid "
'      Sql = Sql & " and a.Finyear=m.Finyear"

      Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.CreatedDate=m.CreatedDate"
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      Sql = Sql & " Order by a.CustomerCode, m.CreatedDate" ', m.ItemCode"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    'All - MODIFIED
      Sql = "Insert Into SAL_SalseDocument_Tmp"
      Sql = Sql & " Select x.* From ("
      Sql = Sql & " Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
End If

      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
      Sql = Sql & " and a.Finyear=m.Finyear"
      Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and m.RevisionNo In(" & i & ") and a.CreatedDate=m.CreatedDate"
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      Sql = Sql & " ) x"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Where x.ItemCode In(Select m.ItemCode From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Where x.SentItemCode In(Select m.SentItemCode From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
End If

      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
      Sql = Sql & " and a.Finyear=m.Finyear"

      Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"

      Sql = Sql & " and m.RevisionNo In(" & j & ")  and a.CreatedDate=m.CreatedDate"
      Sql = Sql & " and x.CustomerCode=a.CustomerCode"
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and x.ItemCode=m.ItemCode"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and x.SentItemCode=m.SentItemCode"
End If
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      Sql = Sql & " )"
      Sql = Sql & " and Not Exists( Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
End If

      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
'      Sql = Sql & " and a.Finyear=m.Finyear"
      Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and m.RevisionNo In(" & j & ")  and a.CreatedDate=m.CreatedDate"

      Sql = Sql & " and x.OrderSource=a.OrderSource and x.RefNo=a.ReferenceNo and x.RefDate=a.ReferenceDate"
      Sql = Sql & " and x.ValidFrom=a.ValidFrom and x.ValidTo=a.ValidTo and x.OrderingType=a.OrderingType and x.Remarks=a.Remarks"
      Sql = Sql & " and x.FormCode=a.FormCode"
      Sql = Sql & " and x.PackingCode=a.PackingCode and x.DeliveryCode=a.DeliveryCode and x.PaymentCode=a.PaymentCode"
      Sql = Sql & " and x.InsuranceCode=a.InsuranceCode and x.ModeCode=a.ModeCode and x.FreightCode=a.FreightCode"
      Sql = Sql & " and x.FormulaCode=a.FormulaCode and x.TotalAmount=a.TotalAmount and x.PostingAc=a.PostingAc"
      Sql = Sql & " and x.Status=a.Status and x.OrdAuth=a.OrderAuthorization and x.OrdAuthDate=a.OrderAuthorizationDate"
      Sql = Sql & " and x.OrdAuthBy=a.OrderAuthorizationBy" '
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and x.PartNo=m.PartNo"
End If
      Sql = Sql & " and x.ScheduleType=m.ScheduleType and m.OrderQty=m.OrderQty"
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and x.GeneralDetails=a.GeneralDetails"
End If
      Sql = Sql & " and x.Rate=m.Rate"
      
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and x.AssRate=m.AssessableRate and x.UOM=m.UOM"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and x.AssRate=m.AssRate"
End If
      
      Sql = Sql & " and x.Branchid='" & gBranchID & "' and x.Compid ='" & gCompID & "'"
'      Sql = Sql & " and x.Finyear='" & gFinyear & "'"
      Sql = Sql & " )"
      Sql = Sql & " and x.Branchid='" & gBranchID & "' and x.Compid ='" & gCompID & "'"
'      Sql = Sql & " and x.Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    'All - DELETED
      Sql = "Insert Into SAL_SalseDocument_Tmp"
      Sql = Sql & " Select " & Trim$(mQry) & ", 'DELETED' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Order_Hdr a"
End If
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Order_Dtl m"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkOrder_Hdr a"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl m"
End If

      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
      Sql = Sql & " and a.OrderDate=m.OrderDate and a.CreatedDate=m.CreatedDate"
      Sql = Sql & " and a.Branchid=m.Branchid and a.Compid=m.Compid"
'      Sql = Sql & " and a.Finyear=m.Finyear"
      Sql = Sql & " Where a.OrderType='" & Trim$(txttype) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and m.RevisionNo In(" & i & ")"
      Sql = Sql & " "

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and m.ItemCode Not In( Select o.ItemCode From SAL_Order_Dtl o"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and m.SentItemCode Not In( Select o.SentItemCode From SAL_JobWorkOrder_Dtl o"
End If

      Sql = Sql & " Where a.OrderType=o.OrderType and a.OrderNo=o.OrderNo"
      Sql = Sql & " and a.OrderDate=o.OrderDate and o.RevisionNo In(" & j & ")"
      Sql = Sql & " and a.CreatedDate=o.CreatedDate"
      Sql = Sql & " and a.Branchid=o.Branchid and a.Compid=o.Compid"
'      Sql = Sql & " and a.Finyear=o.Finyear"
      Sql = Sql & " )"
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

      If j = (cmbRevNo.ListCount - 1) Then
        GoTo ExitHere
      End If

    Next i

'  ElseIf optSelective Then

'    j = IIf(Trim$(lblRevNo.Tag) = "", Trim$(lblRevNo(1).Tag), Trim$(lblRevNo.Tag))
'    For i = j To Val(lblRevNo(1).Tag)
'
'    'Selective - NEW
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select " & Trim$(mQry) & ", 'NEW' As RecStatus"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a"
'      Sql = Sql & " Full Outer Join ("
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Dtl n"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl n"
'
'      Sql = Sql & " Where n.OrderType='" & Trim$(txtType) & "' and n.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and n.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and n.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " and n.ItemCode Not In("

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " and n.SentItemCode Not In("
'
'      Sql = Sql & " Select o.ItemCode From ("
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Dtl o"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl o"
'
'      Sql = Sql & " Where o.OrderType='" & Trim$(txtType) & "' and o.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and o.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and o.RevisionNo In(" & i & ")"
'      Sql = Sql & " and o.Branchid='" & gBranchID & "' and o.Compid ='" & gCompID & "'"
''      Sql = Sql & " and o.Finyear='" & gFinyear & "'"
'      Sql = Sql & " ) o"
'
'      Sql = Sql & " Inner Join ("
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Dtl p"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl p"
'
'      Sql = Sql & " Where p.OrderType='" & Trim$(txtType) & "' and p.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and p.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and p.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
'      Sql = Sql & " and p.Branchid='" & gBranchID & "' and p.Compid ='" & gCompID & "'"
''      Sql = Sql & " and p.Finyear='" & gFinyear & "'"
'      Sql = Sql & " ) p"
'      Sql = Sql & " On o.OrderType=p.OrderType and o.OrderNo=p.OrderNo and o.OrderDate=p.OrderDate"
'      Sql = Sql & " and o.ItemCode=p.ItemCode and o.Branchid=p.Branchid and o.Compid=p.Compid"
'      Sql = Sql & " and o.Finyear=p.Finyear"
'      Sql = Sql & " )"
'      Sql = Sql & " ) m"
'      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
'      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid "
''      Sql = Sql & " and a.Finyear=m.Finyear"
'
'      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and a.CreatedDate=m.CreatedDate"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      Sql = Sql & " Order by a.CustomerCode, m.CreatedDate, m.ItemCode"
'      db.Execute Sql
'
''      Sql = "Insert Into SAL_SalseDocument_Tmp"
''      Sql = Sql & " Select  " & Trim$(mQry) & ",'NEW' As RecStatus"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a"

''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Left Outer Join SAL_Order_Dtl b"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl b"
''
''      Sql = Sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo"
''      Sql = Sql & " and a.OrderDate=b.OrderDate and a.RevisionNo=b.RevisionNo"
''      Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
'''      Sql = Sql & " and a.Finyear=b.Finyear"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Where b.ItemCode Not In(Select c.ItemCode"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Where b.SentItemCode Not In(Select c.SentItemCode"
''
''      Sql = Sql & " From ("
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Dtl c"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl c"
''
''      Sql = Sql & " Where c.OrderType='" & Trim$(txtType) & "' and c.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and c.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''      Sql = Sql & " and c.RevisionNo In(" & i & ")"
''      Sql = Sql & " and c.Branchid='" & gBranchID & "' and c.Compid ='" & gCompID & "'"
''      Sql = Sql & " and c.Finyear='" & gFinyear & "'"
''      Sql = Sql & " ) c"
''      Sql = Sql & " Inner Join ("
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Dtl d"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Dtl d"
''
''      Sql = Sql & " Where d.OrderType='" & Trim$(txtType) & "' and and d.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and d.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''      Sql = Sql & " and d.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
''      Sql = Sql & " and d.Branchid='" & gBranchID & "' and d.Compid ='" & gCompID & "'"
''      Sql = Sql & " and d.Finyear='" & gFinyear & "'"
''      Sql = Sql & " ) d"
''      Sql = Sql & " On c.OrderType=d.OrderType and c.OrderNo=d.OrderNo and c.OrderDate=d.OrderDate"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " and c.ItemCode=d.ItemCode"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " and c.SentItemCode=d.SentItemCode"
''
''      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid"
''      Sql = Sql & " and c.Finyear=d.Finyear"
''      Sql = Sql & " )"
''      Sql = Sql & " and a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''      Sql = Sql & " and a.Revisionno In(" & Val(lblRevNo(1).Tag) & ")"
''      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
''      Sql = Sql & " db.Execute Sql"
'
'    'Selective - MODIFIED
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select x.* From ("
'      Sql = Sql & " Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
'
'      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
'      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
'      Sql = Sql & " and a.Finyear=m.Finyear"
'      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and m.RevisionNo In(" & i & ") and a.CreatedDate=m.CreatedDate"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      Sql = Sql & " ) x"
'      Sql = Sql & " Where x.ItemCode In(Select m.ItemCode"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
'
'      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
'      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
'      Sql = Sql & " and a.Finyear=m.Finyear"
'
'      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'
'      Sql = Sql & " and m.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")  and a.CreatedDate=m.CreatedDate"
'      Sql = Sql & " and x.CustomerCode=a.CustomerCode and x.ItemCode=m.ItemCode "
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      Sql = Sql & " )"
'      Sql = Sql & " and Not Exists( Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a Left Outer Join SAL_Order_Dtl m"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a Left Outer Join SAL_JobWorkOrder_Dtl m"
'
'      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
'      Sql = Sql & " and a.OrderDate=m.OrderDate and a.Branchid=m.Branchid and a.Compid=m.Compid"
''      Sql = Sql & " and a.Finyear=m.Finyear"
'      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and m.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")  and a.CreatedDate=m.CreatedDate"
'
'      Sql = Sql & " and x.OrderSource=a.OrderSource and x.RefNo=a.ReferenceNo and x.RefDate=a.ReferenceDate"
'      Sql = Sql & " and x.ValidFrom=a.ValidFrom and x.ValidTo=a.ValidTo and x.OrderingType=a.OrderingType and x.Remarks=a.Remarks"
'      Sql = Sql & " and x.FormCode=a.FormCode"
'      Sql = Sql & " and x.PackingCode=a.PackingCode and x.DeliveryCode=a.DeliveryCode and x.PaymentCode=a.PaymentCode"
'      Sql = Sql & " and x.InsuranceCode=a.InsuranceCode and x.ModeCode=a.ModeCode and x.FreightCode=a.FreightCode"
'      Sql = Sql & " and x.FormulaCode=a.FormulaCode and x.TotalAmount=a.TotalAmount and x.PostingAc=a.PostingAc"
'      Sql = Sql & " and x.Status=a.Status and x.OrdAuth=a.OrderAuthorization and x.OrdAuthDate=a.OrderAuthorizationDate"
'      Sql = Sql & " and x.OrdAuthBy=a.OrderAuthorizationBy and x.PartNo=m.PartNo"
'      Sql = Sql & " and x.ScheduleType=m.ScheduleType and m.OrderQty=m.OrderQty and x.GeneralDetails=a.GeneralDetails"
'      Sql = Sql & " and x.UOM=m.UOM and x.Rate=m.Rate and x.AssRate=m.AssessableRate"
'      Sql = Sql & " and x.Branchid='" & gBranchID & "' and x.Compid ='" & gCompID & "'"
''      Sql = Sql & " and x.Finyear='" & gFinyear & "'"
'      Sql = Sql & " )"
'      Sql = Sql & " and x.Branchid='" & gBranchID & "' and x.Compid ='" & gCompID & "'"
''      Sql = Sql & " and x.Finyear='" & gFinyear & "'"
'      db.Execute Sql
'
''      Sql = "Insert Into SAL_SalseDocument_Tmp"
''      Sql = Sql & " Select  " & Trim$(mQry) & ",'MODIFIED' As RecStatus"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a"

''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Left Outer Join SAL_Order_Dtl b"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl b"
''
''      Sql = Sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.OrderDate=b.OrderDate"
''      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
''      Sql = Sql & " and a.Finyear=b.Finyear"
''      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''      Sql = Sql & " and a.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
''      Sql = Sql & " and Exists ("
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select * From SAL_Order_Hdr c"

''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Left Outer Join SAL_Order_Dtl d"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select * From SAL_JobWorkOrder_Hdr c"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl d"
''
''      Sql = Sql & " On c.OrderType=d.OrderType and c.OrderNo=d.OrderNo and c.OrderDate=d.OrderDate"
''      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid and c.Finyear=d.Finyear"
''      Sql = Sql & " Where c.OrderType='" & Trim$(txtType) & "' and c.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and c.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''
''      If Trim$(DocFlg) = "Sales"  Then
''        Sql = Sql & " and b.ItemCode=d.ItemCode and c.RevisionNo In(" & i & ")"
''        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
''        Sql = Sql & " or b.UOM<>d.UOM or b.Qty<>d.Qty or b.Volume<>d.Volume"
''        Sql = Sql & " or b.AssessableRate<>d.AssessableRate"
''      ElseIf Trim$(DocFlg) = "Job Work"  Then
''        Sql = Sql & " and b.SentItemCode=d.SentItemCode and c.RevisionNo In(" & i & ")"
''        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
''        Sql = Sql & " or b.ItemGroupKey<>d.ItemGroupKey or b.ReceivedItemCode<>d.ReceivedItemCode"
''        Sql = Sql & " or b.RecdUOM<>d.RecdUOM or b.SentUOM<>d.SentUOM"
''        Sql = Sql & " or b.SendQty<>d.SendQty or b.RatePer<>d.RatePer or b.AssRate<>d.AssRate"
''        Sql = Sql & " or b.RelationType<>d.RelationType or b.isActive<>d.isActive"
''      End If
''
''      Sql = Sql & " or b.VolumePer<>d.VolumePer or b.Rate<>d.Rate"
''      Sql = Sql & " or IsNull(b.TargetDate,'')<>IsNull(d.TargetDate,'')"
''      Sql = Sql & " )"
''      Sql = Sql & " )"
''      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'
'    'Selective - DELETED
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select " & Trim$(mQry) & ", 'DELETED' As RecStatus"
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a"

'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Left Outer Join SAL_Order_Dtl m"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl m"
'
'      Sql = Sql & " On a.OrderType=m.OrderType and a.OrderNo=m.OrderNo"
'      Sql = Sql & " and a.OrderDate=m.OrderDate and a.CreatedDate=m.CreatedDate"
'      Sql = Sql & " and a.Branchid=m.Branchid and a.Compid=m.Compid"
''      Sql = Sql & " and a.Finyear=m.Finyear"
'      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and m.RevisionNo In(" & i & ")"
'      Sql = Sql & " and m.ItemCode Not In("
'
'      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Select o.ItemCode From SAL_Order_Dtl o"

'      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Select o.ItemCode From SAL_JobWorkOrder_Dtl o"
'
'      Sql = Sql & " Where a.OrderType=o.OrderType and a.OrderNo=o.OrderNo"
'      Sql = Sql & " and a.OrderDate=o.OrderDate and o.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
'      Sql = Sql & " and a.CreatedDate=o.CreatedDate"
'      Sql = Sql & " and a.Branchid=o.Branchid and a.Compid=o.Compid"
''      Sql = Sql & " and a.Finyear=o.Finyear"
'      Sql = Sql & " )"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      db.Execute Sql
'
''      Sql = "Insert Into SAL_SalseDocument_Tmp"
''      Sql = Sql & " Select  " & Trim$(mQry) & ", 'DELETED' As RecStatus"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " From SAL_Order_Hdr a"

''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " Left Outer Join SAL_Order_Dtl b"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " From SAL_JobWorkOrder_Hdr a"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " Left Outer Join SAL_JobWorkOrder_Dtl b"
''
''      Sql = Sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo"
''      Sql = Sql & " and a.OrderDate=b.OrderDate"
''      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
''      Sql = Sql & " and a.Finyear=b.Finyear"
''      Sql = Sql & " Where a.OrderType='" & Trim$(txtType) & "' and a.OrderNo='" & Trim$(txtNo) & "'"
''      Sql = Sql & " and a.OrderDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
''      Sql = Sql & " and a.RevisionNo In(" & i & ")"
''
''      If Trim$(DocFlg) = "Sales"  Then: Sql = Sql & " and b.ItemCode Not In( Select c.ItemCode From SAL_Order_Dtl c"

''      If Trim$(DocFlg) = "Job Work"  Then: Sql = Sql & " and b.SentItemCode Not In( Select c.SentItemCode From SAL_JobWorkOrder_Dtl c"
''
''      Sql = Sql & " Where a.OrderType = c.OrderType and a.OrderNo = c.OrderNo"
''      Sql = Sql & " and a.OrderDate = c.OrderDate"
''      Sql = Sql & " and c.RevisionNo In(" & Val(lblRevNo(1).Tag) & ") and a.Branchid=c.Branchid"
''      Sql = Sql & " and a.Compid=b.Compid and a.Finyear=b.Finyear"
''      Sql = Sql & " )"
''      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
''      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
''      db.Execute Sql
'
'    Next i
'  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Header()
  Dim mstr As String
  Dim CurrDate As Date
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
  
    frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    If RptFlg = "Audit" Then
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.MaxCols = 4
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
      Else
        frmPrint.sprPrint.MaxCols = 5
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
      End If
    ElseIf RptFlg = "RevisionSummary" Then
      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
    ElseIf RptFlg = "Listing" Then
      frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j
    ElseIf RptFlg = "View" Or RptFlg = "Print" Or RptFlg = "Obsolute" Then
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
    Else
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
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
      If RptFlg <> "View" And RptFlg <> "Print" And RptFlg <> "Obsolute" Then
        frmPrint.sprPrint.AddCellSpan 1, j, 4, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If RptFlg = "RevisionSummary" Then
          frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
        ElseIf RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j
        Else
          If Trim$(DocFlg) = "Sales" Then
            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
          Else
            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
          End If
        End If
        frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      Else
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
        End If
      End If
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If
    
    If iPageNo = 1 Then
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, gCompName
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

      j = j + 2
GridCount = GridCount + 2
    End If
        
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText 1, j, IIf(Trim$(DocFlg) <> "", Trim$(DocFlg), "") & " Order " & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", "- Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    If Trim$(RptFlg) = "Listing" Or Trim$(RptFlg) = "RevisionSummary" Then 'Listing
    'Top
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ord.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "UOM"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 4
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Status"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
        If Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Revision Status"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
      Else  'Job Work
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ReUOM"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 11

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Status"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
        If Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Revision Status"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
      End If
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
    'Bottom
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ord.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
        If Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
      Else  'Job Work
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemGroupKey"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SeUOM"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate Per"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 11

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
        If Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
      End If
      
      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
          
    ElseIf Trim$(RptFlg) = "Audit" Then 'Audit
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ord. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemCode"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        k = 0
      Else
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ReceivedItemCode"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentItemCode"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        k = 1
      End If
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User ID"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ord.Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemDescription"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      Else
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ReceivedDescription"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentDescription"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      End If
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User Name"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4 + k

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Transaction Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    End If
        
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
      If j > 35 Then
        k = 1
      Else
        k = 0
      End If
  End If
End Sub

Private Function sPrintcurrent_Audit() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim SRow As Long
  Dim CurrDate As Date
  Dim Gtot(2) As Double
  Dim tmp As String
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
    
  Try
  sPrintcurrent_Audit = False
  CurrDate = ServerDate
  
  Sql = "Select Distinct a.OrderType, a.OrderNo, a.OrderDate,"
  Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As Userid,"
  Sql = Sql & " Case When a.Userid='A0007' Then 'ADMIN' Else IsNull(c.UserName,'') End As UserName,"
  Sql = Sql & " IsNull(a.Remarks ,'') As Remarks"
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & ",a.ItemCode,d.Itemdescription From dbo.fn_SAL_SalesOrder( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " ,a.transtamp, a.ReceivedItemCode, a.SentItemCode"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " ,a.ItemGroupKey,d.Itemdescription as RecdDesc,d1.Itemdescription as SentDesc From dbo.fn_SAL_JobworkOrder( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " left outer join INV_Material_Master d on a.ReceivedItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " left outer join INV_Material_Master d1 on a.SentItemCode = d1.ItemCode and a.Branchid = d1.Branchid and a.Compid = d1.Compid"
End If
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " left outer join INV_Material_Master d on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid"
End If
  Sql = Sql & " Left Outer Join dbo.fn_SEC_Users ( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
  Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
  Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  Sql = Sql & " Where a.OrderDate between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTodate.Value, "yyyy-MM-dd") & "'"

  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.OrderNo in("
      SRow = 0
      For i = 1 To fpsList.SelectionCount
        SRow = fpsList.GetMultiSelItem(SRow)
        fpsList.Row = SRow
fpsList.Col = 2
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.OrderType, a.OrderNo, a.OrderDate"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.RecordCount <= 0 Then
    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  If cmbType.ListIndex = 0 Then       'Spread
  
    frmPrint.Show
    If Trim$(DocFlg) = "Sales" Then
      pheight = 54
Orient = 1
    Else
      pheight = 35
Orient = 2
    End If
    j = 1
GridCount = 1
k = 0
iPageNo = 1
    mCustNm = ""
        
    Header
    
    j = j + 1
GridCount = GridCount + 1
RowProcess
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
    
    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
          
        If Trim$(mCustNm) <> Trim$(rs("OrderType")) & Trim$(rs("OrderNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrderType")) & Trim$(rs("OrderNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemCode"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          k = 0
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ReceivedItemCode")) & " [" & Trim$(rs("ItemGroupKey")) & "]"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          k = 1
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Userid")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4 + k

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        If Trim$(mCustNm) <> Trim$(rs("OrderType")) & Trim$(rs("OrderNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrderDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Itemdescription"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdDesc"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("UserName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4+k

'        .SetText .Col, j, Trim$(rs("Transtamp")): .TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        mCustNm = Trim$(rs("OrderType")) & Trim$(rs("OrderNo"))
        rs.MoveNext
      Next i
    
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    Screen.MousePointer = 1
    totalpages = iPageNo
  
  ElseIf cmbType.ListIndex = 1 Then   'File
  
    pheight = 71
    pPaper = vbPRPSFanfoldUS
    Orient = 1
    iPageWidth = 132
    iLineNo = 1
    fOpenPort Me
  
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    gStrReportHeading = "Delivery Challan Report between " & mFromDate & " and " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("OrderNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("OrderDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("DCQty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
      Gtot(1) = Gtot(1) + Val(rs("GoodsValue"))
      Gtot(2) = Gtot(2) + Val(rs("AssValue"))
      i = i + 1
      sSendToPort ""
      rs.MoveNext
    Loop
    
    sSendToPort Space(10) & "Goods Value Grand Total      : " & Gtot(1) & Space(10) & "Assessable Value Grand Total : " & Gtot(2)

'    sSendToPort Space(10) & "Assessable Value Grand Total : " & GTot(2)
    Call EndofReport(iPageWidth, i - 1)
  End If
  
  rs.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function
Private Function sPrintcurrent_SalesView() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  Dim s As Integer
  Dim GenSpec As String
  Dim PartNo As String
  Dim CurDate As Date
  
  Try
  
  CurDate = ServerDate
  Set snap = New AceADODB.Recordset
  Set rs = New AceADODB.Recordset
  Set rs1 = New AceADODB.Recordset
  Set rsDetails = New AceADODB.Recordset
  
  If cmbType.ListIndex = 1 Then
    Sql = "select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, c.PartyName, a.ReferenceNo, "
    Sql = Sql + " a.ReferenceDate,a.RevisionLevel, a.ItemCode, a.ItemDescription, a.PartNo, a.RevisionNo,"
    Sql = Sql + " b.sDesc, a.OrderQty, isNull(d.ItemSpecification,'') ItemSpecification, a.GeneralDetails"
    Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
    Sql = Sql + " left outer join Inv_UOM_Master b on a.UOM = b.Code and a.Branchid = b.Branchid"
    Sql = Sql + " and a.Compid = b.Compid left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
    Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join INV_Material_Master d "
    Sql = Sql + " on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid "
    Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
    Sql = Sql + " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
  
    If RSOpen(snap, Sql) = False Then Exit Function
    If snap.EOF Then
      sShowMessage "No Record(s) Found"
      sPrintcurrent_SalesView = False
      GoTo ExitHere
    End If
    
    pheight = 71
    pPaper = vbPRPSFanfoldStdGerman
    Orient = 1
    iPageWidth = 80
    iLineNo = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "Sales Order")
    sSendToPort ""
    sSendToPort LAlign(17, "SO No.          :") & Space(1) & LAlign(20, snap("OrderType") & snap("OrderNo")) & Space(12) & LAlign(18, "SO Date          : ") & Space(1) & LAlign(10, snap("OrderDate"))
    sSendToPort LAlign(17, "Customer PO No. :") & Space(1) & LAlign(20, snap("ReferenceNo")) & Space(12) & LAlign(18, "Customer PO Date :") & Space(1) & LAlign(10, snap("ReferenceDate"))
    sSendToPort LAlign(17, "Revision Level  :") & Space(1) & LAlign(20, snap("RevisionLevel"))
    sSendToPort LAlign(17, "Customer Name   :") & Space(1) & LAlign(50, snap("PartyName"))
    sPageHeader
    
    i = 1
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      GenSpec = Trim$(snap("GeneralDetails"))
      For i = 1 To snap.RecordCount
        'ProcessFlow
        If Trim$(gReportID) = "PMI" Then
          Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
          Sql = Sql + " left outer join dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b "
          Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
          Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
          If RSOpen(rs1, Sql) = False Then Exit Function
          If rs1.EOF Then
            Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
          End If
          'RawMaterials
          Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
          Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
          Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
          If RSOpen(rsDetails, Sql) = False Then Exit Function
        End If
                
        sSendToPort LAlign(20, snap("PartNo")) & Space(1) & LAlign(40, snap("ItemDescription")) & Space(1) & LAlign(3, snap("RevisionNo")) & Space(1) & LAlign(3, CheckNull(snap("sDesc"))) & Space(1) & RAlign(10, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty")))
        If Trim$(snap("ItemCode")) <> "" Then
 sSendToPort LAlign(20, snap("ItemCode"))
End If
'        If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'          WordSplit Trim$(snap("ItemSpecification")), 40
'          sSendToPort Space(21) & LAlign(40, WSplit(1))
'          For s = 2 To UBound(WSplit)
'            sSendToPort Space(21) & LAlign(40, WSplit(s))
'          Next s
'        End If
        sSendToPort ""
        snap.MoveNext
      Next i
    End If
    'Delivery Details
    Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
    Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
    Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
    Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
    If RSOpen(rs, Sql) = False Then Exit Function
    If Not rs.EOF Then
      PartNo = ""
      sSendToPort String(iPageWidth, "-")
      sSendToPort LAlign(17, "Delivery Details :") & Space(1) & LAlign(30, "Part No") & Space(1) & LAlign(10, "Sch. Date") & Space(1) & RAlign(10, "Sch.Qty")
      rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        If PartNo <> rs("PartNo") Then
          If i > 1 Then
 sSendToPort ""
End If
          sSendToPort Space(18) & LAlign(30, rs("PartNo")) & Space(1) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        Else
          sSendToPort Space(49) & LAlign(10, rs("ScheduleDate")) & Space(1) & RAlign(10, rs("ScheduleQty"))
        End If
        PartNo = rs("PartNo")
        rs.MoveNext
      Next i
    End If
    sSendToPort String(iPageWidth, "-")
    sSendToPort ""
                         'Process Flow Print
    If Trim$(gReportID) = "PMI" Then
      If Not rs1.EOF Then
        sSendToPort LAlign(20, "Operation(s):")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(5, ("S.No")) & Space(2) & LAlign(50, "Description")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rs1.RecordCount
          sSendToPort LAlign(5, rs1("ProcessSequence")) & Space(2) & LAlign(50, rs1("ProcessFlow"))
          rs1.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
        sSendToPort ""
      End If
                        'Raw Material Print
      If Not rsDetails.EOF Then
        sSendToPort LAlign(25, "Rawmaterial Details:")
        sSendToPort String(iPageWidth, "-")
        sSendToPort LAlign(20, "Rawmaterial Code") & Space(1) & LAlign(40, "Description") & Space(1) & LAlign(5, "UOM") & Space(1) & RAlign(12, "Norms")
        sSendToPort String(iPageWidth, "-")
        For i = 1 To rsDetails.RecordCount
          sSendToPort LAlign(20, rsDetails("ChildCode")) & Space(1) & LAlign(40, rsDetails("ItemDescription")) & Space(1) & LAlign(5, rsDetails("PurchaseUOMDesc")) & Space(1) & RAlign(12, rsDetails("Qty"))
          rsDetails.MoveNext
        Next i
        sSendToPort String(iPageWidth, "-")
      End If
    rs1.Close
    rsDetails.Close
    End If
    
    If Trim$(GenSpec) <> "" Then
      WordSplit Trim$(GenSpec), iPageWidth
      sSendToPort "Note :"
      sSendToPort LAlign(iPageWidth, WSplit(1))
      For s = 2 To UBound(WSplit)
        sSendToPort LAlign(iPageWidth, WSplit(s))
      Next s
      sSendToPort String(iPageWidth, "-")
    End If
                         'Order Issued
    Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
    Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
    Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
    Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If RSOpen(snap, Sql) = False Then Exit Function
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      For i = 1 To snap.RecordCount
        If i = 1 Then
          sSendToPort LAlign(16, "Order Issued To:") & Space(1) & LAlign(50, snap("Description"))
        Else
          sSendToPort Space(17) & LAlign(50, snap("Description"))
        End If
        snap.MoveNext
      Next i
    End If
  
  Else

      frmPrint.Show
  
      Sql = "select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, c.PartyName, a.ReferenceNo, "
      Sql = Sql + " a.ReferenceDate,a.RevisionLevel, a.ItemCode, a.ItemDescription, a.PartNo, a.RevisionNo,"
      Sql = Sql + " b.sDesc, a.OrderQty, isNull(d.ItemSpecification,'') ItemSpecification, a.GeneralDetails,"
      Sql = Sql & " PackingDesc, InsuranceDesc, DeliveryDesc, PaymentDesc, FreightDesc, ModeDesc, FormDesc, FormulaName, PostingAcName "
      Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + " left outer join Inv_UOM_Master b on a.UOM = b.Code and a.Branchid = b.Branchid"
      Sql = Sql + " and a.Compid = b.Compid left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
      Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join INV_Material_Master d "
      Sql = Sql + " on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid "
      Sql = Sql & " Left Outer Join dbo.fn_SAL_SO_Terms ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')f"
      Sql = Sql & " On f.OrderType = a.OrderType and f.OrderNo = a.OrderNo and f.OrderDate = a.OrderDate"
      Sql = Sql & " and f.BranchID = a.BranchID and f.CompID = a.CompID"
      Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
      Sql = Sql + " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    
      If RSOpen(snap, Sql) = False Then Exit Function
      If snap.EOF Then
        sShowMessage "No Record(s) Found"
        sPrintcurrent_SalesView = False
        GoTo ExitHere
      End If
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
      frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
  
      frmPrint.sprPrint.Col = 3
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
        If RptFlg <> "View" And RptFlg <> "Print" And RptFlg <> "Obsolute" Then
          frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
      End If
      
      If iPageNo = 1 Then
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

        j = j + 2
GridCount = GridCount + 2
      End If
          
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText 1, j, "Sales Order"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

      j = j + 2
GridCount = GridCount + 2
      
      frmPrint.sprPrint.MaxCols = 5
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(1) = 20

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(2) = 36

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(3) = 10

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(4) = 6

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(5) = 10
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "SO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderType") & snap("OrderNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "SO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("OrderDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceNo")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer PO Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ReferenceDate")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Level"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("RevisionLevel")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Customer Name"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PartyName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Packing & Forwarding"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PackingDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Insurance"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("InsuranceDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DeliveryDesc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("DeliveryDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PaymentDetails"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PaymentDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Freight"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FreightDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Mode"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("ModeDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Form"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormDesc")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "FormulaName"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("FormulaName")
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "PostingAc"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("PostingAcName")
      j = j + 2
GridCount = GridCount + 2
RowProcess
Cont:
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rev. No "
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

      j = j + 1
GridCount = GridCount + 1
RowProcess
      s = s + 1
      If s > 1 Then
 GoTo Cont1
End If
      
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        GenSpec = Trim$(snap("GeneralDetails"))
        For i = 1 To snap.RecordCount
          'ProcessFlow
          If Trim$(gReportID) = "PMI" Then
            Sql = " select distinct a.ProcessSequence,a.ProcessFlow from BOM_Process_Flow a"
            Sql = Sql + " left outer join dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b "
            Sql = Sql + " on a.ItemCode= b.ItemCode and a.BranchId=b.BranchId and a.CompId=b.CompId"
            Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by ProcessSequence"
            If RSOpen(rs1, Sql) = False Then Exit Function
            If rs1.EOF Then
              Sql = " select distinct a.OperationSequence ProcessSequence,a.OperationName ProcessFlow from MPSMaster a"
              Sql = Sql + " where a.ItemCode = '" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' order by a.OperationSequence"
              If RSOpen(rs1, Sql) = False Then Exit Function
            End If
            'RawMaterials
            Sql = " select distinct a.ChildCode,b.ItemDescription,a.Qty,b.PurchaseUOMDesc from BOM_STRUCTURE_DTL a"
            Sql = Sql + " left outer join Material_Master b on a.ChildCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
            Sql = Sql + " where a.ParentCode ='" & snap("ItemCode") & " ' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" ' order by ProcessSequence"
            If RSOpen(rsDetails, Sql) = False Then Exit Function
          End If
Cont1:
          k = 0
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("RevisionNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, CheckNull(snap("sDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, IIf(snap("OrderQty") = 0, "Open", snap("OrderQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
          If pheight > 54 And k = 1 And s = 1 Then
 GoTo Cont
End If
            
'          If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
'            WordSplit Trim$(snap("ItemSpecification")), 40
'            .Col = 1: .Row = j

'            .SetText .Col, .Row, WSplit(1): .TypeHAlign = TypeHAlignLeft

'            j = j + 1: GridCount = GridCount + 1: RowProcess
'
'            For s = 2 To UBound(WSplit)
'              .Col = 1: .Row = j

'              .SetText .Col, .Row, WSplit(s): .TypeHAlign = TypeHAlignLeft

'              j = j + 1: GridCount = GridCount + 1: RowProcess
'            Next s
'          End If
          snap.MoveNext
        Next i
      End If
      'Delivery Details
      Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
      Sql = Sql + " where a.OrderType = '" + Trim$(txttype) + "' and a.OrderNo = " & Val(txtNo) & " "
      Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
      Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
      If RSOpen(rs, Sql) = False Then Exit Function
      If Not rs.EOF Then
        PartNo = ""
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Delivery Details :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Sch. Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        j = j + 1
GridCount = GridCount + 1
RowProcess

        rs.MoveLast
rs.MoveFirst
        For i = 1 To rs.RecordCount
          If PartNo <> rs("PartNo") Then
            If i > 1 Then
 j = j + 1
GridCount = GridCount + 1
RowProcess
End If
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("PartNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs("ScheduleQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          PartNo = rs("PartNo")
          rs.MoveNext
        Next i
      End If
      j = j + 1
GridCount = GridCount + 1
RowProcess
'Process Flow Print
      If Trim$(gReportID) = "PMI" Then
        If Not rs1.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Operation(s):"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "S.No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        
          For i = 1 To rs1.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessSequence")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rs1("ProcessFlow")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rs1.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
                          'Raw Material Print
        If Not rsDetails.EOF Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Details:"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rawmaterial Code"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Norms"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

          j = j + 1
GridCount = GridCount + 1
RowProcess
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

          For i = 1 To rsDetails.RecordCount
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ChildCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("PurchaseUOMDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, rsDetails("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            j = j + 1
GridCount = GridCount + 1
RowProcess
            rsDetails.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
      rs1.Close
      rsDetails.Close
      End If
      
      If Trim$(GenSpec) <> "" Then
        WordSplit Trim$(GenSpec), iPageWidth
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Note :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(1)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        For s = 2 To UBound(WSplit)
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, WSplit(s)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        Next s
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
                           'Order Issued
      Sql = "select b.Description from SAL_Order_Department a left outer join TermsMaster b"
      Sql = Sql + " on a.DepartmentCode = b.Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
      Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txttype) + "' "
      Sql = Sql + " and a.OrderNo = " & Val(txtNo) & " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "'"
      Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
      If RSOpen(snap, Sql) = False Then Exit Function
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        For i = 1 To snap.RecordCount
          If i = 1 Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Issued To :"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          Else
            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("Description")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          snap.MoveNext
        Next i
      End If
      j = j + 3
GridCount = GridCount + 3
RowProcess

      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 3, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "System Administrator"
  frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

      j = j + 1
GridCount = GridCount + 1
RowProcess
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
    Screen.MousePointer = 1
    totalpages = iPageNo
  End If
  
  sPrintcurrent_SalesView = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub txtType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtType_DblClick
End If
End Sub

Private Sub Design_fpsList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsList


fpsList.KeyUp +=  new EventHandler(fpsList_KeyUp)
fpsList.Click +=  new EventHandler(fpsList_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptOrderDocument.frm", fpsList, "")

End Sub
