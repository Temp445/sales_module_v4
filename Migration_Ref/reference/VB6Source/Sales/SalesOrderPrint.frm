VERSION 5.00
Begin VB.Form rptSalesOrderPrint 
   Caption         =   "Sales Order Print"
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
   Begin VB.CheckBox chkPrintSpec 
      Caption         =   "Print Item Specification"
      Height          =   210
      Left            =   4597
      TabIndex        =   9
      Top             =   3210
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   3990
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   1050
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   330
         Width           =   1050
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         Height          =   192
         Left            =   2016
         TabIndex        =   7
         Top             =   156
         Width           =   852
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
         Height          =   192
         Left            =   960
         TabIndex        =   5
         Top             =   156
         Width           =   696
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Order Type"
         ForeColor       =   &H00FF0000&
         Height          =   192
         Left            =   72
         TabIndex        =   3
         Top             =   156
         Width           =   876
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5535
      TabIndex        =   1
      Top             =   2685
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4245
      TabIndex        =   0
      Top             =   2685
      Width           =   1300
   End
End
Attribute VB_Name = "rptSalesOrderPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
 
  Try
  Me.MousePointer = vbHourglass

  Set frmCurrent = Me
  DoEvents
  iLineNo = 1
  iPageNo = 1
  
  If sPrintcurrent = False Then GoTo ExitHere
  sClosePort
  Set frmOutput = New frmReportViewer
  frmOutput.Caption = gStrReportHeading
  frmOutput.txtTotalPages.Text = CStr(iPageNo)
  frmOutput.wbrView.Navigate strFileName
  frmOutput.Tag = strFileName
  gStrReportHeading = ""
  frmOutput.Show vbModal
  DoEvents
  Screen.MousePointer = vbNormal
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Public Sub sPageHeader()
  
  Try
  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(20, "Part No") & Space(1) & LAlign(40, "Item Description") & Space(1) & LAlign(3, "Rev") & Space(1) & LAlign(3, "UOM") & Space(1) & RAlign(10, "Order Qty")
  sSendToPort LAlign(20, "Item Code")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrent() As Boolean
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
  

  Sql = "select a.OrderType, a.OrderNo, a.OrderDate, a.CustomerCode, c.PartyName, a.ReferenceNo, "
  Sql = Sql + " a.ReferenceDate,a.RevisionLevel, a.ItemCode, a.ItemDescription, a.PartNo, a.RevisionNo,"
  Sql = Sql + " b.sDesc, a.OrderQty, isNull(d.ItemSpecification,'') ItemSpecification, a.GeneralDetails"
  Sql = Sql + " from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
  Sql = Sql + " left outer join Inv_UOM_Master b on a.UOM = b.Code and a.Branchid = b.Branchid"
  Sql = Sql + " and a.Compid = b.Compid left outer join FAS_Party_Master c on a.CustomerCode = c.PartyCode"
  Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid left outer join INV_Material_Master d "
  Sql = Sql + " on a.ItemCode = d.ItemCode and a.Branchid = d.Branchid and a.Compid = d.Compid "
  Sql = Sql + " where a.OrderType = '" + Trim$(txtType) + "' and a.OrderNo = " & Val(txtNo) & " "
  Sql = Sql + " and a.OrderDate = '" + Format(txtDate, "yyyy-MM-dd") + "' and d.status = 1 and a.Branchid = '" & gBranchID & "' "
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"

  If RSOpen(snap, Sql) = False Then Exit Function
  If snap.EOF Then
    sShowMessage "No Record(s) Found"
    sPrintcurrent = False
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
      If chkPrintSpec = 1 And Trim$(snap("ItemSpecification")) <> "" Then
        WordSplit Trim$(snap("ItemSpecification")), 40
        sSendToPort Space(21) & LAlign(40, WSplit(1))
        For s = 2 To UBound(WSplit)
          sSendToPort Space(21) & LAlign(40, WSplit(s))
        Next s
      End If
      sSendToPort ""
      snap.MoveNext
    Next i
  End If
  'Delivery Details
  Sql = "select distinct a.ItemCode, a.PartNo, a.ScheduleDate, a.ScheduleQty from SAL_Order_Schedule a"
  Sql = Sql + " where a.OrderType = '" + Trim$(txtType) + "' and a.OrderNo = " & Val(txtNo) & " "
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
  Sql = Sql + " where b.Typeflg = 'Department' and a.OrderType = '" + Trim$(txtType) + "' "
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
  
  snap.Close
  rs.Close
 
  sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set rs1 = Nothing
  Set rs = Nothing
  Set rsDetails = Nothing
End Try
End Function

Private Sub Form_Load()
' On Error Resume Next
  If gReportID = "KUS" Then
    chkPrintSpec.Visible = True
  Else
    chkPrintSpec.Visible = False
  End If
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub txtType_DblClick()
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate

  Sql = "select Distinct a.OrderType, a.OrderNo, a.OrderDate, (Select e.PartyName from "
  Sql = Sql + " FAS_PARTY_MASTER e where e.PartyCode=a.CustomerCode and e.Branchid=a.Branchid"
  Sql = Sql + " and e.compid=a.compid) PartyName from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where a.Branchid='" & gBranchID & "' "
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.OrderDate desc, a.OrderNo Desc"

  If FetchValue_New(False, db.ConnectionString, Sql, "1111", "List of Sales Order's", "Order Type", "Order No", "Order Date", "Customer Name") = True Then
    txtType = getvalue(1)
    txtNo = getvalue(2)
    txtDate = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  If KeyCode = vbKeyF2 Then
 txtType_DblClick
End If
End Sub

