VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptDeliveryChallanDocument 
   Caption         =   "Delivery Challan Document"
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
      Left            =   7530
      TabIndex        =   2
      Top             =   1890
      Width           =   345
   End
   Begin VB.Frame FraStep 
      Caption         =   "Date"
      Height          =   630
      Index           =   1
      Left            =   4215
      TabIndex        =   23
      Top             =   1725
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
         Format          =   108986369
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
         Format          =   108986369
         CurrentDate     =   37273
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   210
         Left            =   1785
         TabIndex        =   25
         Top             =   255
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6495
      TabIndex        =   15
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5200
      TabIndex        =   14
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   3915
      TabIndex        =   13
      Top             =   4095
      Width           =   1300
   End
   Begin VB.Frame FraReportType 
      Height          =   945
      Left            =   4283
      TabIndex        =   18
      Top             =   3090
      Width           =   3135
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   585
         Width           =   990
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptDeliveryChallanDocument.frx":0000
         Left            =   1035
         List            =   "rptDeliveryChallanDocument.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   19
         Top             =   255
         Width           =   885
      End
   End
Begin AceSpread fpsList
      Height          =   1470
      Left            =   7860
      TabIndex        =   5
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
      Left            =   7485
      TabIndex        =   10
      Top             =   2550
      Width           =   345
   End
   Begin VB.Frame FraStep 
      Height          =   780
      Index           =   0
      Left            =   4253
      TabIndex        =   17
      Top             =   2310
      Width           =   3195
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC Date"
         Height          =   210
         Left            =   2040
         TabIndex        =   22
         Top             =   180
         Width           =   585
      End
      Begin VB.Label lblEnqNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC.No."
         Height          =   210
         Left            =   990
         TabIndex        =   21
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblEnqType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DC Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   180
         Width           =   615
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
      Index           =   2
      Left            =   7500
      TabIndex        =   6
      Top             =   2655
      Width           =   345
   End
   Begin VB.Frame FraStep 
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
      Index           =   2
      Left            =   4170
      TabIndex        =   16
      Top             =   2490
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
         Left            =   150
         TabIndex        =   3
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
         Height          =   225
         Left            =   2115
         TabIndex        =   4
         Top             =   255
         Width           =   1050
      End
   End
End
Attribute VB_Name = "rptDeliveryChallanDocument"
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

Public Sub sAudit()
    
 Try

  fraStep(0).Visible = False
  CmdStep(0).Visible = False
  
  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = True
  CmdStep(2).Visible = True
  fraStep(3).Visible = True
  CmdStep(3).Visible = True
  
  fraStep(1).Enabled = True
  CmdStep(1).Enabled = True
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
  
  fraStep(1).Visible = True
  CmdStep(1).Visible = True
  fraStep(2).Visible = True
  CmdStep(2).Visible = True
  
  fraStep(1).Enabled = True
  CmdStep(1).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sprint()
  Try

  fraStep(0).Left = 4240
  fraStep(0).Width = 3160
  CmdStep(0).Left = 7470
  
  fraStep(1).Visible = False
  CmdStep(1).Visible = False
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  
  fraStep(0).Visible = True
  CmdStep(0).Visible = True
  MDISales.Toolbar1.Buttons(5).Enabled = False
  MDISales.Toolbar1.Buttons(4).Enabled = False
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  fraStep(0).Left = 4240
  fraStep(0).Width = 3160
  CmdStep(0).Left = 7470
  
  fraStep(1).Visible = False
  CmdStep(1).Visible = False
  fraStep(2).Visible = False
  CmdStep(2).Visible = False
  
  fraStep(0).Visible = True
  CmdStep(0).Visible = True
  MDISales.Toolbar1.Buttons(5).Enabled = False
  MDISales.Toolbar1.Buttons(4).Enabled = False
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub txttype_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct DCType, DCNo, DCDate From SAL_DC_HDR"
  Sql = Sql & " Where Branchid='" & gBranchID & "'"
  Sql = Sql & " and Compid = '" & gCompID & "' and finyear = '" & gFinyear & "'"
  Sql = Sql & " Order by dcdate desc, dcno desc"

  If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of DC No's", "DC Type", "DC No", "DC Date") = True Then
    txttype.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

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
  
  Sql = "Select Distinct IsNull(b.PartyName,'--') As CustomerName,"
  Sql = Sql & " a.CustomerCode"
'  Sql = Sql & " IsNull(Max(c.Revisionno),'0') As RevNo, Max(c.TranStamp) As TranStamp"
  Sql = Sql & " From SAL_CUSTOMER_ITEMMASTER a"
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On a.CustomerCode=b.PartyCode and a.BranchID=b.Branchid"
  Sql = Sql & " and a.CompID=b.Compid"
'  Sql = Sql & " Left Outer Join SAL_CUSTOMER_ITEMMASTER_His c"
'  Sql = Sql & " On a.CustomerCode=c.CustomerCode and a.ItemCode=c.ItemCode"
'  Sql = Sql & " and a.BranchID=c.Branchid and a.CompID=c.Compid"
  Sql = Sql & " Where a.Status=1 and a.BranchID='" & gBranchID & "'"
  Sql = Sql & " and a.CompID='" & gCompID & "'"
'  Sql = Sql & " Group by a.CustomerCode, b.PartyName, a.FormulaCode"
    
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
'    CmdStep(2).Enabled = False
'    fpsList.Visible = True
'    fpsList.Enabled = True
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
    CmdStep(2).Enabled = False
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

Private Sub fpsList_KeyUp(sender as object, e as KeyUpEventArgs)
  
' On Error Resume Next
  If fpsList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub fpsList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If fpsList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
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
  
  If fraStep(0).Visible = True And CmdStep(0).Enabled = True Then
    If txttype = "" Then
 sShowMessage "Select the DC Type"
Exit Sub
End If
    Call StepProceedNext(Me, Index)
    Index = 2
Call StepProceedNext(Me, Index)
  ElseIf fraStep(1).Visible = True And CmdStep(1).Enabled = True Then
    Call StepProceedNext(Me, Index)
  ElseIf fraStep(2).Visible = True And CmdStep(2).Enabled = True Then
    Call StepProceedNext(Me, Index)
    Index = 2
Call StepProceedNext(Me, Index)
  Else
    Call StepProceedNext(Me, Index)
    cmdview.Enabled = True
  End If
  
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If cmbType.ListIndex = 0 Then
  
    If Trim$(RptFlg) = "View" Then
      If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Print" Then
      If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Listing" Then
      If sPrintcurrent_Listing = False Then
 GoTo ExitHere
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
    
    If mDocSeq(1) = 1 Then
      If sPrintcurrent_Listing = False Then
 GoTo ExitHere
End If
    ElseIf mDocSeq(1) = 2 Then
      
    ElseIf mDocSeq(1) = 2 Then    'Revision Details
    
    ElseIf mDocSeq(1) = 3 Then    'Audit
    
    ElseIf mDocSeq(1) = 4 Then    'View
      If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
    ElseIf mDocSeq(1) = 5 Then    'Print
      If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
    Else
      'Skip
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

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call ClearControls
  Call StepClear(Me)
  Call StepEnable(Me)
  
  If RptFlg = "View" Then
    Call sView
  ElseIf RptFlg = "Print" Then
    Call sprint
  ElseIf RptFlg = "Listing" Then
    Call sListing
  End If
  Call sButtonsAction
'  MDISales.Toolbar1.Buttons(5).Enabled = False
'  MDISales.Toolbar1.Buttons(4).Enabled = False
  cmbType.ListIndex = 0
  
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
    sSendToPort LAlign(10, "DC  No.") & Space(1) & LAlign(30, "Customer") & Space(1) & LAlign(20, "Part No.") & Space(1) & LAlign(4, "UOM") & Space(1) & LAlign(25, "Item Code") & Space(1) & LAlign(12, "UOM") & Space(1) & RAlign(12, "Rate") & Space(1) & RAlign(12, "Goods Value")
    sSendToPort LAlign(10, "DC  Date") & Space(32) & LAlign(20, "Part Description") & Space(6) & LAlign(25, "Item Description") & Space(1) & LAlign(12, "Qty") & Space(1) & RAlign(12, "Ass.Rate") & Space(1) & RAlign(12, "Ass. Value")
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
  Dim sRow As Long
  Dim Gtot(2) As Double
  Dim tmp As String
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
  Dim mDocType As String
    
  Try
  
  sPrintcurrent_Listing = False
  
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
  End If
  
  Sql = "Select Distinct a.DCType, a.DCNo, a.DCDate, a.CustomerCode,"
  Sql = Sql & " b.PartyName As CustomerName, a.Remarks,"
  Sql = Sql & " Case When a.DCSource='D' Then 'Direct' Else 'Order' End As DCSource,"
  Sql = Sql & " IsNull(a.OrderType,'') As OrdType, IsNull(a.OrderNo,'') As OrdNo,"
  Sql = Sql & " IsNull(a.OrderDate,'') As OrdDate, IsNull(a.InvoiceType,'') As InvType,"
  Sql = Sql & " IsNull(a.InvoiceNo,'') As InvNo, IsNull(a.InvoiceDate,'') As InvDate,"
  Sql = Sql & " a.ItemCode, IsNull(c.ItemDescription,'') As ItemDescription,"
  Sql = Sql & " f.SDesc As UOM, IsNull(d.PartNo,'') As PartNo,"
  Sql = Sql & " h.SDesc As PartUOM, IsNull(e.ItemDescription,'') As PartDescription,"
  Sql = Sql & " a.DCQty As Qty, Round(a.Rate,0) As Rate, Round(a.DCQty*a.Rate,0) As GoodsValue,"
  Sql = Sql & " Round(a.AssessableRate,0) As AssRate,"
  Sql = Sql & " Round(a.DCQty*a.AssessableRate,0) As AssValue, a.ItemGroupKey"
  Sql = Sql & " From SalesDC a"
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid"
  Sql = Sql & " and a.Compid=b.Compid"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER c"
  Sql = Sql & " On a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " and c.status=1"
  Sql = Sql & " Left Outer Join PartMast d"
  Sql = Sql & " On a.ItemCode=d.StageCode and a.Branchid=d.BranchID and a.Compid=d.CompID"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER e"
  Sql = Sql & " On d.ItemCode=e.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " Left Outer Join INV_UOM_Master f"
  Sql = Sql & " On c.UOM=f.Code and a.Branchid=f.Branchid and a.Compid=f.CompId and f.Status=1"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER g"
  Sql = Sql & " On d.ItemCode=g.ItemCode and a.Branchid=g.Branchid and a.Compid=g.Compid"
  Sql = Sql & " and g.status=1"
  Sql = Sql & " Left Outer Join INV_UOM_Master h"
  Sql = Sql & " On g.UOM=h.Code and a.Branchid=h.Branchid and a.Compid=h.CompId"
  Sql = Sql & " and h.Status=1"
  Sql = Sql & " Where a.DCDate Between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTodate.Value, "yyyy-MM-dd") & "'"
  
  If optCustSelective Then
      Sql = Sql & " and a.CustomerCode in "
      Sql = Sql & Spread_Get_SelectionList(fpsList, 2)  End If
    
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.DCDate, a.DCNo"
  
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

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
      'Top
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mDocType) <> Trim$(rs("DCType")) & Trim$(rs("DCNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DCType")) & Trim$(rs("DCNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("OrdType")) & Trim$(rs("OrdNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          mCustNm = ""
        End If
        If Trim$(mCustNm) <> Trim$(rs("CustomerName")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("CustomerName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Uom"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("GoodsValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DCSource"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
      'Botom
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
         
        If Trim$(mDocType) <> Trim$(rs("DCType")) & Trim$(rs("DCNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("DCDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("OrdDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
                
        mDocType = Trim$(rs("DCType")) & Trim$(rs("DCNo"))
        mCustNm = Trim$(rs("CustomerName"))
        rs.MoveNext
      Next i
    
      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
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
    gStrReportHeading = "SALES DC  (LISTING) BETWEEN " & mFromDate & " AND " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("DC No")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("DC Date"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("Qty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
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

Private Sub Header()
  Dim mstr As String
  Dim CurrDate As Date
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    If RptFlg = "Audit" Then
      frmPrint.sprPrint.MaxCols = 5
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
    ElseIf RptFlg = "Listing" Then
      frmPrint.sprPrint.MaxCols = 10
      frmPrint.sprPrint.Col = 10
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
        If RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
        End If
        frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      Else
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
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

    frmPrint.sprPrint.SetText 1, j, IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Delivery Challan" & IIf(Trim$(RptFlg) = "Listing", " - Listing", " - Audit")
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    If Trim$(RptFlg) = "Listing" Then
    'Top
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "DC. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order No."
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No."
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
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

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Delivery Challan Source"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
    'Bottom
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "DC Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Order Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
              
      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
            
    ElseIf Trim$(RptFlg) = "Audit" Then 'Audit
    
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Type"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Date"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User ID"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Transaction Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
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
  End If
End Sub

Private Function sPrintcurrent_Audit() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim CurrDate As Date
  Dim Gtot(2) As Double
  Dim tmp As String
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
    
  Try
  sPrintcurrent_Audit = False
  CurrDate = ServerDate
    
  Sql = "Select Distinct a.DCType, a.DCNo, a.DCDate,"
  Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As Userid,"
  Sql = Sql & " Case When a.Userid='ADMIN' Then 'ADMIN' Else IsNull(c.UserName,'') End As UserName,"
  Sql = Sql & " IsNull(a.Remarks ,'') As Remarks" ',a.TranStamp"
  Sql = Sql & " From SAL_DC_HDR a"
  Sql = Sql & " Left Outer Join dbo.fn_SEC_Users ( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
  Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
  Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  Sql = Sql & " Where a.DCDate between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpTodate.Value, "yyyy-MM-dd") & "'"
  
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.DCNo in("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.DCDate, a.DCNo"
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
pheight = 54
Orient = 1
    mCustNm = ""
    
    Header
    j = j + 1
GridCount = GridCount + 1
RowProcess

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        If Trim$(mCustNm) <> Trim$(rs("DCType")) & Trim$(rs("DCNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DCType")) & Trim$(rs("DCNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DCDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Userid")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("UserName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
'        .Row = j: .Col = 5

'        .SetText .Col, j, CheckNull(rs("TranStamp")): .TypeHAlign = TypeHAlignLeft
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        mCustNm = Trim$(rs("DCType")) & Trim$(rs("DCNo"))
        rs.MoveNext
      Next i
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
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
      sSendToPort LAlign(10, rs("DC No")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("DC Date"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("DCQty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
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

Private Function sPrintCurrent_Print() As Boolean
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  Dim CurrDate As Date
  Dim i As Integer
  Dim k As Integer
  
  Try
  sPrintCurrent_Print = False
  
  CurrDate = ServerDate
  
  Sql = "Select Distinct a.DCType, a.DCNo,a.DCDate, a.DCSource, a.CustomerCode, c.PartyName As CustomerName,"
  Sql = Sql & " a.Remarks,a.ItemCode, b.ItemDescription, b.UOMDesc, b.UOMDecimal,"
  Sql = Sql & " a.DCQty, a.Rate, a.AssessableRate As AssRate, a.OrderType,"
  Sql = Sql & " a.OrderNo, a.OrderDate, Case When a.DCSource='D' Then 'Direct' Else 'Order' End As DCSource"
  Sql = Sql & " ,d.Partno From SalesDC a"
  Sql = Sql & " Left Outer Join Material_Master b On a.ItemCode=b.ItemCode"
  Sql = Sql & " and a.Branchid = b.Branchid and a.Compid = b.Compid"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid"
  Sql = Sql & " and a.Compid=c.Compid"
  Sql = Sql & " Left Outer Join PartMast d"
  Sql = Sql & " On a.ItemCode=d.StageCode and a.Branchid=d.BranchID and a.Compid=d.CompID"
  Sql = Sql & " Where a.DCType='" & Trim$(txttype) & "'"
  Sql = Sql & " and a.DCNo='" & Trim$(txtNo) & "' and a.DCDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " "
  Sql = Sql & " and b.Status=1 and a.Branchid='" & gBranchID & "'"
  Sql = Sql & " and a.Compid ='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then Exit Function
  If snap.RecordCount <= 0 Then
    sShowMessage "Record(s) Not Found"
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

    snap.MoveLast
snap.MoveFirst
    
    If Not snap.EOF() Then
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
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

        frmPrint.sprPrint.SetText 1, j, "Delivery Challan"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

        j = j + 2
GridCount = GridCount + 2

        frmPrint.sprPrint.MaxCols = 7
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DC  No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & Trim$(snap("DCType")) & Trim$(snap("DCNo"))
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DC Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & Format(snap("DCDate"), "dd/MM/yyyy")
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order  No."
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & Trim$(snap("OrderType")) & Trim$(snap("OrderNo"))
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Order Date"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & Format(snap("OrderDate"), "dd/MM/yyyy")
        j = j + 1
GridCount = GridCount + 1
RowProcess

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "DC Source"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & snap("DCSource")
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
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, ":" & Space(1) & Trim$(snap("CustomerName"))
        j = j + 2
GridCount = GridCount + 2
RowProcess
      
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item code"
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "UOM"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Goods"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Remarks"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Ass. Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      
        For i = 1 To snap.RecordCount
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, snap("PartNo")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("UOMDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("DCQty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Format(CheckNull(snap("Rate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Format(Val(snap("DCQty") * snap("Rate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Remarks"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, snap("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Format(CheckNull(snap("AssRate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
          snap.MoveNext
        Next i
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 4
GridCount = GridCount + 4
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True
        frmPrint.sprPrint.AddCellSpan 3, j, 2, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "System Administrator"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

        j = j + 4
GridCount = GridCount + 4
RowProcess
    End If
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 4
    Screen.MousePointer = 1
    totalpages = iPageNo
        
  ElseIf cmbType.ListIndex = 1 Then   'File
  
    If mDocSeq(1) = 4 Then
    
      pheight = 75
      pPaper = vbPRPSA4
      Orient = 1
      iPageWidth = 80
      iLineNo = 1
      fOpenPort Me
      
    ElseIf mDocSeq(1) = 5 Then
    
      Printer.FontName = "Draft 10cpi"
      If Printer.FontName <> "Draft 10cpi" Then
        Printer.FontName = "Roman 10cpi"
        If Printer.FontName <> "Roman 10cpi" Then
          Printer.FontName = "Courier New"
          Printer.FontSize = 11
        End If
      End If
    
      k = 0
      iPageNo = 1
      iPageWidth = 80
      pheight = 86
    End If
      
    If gReportID = "KUS" Then
      Dumrow 11
      k = k + 11
RowProcess
    Else
      Dumrow 5
      k = k + 5
RowProcess
    End If
    
    If mDocSeq(1) = 4 Then
      sSendToPort CAlign(iPageWidth, "SALES DC ")
      sSendToPort ""
    ElseIf mDocSeq(1) = 5 Then
      Printer.Print CAlign(iPageWidth, "SALES DC ")
      Printer.Print
    End If
    k = k + 2
RowProcess
    
    Sql = "Select b.add1,b.add2, b.city, b.pincode From FAS_PARTY_ADDRESS b"
    Sql = Sql & " Where b.PartyCode=" & Val(snap("CustomerCode")) & ""
    Sql = Sql & " and b.Compid='" & gCompID & "'"
    
    Set snap1 = New AceADODB.Recordset
    If RSOpen(snap1, Sql) = False Then GoTo ExitHere
    
    If snap1.RecordCount >= 1 Then
      If mDocSeq(1) = 4 Then
        
        If gReportID = "KUS" Then
          sSendToPort Space(2) & LAlign(5, "To:")
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, snap("CustomerName"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
        Else
          sSendToPort Space(2) & LAlign(40, gCompName) & Space(1) & LAlign(40, snap("CustomerName"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, Cadd1) & Space(1) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, Cadd2) & Space(1) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

          sSendToPort Space(2) & LAlign(40, Ccity & "-" & Cpincode) & Space(1) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
        End If
        sSendToPort Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess

        sSendToPort Space(2) & LAlign(11, "DC  No.    :") & Space(1) & LAlign(10, snap("DC Type") & snap("DC No")) & Space(22) & LAlign(12, "DC  Date    :") & Space(1) & LAlign(10, snap("DC Date"))
k = k + 1
RowProcess

'        sSendToPort Space(2) & LAlign(11, "Order No. :") & Space(1) & LAlign(10, snap("OrderType") & snap("OrderNo")) & Space(22) & LAlign(12, "Order Date :") & Space(1) & LAlign(10, CheckNull(snap("OrderDate"))): k = k + 1: RowProcess

        sSendToPort Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
        
      
      ElseIf mDocSeq(1) = 5 Then
      
        If gReportID = "KUS" Then
          Printer.Print Space(2) & LAlign(5, "To:")
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, snap("CustomerName"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
        Else
          Printer.Print Space(2) & LAlign(40, gCompName) & Space(1) & LAlign(40, snap("CustomerName"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, Cadd1) & Space(1) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, Cadd2) & Space(1) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

          Printer.Print Space(2) & LAlign(40, Ccity & "-" & Cpincode) & Space(1) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
        End If
        Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess

        Printer.Print Space(2) & LAlign(11, "DC  No.    :") & Space(1) & LAlign(10, snap("DC Type") & snap("DC No")) & Space(22) & LAlign(12, "DC  Date    :") & Space(1) & LAlign(10, snap("DC Date"))
k = k + 1
RowProcess

        Printer.Print Space(2) & LAlign(11, "Order No. :") & Space(1) & LAlign(10, snap("OrderType") & snap("OrderNo")) & Space(22) & LAlign(12, "Order Date :") & Space(1) & LAlign(10, CheckNull(snap("OrderDate")))
k = k + 1
RowProcess

        Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
        
      Else
        'Skip
      End If
    Else
      'Skip
    End If
    sPageHeader
    k = k + 4
RowProcess
     
    If snap.RecordCount >= 1 Then
      snap.MoveLast
snap.MoveFirst
      For i = 1 To snap.RecordCount
      
        If mDocSeq(1) = 4 Then
          sSendToPort Space(2) & LAlign(20, snap("ItemCode")) & Space(1) & LAlign(33, snap("ItemDescription")) & Space(1) & LAlign(3, snap("UOM")) & Space(1) & RAlign(10, Round(snap("Rate"), 2)) & Space(1) & RAlign(10, Round(Val(snap("Qty")) * Val(snap("Rate")), 2))
          sSendToPort Space(56) & Space(1) & LAlign(3, snap("Qty")) & Space(1) & RAlign(10, Round(snap("AssRate"), 2)) & Space(1) & RAlign(10, Round(Val(snap("Qty")) * Val(snap("AssRate")), 2))
k = k + 1
RowProcess

          sSendToPort ""
k = k + 1
RowProcess
        ElseIf mDocSeq(1) = 5 Then
          Printer.Print Space(2) & LAlign(20, snap("ItemCode")) & Space(1) & LAlign(33, snap("ItemDescription")) & Space(1) & LAlign(3, snap("UOM")) & Space(1) & RAlign(10, Round(snap("Rate"), 2)) & Space(1) & RAlign(10, Round(Val(snap("Qty")) * Val(snap("Rate")), 2))
          Printer.Print Space(56) & Space(1) & LAlign(3, snap("Qty")) & Space(1) & RAlign(10, Round(snap("AssRate"), 2)) & Space(1) & RAlign(10, Round(Val(snap("Qty")) * Val(snap("AssRate")), 2))
k = k + 1
RowProcess

          Printer.Print
k = k + 1
RowProcess
        End If
        
        snap.MoveNext
      Next i
    End If
    
    If mDocSeq(1) = 4 Then
      sSendToPort Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
      snap.MoveFirst
      sSendToPort Space(2) & LAlign(10, "Remarks :") & Space(1) & LAlign(50, snap("Remarks"))
k = k + 1
RowProcess

      sSendToPort Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
    ElseIf mDocSeq(1) = 5 Then
      Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
      snap.MoveFirst
      Printer.Print Space(2) & LAlign(10, "Remarks :") & Space(1) & LAlign(50, snap("Remarks"))
k = k + 1
RowProcess

      Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
      Printer.EndDoc
      sShowMessage "Print Over"
    End If
    
  End If
  snap.Close
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set snap1 = Nothing
End Try
End Function

Private Sub txtType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txttype_DblClick
End If
End Sub

Private Sub Design_fpsList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsList


fpsList.KeyUp +=  new EventHandler(fpsList_KeyUp)
fpsList.Click +=  new EventHandler(fpsList_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptDeliveryChallanDocument.frm", fpsList, "")

End Sub
