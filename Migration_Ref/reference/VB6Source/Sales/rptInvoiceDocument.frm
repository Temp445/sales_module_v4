VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptInvoiceDocument 
   Caption         =   "Sales Invoice Document"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
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
   ScaleHeight     =   11055
   ScaleWidth      =   19080
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
      Left            =   7680
      TabIndex        =   2
      Top             =   1635
      Width           =   345
   End
   Begin VB.Frame FraStep 
      Caption         =   "Date"
      Height          =   630
      Index           =   1
      Left            =   4365
      TabIndex        =   23
      Top             =   1440
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
         Format          =   196083713
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
         Format          =   196083713
         CurrentDate     =   37273
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   210
         Left            =   1785
         TabIndex        =   25
         Top             =   255
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   360
      End
   End
Begin AceSpread fpsList
      Height          =   1470
      Left            =   8055
      TabIndex        =   5
      Top             =   2310
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
   Begin VB.Frame FraReportType 
      Height          =   945
      Left            =   4425
      TabIndex        =   17
      Top             =   2790
      Width           =   3180
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptInvoiceDocument.frx":04D4
         Left            =   1035
         List            =   "rptInvoiceDocument.frx":04DE
         Style           =   2  'Dropdown List
         TabIndex        =   11
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
         Width           =   990
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   18
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6645
      TabIndex        =   15
      Top             =   3810
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5355
      TabIndex        =   14
      Top             =   3810
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   4065
      TabIndex        =   13
      Top             =   3810
      Width           =   1300
   End
   Begin VB.Frame FraStep 
      Height          =   780
      Index           =   0
      Left            =   4425
      TabIndex        =   16
      Top             =   2010
      Width           =   3180
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
         Caption         =   "Inv. Date"
         Height          =   210
         Left            =   2040
         TabIndex        =   21
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lblEnqNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv.No."
         Height          =   210
         Left            =   990
         TabIndex        =   20
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lblEnqType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv.Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   19
         Top             =   180
         Width           =   615
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
      Left            =   7665
      TabIndex        =   10
      Top             =   2265
      Width           =   345
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
      Left            =   7665
      TabIndex        =   6
      Top             =   2370
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
      Height          =   585
      Index           =   2
      Left            =   4365
      TabIndex        =   22
      Top             =   2205
      Width           =   3270
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
         Left            =   2160
         TabIndex        =   4
         Top             =   255
         Width           =   1050
      End
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
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
   End
End
Attribute VB_Name = "rptInvoiceDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim mFromDate As Date
Dim mToDate As Date
Dim j As Long

Public Sub sAudit()
    
  Try

  frastep(0).Visible = False
  CmdStep(0).Visible = False
  
  frastep(1).Visible = True
  CmdStep(1).Visible = True
  frastep(2).Visible = True
  CmdStep(2).Visible = True
  
  frastep(1).Enabled = True
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

  frastep(0).Visible = False
  CmdStep(0).Visible = False
  
  frastep(1).Visible = True
  CmdStep(1).Visible = True
  frastep(2).Visible = True
  CmdStep(2).Visible = True
  
  frastep(1).Enabled = True
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
  
  frastep(0).Visible = True
  CmdStep(0).Visible = True
  frastep(1).Visible = False
  CmdStep(1).Visible = False
  frastep(2).Visible = False
  CmdStep(2).Visible = False
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  frastep(0).Visible = True
  CmdStep(0).Visible = True
  frastep(1).Visible = False
  CmdStep(1).Visible = False
  frastep(2).Visible = False
  CmdStep(2).Visible = False
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frmPrint.cmdPPrint.Enabled = False
  frmPrint.cmdExcel.Enabled = False
  MDISales.Toolbar1.Buttons(4).Enabled = False
  MDISales.Toolbar1.Buttons(5).Enabled = False
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
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub txttype_DblClick()
  Dim Sql As String
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select Distinct InvoiceType, InvoiceNo, InvoiceDate From SAL_Invoice_HDR"
  Sql = Sql & " Where Branchid='" & gBranchID & "' and Compid='" & gCompID & "' and finyear='" & gFinyear & "'"
  Sql = Sql & " Group by InvoiceType, InvoiceNo, InvoiceDate"
  Sql = Sql & " Order by InvoiceDate desc, InvoiceNo desc"

  If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of Invoice No's", "Invoice Type", "Invoice No", "Invoice Date", "Rev. No.", "Rev. Date") = True Then

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
  
  If Trim$(RptFlg) = "Listing" Then    'Listing
    Sql = "Select Distinct a.CustomerCode As [Customer Code], b.PartyName As [Customer Name]"
    Sql = Sql & " From SAL_INVOICE_HDR a"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
    Sql = Sql & " On b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid"
    
  ElseIf Trim$(RptFlg) = "Audit" Then  'Audit
  
    Sql = "Select Distinct a.InvoiceType As [Inv.Type], a.InvoiceNo As [Inv.No.],"
    Sql = Sql & " a.InvoiceDate  As [Inv.Date],"
    Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As [User ID],"
    Sql = Sql & " Case When a.Userid='A0007' Then 'ADMIN' Else IsNull(c.UserName,'')"
    Sql = Sql & " End As [User Name]"
    Sql = Sql & " From SAL_INVOICE_HDR a"

    Sql = Sql & " Left Outer Join dbo.fn_SEC_Users ('" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
    Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
    Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  End If
    
  Sql = Sql & " Where a.InvoiceDate Between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  
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
'    fpsList.Visible = True
'  End If
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
    fpsList.Enabled = True
 fpsList.Visible = True
    fpsList.ClearSelection
    Call ADORS_Spread_DataFill(rs, fpsList, True)
    
    fpsList.MaxRows = fpsList.DataRowCnt
    Call Spread_Set_Focus(fpsList, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub fpsList_KeyUp(sender as object, e as KeyUpEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  Call fpsList_Click(fpsList.ActiveCol, fpsList.ActiveRow)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub fpsList_Click(sender as object, e as ClickEventArgs)
  Dim i As Integer
  Dim retval As Variant
  
' On Error Resume Next
  If Row = 0 And Col > 0 Then
    fpsList.Tag = Col
      fpsList.Row = 0
fpsList.Row2 = 0
      
      fpsList.BlockMode = True
      For i = 1 To fpsList.MaxCols
        fpsList.Col = i
fpsList.Col2 = i
fpsList.GetText i, 0, retval
        fpsList.CellType = CellTypeButton
        fpsList.TypeButtonText = Trim$(retval)
        fpsList.TypeButtonTextColor = RGB(0, 0, 0)
      Next i
      
      fpsList.Col = Col
fpsList.Col2 = Col
      fpsList.BlockMode = False
      
  End If
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
        
  If optCustSelective.Value = True Then
   If fpsList.SelectionCount = 0 Then
     sShowMessage "Select atleast one Customer"
     Exit Sub
   End If
  End If
  
  If frastep(0).Visible = True And CmdStep(0).Enabled = True Then
    If txttype = "" Then
 sShowMessage "Select the Order Type"
Exit Sub
End If
    Call StepProceedNext(Me, Index)
    Index = 2
Call StepProceedNext(Me, Index)
  ElseIf frastep(1).Visible = True And CmdStep(1).Enabled = True Then
    Call StepProceedNext(Me, Index)
  ElseIf frastep(2).Visible = True And CmdStep(2).Enabled = True Then
    Call StepProceedNext(Me, Index)
    Index = 2
Call StepProceedNext(Me, Index)
  Else
    Call StepProceedNext(Me, Index)
    cmdView.Enabled = True
  End If
  
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If cmbType.ListIndex = 0 Then
  
    If Trim$(RptFlg) = "View" Then
      If Trim$(txttype) = "" Then
 sShowMessage "Please Select Invoice Type"
ClearControls
GoTo ExitHere
End If
      If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Print" Then
      If Trim$(txttype) = "" Then
 sShowMessage "Please Select Invoice Type"
ClearControls
GoTo ExitHere
End If
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
  If RptFlg = "View" Then
    Call sView
  ElseIf RptFlg = "Print" Then
    Call sprint
  ElseIf RptFlg = "Listing" Then
    Call sListing
  Else
    Call sAudit
  End If
  Call sButtonsAction
  cmbType.ListIndex = 0
  MDISales.Toolbar1.Buttons(4).Enabled = False
  MDISales.Toolbar1.Buttons(5).Enabled = False
  
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
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frastep(1).Enabled = True
  CmdStep(1).Enabled = True
  
  cmbType.ListIndex = 0
  chkPrintdate.Value = 0
  txttype.Text = ""
txttype.Tag = ""
txtNo.Text = ""
txtDate.Text = ""
  
  cmdView.Enabled = False
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
    sSendToPort LAlign(10, "Invoice No.") & Space(1) & LAlign(30, "Customer") & Space(1) & LAlign(20, "Part No.") & Space(1) & LAlign(4, "UOM") & Space(1) & LAlign(25, "Item Code") & Space(1) & LAlign(12, "UOM") & Space(1) & RAlign(12, "Rate") & Space(1) & RAlign(12, "Goods Value")
    sSendToPort LAlign(10, "Invoice Date") & Space(32) & LAlign(20, "Part Description") & Space(6) & LAlign(25, "Item Description") & Space(1) & LAlign(12, "Qty") & Space(1) & RAlign(12, "Ass.Rate") & Space(1) & RAlign(12, "Ass. Value")
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

  Sql = "Select Distinct a.InvoiceType, a.InvoiceNo, a.InvoiceDate, a.InvoiceTime,"
  Sql = Sql & " a.CustomerCode, IsNull(b.PartyName,'') As CustomerName,"
  Sql = Sql & " a.ItemCode, IsNull(c.ItemDescription,'') As ItemDescription,"
  Sql = Sql & " IsNull(f.SDesc,'') As UOM, IsNull(d.PartNo,'') As PartNo,"
  Sql = Sql & " IsNull(h.SDesc,'') As PartUOM, IsNull(e.ItemDescription,'') As PartDescription,"
  Sql = Sql & " a.InvoiceQty As Qty, Round(a.Rate,0) As Rate,"
  Sql = Sql & " Round(a.InvoiceQty*a.Rate,0) As GoodsValue,"
  Sql = Sql & " Round(a.AssessableRate,0) As AssRate,"
  Sql = Sql & " Round(a.InvoiceQty*a.AssessableRate,0) As AssValue,a.TotalAmount"
  Sql = Sql & " From SalesInvoice a"
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid"
  Sql = Sql & " and a.Compid=b.Compid"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER c"
  Sql = Sql & " On a.ItemCode=c.ItemCode"
  Sql = Sql & " and a.Branchid=c.Branchid and a.Compid=c.Compid and c.status=1"
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
  Sql = Sql & " Where InvoiceDate Between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpToDate.Value, "yyyy-MM-dd") & "'"

  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = IIf(mDocSeq(6) = 3, 2, 1)
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.InvoiceDate, a.InvoiceNo"

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
        
        If Trim$(mDocType) <> Trim$(rs("InvoiceType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText 1, j, Trim$(rs("InvoiceType")) & Trim$(rs("InvoiceNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(mCustNm) <> Trim$(rs("CustomerName")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("CustomerName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("UOM")
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        j = j + 1
GridCount = GridCount + 1
RowProcess
      'Botom
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        If Trim$(mDocType) <> Trim$(rs("InvoiceType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("InvoiceDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        j = j + 1
GridCount = GridCount + 1
RowProcess
                
        mDocType = Trim$(rs("InvoiceType"))
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
    gStrReportHeading = "SALES Invoice (LISTING) BETWEEN " & mFromDate & " AND " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("InvoiceNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("InvoiceDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("Qty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
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
  CurrDate = ServerDate
  
  
    frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    If RptFlg = "Audit" Then
      frmPrint.sprPrint.MaxCols = 4
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
    ElseIf RptFlg = "Listing" Then
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
    Else
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
    End If
    
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintdate.Value = 1 Then
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
         If RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
        ElseIf RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
        End If
        frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      Else
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 4
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sales Invoice" & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    If Trim$(RptFlg) = "Listing" Then 'Listing
    'Top
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
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

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods Value"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
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

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv. Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Amount"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
            
    ElseIf Trim$(RptFlg) = "Audit" Then 'Audit
    
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemCode"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User ID"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Inv.Date":
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemDescription"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User Name"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Transaction Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

    End If
        
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub RowProcess()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
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
  
  Sql = "Select Distinct a.InvoiceType, a.InvoiceNo, a.InvoiceDate, ItemGroupKey,b.ItemCode,ItemDescription,"
  Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As Userid,"
  Sql = Sql & " Case When a.Userid='A0007' Then 'ADMIN' Else IsNull(c.UserName,'') End As UserName"
  Sql = Sql & ",a.TranStamp" ', IsNull(a.Remarks ,'') As Remarks
  Sql = Sql & " From SAL_INVOICE_HDR a Left Outer Join SAL_INVOICE_Dtl b on b.InvoiceType = a.InvoiceType and b.InvoiceNo = a.InvoiceNo"
  Sql = Sql & " and b.InvoiceDate = a.InvoiceDate and a.Branchid=b.BranchID and a.Compid=b.CompId"
  Sql = Sql & " Left Outer Join dbo.fn_SEC_Users ( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
  Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
  Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  Sql = Sql & " left outer join INV_Material_Master d on b.ItemCode = d.ItemCode"
  Sql = Sql & " and b.Branchid = d.Branchid and b.Compid = d.Compid"
  Sql = Sql & " Where a.InvoiceDate Between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.InvoiceNo In("
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
  Sql = Sql & " Order by a.InvoiceType, a.InvoiceNo, a.InvoiceDate"
  
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
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
  
        If Trim$(mCustNm) <> Trim$(rs("InvoiceType")) & Trim$(rs("InvoiceNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("InvoiceType")) & Trim$(rs("InvoiceNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Itemcode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Userid")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4

'        .SetText .Col, j, Trim$(rs("Remarks")): .ColWidth(6) = 25: .TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mCustNm) <> Trim$(rs("InvoiceType")) & Trim$(rs("InvoiceNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("InvoiceDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("UserName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4

'        .SetText .Col, j, Trim$(rs("Tran/")): .ColWidth(6) = 25: .TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        mCustNm = Trim$(rs("InvoiceType")) & Trim$(rs("InvoiceNo"))
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
    gStrReportHeading = "Delivery Challan Report between " & mFromDate & " and " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("InvoiceNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("InvoiceDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("DCQty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
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
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim tmp As String
  Dim tmp1 As String
  Dim Sql As String
  Dim X As Integer
  Dim k As Integer
  Dim i As Integer
  Dim tot As Double
  Dim EccNo As String
  Dim Rule As String
  Dim Place As String 'Postal Address Range and Division printing
  Dim Range As String
  Dim Division As String
  Dim City As String
  
  Try
  sPrintCurrent_Print = False
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
   
  If Trim$(txttype) = "" Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  End If

  Sql = "Select a.InvoiceType, a.InvoiceNo, a.InvoiceDate, a.InvoiceTime, a.TotalAmount,"
  Sql = Sql & " (Select b.PartyName From FAS_PArty_Master b Where b.PartyCode = a.CustomerCode"
  Sql = Sql & " and b.Branchid = a.Branchid and b.Compid = a.Compid) AS CustomerName, a.ModeCode, "
  Sql = Sql & " (select h.Description from TermsMaster h where h.TypeFlg = 'Mode' and h.Code = a.ModeCode"
  Sql = Sql & " and h.Branchid = a.Branchid and h.Compid = a.Compid) Mode,"
  Sql = Sql & " a.RemovalTime, f.TariffNo, isNull((select c.Commodity from INV_Commodity_Master c "
  Sql = Sql & " where c.TariffNo = f.TariffNo and c.Branchid = '" & gBranchID & "'"
  Sql = Sql & " and c.compid = '" & gCompID & "'), '') Commodity, a.VehicleNo, a.ItemCode, "
  Sql = Sql & " (select g.PartNo from SAL_Customer_ItemMaster g where g.CustomerCode = a.CustomerCode"
  Sql = Sql & " and g.ItemCode = a.ItemCode and g.Branchid = '" & gBranchID & "'"
  Sql = Sql & " and g.Compid = '" & gCompID & "') PartNo, f.ItemDescription, a.InvoiceQty , a.Rate, "
  Sql = Sql & " a.AssessableRate, a.NoofPackets, a.OrderType,a.OrderNo,a.OrderDate, e.Add1,"
  Sql = Sql & " e.Add2,e.Add3,e.City,e.Pincode, e.LST, e.CST, e.ECCNo, "
  Sql = Sql & " isNull((select top 1 ReferenceNo from SAL_ORDER_HDR d where "
  Sql = Sql & " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql & " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql & " and d.finyear = '" & gFinyear & "'), (select top 1 ReferenceNo from SAL_JOBWORKORDER_HDR d where "
  Sql = Sql & " d.OrderType = a.OrderType and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate "
  Sql = Sql & " and d.Branchid = '" & gBranchID & "' and d.Compid = '" & gCompID & "'"
  Sql = Sql & " and d.finyear = '" & gFinyear & "')) RefNo, a.AInformation1,a.Type,f.UOMDecimal from "
  Sql = Sql & " SalesInvoice a left outer join FAS_PARTY_ADDRESS e"
  Sql = Sql & " on a.CustomerCode = e.PartyCode and a.Branchid = e.Branchid "
  Sql = Sql & " and a.Compid = e.Compid left outer join Material_Master f "
  Sql = Sql & " on f.ItemCode = a.ItemCode and f.Branchid = a.Branchid and f.Compid = a.Compid"
  Sql = Sql & " where a.InvoiceType = '" & Trim$(txttype) & "'"
  Sql = Sql & " and a.InvoiceNo = " & Val(txtNo) & " and a.Branchid = '" & gBranchID & "'"
  Sql = Sql & " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If

  Sql = "Select (Address1+', '+Address2) as add1, (City+'-'+Pincode)as Add2,"
  Sql = Sql & " ('TNGST No. '+ LST+' CST No. : ' + Cst +' Dt : '+ convert(varchar(10), CstDate, 105)+' '+'AreaCode No.: '+AreaCode)as ApprovalDet,"
  Sql = Sql & " NotificationNo, NotificationDate from company_master where compid = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
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

    snap.MoveLast
snap.MoveFirst
    
    Header

    If Not snap.EOF() Then
        frmPrint.sprPrint.MaxCols = 5
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 30

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Invoice No. & Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & Trim$(snap("InvoiceType")) & "-" & Trim$(snap("InvoiceNo")) & " & " & Format(snap("InvoiceDate"), "dd/MM/yyyy")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Order No. & Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & Trim$(snap("OrderType")) & "-" & Trim$(snap("OrderNo")) & " & " & Format(snap("InvoiceDate"), "dd/MM/yyyy")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
                
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Customer Name"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & Trim$(snap("CustomerName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
                        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Noof Packets"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("NoofPackets")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Terms & Ccondition"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Tarrif No"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("TariffNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Commodity"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("TariffNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Mode"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("Mode")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 1
RowProcess
        
         frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "CST"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("CST")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 1
RowProcess
        
         frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "LST"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("LST")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 1
RowProcess
        
         frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "ECCNo."
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("ECCNo")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(10) & "Address : "
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, ": " & snap("Add1")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(2) & snap("Add2")
  frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        If Trim$(snap("Add2")) <> "" Then
          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(2) & snap("Add2")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
        
        If Trim$(snap("Add3")) <> "" Then
          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(2) & snap("Add3")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
                
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(2) & snap("City") & " - " & snap("PinCode")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 2
GridCount = GridCount + 2
RowProcess
      
      'Header
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "GoodsValue"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Value"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1

        If snap.RecordCount >= 1 Then
          snap.MoveLast
snap.MoveFirst
          For i = 1 To snap.RecordCount
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("ItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("PartNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("InvoiceQty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format((snap("Rate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(snap("InvoiceQty") * snap("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
  
            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(snap("AssessableRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format((snap("InvoiceQty") * snap("AssessableRate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

            snap.MoveNext
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
          Next i
        End If
        snap.MoveLast
snap.MoveFirst
        Sql = "select distinct a.ExprSeq, a.ExprName, a.Resultvalue, a.GrandTotalFlag "
        Sql = Sql & " from dbo.fn_COM_GetExprResults_Saved ('" & Trim$(txttype.Text) & "'," & Val(txtNo) & ", '" & Format(txtDate, "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, gModuleCode) & ", '" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        Sql = Sql & " order by a.ExprSeq"
        If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
        If Not rs.EOF Then
          snap.MoveLast
snap.MoveFirst

          k = 0
tot = snap("TotalAmount")
X = 1

          rs.MoveLast
rs.MoveFirst
          For i = 1 To 14
            If i = 7 Then
              If Not rs.EOF Then
                frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                rs.MoveNext
              End If
              If Not rs.EOF Then
                If rs("ExprName") = "Sub Total" Then
                  Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, frmPrint.sprPrint.MaxCols, j
                End If
                frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                If rs("ExprName") = "Sub Total" Then
                  Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, frmPrint.sprPrint.MaxCols, j
                End If
                rs.MoveNext
              End If
            ElseIf i >= 10 And i <= 11 Then
              If Not rs.EOF Then
                If rs("ExprName") = "GRAND TOTAL" Then
                   Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, frmPrint.sprPrint.MaxCols, j
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
frmPrint.sprPrint.FontBold = True

                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
                  j = j + 1
GridCount = GridCount + 1
RowProcess

                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                ElseIf CheckNull(rs("ExprName")) = "VAT" Then
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                Else
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                End If
                rs.MoveNext
              End If
            ElseIf i Mod 2 = 1 Then
              If Not rs.EOF Then
                If CheckNull(rs("ExprName")) = "GRAND TOTAL" Then
                   Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, frmPrint.sprPrint.MaxCols, j
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
frmPrint.sprPrint.FontBold = True

                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                ElseIf CheckNull(rs("ExprName")) = "GOODS VALUE" Then
                  j = j + 1
GridCount = GridCount + 1
RowProcess

                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                ElseIf CheckNull(rs("ExprName")) = "VAT" Then
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                Else
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ExprName")
                  frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("ResultValue"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
j = j + 1
GridCount = GridCount + 1
                End If
                rs.MoveNext
              End If
            End If
          Next i
        End If
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, frmPrint.sprPrint.MaxCols, j

        j = j + 4
GridCount = GridCount + 4
RowProcess
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Prepared By"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Authorized By"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
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
      sSendToPort CAlign(iPageWidth, "SALES Invoice")
      sSendToPort ""
    ElseIf mDocSeq(1) = 5 Then
      Printer.Print CAlign(iPageWidth, "SALES Invoice")
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

        sSendToPort Space(2) & LAlign(11, "Invoice No.    :") & Space(1) & LAlign(10, snap("InvoiceType") & snap("InvoiceNo")) & Space(22) & LAlign(12, "Invoice Date    :") & Space(1) & LAlign(10, snap("InvoiceDate"))
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

        Printer.Print Space(2) & LAlign(11, "Invoice No.    :") & Space(1) & LAlign(10, snap("InvoiceType") & snap("InvoiceNo")) & Space(22) & LAlign(12, "Invoice Date    :") & Space(1) & LAlign(10, snap("InvoiceDate"))
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
'  Set snap1 = Nothing
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

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptInvoiceDocument.frm", fpsList, "")

End Sub
