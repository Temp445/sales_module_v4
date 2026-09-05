VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLCNoInvoiceUpdate 
   Caption         =   "LCNo Invoice Update"
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
   ScaleHeight     =   11115
   ScaleWidth      =   20280
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   7770
      MaxLength       =   20
      TabIndex        =   20
      ToolTipText     =   "Enter the No of Packets"
      Top             =   4620
      Width           =   1710
   End
   Begin VB.Frame fraCustomer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   75
      TabIndex        =   8
      Top             =   660
      Width           =   13170
      Begin VB.TextBox txtLCBalAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   10815
         MaxLength       =   20
         TabIndex        =   18
         ToolTipText     =   "Enter the No of Packets"
         Top             =   375
         Width           =   1710
      End
      Begin VB.TextBox txtLCAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9090
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Enter the No of Packets"
         Top             =   375
         Width           =   1710
      End
      Begin VB.TextBox txtLCNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6075
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Enter the PLA Sl No"
         Top             =   360
         Width           =   1710
      End
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         TabIndex        =   0
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help"
         Top             =   345
         Width           =   3435
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   12585
         TabIndex        =   3
         ToolTipText     =   "Click to OK"
         Top             =   375
         Width           =   465
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   315
         Left            =   3555
         TabIndex        =   1
         ToolTipText     =   "Select the Invoice From Date"
         Top             =   345
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22413313
         CurrentDate     =   37643
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         ToolTipText     =   "Select the Invoice To Date"
         Top             =   345
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22413313
         CurrentDate     =   37643
      End
      Begin MSComCtl2.DTPicker dtpLCDate 
         Height          =   330
         Left            =   7815
         TabIndex        =   14
         ToolTipText     =   "Enter the PLA Date"
         Top             =   360
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   22413313
         CurrentDate     =   41759.5153240741
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Balance LC Amt."
         Height          =   210
         Left            =   10815
         TabIndex        =   19
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "LC Amt."
         Height          =   210
         Left            =   9090
         TabIndex        =   17
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "LC Date"
         Height          =   210
         Left            =   7845
         TabIndex        =   15
         Top             =   165
         Width           =   570
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "LC No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   6090
         TabIndex        =   13
         Top             =   165
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date To"
         Height          =   210
         Left            =   4800
         TabIndex        =   11
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date From"
         Height          =   210
         Left            =   3540
         TabIndex        =   10
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   105
         TabIndex        =   9
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   7305
      TabIndex        =   7
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5115
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   6017
      TabIndex        =   6
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5115
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4725
      TabIndex        =   5
      ToolTipText     =   "Click to Save the Data"
      Top             =   5115
      Width           =   1300
   End
Begin AceSpread SpInv
      Height          =   2955
      Left            =   4065
      TabIndex        =   4
      ToolTipText     =   "Enter the RR Details"
      Top             =   1545
      Width           =   5610
      _Version        =   458752
      _ExtentX        =   9895
      _ExtentY        =   5212
      _StockProps     =   64
      Enabled         =   0   'False
      ColsFrozen      =   4
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   1
      ProcessTab      =   -1  'True
      UserResize      =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   210
      Left            =   7365
      TabIndex        =   21
      Top             =   4680
      Width           =   345
   End
End
Attribute VB_Name = "frmLCNoInvoiceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Enum eInv
  mInvType = 1
  mInvNo = 2
  mInvDate = 3
  mInvAmt = 4
End Enum

Private Sub cmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  fraCustomer.Enabled = True
  SpInv.Enabled = False
  cmdSave.Enabled = False
  txtCustName.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdOk_Click()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
 
  Try
  Me.MousePointer = vbHourglass
    
  Set rs = New AceADODB.Recordset
  If contrl_valid(txtCustName, "Select the Custer Name") = False Then GoTo ExitHere
  If CDate(DTPFrom.Value) > CDate(DTPTo.Value) Then
    sShowMessage "From date can't be greater than to date"
    DTPFrom.SetFocus
    DTPFrom.Value = Format(DTPFrom, "dd/mm/yyyy")
    GoTo ExitHere
  End If

  fraCustomer.Enabled = False
  SpInv.Enabled = True
  cmdSave.Enabled = True
  SpInv.SetFocus
  Spread_Set_Focus SpInv, 1, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdSave_Click()
  Dim i As Long
  Dim tmp As Variant
  Dim cmd As AceADODB.Command
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  If Validate = False Then GoTo ExitHere
  
  CurDate = ServerDate
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  For i = 1 To SpInv.DataRowCnt
    Sql = " Execute SAL_Invoice_LCRef_Ins"
    SpInv.GetText eInv.mInvType, i, tmp
    Sql = Sql & " @InvType = '" & Trim$(tmp) & "',"
    SpInv.GetText eInv.mInvNo, i, tmp
    Sql = Sql & " @InvNo = " & Val(Trim$(tmp)) & ","
    SpInv.GetText eInv.mInvDate, i, tmp
    Sql = Sql & " @InvDate = '" & Format(tmp, "yyyy-MM-dd") & "',"
    Sql = Sql & " @CustomerCode = " & Val(Trim$(txtCustName.Tag)) & ","
    Sql = Sql & " @LCNo = '" & Trim$(txtLCNo.Text) & "',"
    Sql = Sql & " @LCDate = '" & Format(dtpLCDate.Value, "yyyy-MM-dd") & "',"
    SpInv.GetText eInv.mInvAmt, i, tmp
    Sql = Sql & " @InvAmount = " & Val(tmp) & ","
    Sql = Sql & ReturnCommonNamedParametersSqlForSave(CurDate, , , , , , False)
    If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  sShowMessage "Record(s) Saved"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub DTPFrom_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub DTPTo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If TypeOf Me.ActiveControl Is fpSpread Then
 GoTo ExitHere
End If
  If KeyAscii = 13 Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurrDate As Date
  
  Try
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  DTPFrom = CurrDate
  DTPTo = CurrDate
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub SpInv_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim i As Long
  Dim sRow As Long
  
  Try
  Me.MousePointer = vbHourglass
    
  Sql = "select a.InvoiceType,a.InvoiceNo,a.InvoiceDate, a.TotalAmount from SalesInvoice a "
  Sql = Sql + " where a.CustomerCode = " & Val(txtCustName.Tag) & ""
  Sql = Sql + " and a.InvoiceDate between convert(datetime,'" & Format(DTPFrom.Value, "yyyy-MM-dd") & "',111)"
  Sql = Sql + " and convert(datetime,'" & Format(DTPTo.Value, "yyyy-MM-dd") & "',111)"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  Sql = Sql + " and not exists (select * from SAL_Invoice_LCRef d where d.InvNo=a.InvoiceNo and d.InvDate=a.InvoiceDate and d.InvType=a.InvoiceType"
  Sql = Sql + " and d.CustomerCode=a.CustomerCode and d.Branchid=a.Branchid and d.Compid=a.Compid) "
  Sql = Sql + " order by a.InvoiceDate,a.InvoiceType,a.InvoiceNo"
  If FetchValue_New(True, db.ConnectionString, Sql, "1111", "List of Invoice", "Invoice Type", "Invoice No", "Invoice Date", "Invoice Amount") = True Then
    SpInv.MaxRows = SpInv.DataRowCnt + UBound(getvalue()) + 1
    sRow = SpInv.DataRowCnt + 1
    For i = 1 To UBound(getvalue, 1)
      SpInv.SetText eInv.mInvType, sRow, Trim(getvalue(i, 1))
      SpInv.SetText eInv.mInvNo, sRow, getvalue(i, 2)
      SpInv.SetText eInv.mInvDate, sRow, Format(getvalue(i, 3), "dd/MM/yyyy")
      SpInv.SetText eInv.mInvAmt, sRow, getvalue(i, 4)
      If Spread_Clear_Duplicate(SpInv, sRow, eInv.mInvType, eInv.mInvNo, eInv.mInvDate) = False Then
        SpInv.SetText eInv.mInvType, sRow, Trim(getvalue(i, 1))
        SpInv.SetText eInv.mInvNo, sRow, getvalue(i, 2)
        SpInv.SetText eInv.mInvDate, sRow, Format(getvalue(i, 3), "dd/MM/yyyy")
        SpInv.SetText eInv.mInvAmt, sRow, getvalue(i, 4)
        sRow = sRow + 1
      End If
    Next i
    txtTotal.Text = Val(Spread_Col_Total(SpInv, eInv.mInvAmt, 1, SpInv.DataRowCnt))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub SpInv_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  
  If Shift = 1 And KeyCode = vbKeyDelete Then
      SpInv.DeleteRows SpInv.ActiveRow, 1
      SpInv.MaxRows = SpInv.DataRowCnt + 1
    txtTotal.Text = Val(Spread_Col_Total(SpInv, eInv.mInvAmt, 1, SpInv.DataRowCnt))
  End If
  
End Sub

Private Sub txtcustname_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select distinct (select b.PartyName from FAS_PARTY_MASTER b where "
  Sql = Sql + " b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.compid=a.Compid) PartyName,a.CustomerCode "
  Sql = Sql + " from SAL_Customer_ItemMaster a where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer Name") = True Then
    txtCustName.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtCustName.Text = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

  Me.MousePointer = 0
End Try
End Sub

Private Function Validate() As Boolean
  Dim i As Long
  Dim temp As Variant

  Try
  
  Validate = False
  
  If SpInv.DataRowCnt = 0 Then
    sShowMessage "Please select the Invoice Details"
    GoTo ExitHere
  End If
    
  If Val(Spread_Col_Total(SpInv, eInv.mInvAmt, 1, SpInv.DataRowCnt)) > Val(txtLCBalAmt.Text) Then
    sShowMessage "Sum Invoice Amount (" & Val(Spread_Col_Total(SpInv, eInv.mInvAmt, 1, SpInv.DataRowCnt)) & ") should be less than or equal to LC Balance Amount"
    Spread_Set_Focus SpInv, 1, 1
    GoTo ExitHere
  End If

  Validate = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub txtCustName_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
     txtcustname_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtLCNo_DblClick()
  Dim Sql As String
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " select * from ("
  Sql = Sql & " select a.LCNo, a.LCDate, a.LCAmount, (a.LCAmount -"
  Sql = Sql & " isnull((select "
  Sql = Sql & " sum(case when c.AcceptedQty > 0 then ((c.AcceptedQty/c.InvoiceQty)*b.InvAmount) else b.InvAmount End) InvAmount"
  Sql = Sql & " from SAL_Invoice_LCRef b"
  Sql = Sql & " left outer join SAL_Invoice_Dtl c"
  Sql = Sql & " on c.InvoiceNo=b.InvNo and c.InvoiceDate=b.InvDate and c.InvoiceType=b.InvType"
  Sql = Sql & " and c.Branchid=b.Branchid and c.Compid=b.Compid "
  Sql = Sql & " where b.CustomerCode=a.CustomerCode"
  Sql = Sql & " and b.LCNo=a.LCNo and b.LCDate=a.LCDate and b.Branchid=a.Branchid and b.Compid=a.Compid),0)) as BalanceAmt"
  Sql = Sql & " from SAL_LC_Dtl a"
  Sql = Sql & " where a.CustomerCode=" & Val(txtCustName.Tag) & " and a.Branchid ='" & gBranchID & "' and a.Compid = '" & gCompID & "') x"
  Sql = Sql & " where BalanceAmt > 0"
  If FetchValue_New(False, db.ConnectionString, Sql, "1110", "List of LC's", "LC No", "LC Date", "LC Amount") = True Then
    txtLCNo.Text = getvalue(1)
    dtpLCDate.Value = getvalue(2)
    txtLCAmt.Text = Format(getvalue(3), "0.00")
    txtLCBalAmt.Text = Format(getvalue(4), "0.00")
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtLCNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  If KeyCode = vbKeyF2 Then
 txtLCNo_DblClick
End If
End Sub

Private Sub Design_SpInv(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SpInv


SpInv.DblClick +=  new EventHandler(SpInv_DblClick)
SpInv.KeyDown +=  new EventHandler(SpInv_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_SpInv("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\LCNoInvoiceUpdate.frm", SpInv, "")

End Sub
