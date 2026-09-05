VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCForm 
   Caption         =   "C Form"
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
   Begin VB.Frame FraInvoiceDet 
      Caption         =   "Bill Details"
      Height          =   2025
      Left            =   2790
      TabIndex        =   14
      Top             =   3735
      Width           =   6135
Begin AceSpread sprInv
         Height          =   1695
         Left            =   135
         TabIndex        =   15
         Top             =   210
         Width           =   5850
         _Version        =   458752
         _ExtentX        =   10319
         _ExtentY        =   2990
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
   Begin VB.Frame fraDet 
      Caption         =   "Party Details"
      Height          =   2295
      Left            =   2985
      TabIndex        =   3
      Top             =   1410
      Width           =   5730
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   1800
         Width           =   660
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Top             =   1410
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         Format          =   166395907
         CurrentDate     =   39291
      End
      Begin MSComCtl2.DTPicker dtpregistraionDate 
         Height          =   330
         Left            =   1440
         TabIndex        =   11
         Top             =   1005
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         Format          =   166395907
         CurrentDate     =   39291
      End
      Begin VB.TextBox txtRegistrationNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   7
         Top             =   630
         Width           =   2490
      End
      Begin VB.TextBox txtPartyName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   255
         Width           =   4140
      End
      Begin MSComCtl2.DTPicker dtpValidTo 
         Height          =   330
         Left            =   1440
         TabIndex        =   13
         Top             =   1815
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         Format          =   166395907
         CurrentDate     =   39291
      End
      Begin VB.Label lblValidTo 
         AutoSize        =   -1  'True
         Caption         =   "Valid To"
         Height          =   210
         Left            =   795
         TabIndex        =   10
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblValidFrom 
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
         Height          =   210
         Left            =   615
         TabIndex        =   9
         Top             =   1395
         Width           =   765
      End
      Begin VB.Label lblRegistrationDate 
         AutoSize        =   -1  'True
         Caption         =   "Registration Date"
         Height          =   210
         Left            =   150
         TabIndex        =   8
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label lblRegistrationNo 
         AutoSize        =   -1  'True
         Caption         =   "* Registration No"
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   645
         Width           =   1200
      End
      Begin VB.Label lblPartyName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "* Party Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   450
         TabIndex        =   4
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6525
      TabIndex        =   0
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5850
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5235
      TabIndex        =   1
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5850
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3945
      TabIndex        =   2
      ToolTipText     =   "Click to Save the Data"
      Top             =   5850
      Width           =   1300
   End
End
Attribute VB_Name = "frmCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eBill
  mInvType = 1
  mInvNo = 2
  mInvDate = 3
  mInvAmount = 4
End Enum

Private Sub CmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  txtPartyName.Text = ""
  txtRegistrationNo.Text = ""
  
  dtpregistraionDate.Value = ServerDate
  dtpValidFrom.Value = dtpregistraionDate.Value
  dtpValidTo.Value = dtpValidFrom.Value
  
  Call Spread_Clear(sprInv)
  sprInv.MaxRows = 1
 
  fraDet.Enabled = True
  FraInvoiceDet.Enabled = False
  cmdSave.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub Cmdok_Click()
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Integer
  
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  If contrl_valid(txtPartyName, "Select PartyName") = False Then GoTo ExitHere
  If contrl_valid(txtRegistrationNo, "Enter RegistrationNo") = False Then GoTo ExitHere
  
  CurDate = ServerDate
  
  If gModuleCode = eMdlCode.mPUR Then
    sql = " Select distinct a.DocType, a.DocNo, a.DocDate, a.InvoiceAmount as Amt"
    sql = sql & " From Pur_PurchaseInvoice_Hdr a"
    sql = sql & " Left Outer Join dbo.fn_PUR_PurchaseOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
    sql = sql & " On a.RefNo = b.OrderNo and a.RefDate = b.OrderDate and a.RefType = b.OrderType"
    sql = sql & " and a.VendorCode = b.VendorCode and a.BranchID = b.BranchID and a.CompID = b.CompID"
    sql = sql & " Where a.RefType <> '' and a.RefNo <> 0 and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
    sql = sql & " and (a.DocType + Convert(varchar, a.DocNo) + Convert(Varchar(10), Convert(datetime, a.DocDate), 111))"
    sql = sql & " not in (Select (DocType + Convert(varchar, DocNo) + Convert(Varchar(10), Convert(datetime, DocDate), 111))"
    sql = sql & " From Com_CForm_Tran Where ModuleCode = " & Val(gModuleCode) & "  and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and PartyCode = " & txtPartyName.Tag & ")"
    sql = sql & " and b.FormCode in (Select OptionValue From Com_Report_Option Where Code = 1 and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "')"
    sql = sql & " and a.VendorCode =  = " & txtPartyName.Tag & " " 'and a.Finyear = '" & gFinyear & "'"
  ElseIf gModuleCode = eMdlCode.mSAL Then
    sql = " Select distinct a.InvoiceType as DocType, a.InvoiceNo as DocNo, a.InvoiceDate as DocDate, b.TotalAmount as Amt"
    sql = sql & " From Sal_Invoice_Dtl a"
    sql = sql & " Left Outer Join Sal_Invoice_Hdr b"
    sql = sql & " On a.InvoiceDate = b.InvoiceDate and  a.InvoiceNo = b.InvoiceNo and a.InvoiceType = b.InvoiceType"
    sql = sql & " and a.BranchID = b.BranchID and a.CompID = b.CompID"
    sql = sql & " Left Outer Join dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b1"
    sql = sql & " On a.OrderNo = b1.OrderNo and a.OrderDate = b1.OrderDate and a.OrderType = b1.OrderType"
    sql = sql & " and b.CustomerCode = b1.CustomerCode and a.BranchID = b1.BranchID and a.CompID = b1.CompID"
    sql = sql & " Where a.OrderType <> '' and a.OrderNo <> 0 and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
    sql = sql & " and (a.InvoiceType + Convert(varchar, a.InvoiceNo) + Convert(Varchar(10), Convert(datetime, a.InvoiceDate), 111))"
    sql = sql & " not in (Select (DocType + Convert(varchar, DocNo) + Convert(Varchar(10), Convert(datetime, DocDate), 111))"
    sql = sql & " From Com_CForm_Tran Where ModuleCode = " & Val(gModuleCode) & "  and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and PartyCode = " & txtPartyName.Tag & ")"
    sql = sql & " and b1.FormCode in (Select OptionValue From Com_Report_Option Where Code = 1"
    sql = sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "') and b.CustomerCode = " & txtPartyName.Tag & " " ' and a.Finyear = '" & gFinyear & "'"
  End If
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
      sprInv.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprInv.SetText eBill.mInvType, rCnt, rs("DocType")
        sprInv.SetText eBill.mInvNo, rCnt, rs("DocNo")
        sprInv.SetText eBill.mInvDate, rCnt, rs("DocDate")
        sprInv.SetText eBill.mInvAmount, rCnt, rs("Amt")
        rs.MoveNext
      Next rCnt
    fraDet.Enabled = False
    FraInvoiceDet.Enabled = True
    cmdSave.Enabled = True
  Else
    sShowMessage "No Bill Details Found For the Party " & Trim$(txtPartyName.Text)
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim Sql As String
  Dim CurDate As Date
  Dim rCnt As Integer
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
    For rCnt = 1 To sprInv.DataRowCnt
      sql = " Execute Com_Ins_CForm_Tran "
      sql = sql & " @ModuleCode = " & Val(gModuleCode) & ","
      sprInv.GetText eBill.mInvType, rCnt, retVal
      sql = sql & " @DocType = '" & Trim$(retVal) & "',"
      sprInv.GetText eBill.mInvNo, rCnt, retVal
      sql = sql & " @DocNo = " & Val(retVal) & ","
      sprInv.GetText eBill.mInvDate, rCnt, retVal
      sql = sql & " @DocDate = '" & Format(retVal, "yyyy-MM-dd") & "',"
      sql = sql & " @PartyCode = " & Val(txtPartyName.Tag) & ","
      sql = sql & " @RegistrationNo = '" & Trim$(txtRegistrationNo.Text) & "',"
      sql = sql & " @RegistrationDate = '" & Format(dtpregistraionDate.Value, "yyyy-MM-dd") & "',"
      sql = sql & " @ValidFrom = '" & Format(dtpValidFrom.Value, "yyyy-MM-dd") & "',"
      sql = sql & " @ValidTo = '" & Format(dtpValidTo.Value, "yyyy-MM-dd") & "',"
      sql = sql & " @isIssued = " & IIf(gModuleCode = eMdlCode.mSAL, 1, 0) & ","
      sprInv.GetText eBill.mInvAmount, rCnt, retVal
      sql = sql & " @InvoiceAmount = " & Val(retVal) & ","
      sql = sql & " @UserID = '" & gUserID & "',"
      sql = sql & " @BranchID = '" & gBranchID & "',"
      sql = sql & " @CompID = '" & gCompID & "',"
      sql = sql & " @TranStamp = '" & Format(CurDate, "yyyy-MM-dd HH:MM:ss") & "',"
      sql = sql & " @FinYear = '" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
    Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage " Record Saved."
  Call CmdClear_Click
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
  Me.MousePointer = vbHourglass
  
  dtpregistraionDate.Value = ServerDate
  dtpValidFrom.Value = dtpregistraionDate.Value
  dtpValidTo.Value = dtpValidFrom.Value

  If gModuleCode = eMdlCode.mSAL Then 'Sales
    lblPartyName.Caption = "  Issued To"
  ElseIf gModuleCode = eMdlCode.mPUR Then  'Purchase
    lblPartyName.Caption = "Received From"
  End If
  
  FraInvoiceDet.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub txtPartyName_DblClick()
  Dim sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  
'  Sql = " Select distinct b.PartyName, x.Party From ("
'  If gModuleCode = eMdlCode.mPUR Then
'    Sql = Sql & " Select distinct a.VendorCode as Party, FormCode, BranchID, CompID, Finyear"
'    Sql = Sql & " From dbo.fn_PUR_PurchaseOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
'  ElseIf gModuleCode = eMdlCode.mSAL Then
'    Sql = Sql & " Select distinct a.CustomerCode as Party, FormCode, BranchID, CompID, Finyear"
'    Sql = Sql & " From dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
'  End If
'  Sql = Sql & " Where a.FormCode in (Select OptionValue From Com_Report_Option Where Code = 1"
'  Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "') "
'  Sql = Sql & " )x"
'  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
'  Sql = Sql & " On x.Party = b.PartyCode and x.BranchID = b.BranchID and x.CompID = b.CompID"
'  Sql = Sql & " Where b.Branchid ='" & gBranchID & "' and b.Compid ='" & gCompID & "'" 'and x.Finyear = '" & gFinyear & "'
  
'  If gModuleCode = eMdlCode.mPUR Then
'    Sql = Sql & " and b.Actype in ('P','Y') and b.EffStatus = 1"
'  ElseIf gModuleCode = eMdlCode.mSAL Then
'    Sql = Sql & " and b.Actype in ('D','Y') and b.EffStatus = 1"
'  End If
'  Sql = Sql & " Order by b.PartyName"
  
  sql = " Select LedgerName, LedgerCode From FAS_LedgerMaster"
  sql = sql & " Where CompCode = '" & gFasCompCode & "'"
  If gModuleCode = eMdlCode.mPUR Then
    sql = sql & " and LrRoute = '0303'"
  ElseIf gModuleCode = eMdlCode.mSAL Then
    sql = sql & " and LrRoute = '0206'"
  End If
  If FetchValue_New(False, db.ConnectionString, sql, "100000", "List of Accounts", "Account Name") = True Then
    txtPartyName.Text = getvalue(1)
    txtPartyName.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function ValidateData() As Boolean
  
  Try
  ValidateData = False
  
  If contrl_valid(txtPartyName, "Select Party Name") = False Then GoTo ExitHere
  If contrl_valid(txtRegistrationNo, "Enter RegistrationNo") = False Then GoTo ExitHere

  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprInv(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprInv

End Sub

Private Sub InitializeSpreads()

	Call Design_sprInv("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\frmCForm.frm", sprInv, "")

End Sub
