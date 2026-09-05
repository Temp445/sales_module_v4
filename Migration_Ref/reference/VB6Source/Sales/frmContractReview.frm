VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContractReview 
   Caption         =   "Contract Review"
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6690
      Width           =   1300
   End
   Begin VB.Frame fraHdr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   2345
      TabIndex        =   3
      Top             =   105
      Width           =   7140
      Begin VB.TextBox txtItemcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3450
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
         Width           =   1890
      End
      Begin VB.TextBox txtIEItemcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5370
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
         Width           =   1425
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   6555
         TabIndex        =   10
         ToolTipText     =   "Click to OK"
         Top             =   1455
         Width           =   495
      End
      Begin VB.TextBox txtDrawingno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5310
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Enter the Remarks"
         Top             =   1455
         Width           =   1185
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "Enter the Ref. No"
         Top             =   915
         Width           =   3330
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2955
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   330
         Width           =   3825
      End
      Begin VB.TextBox txtDocNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   825
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   6
         ToolTipText     =   "Jobwork Enquiry No"
         Top             =   345
         Width           =   795
      End
      Begin VB.ComboBox cmbDocType 
         Height          =   330
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select Jobwork Enquiry Type"
         Top             =   330
         Width           =   720
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Enter the Remarks"
         Top             =   1455
         Width           =   5190
      End
      Begin MSComCtl2.DTPicker dtpDocDate 
         Height          =   315
         Left            =   1665
         TabIndex        =   13
         ToolTipText     =   "Jobwork Enquiry Date"
         Top             =   345
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   22347777
         CurrentDate     =   37258
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order Ref No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "PartNo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3450
         TabIndex        =   21
         Top             =   735
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2955
         TabIndex        =   20
         Top             =   135
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "IE PartNo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5370
         TabIndex        =   19
         Top             =   735
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "Drawing Rev.No"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5295
         TabIndex        =   18
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc.Type"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   135
         Width           =   690
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Doc.No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   825
         TabIndex        =   16
         Top             =   135
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Doc.Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1680
         TabIndex        =   15
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label11 
         Caption         =   "Description"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   1275
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5265
      TabIndex        =   2
      ToolTipText     =   "Click to Clear the Data"
      Top             =   6690
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3975
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to Save the Data"
      Top             =   6690
      Width           =   1300
   End
   Begin VB.Frame fradtl 
      Height          =   4425
      Left            =   1445
      TabIndex        =   23
      Top             =   2040
      Width           =   8940
Begin AceSpread spDetails
         Height          =   4140
         Left            =   60
         TabIndex        =   24
         Top             =   150
         Width           =   8805
         _Version        =   458752
         _ExtentX        =   15531
         _ExtentY        =   7302
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
         MaxCols         =   6
         ScrollBars      =   2
      End
   End
End
Attribute VB_Name = "frmContractReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim CurrDate As Date
  Dim ret As Integer
  Dim tmp As Variant
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  
  Try
  
  ret = 0
    For i = 1 To spDetails.DataRowCnt
      spDetails.GetText 3, i, tmp
      ret = ret + Val(tmp)
      
      spDetails.GetText 4, i, tmp
      ret = ret + Val(tmp)
      
      spDetails.GetText 5, i, tmp
      ret = ret + Val(tmp)
      
      If Val(ret) <> 1 Then
        ret = 0
        sShowMessage "Please select any one option"
        Spread_Set_Focus spDetails, i, 3
        Chkbefsave = False
        GoTo ExitHere
      End If
      ret = 0
    Next i
  
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Load_Spr()
  Dim sql As String
  Dim i As Integer
  Dim rs As AceADODB.Recordset
  
  Try

  sql = "Select Code,Description from TermsMaster where TypeFlg = 'Contract Review' "
  sql = sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
  i = 1
  spDetails.MaxRows = rs.RecordCount
  Do While Not rs.EOF
    spDetails.SetText 1, i, rs("Code")
    spDetails.SetText 2, i, rs("Description")
    spDetails.Row = i
spDetails.Col = 3
spDetails.CellType = CellTypeCheckBox
spDetails.TypeVAlign = TypeVAlignCenter
spDetails.TypeHAlign = TypeHAlignCenter

    spDetails.Row = i
spDetails.Col = 4
spDetails.CellType = CellTypeCheckBox
spDetails.TypeVAlign = TypeVAlignCenter
spDetails.TypeHAlign = TypeHAlignCenter

    spDetails.Row = i
spDetails.Col = 5
spDetails.CellType = CellTypeCheckBox
spDetails.TypeVAlign = TypeVAlignCenter
spDetails.TypeHAlign = TypeHAlignCenter
    rs.MoveNext
    i = i + 1
  Loop
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmbDocType_Click()
  Try
  Me.MousePointer = vbHourglass
  
  Call SetDocNo
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdClear_Click()
  Dim CurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  txtDocNo.Text = ""
  txtDrawingno.Text = ""
  txtIEItemcode.Text = ""
  txtItemcode.Text = ""
  txtRefNo.Text = ""
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpDocDate = CurrDate
  fraHdr.Enabled = True
  fradtl.Enabled = False
  
  Call Spread_Clear(spDetails)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetDocNo()
  Try

  txtDocNo = GetDocNo("SAL_ContractReview_HDR", "DocType", "DocNo", Trim$(cmbDocType), "CR")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdExit_Click()
  Try
  
  Unload Me
  
  Exit Sub
Catch ex As Exception
  Call ErrMsgShow(ex)
  
End Sub

Private Sub cmdOK_Click()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If contrl_valid(txtCustomer, "Select the Customer Name") = False Then
 GoTo ExitHere
End If
 
  Call Load_Spr
  
  fraHdr.Enabled = False
  fradtl.Enabled = True
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim sql As String
  Dim i As Long
  Dim tmp As Variant
  Dim rCnt As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  If Chkbefsave = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetDocNo
  
  sql = "Execute SAL_Ins_ContractReview_Hdr "
  sql = sql & " @DocType='" & Trim$(cmbDocType) & "',"
  sql = sql & " @DocNo=" & Val(txtDocNo) & ","
  sql = sql & " @DocDate='" & Format(dtpDocDate, "yyyy-MM-dd") & "',"
  sql = sql & " @CustomerCode=" & Val(txtCustomer.Tag) & ","
  sql = sql & " @RefNo='" & Trim$(txtRefNo) & "',"
  sql = sql & " @ItemCode='" & Trim$(txtIEItemcode) & "',"
  sql = sql & " @UserID='" & gUserID & "',"
  sql = sql & " @BranchID='" & gBranchID & "',"
  sql = sql & " @CompID='" & gCompID & "',"
  sql = sql & " @Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    For rCnt = 1 To spDetails.DataRowCnt
      sql = "execute SAL_Ins_ContractReview_Dtl"
      sql = sql & " @DocType='" & Trim$(cmbDocType) & "',"
      sql = sql & " @DocNo=" & Val(txtDocNo) & ","
      sql = sql & " @DocDate='" & Format(dtpDocDate, "yyyy-MM-dd") & "',"
      spDetails.GetText 1, rCnt, tmp
      sql = sql & " @ReviewNo=" & Val(tmp) & ","
      
      spDetails.GetText 3, rCnt, tmp
      sql = sql & " @OK=" & Val(tmp) & ","
      
      spDetails.GetText 4, rCnt, tmp
      sql = sql & " @AMD=" & Val(tmp) & ","
      
      spDetails.GetText 5, rCnt, tmp
      sql = sql & " @NA=" & Val(tmp) & ","
      
      spDetails.GetText 6, rCnt, tmp
      sql = sql & " @ActionNeed='" & Trim$(tmp) & "',"
      
      sql = sql & " @UserID='" & gUserID & "',"
      sql = sql & " @BranchID='" & gBranchID & "',"
      sql = sql & " @CompID='" & gCompID & "',"
      sql = sql & " @Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If

  sShowMessage "Order Number :" & Trim$(cmbDocType) & Trim$(txtDocNo) & "  Is Generated"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurrDate As Date
  
  Try

  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpDocDate = CurrDate
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtCustomer_DblClick()
  Dim sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  
  sql = "select distinct b.PartyName, a.CustomerCode, a.ReferenceNo, a.PartNo, a.ItemDescription, "
  sql = sql & " a.ItemCode, c.RevisionNo as DrawingRevNo from "
  sql = sql & " fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')  a"
  sql = sql & " left outer join FAS_Party_Master b on a.CustomerCode = b.PartyCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " left outer join DrawingMaster c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  sql = sql & " where a.Status='A' and a.ItemStatus <> 'C' and a.Finyear = '" & gFinyear & "'"
  sql = sql & " and a.ItemCode not in (select d.ItemCode from SAL_ContractReview_Hdr d where d.ItemCode = a.Itemcode and d.Branchid =a.Branchid and d.Compid = a.Compid)"
  If FetchValue_New(False, db.ConnectionString, sql, "101111", "List of Customers", "Customer Name", "PO Ref No", "Part No", "Part Description", "Item Code") = True Then
    txtCustomer.Tag = getvalue(2)
    txtCustomer = getvalue(1)
    txtRefNo = getvalue(3)
    txtItemcode = getvalue(4)
    txtdesc = getvalue(5)
    txtIEItemcode = getvalue(6)
    txtDrawingno = getvalue(7)
  End If
  
  
  
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spDetails(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDetails

End Sub

Private Sub InitializeSpreads()

	Call Design_spDetails("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmContractReview.frm", spDetails, "")

End Sub
