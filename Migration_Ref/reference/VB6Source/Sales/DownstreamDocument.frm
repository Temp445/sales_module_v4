VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDownstreamDocument 
   Caption         =   "Customer Drawing Downstream Document"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6300
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
      Left            =   795
      TabIndex        =   3
      Top             =   75
      Width           =   10245
      Begin VB.ComboBox cmbDisposition 
         Height          =   330
         ItemData        =   "DownstreamDocument.frx":0000
         Left            =   2805
         List            =   "DownstreamDocument.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1455
         Width           =   3765
      End
      Begin VB.TextBox txtRevDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8445
         MaxLength       =   50
         TabIndex        =   24
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
         Width           =   1470
      End
      Begin VB.TextBox txtItemcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
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
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Enter the Remarks"
         Top             =   1470
         Width           =   1425
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         ToolTipText     =   "Click to OK"
         Top             =   1455
         Width           =   495
      End
      Begin VB.TextBox txtDrawingNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7215
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
         Width           =   1215
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6795
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "Enter the Ref. No"
         Top             =   330
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
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
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
      Begin MSComCtl2.DTPicker dtpEffDate 
         Height          =   315
         Left            =   105
         TabIndex        =   26
         ToolTipText     =   "Jobwork Enquiry Date"
         Top             =   1470
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22347777
         CurrentDate     =   37258
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Obsolete Drawing Disposition"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2805
         TabIndex        =   29
         Top             =   1245
         Width           =   2130
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Effective From"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   105
         TabIndex        =   27
         Top             =   1260
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Drawing Rev. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   8445
         TabIndex        =   25
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order Ref No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6795
         TabIndex        =   22
         Top             =   135
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "PartNo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   105
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
         Left            =   1365
         TabIndex        =   19
         Top             =   1290
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "Drawing Rev.No"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   7215
         TabIndex        =   18
         Top             =   720
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
         Left            =   2010
         TabIndex        =   14
         Top             =   735
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5267
      TabIndex        =   2
      ToolTipText     =   "Click to Clear the Data"
      Top             =   6300
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3975
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to Save the Data"
      Top             =   6300
      Width           =   1300
   End
   Begin VB.Frame fradtl 
      Height          =   4230
      Left            =   1605
      TabIndex        =   23
      Top             =   1965
      Width           =   8625
Begin AceSpread spDocument
         Height          =   3885
         Left            =   120
         TabIndex        =   30
         Top             =   225
         Width           =   8370
         _Version        =   458752
         _ExtentX        =   14764
         _ExtentY        =   6853
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
         MaxCols         =   3
         ScrollBars      =   2
      End
   End
End
Attribute VB_Name = "frmDownstreamDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Function Chkbefsave() As Boolean
  
  Try
  
  If spDocument.DataRowCnt = 0 Then
    sShowMessage "Please select alteast one PPAP Document"
    Chkbefsave = False
    spDocument.SetFocus
    Spread_Set_Focus spDocument, 1, 2
    GoTo ExitHere
  End If
  
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
  spDocument.MaxRows = rs.RecordCount
  Do While Not rs.EOF
    spDocument.SetText 1, i, rs("Code")
    spDocument.SetText 2, i, rs("Description")
    spDocument.Row = i
spDocument.Col = 3
spDocument.CellType = CellTypeCheckBox
spDocument.TypeVAlign = TypeVAlignCenter
spDocument.TypeHAlign = TypeHAlignCenter

    spDocument.Row = i
spDocument.Col = 4
spDocument.CellType = CellTypeCheckBox
spDocument.TypeVAlign = TypeVAlignCenter
spDocument.TypeHAlign = TypeHAlignCenter

    spDocument.Row = i
spDocument.Col = 5
spDocument.CellType = CellTypeCheckBox
spDocument.TypeVAlign = TypeVAlignCenter
spDocument.TypeHAlign = TypeHAlignCenter
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
  txtDrawingNo.Text = ""
  txtIEItemcode.Text = ""
  txtItemcode.Text = ""
  txtRefNo.Text = ""
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpDocDate = CurrDate
  fraHdr.Enabled = True
  fradtl.Enabled = False
  
  Call Spread_Clear(spDocument)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetDocNo()
  Try

  txtDocNo = GetDocNo("SAL_DownstreamDocument_HDR", "DocType", "DocNo", Trim$(cmbDocType), "DD")
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

Private Sub cmdOk_Click()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If contrl_valid(txtCustomer, "Select the Customer Name") = False Then
 GoTo ExitHere
End If
 
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
  
  sql = "Execute SAL_Ins_DownstreamDocument_Hdr "
  sql = sql & " @DocType='" & Trim$(cmbDocType) & "',"
  sql = sql & " @DocNo=" & Val(txtDocNo) & ","
  sql = sql & " @DocDate='" & Format(dtpDocDate, "yyyy-MM-dd") & "',"
  sql = sql & " @CustomerCode=" & Val(txtCustomer.Tag) & ","
  sql = sql & " @RefNo='" & Trim$(txtRefNo) & "',"
  sql = sql & " @ItemCode='" & Trim$(txtIEItemcode) & "',"
  sql = sql & " @EffectiveFrom='" & Trim$(dtpEffDate) & "',"
  sql = sql & " @Disposition='" & Trim$(cmbDisposition) & "',"
  sql = sql & " @UserID='" & gUserID & "',"
  sql = sql & " @BranchID='" & gBranchID & "',"
  sql = sql & " @CompID='" & gCompID & "',"
  sql = sql & " @Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    For rCnt = 1 To spDocument.DataRowCnt
      sql = "execute SAL_Ins_DownstreamDocument_Dtl"
      sql = sql & " @DocType='" & Trim$(cmbDocType) & "',"
      sql = sql & " @DocNo=" & Val(txtDocNo) & ","
      sql = sql & " @DocDate='" & Format(dtpDocDate, "yyyy-MM-dd") & "',"
      sql = sql & " @ItemCode='" & Trim$(txtIEItemcode) & "',"
      
      spDocument.GetText 1, rCnt, tmp
      sql = sql & " @DocumentNo='" & Trim$(tmp) & "',"
      
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

  sShowMessage "Doc No :" & Trim$(cmbDocType) & Trim$(txtDocNo) & "  Is Generated"
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
  dtpEffDate = CurrDate
  cmbDisposition.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spDocument_DblClick(sender as object, e as DblClickEventArgs)
  Dim sql As String
  Dim CurDate As Date
  Dim i As Long

  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate

  sql = " select a.StandardLName, a.DocType, a.PreparedBy, b.UserNAme from PPAP a"
  sql = sql & " left outer join fn_SEC_Users('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "') b"
  sql = sql & " on a.PreparedBy = b.Userid and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " where a.CustomerCode = " & Val(txtCustomer.Tag) & " and a.ItemCode ='" & Trim$(txtIEItemcode) & "'"
  If FetchValue_New(True, db.ConnectionString, sql, "1001", "List of PPAP Documents", "Standard Name", "Prepared By") = True Then
    spDocument.MaxRows = spDocument.DataRowCnt + UBound(getvalue, 1) + 1
    For i = 1 To UBound(getvalue, 1)
      spDocument.SetText 1, i, getvalue(i, 2)
      spDocument.SetText 2, i, getvalue(i, 1)
      spDocument.SetText 3, i, getvalue(i, 3)
      spDocument.SetText 4, i, getvalue(i, 4)
      Avoid_Duplicate spDocument, 1
    Next i
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustomer_DblClick()
  Dim sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  
  sql = "select distinct b.PartyName, a.CustomerCode, a.ReferenceNo, a.PartNo, a.ItemDescription, "
  sql = sql & " a.ItemCode, c.RevisionNo as DrawingRevNo, c.RevisionDate from "
  sql = sql & " fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')  a"
  sql = sql & " left outer join FAS_Party_Master b on a.CustomerCode = b.PartyCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " left outer join DrawingMaster c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  sql = sql & " where a.Status='A' and a.ItemStatus <> 'C' and a.Finyear = '" & gFinyear & "'"
  sql = sql & " and a.ItemCode not in (select d.ItemCode from SAL_DownstreamDocument_Hdr d where d.ItemCode = a.Itemcode and d.Branchid =a.Branchid and d.Compid = a.Compid)"
  If FetchValue_New(False, db.ConnectionString, sql, "101111", "List of Customers", "Customer Name", "PO Ref No", "Part No", "Part Description", "Item Code") = True Then
    txtCustomer.Tag = getvalue(2)
    txtCustomer = getvalue(1)
    txtRefNo = getvalue(3)
    txtItemcode = getvalue(4)
    txtdesc = getvalue(5)
    txtIEItemcode = getvalue(6)
    txtDrawingNo = getvalue(7)
    txtRevDate = getvalue(8)
  End If
  
  
  
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spDocument(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spDocument


spDocument.DblClick +=  new EventHandler(spDocument_DblClick)

End Sub

Private Sub InitializeSpreads()

	Call Design_spDocument("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\DownstreamDocument.frm", spDocument, "")

End Sub
