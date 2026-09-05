VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptDownstreamDocument 
   Caption         =   "Downstream Document"
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
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6315
      TabIndex        =   0
      Top             =   3555
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5025
      TabIndex        =   1
      Top             =   3555
      Width           =   1300
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
      Left            =   7350
      TabIndex        =   14
      Top             =   1920
      Width           =   345
   End
   Begin VB.Frame frastep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   4080
      TabIndex        =   7
      Top             =   1695
      Width           =   3195
      Begin VB.TextBox txtDocNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   9
         ToolTipText     =   "Enter Request No"
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtdoctype 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   8
         ToolTipText     =   "Enter Request No"
         Top             =   330
         Width           =   885
      End
      Begin MSComCtl2.DTPicker dtpDocDate 
         Height          =   315
         Left            =   1905
         TabIndex        =   10
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   56033281
         CurrentDate     =   37313
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   13
         Top             =   135
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1905
         TabIndex        =   12
         Top             =   135
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   990
         TabIndex        =   11
         Top             =   135
         Width           =   570
      End
   End
   Begin VB.Frame fraReportType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4110
      TabIndex        =   3
      Top             =   2475
      Width           =   3135
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   585
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.ComboBox cmbType 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "rptDownstreamDocument.frx":0000
         Left            =   1035
         List            =   "rptDownstreamDocument.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3735
      TabIndex        =   2
      Top             =   3555
      Width           =   1300
   End
End
Attribute VB_Name = "rptDownstreamDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sprow As Long

Private Erpno As String
Private TsNo As String

Public Sub Header()
  Dim Spr As AceSpread
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim sdate As Date
  Dim SValue As Double
  Dim tmp As Variant
  Dim CurDate As Date
  
  Try
  
  Set Spr = frmprint.sprPrint
  
  CurDate = ServerDate
  
    
    sql = "select distinct b.PartyName, a.ItemCode, d.PartNo, d.ItemDescription, c.RevisionNo as DrawingRevNo, c.RevisionDate, a.EffectiveFrom, a.Disposition from "
    sql = sql & " fn_SAL_DownstreamDocument_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join "
    sql = sql & " fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')  d"
    sql = sql & " on a.CustomerCode = d.CustomerCode and a.ItemCode = d.ItemCode and a.Branchid=d.Branchid and a.Compid = d.Compid and a.Finyear = d.Finyear"
    sql = sql & " left outer join FAS_Party_Master b on a.CustomerCode = b.PartyCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
    sql = sql & " left outer join DrawingMaster c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
    sql = sql & " where a.DocType='" & txtdoctype & "' and a.Docno=" & txtDocNo & " and a.DocDate='" & Format(dtpDocDate, "yyyy-MM-dd") & "' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      
      Call spSetColWidth(Spr, 1, 6, 2, 34, 3, 20, 4, 20)
      
      Call RptStandard_Header(4, sprow, Erpno, TsNo, IIf(chkPrintDate, True, False))
      GridCount = GridCount + 6
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 4, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Customer Name: " & rs("PartyName"), mAlignLeftCenter, True

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderTop)
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 4, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Part No.: " & rs("PartNo"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 4, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Part Description: " & rs("ItemDescription"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Latest Drawing Revision No.: " & rs("DrawingRevNo"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      
      SpSetCellSpan frmprint.sprPrint, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 3, sprow, "Drawing Received on: " & rs("RevisionDate"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Effective From: " & rs("EffectiveFrom"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      
      SpSetCellSpan frmprint.sprPrint, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 3, sprow, "IE Product No.: " & rs("ItemCode"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 4, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Obsolete Drawing Disposition: " & rs("Disposition"), mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      SpSetCellSpan frmprint.sprPrint, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "Sl No.", mAlignCenterCenter, True
      spSetText frmprint.sprPrint, 2, sprow, "Related Documents", mAlignCenterCenter, True
      spSetText frmprint.sprPrint, 3, sprow, "Prepared By", mAlignCenterCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 2, 1, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "UPDATION OF:", mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrent() As Boolean
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim i As Integer
  Dim tmp As Variant
  Dim Spr As AceSpread
  Dim CurDate As Date
  Dim j As Long
  
  
  Try
  
  
  CurDate = ServerDate

  sql = " select a.DocumentNo, b.StandardLName, b.PreparedBy, c.UserName from  "
  sql = sql & " fn_SAL_DownstreamDocument_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  sql = sql & " left outer join PPAP b on a.ItemCode = b.ItemCode and a.DocumentNo = b.DocType and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " left outer join fn_SEC_Users('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "') c"
  sql = sql & " on c.Userid=b.PreparedBy and c.Branchid = b.Branchid and c.Compid = b.Compid where a.ItemCode = '" & Trim$(txtdoctype.Tag) & "'"
  sql = sql & " and a.DocType = '" & Trim$(txtdoctype.Text) & "' and a.DocNo = " & Val(txtDocNo.Text) & " and a.DocDate = '" & Format(dtpDocDate, "yyyy-MMM-dd") & "'"
  sql = sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 41
Orient = 1
    Set Spr = frmprint.sprPrint

    gStrReportHeading = "CUSTOMER DRAWING  DOWNSTREAM DOCUMENT UPDATION"
  
    Call Header

      For j = 1 To rs.RecordCount
        SpSetCellSpan frmprint.sprPrint, 3, sprow, 2, 1
        spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, False
        spSetText frmprint.sprPrint, 1, sprow, Val(j), mAlignLeftCenter, True
        spSetText frmprint.sprPrint, 2, sprow, rs("StandardLName"), mAlignLeftCenter, True
        spSetText frmprint.sprPrint, 3, sprow, CheckNull(rs("UserName")), mAlignLeftCenter, True
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        rs.MoveNext
      Next j
          
    
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 2, 1, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "*RIS-Receiving Inspection Standards", mAlignLeftCenter, True
      spSetText frmprint.sprPrint, 3, sprow, "Mention NR-If Not Required for", mAlignLeftCenter, False
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 2, 1, 3, sprow, 2, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 3, sprow, "enabling Futher Processes", mAlignLeftCenter, False
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderRight)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      
      SpSetCellSpan frmprint.sprPrint, 1, sprow, 4, 1
      spSetFont frmprint.sprPrint, 1, Spr.MaxCols, sprow, sprow, "Times New Roman", 12, True
      spSetText frmprint.sprPrint, 1, sprow, "All the Above Updation & Other Requirements are Complied with on (Date)", mAlignLeftCenter, True
      sprow = sprow + 4
GridCount = GridCount + 4
RowProcess Me, sprow
      
      
      spSetText frmprint.sprPrint, 1, sprow, "Quality Dept", mAlignLeftCenter, True
      spSetText frmprint.sprPrint, 3, sprow, "Manager-Operation", mAlignLeftCenter, True
      spSetText frmprint.sprPrint, 4, sprow, "Date:", mAlignLeftCenter, True
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      Spr.MaxRows = Spr.DataRowCnt
    
    TotalPages = iPageNo
    frmprint.Show
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set snap = Nothing
End Try
End Function

Public Sub sView()
  'No Change
End Sub

Public Sub sPrint()
  'No Change
End Sub

Private Sub CmdClear_Click()
  ClearForm Me
  StepClear Me
  StepEnable Me
  
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)

  Try
  Me.MousePointer = vbHourglass
  
  If Index = 0 Then
    If contrl_valid(txtDocNo, "Select the DocNo") = False Then
 GoTo ExitHere
End If
  End If
  
  Call StepProceedNext(Me, Index)
  If frastep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdView.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass
  

  Erpno = Get_ERPandTSNo("mnuDownstreamDocumentReport", "SAL", mErp)
  TsNo = Get_ERPandTSNo("mnuDownstreamDocumentReport", "SAL", mTs)
  If sPrintcurrent = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
  
  Try

  cmbType.ListIndex = 0
  chkPrintDate.Value = 1
  dtpDocDate.Value = Format(ServerDate, "yyyy-MM-dd")
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

Private Sub txtDocNo_DblClick()
  Dim sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  sql = "select distinct a.DocType,DocNo,a.DocDate, b.PartyName, a.CustomerCode, a.RefNo, d.PartNo, d.ItemDescription, "
  sql = sql & " a.ItemCode, c.RevisionNo as DrawingRevNo, c.RevisionDate from "
  sql = sql & " fn_SAL_DownstreamDocument_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join "
  sql = sql & " fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')  d"
  sql = sql & " on a.CustomerCode = d.CustomerCode and a.ItemCode = d.ItemCode and a.Branchid=d.Branchid and a.Compid = d.Compid and a.Finyear = d.Finyear"
  sql = sql & " left outer join FAS_Party_Master b on a.CustomerCode = b.PartyCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " left outer join DrawingMaster c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  sql = sql & " where a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "11110011", "List of Downstream Documents", "Doc Type", "Doc No", "Doc Date", "Customer Name", "Part No", "Part Description") = True Then
    txtdoctype.Text = getvalue(1)
    txtDocNo.Text = getvalue(2)
    dtpDocDate = Format(getvalue(3), "dd/MM/yyyy")
    txtdoctype.Tag = getvalue(9)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

