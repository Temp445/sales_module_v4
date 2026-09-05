VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptContractReviewReport 
   Caption         =   "Contract Review Report"
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
      Left            =   5027
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
         Format          =   59113473
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
         ItemData        =   "rptContractReviewReport.frx":0000
         Left            =   1035
         List            =   "rptContractReviewReport.frx":000A
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
Attribute VB_Name = "rptContractReviewReport"
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
  Dim CurDate As Date
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim sdate As Date
  Dim SValue As Double
  Dim tmp As Variant
  
  Try
  
  Set Spr = frmprint.sprPrint
  
  CurDate = ServerDate
    
    sql = " Select a.DocDate, a.RefNo,a.ItemCode,b.ItemDescription,b.PartNo from SAL_ContractReview_Hdr a"
    sql = sql & " left outer join fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') b"
    sql = sql & " On a.CustomerCode = b.CustomerCode and a.RefNo = b.ReferenceNo and a.ItemCode = b.ItemCode"
    sql = sql & " and a.Branchid = b.Branchid and a.Compid = b.Compid"
    sql = sql & " Where a.DocType = '" & Trim$(txtdoctype.Text) & "' and a.DocNo = '" & Val(txtDocNo.Text) & "' and a.DocDate = '" & Format(dtpDocDate, "yyyy-MM-dd") & "' "
    sql = sql & " and a.CustomerCode = " & Val(txtDocNo.Tag) & " and b.Status='A' "
    sql = sql & " and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
    If Not rs.EOF Then
      
      Call spSetColWidth(Spr, 1, 3, 2, 21, 3, 23, 4, 5, 5, 5, 6, 5, 7, 6, 8, 6)
      
      Spr.MaxCols = 8
      Call spSetText(Spr, 1, sprow, gCompName, mAlignLeftCenter, True, 6, 1)
      Call spSetText(Spr, 7, sprow, rs("Itemcode"), mAlignCenterCenter, True, 2, 1)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Call spSetText(Spr, 1, sprow, gStrReportHeading, mAlignCenterCenter, True, Spr.MaxCols, 1)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetText(Spr, 8, sprow, "Date : " & rs("DocDate"), mAlignRightCenter)

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow

      Call spSetText(Spr, 1, sprow, "Name Of the Customer", mAlignLeftCenter, True)
      Call spSetText(Spr, 3, sprow, ": " & Trim$(Label10.Tag), mAlignLeftCenter, False)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderTop)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetText(Spr, 1, sprow, "Purchase Order / Amendment Ref", mAlignLeftCenter, True)
      Call spSetText(Spr, 3, sprow, ": " & Trim$(rs("RefNo")), mAlignLeftCenter, False)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderTop)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetText(Spr, 1, sprow, "Part No", mAlignLeftCenter, True, 2, 1)
      Call spSetText(Spr, 3, sprow, ": " & Trim$(rs("PartNo")), mAlignLeftCenter, False)
      Call spSetText(Spr, 4, sprow, "Part Description", mAlignLeftCenter, True, 3, 1)
      Call spSetText(Spr, 7, sprow, ": " & Trim$(rs("Itemdescription")), mAlignLeftCenter, False, 2, 1)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow

      Call spSetText(Spr, 1, sprow, "IE Part No", mAlignLeftCenter, True, 2, 1)
      Call spSetText(Spr, 3, sprow, ": " & Trim$(rs("Itemcode")), mAlignLeftCenter, False)
      Call spSetText(Spr, 4, sprow, "Drawing Rev. No", mAlignLeftCenter, True, 3, 1)
      Call spSetText(Spr, 7, sprow, ": " & Trim$(txtdoctype.Tag), mAlignLeftCenter, False, 2, 1)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      sprow = sprow + 2
GridCount = GridCount + 2
RowProcess Me, sprow
      
      Spr.Row = sprow
Spr.Col = 1
Spr.TypeEditMultiLine = True
      Call spSetText(Spr, 1, sprow, "Sl. No", mAlignCenterCenter, True, 1, 3)
      Call spSetText(Spr, 2, sprow, "Details / Conditions", mAlignCenterCenter, True, 2, 3)
      Spr.Row = sprow
Spr.Col = 4
Spr.TypeEditMultiLine = True
      Call spSetText(Spr, 4, sprow, "Customer Request for Contracts", mAlignCenterCenter, True, 3, 2)
      Spr.Row = sprow
Spr.Col = 7
Spr.TypeEditMultiLine = True
      Call spSetText(Spr, 7, sprow, "Action Needed", mAlignCenterCenter, True, 1, 3)
      Spr.Row = sprow
Spr.Col = 8
Spr.TypeEditMultiLine = True
      Call spSetText(Spr, 8, sprow, "Resp. Sign", mAlignCenterCenter, True, 1, 3)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow + 1, mBorderAllSides)
      sprow = sprow + 2
GridCount = GridCount + 2
RowProcess Me, sprow
      
      Call spSetText(Spr, 4, sprow, "OK", mAlignCenterCenter, True)
      Call spSetText(Spr, 5, sprow, "AMD", mAlignCenterCenter, True)
      Call spSetText(Spr, 6, sprow, "NA", mAlignCenterCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
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
  Dim snap1 As AceADODB.Recordset
  Dim i As Integer
  Dim j As Integer
  Dim k As Long
  Dim L As Long
  Dim n As Integer
  Dim X As Integer
  Dim tmp As Variant
  Dim Spr As AceSpread

  Try
  
  sql = " Select ReviewNo,OK,AMD,NA,ActionNeed, Description from SAL_ContractReview_Dtl a Left Outer Join TermsMaster b"
  sql = sql & " On convert(varchar,ReviewNo) = Code and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " Where Typeflg = 'Contract Review' and DocType = '" & Trim$(txtdoctype.Text) & "' and DocNo= '" & Trim$(txtDocNo.Text) & "' and DocDate = '" & Format(dtpDocDate, "yyyy-MM-dd") & "'"

  sql = sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
 GoTo ExitHere
End If
  If Not rs.EOF Then
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 37
Orient = 2
    Set Spr = frmprint.sprPrint

    gStrReportHeading = "CONTRACT REVIEW & AMENDMENT CHECK LIST"

    Call Header

      For j = 1 To rs.RecordCount
        Call spSetText(Spr, 1, sprow, rs("ReviewNo"), mAlignCenterCenter, False)
        Call spSetText(Spr, 2, sprow, rs("Description"), mAlignLeftCenter, False, 2, 1)
        
        Spr.Row = sprow
Spr.Col = 4
Spr.CellType = CellTypeCheckBox
Spr.TypeVAlign = TypeVAlignCenter
Spr.TypeHAlign = TypeHAlignCenter
        tmp = IIf(rs("OK") = True, 1, 0)
        Call spSetText(Spr, Spr.Col, Spr.Row, Val(tmp), mAlignCenterCenter)
        Spr.Row = sprow
Spr.Col = 5
Spr.CellType = CellTypeCheckBox
Spr.TypeVAlign = TypeVAlignCenter
Spr.TypeHAlign = TypeHAlignCenter
        tmp = IIf(rs("AMD") = True, 1, 0)
        Call spSetText(Spr, Spr.Col, Spr.Row, Val(tmp), mAlignCenterCenter)
        Spr.Row = sprow
Spr.Col = 6
Spr.CellType = CellTypeCheckBox
Spr.TypeVAlign = TypeVAlignCenter
Spr.TypeHAlign = TypeHAlignCenter
        tmp = IIf(rs("NA") = True, 1, 0)
        Call spSetText(Spr, Spr.Col, Spr.Row, Val(tmp), mAlignCenterCenter)

        Call spSetText(Spr, 7, sprow, rs("ActionNeed"), mAlignLeftCenter)
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        rs.MoveNext
      Next j
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetText(Spr, 1, sprow, "Prepared By:", mAlignLeftCenter, True)
      Call spSetText(Spr, 4, sprow, "Reviwed / Approved By:", mAlignLeftCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetText(Spr, 1, sprow, "*OK     - Accepted as per Purchase Order", mAlignLeftCenter, True)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Call spSetText(Spr, 1, sprow, "*AMD - Purchase Order Amendmend Required", mAlignLeftCenter, True)
      sprow = sprow + 1
GridCount = GridCount + 1 ': RowProcess Me, sprow
      Call spSetText(Spr, 1, sprow, "*NA     - No Requirement by Customer", mAlignLeftCenter, True)
      Spr.MaxRows = Spr.DataRowCnt
    
    
'    TotalPages = iPageNo
    frmprint.Show
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set snap1 = Nothing
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

  
  Try
  Me.MousePointer = vbHourglass

  Call Form_Clear(Me)
  
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
  

  Erpno = Get_ERPandTSNo("mnuSkillSetMasterReport", "HRD", mErp)
  TsNo = Get_ERPandTSNo("mnuSkillSetMasterReport", "HRD", mTs)
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
  
  Try
  Me.MousePointer = vbHourglass
  
  sql = " Select DocType,DocNo,DocDate,PartyName,a.Itemcode,CustomerCode,c.RevisionNo from SAL_ContractReview_Hdr a"
  sql = sql & " Left outer join FAS_Party_Master b on a.CustomerCode = b.PartyCode and a.Branchid = b.Branchid and a.Compid = b.Compid"
  sql = sql & " left outer join DrawingMaster c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  sql = sql & " Where a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "11111", "List of Enquiry No's", "Enquiry Type", "Enquiry No", "Enquiry Date") = True Then
    txtdoctype.Text = getvalue(1)
    txtDocNo.Text = getvalue(2)
    dtpDocDate = Format(getvalue(3), "dd/MM/yyyy")
    Label10.Tag = getvalue(4)
    txtDocNo.Tag = getvalue(6)
    txtdoctype.Tag = getvalue(7)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

