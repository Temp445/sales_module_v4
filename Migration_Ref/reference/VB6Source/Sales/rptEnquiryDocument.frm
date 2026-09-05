VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptEnquiryDocument 
   Caption         =   "Enquiry Document"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
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
   ScaleHeight     =   8490
   ScaleWidth      =   11430
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
      Index           =   0
      Left            =   8175
      TabIndex        =   11
      Top             =   2655
      Width           =   345
   End
   Begin VB.Frame FraStep 
      Caption         =   "Date"
      Height          =   630
      Index           =   1
      Left            =   4260
      TabIndex        =   24
      Top             =   1890
      Width           =   3270
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1995
         TabIndex        =   1
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   92733441
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
         Format          =   92733441
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   360
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
      Index           =   1
      Left            =   7575
      TabIndex        =   2
      Top             =   2085
      Width           =   345
   End
   Begin VB.Frame FraReportType 
      Height          =   945
      Left            =   4350
      TabIndex        =   17
      Top             =   3135
      Width           =   3135
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "rptEnquiryDocument.frx":0000
         Left            =   1035
         List            =   "rptEnquiryDocument.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   210
         Width           =   1995
      End
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   585
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6570
      TabIndex        =   16
      Top             =   4140
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5280
      TabIndex        =   15
      Top             =   4140
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   3990
      TabIndex        =   14
      Top             =   4140
      Width           =   1300
   End
Begin AceSpread fpsList
      Height          =   1470
      Left            =   8070
      TabIndex        =   5
      Top             =   2610
      Visible         =   0   'False
      Width           =   3225
      _Version        =   458752
      _ExtentX        =   5689
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
   Begin VB.Frame FraStep 
      Height          =   780
      Index           =   0
      Left            =   3975
      TabIndex        =   18
      Top             =   2355
      Width           =   4110
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   855
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
      Begin VB.ComboBox cmbRevNo 
         Height          =   330
         Left            =   3165
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   345
         Width           =   840
      End
      Begin VB.Label lblEnqType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   23
         Top             =   180
         Width           =   675
      End
      Begin VB.Label lblEnqNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.No."
         Height          =   210
         Left            =   990
         TabIndex        =   22
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq. Date"
         Height          =   210
         Left            =   2040
         TabIndex        =   21
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblRevNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev. No."
         Height          =   210
         Left            =   3165
         TabIndex        =   19
         Top             =   165
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
      Left            =   7575
      TabIndex        =   6
      Top             =   2715
      Width           =   345
   End
   Begin VB.Frame FraStep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   600
      Index           =   2
      Left            =   4260
      TabIndex        =   27
      Top             =   2535
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
         Height          =   195
         Left            =   1980
         TabIndex        =   4
         Top             =   255
         Width           =   1050
      End
   End
End
Attribute VB_Name = "rptEnquiryDocument"
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
Public Sub sAudit()
    
 Try

  FraStep(0).Visible = False
  cmdStep(0).Visible = False
  
  FraStep(1).Visible = True
  cmdStep(1).Visible = True
  FraStep(2).Visible = True
  cmdStep(2).Visible = True
  
  FraStep(1).Enabled = True
  cmdStep(1).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Public Sub sObsolute()
    
  Try

  FraStep(0).Left = 3815
  FraStep(0).Width = 4055
  cmdStep(0).Left = 7930
  
  FraStep(1).Visible = False
  cmdStep(1).Visible = False
  FraStep(2).Visible = False
  cmdStep(2).Visible = False
  
  FraStep(0).Visible = True
  cmdStep(0).Visible = True
  
  lblRevNo.Enabled = True
  cmbRevNo.Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sRevisionSummary()
    
  Try

  FraStep(0).Left = 3815
  FraStep(0).Width = 4055
  cmdStep(0).Left = 7930

  FraStep(1).Visible = False
  cmdStep(1).Visible = False
  FraStep(2).Visible = False
  cmdStep(2).Visible = False
  
  FraStep(0).Enabled = True
  FraStep(0).Visible = True
  cmdStep(0).Visible = True
  cmdStep(0).Enabled = True

  lblRevNo.Enabled = False
  cmbRevNo.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sListing()
    
  Try

  FraStep(0).Visible = False
  cmdStep(0).Visible = False
  
  FraStep(1).Visible = True
  cmdStep(1).Visible = True
  FraStep(2).Visible = True
  cmdStep(2).Visible = True
  
  FraStep(1).Enabled = True
  cmdStep(1).Enabled = True
  
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sprint()
  Try

  FraStep(0).Left = 4330
  FraStep(0).Width = 3160
  cmdStep(0).Left = 7590
  
  FraStep(1).Visible = False
  cmdStep(1).Visible = False
  FraStep(2).Visible = False
  cmdStep(2).Visible = False
  
  cmdStep(0).Enabled = True
  FraStep(0).Enabled = True
  FraStep(0).Visible = True
  cmdStep(0).Visible = True
  
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  FraStep(0).Left = 4330
  FraStep(0).Width = 3160
  cmdStep(0).Left = 7590
  
  FraStep(1).Visible = False
  cmdStep(1).Visible = False
  FraStep(2).Visible = False
  cmdStep(2).Visible = False
  
  FraStep(0).Enabled = True
  cmdStep(0).Enabled = True
  
  FraStep(0).Visible = True
  cmdStep(0).Visible = True
  frmPrint.cmdPPrint.Enabled = False
  frmPrint.cmdExcel.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

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
  cmbRevNo.Clear
  
  Sql = "Select Distinct EnquiryType, EnquiryNo, EnquiryDate,"
  Sql = Sql & " Max(RevisionNo) As RevNo, Max(RevisionDate) As RevDate"

  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_ENQUIRY_HDR"
End If
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobworkEnquiry_Hdr"
End If

  Sql = Sql & " Where Branchid='" & gBranchID & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and finyear='" & gFinyear & "'"
  Sql = Sql & " Group by EnquiryType, EnquiryNo, EnquiryDate"
  Sql = Sql & " Order by EnquiryDate desc, EnquiryNo desc"

  If FetchValue_New(False, db.ConnectionString, Sql, "11100", "List of Enquiry No's", "Enquiry Type", "Enquiry No", "Enquiry Date") = True Then

    If Trim$(RptFlg) = "Obsolute" Then
      If Val(getvalue(4)) < 1 Then
        sShowMessage "Obsolete Document did not exist"
        GoTo ExitHere
      End If
    End If

    txtType.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)

    If Trim$(txtType.Text) <> "" Then
 cmdStep(1).Enabled = True
End If
    If Trim$(RptFlg) = "Obsolute" Then
      getvalue(4) = Val(getvalue(4)) - 1
      If Trim$(cmbRevNo.Text) <> Trim$(getvalue(4)) Then  'Obsolete Document - Eg. 0,1, To 5 - Only 0,1, To 4 (Last One Restriction)
        cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(getvalue(4))) / 1))) & Val(Val(getvalue(4)))
      End If
      For i = Val(getvalue(4)) - 1 To 0 Step -1
        cmbRevNo.AddItem Space(5 - Len(Trim$(Int(i) / 1))) & Val(i)
      Next i

      If cmbRevNo.ListCount >= 1 Then
 cmbRevNo.ListIndex = 0
End If

    Else
      If Val(getvalue(4)) = 0 Then
        cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(getvalue(4))) / 1))) & Val(getvalue(4))
      ElseIf Val(cmbRevNo.Text) <> Val(getvalue(4)) Then
        cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(getvalue(4))) / 1))) & Val(getvalue(4))
      End If
      For i = Val(getvalue(4)) - 1 To 0 Step -1
        cmbRevNo.AddItem Space(5 - Len(Trim$(Int(i) / 1))) & Val(i)
      Next i

      If cmbRevNo.ListCount >= 1 Then
 cmbRevNo.ListIndex = 0
End If
    End If
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
  cmdStep(2).Enabled = True
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
    If DocFlg = "Sales" Then
      Sql = Sql & " From SAL_CUSTOMER_ITEMMASTER a"
    Else
      Sql = Sql & " From SAL_JobWork_ITEMMASTER a"
    End If
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
    Sql = Sql & " On a.CustomerCode=b.PartyCode and a.BranchID=b.Branchid"
    Sql = Sql & " and a.CompID=b.Compid"
  '  Sql = Sql & " Left Outer Join SAL_CUSTOMER_ITEMMASTER_His c"
  '  Sql = Sql & " On a.CustomerCode=c.CustomerCode and a.ItemCode=c.ItemCode"
  '  Sql = Sql & " and a.BranchID=c.Branchid and a.CompID=c.Compid"
    Sql = Sql & " Where  a.BranchID='" & gBranchID & "'and a.CompID='" & gCompID & "'"
    If DocFlg = "Sales" Then
      Sql = Sql & " and a.CustomerCode in (Select CustomerCode From SAL_Enquiry_Hdr "
      Sql = Sql & " Where CustomerCode = a.CustomerCode and BranchID = a.BranchID and CompID = a.CompID)"
      Sql = Sql & " and a.Status=1 "
    Else
      Sql = Sql & " and a.CustomerCode in (Select CustomerCode From SAL_JobWorkEnquiry_Hdr "
      Sql = Sql & " Where CustomerCode = a.CustomerCode and BranchID = a.BranchID and CompID = a.CompID)"
      Sql = Sql & " and a.IsActive=1 "
    End If
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
    cmdStep(2).Enabled = False
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
' On Error Resume Next
  
  If fpsList.SelectionCount >= 1 Then
    cmdStep(2).Enabled = True
  Else
    cmdStep(2).Enabled = False
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
  
  If Index = 0 Then
    If RptFlg <> "Listing" And Trim$(txtType) = "" Then
      sShowMessage "Select the Enquiry Type"
      Exit Sub
    End If
  ElseIf Index = 1 Then
    If RptFlg = "Listing" Then
      If dtpToDate < dtpFromDate Then
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
  Else
    If dtpToDate < dtpFromDate Then
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
  
'  If optCustSelective.Value = True Then
'    If fpsList.SelectionCount = 0 Then
'      sShowMessage "Select atleast one Customer"
'      Exit Sub
'    End If
'  End If
'
'  If FraStep(0).Visible = True And cmdStep(0).Enabled = True Then
'    If txtType = "" Then: sShowMessage "Select the Enquiry Type": Exit Sub
'    Call StepProceedNext(Me, Index)
'    Index = 2: Call StepProceedNext(Me, Index)
'  ElseIf FraStep(1).Visible = True And cmdStep(1).Enabled = True Then
'    Call StepProceedNext(Me, Index)
'  ElseIf FraStep(2).Visible = True And cmdStep(2).Enabled = True Then
'    Call StepProceedNext(Me, Index)
'    Index = 2: Call StepProceedNext(Me, Index)
'  Else
'    Call StepProceedNext(Me, Index)
'    cmdView.Enabled = True
'  End If
'  fpsList.Enabled = False
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
    ElseIf Trim$(RptFlg) = "RevisionSummary" Then
      If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Obsolute" Then
      If sPrintCurrent_Print = False Then
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
  cmdStep(1).Enabled = False
  cmbRevNo.Clear
  
  FraStep(0).Enabled = True
  cmdStep(0).Enabled = True
  FraStep(1).Enabled = True
  cmdStep(1).Enabled = True
  
  cmbType.ListIndex = 0
  chkPrintDate.Value = 0
  txtType.Text = ""
txtType.Tag = ""
txtNo.Text = ""
txtDate.Text = ""
mRevNo = ""
mDate = ""
  
  cmdView.Enabled = False
  If txtType.Visible = True Then
 txtType.SetFocus
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
    sSendToPort LAlign(10, "Enquiry No.") & Space(1) & LAlign(30, "Customer") & Space(1) & LAlign(20, "Part No.") & Space(1) & LAlign(4, "UOM") & Space(1) & LAlign(25, "Item Code") & Space(1) & LAlign(12, "UOM") & Space(1) & RAlign(12, "Rate") & Space(1) & RAlign(12, "Goods Value")
    sSendToPort LAlign(10, "Enquiry Date") & Space(32) & LAlign(20, "Part Description") & Space(6) & LAlign(25, "Item Description") & Space(1) & LAlign(12, "Qty") & Space(1) & RAlign(12, "Ass.Rate") & Space(1) & RAlign(12, "Ass. Value")
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
  Dim mSNo As Integer
  Dim mEnqType As String
  Dim mCustNm As String
      
  Try
  sPrintcurrent_Listing = False
  
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
  End If
  
  If Trim$(DocFlg) = "Job Work" Then
    Sql = "Select Distinct a.EnquiryType, a.EnquiryNo, a.EnquiryDate, a1.CustomerCode, IsNull(b.PartyName,'') As CustomerName,"
    Sql = Sql & " IsNull(a1.Remarks,'') As Remarks, IsNull(a1.ReferenceNo,'') As RefNo, IsNull(a1.ReferenceDate,'') As RefDate,"
    Sql = Sql & " a.ReceivedItemCode As RecdItemCode, IsNull(i.ItemDescription,'') As RecdItemDesc, IsNull(j.SDesc,'') As RecdUOM,"
    Sql = Sql & " a.SentItemCode As SentItemCode, IsNull(c.ItemDescription,'') As SentItemDesc, IsNull(f.SDesc,'') As SentUOM,"
    Sql = Sql & " a.SentItemCode As ItemCode, IsNull(c.ItemDescription,'') As ItemDescription, IsNull(f.SDesc,'') As UOM,"
    Sql = Sql & " a.Itemgroupkey, IsNull(d.PartNo,'') As PartNo, a.SendQty As Qty, Cast(a.Rate As decimal(15,2)) As Rate,"
    Sql = Sql & " Cast((a.SendQty*a.Rate) As Decimal(15,2)) As GoodsValue, Round(a.AssRate,0) As AssRate, a.RatePer,"
    Sql = Sql & " Round(a.SendQty*a.AssRate,0) As AssValue From SAL_JobworkEnquiry_Dtl a"
    Sql = Sql & " Left Outer Join SAL_JobworkEnquiry_Hdr a1 On a.EnquiryType = a1.EnquiryType and a.EnquiryNo = a1.EnquiryNo"
    Sql = Sql & " and a.EnquiryDate = a1.EnquiryDate and a.Branchid=a1.Branchid and a.Compid=a1.Compid"
    Sql = Sql & " Left Outer Join Material_Master i On a.ReceivedItemCode=i.ItemCode and a.Branchid=i.Branchid and a.Compid=i.Compid"
    Sql = Sql & " Left Outer Join INV_UOM_Master j On i.UOM=j.Code and a.Branchid=j.Branchid and a.Compid=j.Compid and j.status=1"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b On a1.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER c On a.SentItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid and c.status=1"
    Sql = Sql & " Left Outer Join PartMast d On a.SentItemCode=d.StageCode and a.Branchid=d.BranchID and a.Compid=d.CompID"
    Sql = Sql & " Left Outer Join INV_UOM_Master f On c.UOM=f.Code and a.Branchid=f.Branchid and a.Compid=f.CompId and f.Status=1 "
    mCustNm = "SAL_JobworkEnquiry_Dtl"
  Else
    Sql = " Select Distinct a.EnquiryType, a.EnquiryNo, a.EnquiryDate, a1.CustomerCode, IsNull(b.PartyName,'') As CustomerName,"
    Sql = Sql & " IsNull(a1.Remarks,'') As Remarks, IsNull(a1.ReferenceNo,'') As RefNo, IsNull(a1.ReferenceDate,'') As RefDate,"
    Sql = Sql & " a.ItemCode, IsNull(c.ItemDescription,'') As ItemDescription, IsNull(f.SDesc,'') As UOM, a.TargetDate,"
    Sql = Sql & " IsNull(d.PartNo,'') As PartNo, a.Qty, a.Rate ,a.AssessableRate  As AssRate From SAL_Enquiry_Dtl a"
    Sql = Sql & " Left Outer Join SAL_Enquiry_Hdr a1 On a.EnquiryType = a1.EnquiryType and a.EnquiryNo = a1.EnquiryNo"
    Sql = Sql & " and a.EnquiryDate = a1.EnquiryDate and a.Revisionno = a1.Revisionno and a.Branchid=a1.Branchid and a.Compid=a1.Compid"
    Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b On a1.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER c On a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid and c.status=1"
    Sql = Sql & " Left Outer Join PartMast d On a.ItemCode=d.StageCode and a.Branchid=d.BranchID and a.Compid=d.CompID"
    Sql = Sql & " Left Outer Join INV_UOM_Master f On c.UOM=f.Code and a.Branchid=f.Branchid and a.Compid=f.CompId and f.Status=1"
    mCustNm = "SAL_Enquiry_Dtl"
  End If
  Sql = Sql & " Where a.EnquiryDate Between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpToDate.Value, "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.Revisionno = (Select max(Revisionno) as Revisionno from " & Trim$(mCustNm) & " a2 Where a.EnquiryType = a2.EnquiryType and a.EnquiryNo = a1.EnquiryNo"
  Sql = Sql & " and a.EnquiryDate = a2.EnquiryDate and a.Revisionno = a2.Revisionno and a.Branchid=a2.Branchid and a.Compid=a2.Compid)"
  If optCustSelective Then
'    If fpsList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": Exit Function
    
      Sql = Sql & " and a1.CustomerCode in "
      Sql = Sql & Spread_Get_SelectionList(fpsList, 2)
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = IIf(mDocSeq(1) = 3, 2, 1)
'        Sql = Sql + "'" + Trim$(.Text) + "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)'    Sql = Sql + ")"
  End If
  Sql = Sql & " Order by a.EnquiryType, a.EnquiryNo,a.EnquiryDate, a1.CustomerCode"
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
        
    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
      'Top
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If GridCount = 53 Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
        
        If Trim$(mEnqType) <> Trim$(rs("EnquiryType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("EnquiryType")) & Trim$(rs("EnquiryNo"))
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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("UOM"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty") * rs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("RefNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Qty"))
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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
      'Botom
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If GridCount = 53 Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
        
        If Trim$(mEnqType) <> Trim$(rs("EnquiryType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("EnquiryDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty") * rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("TargetDate"), "dd/MM/yyyy")
frmPrint.sprPrint.ColWidth(11) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else  'Job Work
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
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RatePer"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
                
        mEnqType = Trim$(rs("EnquiryType"))
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
    gStrReportHeading = "SALES ENQUIRY (LISTING) BETWEEN " & mFromDate & " AND " & mToDate
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
tmp = ""
Gtot(1) = 0
Gtot(2) = 0
  
    Do While Not rs.EOF
      sSendToPort LAlign(10, rs("EnquiryNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("EnquiryDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("Qty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
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
  Dim Ln As Integer
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
  Dim mEnqType As String

  Try
  sPrintcurrent_RevisionSummary = False

  Call RecordInsertion  'Process

  Sql = "Select Distinct a.EnquiryType, a.EnquiryNo, a.EnquiryDate, IsNull(b.PartNo,'') As PartNo, IsNull(a.RevisionNo,'') As RevisionNo,"
  Sql = Sql & " IsNull(a.RevisionDate,'') As RevisionDate, IsNull(a.CustomerCode,'') As CustomerCode, IsNull(a.CustomerName,'') As CustomerName,"
  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " a.ItemCode, IsNull(a.ItemDescription,'') As ItemDescription, a.UOM, a.TargetDate, a.Qty,"
    Sql = Sql & " a.Rate, IsNull(a.AssessableRate,'') As AssRate, IsNull(RecStatus,'') As RecStatus,"
  Else  'Job Work
    Sql = Sql & " Itemgroupkey,a.ReceivedItemCode As RecdItemCode, IsNull(a.RecdDescription,'') As RecdItemDesc, IsNull(i.UOMDesc,'') As RecdUOM,"
    Sql = Sql & " a.SentItemCode,  IsNull(a.SentDescription,'') As SentItemDesc, a.SendQty As Qty, a.RatePer, a.Rate, a.AssRate,"
  End If
  Sql = Sql & " IsNull(a.Remarks,'') As Remarks, IsNull(a.ReferenceNo,'') As RefNo, IsNull(a.ReferenceDate,'') As RefDate, "
  Sql = Sql & " IsNull(RecStatus,'') As RecStatus From SAL_SalseDocument_Tmp a Left Outer Join PartMast b"
  If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " On a.ItemCode=b.StageCode "
End If
  If Trim$(DocFlg) = "Job work" Then
 Sql = Sql & " On a.SentItemCode=b.StageCode "
End If
  Sql = Sql & " and a.Branchid=b.BranchID and a.Compid=b.CompID"
  If Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " Left Outer Join Material_Master i On a.ReceivedItemCode=i.ItemCode and a.Branchid=i.Branchid and a.Compid=i.Compid"
  End If
  Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
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
    frmPrint.sprPrint.MaxCols = 12
    mCustNm = ""
    
    Header
        
    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
       'Top
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("EnquiryType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("EnquiryType")) & Trim$(rs("EnquiryNo"))
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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("UOM"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("TargetDate"), "dd/MM/yyyy")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Qty"))
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
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
      'Bottom
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mEnqType) <> Trim$(rs("EnquiryType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("EnquiryDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.ColWidth(10) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Else  'Job Work
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Itemgroupkey"))
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

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RatePer"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("RefDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        mEnqType = Trim$(rs("EnquiryType"))
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
  Sql = "Create Table SAL_SalseDocument_Tmp(EnquiryType varchar(3) Not Null,"
  Sql = Sql & " EnquiryNo int Not Null,"
  Sql = Sql & " EnquiryDate datetime Not Null, RevisionNo int Not Null,"
  Sql = Sql & " RevisionDate datetime Not Null, ReferenceNo varchar(15) Not Null,"
  Sql = Sql & " ReferenceDate datetime Not Null, CustomerType char(1) Not Null,"
  Sql = Sql & " CustomerCode int Not Null, CustomerName varchar(50) Not Null,"
  Sql = Sql & " Remarks varchar(50) Not Null, ItemType char(1) Not Null,"

  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " ItemCode varchar(20) Not Null, ItemDescription varchar(50) Not Null,"
    Sql = Sql & " UOM varchar(3) Not Null, Qty float Not Null, Rate float Not Null,"
    Sql = Sql & " AssessableRate float Not Null, TargetDate datetime Not Null,"
    Sql = Sql & " Specification varchar(30) Not Null,"
    Sql = Sql & " QuotationType varchar(3) Not Null, QuotationNo int Not Null,"
    Sql = Sql & " QuotationDate datetime Not Null,"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " ItemGroupKey int Not Null, ReceivedItemCode varchar(20) Not Null,"
    Sql = Sql & " RecdDescription varchar(50) Not Null, RecdUom varchar(3) Not Null,"
    Sql = Sql & " SentItemCode varchar(20) Not Null, SentDescription varchar(50) Not Null,"
    Sql = Sql & " SentUOM varchar(3) Not Null, SendQty float Not Null,"
    Sql = Sql & " VolumePer varchar(15) Not Null, RatePer float Not Null,"
    Sql = Sql & " Rate money Not Null, AssRate money Not Null,"
    Sql = Sql & " TargetDate datetime Not Null, RelationType varchar(3) Not Null,"
'    Sql = Sql & " RateEffDate datetime Not Null,"
    Sql = Sql & " QuotationType varchar(3) Not Null, QuotationNo int Not Null,"
    Sql = Sql & " QuotationDate datetime Not Null, isActive bit Not Null,"
    Sql = Sql & " Status char(20) Not Null, StatusDate datetime Not Null,"
    Sql = Sql & " CreatedDate datetime Not Null,"
  End If
  Sql = Sql & " Userid varchar(10) Not Null, Branchid varchar(3) Not Null,"
  Sql = Sql & " Compid varchar(3) Not Null, Finyear varchar(20) Not Null,"
  Sql = Sql & " RecStatus varchar(20) Not Null)"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

'For Select Fields
  mQry = "a.EnquiryType, a.EnquiryNo, a.EnquiryDate,"
  mQry = mQry & " IsNull(a.RevisionNo,'') As RevisionNo,"
  mQry = mQry & " IsNull(a.RevisionDate,'') As RevisionDate,"
  mQry = mQry & " IsNull(a.ReferenceNo,'') As ReferenceNo,"
  mQry = mQry & " IsNull(a.ReferenceDate,'') As ReferenceDate,"
  mQry = mQry & " IsNull(a.CustomerType,'') As CustomerType, a.CustomerCode,"
  mQry = mQry & " IsNull(a.CustomerName,'') As CustomerName,"
  mQry = mQry & " IsNull( a.Remarks,'') As Remarks, IsNull(b.ItemType,'') As ItemType,"

  If Trim$(DocFlg) = "Sales" Then
    mQry = mQry & " b.ItemCode, IsNull(b.ItemDescription,'') As ItemDescription,"
    mQry = mQry & " IsNull(b.UOM,'') As UOM, b.Qty, b.Rate,"
    mQry = mQry & " b.AssessableRate, IsNull(b.TargetDate,'') As TargetDate,"
    mQry = mQry & " IsNull(b.Specification,'') As Specification,"
    mQry = mQry & " IsNull(b.QuotationType,'') As QuotationType,"
    mQry = mQry & " IsNull(b.QuotationNo,'') As QuotationNo,"
    mQry = mQry & " IsNull(b.QuotationDate,'') As QuotationDate,"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    mQry = mQry & " IsNull(b.ItemGroupKey,'') As ItemGroupKey, b.ReceivedItemCode,"
    mQry = mQry & " IsNull(b.RecdDescription,'') As RecdItemDescription,"
    mQry = mQry & " IsNull(b.RecdUOM,'') As RecdUOM,"
    mQry = mQry & " b.SentItemCode, IsNull(b.SentDescription,'') As SentDescription,"
    mQry = mQry & " IsNull(b.SentUOM,'') As SentUOM, SendQty, b.VolumePer,"
    mQry = mQry & " b.RatePer, b.Rate, b.AssRate,"
    mQry = mQry & " IsNull(b.TargetDate,'') As TargetDate,"
    mQry = mQry & " IsNull(b.RelationType,'') As RelationType,"
    mQry = mQry & " IsNull(b.QuotationType,'') As QuotationType,"
    mQry = mQry & " IsNull(b.QuotationNo,'') As QuotationNo,"
    mQry = mQry & " IsNull(b.QuotationDate,'') As QuotationDate,"
    mQry = mQry & " b.isActive, b.Status, b.StatusDate, b.CreatedDate,"
  End If
  mQry = mQry & " b.Userid, b.Branchid, b.Compid, b.Finyear"

'Insert - From SAL_SalesPlan To SAL_SalesPlan_Tmp
  i = 0
  If i = 0 Then
  'All - INITIAL
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select " & Trim$(mQry) & ", 'INITIAL STATE' As RecStatus"

    If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Enquiry_Hdr a"
End If
    If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"
End If
    If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"
End If
    If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
End If

    Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo"
    Sql = Sql & " and a.EnquiryDate=b.EnquiryDate"
    Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid"
    Sql = Sql & " and a.Compid=b.Compid and a.Finyear=b.Finyear"
    Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
    Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and a.Revisionno=0"

'    If optAll Then: Sql = Sql & " and a.Revisionno=0"

'    If optSelective Then: Sql = Sql & " and a.Revisionno=" & val(Revprev) & ""

    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " Order by a.RevisionNo"
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
 Sql = Sql & " From SAL_Enquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
End If

      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo and a.EnquiryDate=b.EnquiryDate"
      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
      Sql = Sql & " and a.Finyear=b.Finyear"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Where b.ItemCode Not In(Select c.ItemCode"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Where b.SentItemCode Not In(Select c.SentItemCode"
End If

      Sql = Sql & " From ("

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Select * From SAL_Enquiry_Dtl c"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Select * From SAL_JobWorkEnquiry_Dtl c"
End If

      Sql = Sql & " Where c.EnquiryType='" & Trim$(txtType) & "' and c.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and c.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and c.RevisionNo In(" & i & ")"
      Sql = Sql & " and c.Branchid='" & gBranchID & "' and c.Compid ='" & gCompID & "'"
      Sql = Sql & " and c.Finyear='" & gFinyear & "'"
      Sql = Sql & " ) c"
      Sql = Sql & " Inner Join ("

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Select * From SAL_Enquiry_Dtl d"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Select * From SAL_JobWorkEnquiry_Dtl d"
End If

      Sql = Sql & " Where d.EnquiryType='" & Trim$(txtType) & "' and d.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and d.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and d.RevisionNo In(" & j & ")"
      Sql = Sql & " and d.Branchid='" & gBranchID & "' and d.Compid ='" & gCompID & "'"
      Sql = Sql & " and d.Finyear='" & gFinyear & "'"
      Sql = Sql & " ) d"
      Sql = Sql & " On c.EnquiryType=d.EnquiryType and c.EnquiryNo=d.EnquiryNo and c.EnquiryDate=d.EnquiryDate"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and c.ItemCode=d.ItemCode"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and c.SentItemCode=d.SentItemCode"
End If

      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid"
      Sql = Sql & " and c.Finyear=d.Finyear"
      Sql = Sql & " )"
      Sql = Sql & " and a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.Revisionno In(" & j & ") "
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    'All - MODIFIED
      Sql = "Insert Into SAL_SalseDocument_Tmp"
      Sql = Sql & " Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Enquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
End If

      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo"
      Sql = Sql & " and a.EnquiryDate=b.EnquiryDate and a.RevisionNo=b.RevisionNo "
      Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
      Sql = Sql & " and a.Finyear=b.Finyear"
      Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.RevisionNo In(" & j & ")"
      Sql = Sql & " and Exists ("

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Select * From SAL_Enquiry_Hdr c"
End If
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl d"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Select * From SAL_JobWorkEnquiry_Hdr c"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl d"
End If

      Sql = Sql & " On c.EnquiryType=d.EnquiryType and c.EnquiryNo=d.EnquiryNo and c.EnquiryDate=d.EnquiryDate"
      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid and c.Finyear=d.Finyear"
      Sql = Sql & " Where c.EnquiryType='" & Trim$(txtType) & "' and c.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and c.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"

      If Trim$(DocFlg) = "Sales" Then
        Sql = Sql & " and b.ItemCode=d.ItemCode and c.RevisionNo In(" & i & ")"
        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
        Sql = Sql & " or b.UOM<>d.UOM or b.Qty<>d.Qty or b.Volume<>d.Volume"
        Sql = Sql & " or b.AssessableRate<>d.AssessableRate"
      ElseIf Trim$(DocFlg) = "Job Work" Then
        Sql = Sql & " and b.SentItemCode=d.SentItemCode and c.RevisionNo In(" & i & ")"
        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
        Sql = Sql & " or b.ItemGroupKey<>d.ItemGroupKey or b.ReceivedItemCode<>d.ReceivedItemCode"
        Sql = Sql & " or b.RecdUOM<>d.RecdUOM or b.SentUOM<>d.SentUOM"
        Sql = Sql & " or b.SendQty<>d.SendQty or b.RatePer<>d.RatePer or b.AssRate<>d.AssRate"
        Sql = Sql & " or b.RelationType<>d.RelationType or b.isActive<>d.isActive"
      End If

      Sql = Sql & " or b.VolumePer<>d.VolumePer or b.Rate<>d.Rate"
      Sql = Sql & " or IsNull(b.TargetDate,'')<>IsNull(d.TargetDate,'')"
      Sql = Sql & " )"
      Sql = Sql & " )"
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    'All - DELETED
      Sql = "Insert Into SAL_SalseDocument_Tmp"
      Sql = Sql & " Select " & Trim$(mQry) & ", 'DELETED' As RecStatus"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " From SAL_Enquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
End If

      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo"
      Sql = Sql & " and a.EnquiryDate=b.EnquiryDate"
      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
      Sql = Sql & " and a.Finyear=b.Finyear"
      Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.RevisionNo In(" & i & ")"

      If Trim$(DocFlg) = "Sales" Then
 Sql = Sql & " and b.ItemCode Not In( Select c.ItemCode From SAL_Enquiry_Dtl c"
End If
      If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " and b.SentItemCode Not In( Select c.SentItemCode From SAL_JobWorkEnquiry_Dtl c"
End If

      Sql = Sql & " Where a.EnquiryType = c.EnquiryType and A.EnquiryNo = c.EnquiryNo"
      Sql = Sql & " and a.EnquiryDate = c.EnquiryDate"
      Sql = Sql & " and c.RevisionNo In(" & j & ") and a.Branchid=c.Branchid"
      Sql = Sql & " and a.Compid=b.Compid and a.Finyear=b.Finyear"
      Sql = Sql & " )"
      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

      If j = (cmbRevNo.ListCount - 1) Then
        GoTo ExitHere
      End If

    Next i

'  ElseIf optSelective Then

'    j = IIf(Trim$(lblRevNo.Tag) = "", Trim$(mRevCurr), Trim$(lblRevNo.Tag))
'    For i = j To Val(mRevCurr)

'    'Selective - NEW
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select  " & Trim$(mQry) & ",'NEW' As RecStatus"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " From SAL_Enquiry_Hdr a"

'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
'
'      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo"
'      Sql = Sql & " and a.EnquiryDate=b.EnquiryDate and a.RevisionNo=b.RevisionNo"
'      Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
'      Sql = Sql & " and a.Finyear=b.Finyear"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Where b.ItemCode Not In(Select c.ItemCode"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Where b.SentItemCode Not In(Select c.SentItemCode"
'
'      Sql = Sql & " From ("
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Select * From SAL_Enquiry_Dtl c"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Select * From SAL_JobWorkEnquiry_Dtl c"
'
'      Sql = Sql & " Where c.EnquiryType='" & Trim$(txtType) & "' and c.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and c.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and c.RevisionNo In(" & i & ")"
'      Sql = Sql & " and c.Branchid='" & gBranchID & "' and c.Compid ='" & gCompID & "'"
'      Sql = Sql & " and c.Finyear='" & gFinyear & "'"
'      Sql = Sql & " ) c"
'      Sql = Sql & " Inner Join ("
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Select * From SAL_Enquiry_Dtl d"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Select * From SAL_JobWorkEnquiry_Dtl d"
'
'      Sql = Sql & " Where d.EnquiryType='" & Trim$(txtType) & "' and and d.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and d.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and d.RevisionNo In(" & Val(lblRevNo.Tag) & ")"
'      Sql = Sql & " and d.Branchid='" & gBranchID & "' and d.Compid ='" & gCompID & "'"
'      Sql = Sql & " and d.Finyear='" & gFinyear & "'"
'      Sql = Sql & " ) d"
'      Sql = Sql & " On c.EnquiryType=d.EnquiryType and c.EnquiryNo=d.EnquiryNo and c.EnquiryDate=d.EnquiryDate"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " and c.ItemCode=d.ItemCode"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " and c.SentItemCode=d.SentItemCode"
'
'      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid"
'      Sql = Sql & " and c.Finyear=d.Finyear"
'      Sql = Sql & " )"
'      Sql = Sql & " and a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and a.Revisionno In(" & Val(lblRevNo(1).Tag) & ")"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      Sql = Sql & " db.Execute Sql"
'
'    'Selective - MODIFIED
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select  " & Trim$(mQry) & ",'MODIFIED' As RecStatus"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " From SAL_Enquiry_Hdr a"

'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
'
'      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo and a.EnquiryDate=b.EnquiryDate"
'      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
'      Sql = Sql & " and a.Finyear=b.Finyear"
'      Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and a.RevisionNo In(" & Val(lblRevNo(1).Tag) & ")"
'      Sql = Sql & " and Exists ("
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Select * From SAL_Enquiry_Hdr c"

'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl d"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Select * From SAL_JobWorkEnquiry_Hdr c"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl d"
'
'      Sql = Sql & " On c.EnquiryType=d.EnquiryType and c.EnquiryNo=d.EnquiryNo and c.EnquiryDate=d.EnquiryDate"
'      Sql = Sql & " and c.Branchid=d.Branchid and c.Compid=d.Compid and c.Finyear=d.Finyear"
'      Sql = Sql & " Where c.EnquiryType='" & Trim$(txtType) & "' and c.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and c.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'
'      If Trim$(docflg) = "Sales" Then
'        Sql = Sql & " and b.ItemCode=d.ItemCode and c.RevisionNo In(" & i & ")"
'        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
'        Sql = Sql & " or b.UOM<>d.UOM or b.Qty<>d.Qty or b.Volume<>d.Volume"
'        Sql = Sql & " or b.AssessableRate<>d.AssessableRate"
'      ElseIf Trim$(docflg) = "Job Work" Then
'        Sql = Sql & " and b.SentItemCode=d.SentItemCode and c.RevisionNo In(" & i & ")"
'        Sql = Sql & " and (a.CustomerCode<>c.CustomerCode or a.Remarks<>c.Remarks or b.ItemType<>d.ItemType"
'        Sql = Sql & " or b.ItemGroupKey<>d.ItemGroupKey or b.ReceivedItemCode<>d.ReceivedItemCode"
'        Sql = Sql & " or b.RecdUOM<>d.RecdUOM or b.SentUOM<>d.SentUOM"
'        Sql = Sql & " or b.SendQty<>d.SendQty or b.RatePer<>d.RatePer or b.AssRate<>d.AssRate"
'        Sql = Sql & " or b.RelationType<>d.RelationType or b.isActive<>d.isActive"
'      End If
'
'      Sql = Sql & " or b.VolumePer<>d.VolumePer or b.Rate<>d.Rate"
'      Sql = Sql & " or IsNull(b.TargetDate,'')<>IsNull(d.TargetDate,'')"
'      Sql = Sql & " )"
'      Sql = Sql & " )"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'
'    'Selective - DELETED
'      Sql = "Insert Into SAL_SalseDocument_Tmp"
'      Sql = Sql & " Select  " & Trim$(mQry) & ", 'DELETED' As RecStatus"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " From SAL_Enquiry_Hdr a"

'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl b"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " From SAL_JobWorkEnquiry_Hdr a"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " Left Outer Join SAL_JobWorkEnquiry_Dtl b"
'
'      Sql = Sql & " On a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo"
'      Sql = Sql & " and a.EnquiryDate=b.EnquiryDate"
'      Sql = Sql & " and a.RevisionNo=b.RevisionNo and a.Branchid=b.Branchid and a.Compid=b.Compid"
'      Sql = Sql & " and a.Finyear=b.Finyear"
'      Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "' and a.EnquiryNo='" & Trim$(txtNo) & "'"
'      Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
'      Sql = Sql & " and a.RevisionNo In(" & i & ")"
'
'      If Trim$(docflg) = "Sales" Then: Sql = Sql & " and b.ItemCode Not In( Select c.ItemCode From SAL_Enquiry_Dtl c"

'      If Trim$(docflg) = "Job Work" Then: Sql = Sql & " and b.SentItemCode Not In( Select c.SentItemCode From SAL_JobWorkEnquiry_Dtl c"
'
'      Sql = Sql & " Where a.EnquiryType = c.EnquiryType and a.EnquiryNo = c.EnquiryNo"
'      Sql = Sql & " and a.EnquiryDate = c.EnquiryDate"
'      Sql = Sql & " and c.RevisionNo In(" & Val(lblRevNo(1).Tag) & ") and a.Branchid=c.Branchid"
'      Sql = Sql & " and a.Compid=b.Compid and a.Finyear=b.Finyear"
'      Sql = Sql & " )"
'      Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
'      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
'      db.Execute Sql
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
  
  
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
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
      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
    Else
      If Trim$(DocFlg) = "Job Work" And Trim$(RptFlg) = "Obsolute" Then
        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
      Else
        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
      End If
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
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If RptFlg = "RevisionSummary" Then
          frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
        ElseIf RptFlg = "Listing" Then
          frmPrint.sprPrint.Col = 9
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
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If Trim$(DocFlg) = "Job Work" And Trim$(RptFlg) = "Obsolute" Then
          frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 6
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

    frmPrint.sprPrint.SetText 1, j, IIf(Trim$(DocFlg) <> "", Trim$(DocFlg), "") & Space(1) & "Enquiry " & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", " - Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.FontSize = 11
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    If Trim$(RptFlg) = "View" Or Trim$(RptFlg) = "Print" Or Trim$(RptFlg) = "Obsolute" Then
      If Trim$(DocFlg) = "Job Work" Then
        If iPageNo = 1 Then
          frmPrint.sprPrint.MaxCols = 8
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 7

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 7

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

          frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 24
          
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
        frmPrint.sprPrint.Tag = j
        j = j + 16
GridCount = GridCount + 16
RowProcess
      Else
        If iPageNo = 1 Then
          frmPrint.sprPrint.MaxCols = 6
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25

          frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20

          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
          
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
        frmPrint.sprPrint.Tag = j
        j = j + 12
GridCount = GridCount + 12
RowProcess
      End If
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      If Trim$(DocFlg) = "Job Work" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Code"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Code"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Recd.Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      Else
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      End If
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
    'Bottom
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      If Trim$(DocFlg) = "Job Work" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "GroupKey"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentUom"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Else
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      End If
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    ElseIf Trim$(RptFlg) = "Listing" Or Trim$(RptFlg) = "RevisionSummary" Then 'Listing
    'Top
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 30
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "UOM"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 4
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "GoodsValue"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        ElseIf Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Tartget Date"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "RecdUom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "GoodsValue"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        ElseIf Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

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

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Target Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        ElseIf Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. Date"
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

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Group Key"
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
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentUom"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate Per"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass. Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        If Trim$(RptFlg) = "Listing" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        ElseIf Trim$(RptFlg) = "RevisionSummary" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ref. Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        End If
      End If
      
      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
            
    ElseIf Trim$(RptFlg) = "Audit" Then 'Audit
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
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
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 30
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Enq. Date"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      If Trim$(DocFlg) = "Sales" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        k = 0
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
        k = 1
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
  Dim mCustNm As String
    
  Try
  sPrintcurrent_Audit = False
  
  CurrDate = ServerDate
  
  If optCustSelective Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
  End If
  
  Sql = "Select Distinct a.EnquiryType, a.EnquiryNo, a.EnquiryDate,"
  Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As Userid,"
  Sql = Sql & " Case When a.Userid='ADMIN' Then 'ADMIN' Else IsNull(c.UserName,'') End As UserName,"
  Sql = Sql & " IsNull(a.Remarks ,'') As Remarks"
  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " ,Itemcode,ItemDescription From SAL_ENQUIRY_HDR a"
    Sql = Sql & " Left Outer Join SAL_ENQUIRY_Dtl a1 on a.EnquiryType = a1.EnquiryType"
    Sql = Sql & " and a.EnquiryNo= a1.EnquiryNo and a.EnquiryDate=a1.EnquiryDate and a.Branchid=a1.BranchID and a.Compid=a1.CompID"
  ElseIf Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & ",ItemGroupKey,ReceivedItemCode,RecdDescription,SentItemCode,SentDescription From SAL_JobworkEnquiry_Hdr a"
    Sql = Sql & " Left Outer Join SAL_JobworkEnquiry_Dtl a1 on a.EnquiryType = a1.EnquiryType"
    Sql = Sql & " and a.EnquiryNo= a1.EnquiryNo and a.EnquiryDate=a1.EnquiryDate and a.Branchid=a1.BranchID and a.Compid=a1.CompID"
  End If
  Sql = Sql & " Left Outer Join dbo.fn_SEC_Users ( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
  Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
  Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  Sql = Sql & " Where a.EnquiryDate between '" & Format(dtpFromDate.Value, "yyyy-MM-dd") & "' and '" & Format(dtpToDate.Value, "yyyy-MM-dd") & "'"

  If optCustSelective Then
'    If fpsList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": Exit Function
'    Sql = Sql & " and a.EnquiryNo in("
      Sql = Sql & " and a.EnquiryNo in "
      Sql = Sql & Spread_Get_SelectionList(fpsList, 2)
    
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = 2
'        Sql = Sql + "'" + Trim$(.Text) + "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)'    Sql = Sql + ")"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.EnquiryType, a.EnquiryNo, a.EnquiryDate"
  If Trim$(DocFlg) = "Job Work" Then
 Sql = Sql & " ,ItemGroupkey"
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
        
        If Trim$(mCustNm) <> Trim$(rs("EnquiryType")) & Trim$(rs("EnquiryNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("EnquiryType")) & Trim$(rs("EnquiryNo"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("EnquiryDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Itemcode"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          k = 0
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ReceivedItemcode")) & " [" & rs("ItemGroupkey") & "]"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentItemcode"))
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

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        If Trim$(DocFlg) = "Sales" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          k = 0
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecdDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("SentDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          k = 1
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3 + k

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("UserName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4+k

'        frmPrint.sprPrint.SetText .Col, j, Trim$(rs("Transtamp")):  .TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        mCustNm = Trim$(rs("EnquiryType")) & Trim$(rs("EnquiryNo"))
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
      sSendToPort LAlign(10, rs("EnquiryNo")) & Space(1) & LAlign(30, Trim$(rs("CustomerName"))) & Space(1) & LAlign(20, Trim$(rs("PartNo"))) & Space(1) & LAlign(4, Trim$(rs("PartUOM"))) & Space(1) & LAlign(25, Trim$(rs("ItemCode"))) & Space(1) & LAlign(12, Trim$(rs("ItemUOM"))) & Space(1) & RAlign(12, rs("Rate")) & Space(1) & RAlign(12, rs("GoodsValue"))
      sSendToPort LAlign(10, Format(rs("EnquiryDate"), "dd/mm/yyyy")) & Space(32) & LAlign(20, Trim$(rs("PartDescription"))) & Space(6) & LAlign(25, Trim$(rs("ItemDescription"))) & Space(1) & LAlign(12, rs("DCQty")) & Space(1) & RAlign(12, rs("AssRate")) & Space(1) & RAlign(12, rs("AssValue"))
      Gtot(1) = Gtot(1) + Val(rs("GoodsValue"))
      Gtot(2) = Gtot(2) + Val(rs("AssValue"))
      i = i + 1
      sSendToPort ""
      rs.MoveNext
    Loop
    
    sSendToPort Space(10) & "Goods Value Grand Total      : " & Gtot(1) & Space(10) & "Assessable Value Grand Total : " & Gtot(2)
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
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  Dim CurrDate As Date
  Dim i As Integer
  Dim k As Integer
  
  Try
  sPrintCurrent_Print = False
  
  CurrDate = ServerDate
  
  Sql = "Select Distinct a.EnquiryType, a.EnquiryNo, a.EnquiryDate,"
  Sql = Sql & " a.CustomerCode, IsNull(a.Remarks,'') As Remarks,"
  Sql = Sql & " a.RevisionNo, a.RevisionDate,"
  Sql = Sql & " Case When a.CustomerType='E' Then 'Existing Cutomer' Else 'New Customer'"
  Sql = Sql & " End As CustomerType,"
  Sql = Sql & " Case When a1.ItemType='E' Then 'Existing Item' Else 'New Item' End As ItemType, "

  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " a1.Specification,"
    Sql = Sql & " a.CustomerName, a1.ItemCode,"
    Sql = Sql & " IsNull(b.ItemDescription,'') As ItemDescription, a1.UOM,"
    Sql = Sql & " a1.Qty, a1.AssessableRate As AssRate,  "
    Sql = Sql & " IsNull(d.PartNo,'') As PartNo, IsNull(h.SDesc,'') As PartUOM, "
    Sql = Sql & " IsNull(g.ItemDescription,'') As PartDescription,"
  End If

  If Trim$(DocFlg) = "Job Work" Then
'    Sql = Sql & " a.RevisionNo, a.RevisionDate,"
    Sql = Sql & " a1.Itemgroupkey,IsNull(a.CustomerName,'') As CustomerName, a1.ReceivedItemCode As RecdItemCode, "
    Sql = Sql & " IsNull(a1.RecdDescription,'') AS RecdItemDesc, IsNull(c.SDesc,'') As RecdUOM,"
    Sql = Sql & " a1.SentItemCode, IsNull(a1.SentDescription,'') AS SentItemDesc, "
    Sql = Sql & " IsNull(e.PartNo,'') As PartNo, IsNull(g.ItemDescription,'') As PartDescription,"
    Sql = Sql & " a1.AssRate, IsNull(d.SDesc,'') As SentUOM, a1.SendQty As Qty,"
  End If

  Sql = Sql & " a1.Rate, a.ReferenceNo, a.ReferenceDate,"
  Sql = Sql & " IsNull(QuotationType,'') As QuotationType, IsNull(QuotationNo,'') As QuotationNo, IsNull(QuotationDate,'') As QuotationDate"

  If Trim$(DocFlg) = "Sales" Then
    Sql = Sql & " From SAL_Enquiry_Hdr a"
    Sql = Sql & " Left Outer Join SAL_Enquiry_Dtl a1 on a.EnquiryType = a1.EnquiryType and a.EnquiryNo= a1.EnquiryNo and a.EnquiryDate=a1.EnquiryDate"
    Sql = Sql & " and a.Branchid=a1.BranchID and a.Compid=a1.CompID"
    Sql = Sql & " Left Outer Join FAS_Party_Master c"
    Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid"
    Sql = Sql & " and a.Compid=c.Compid"
    Sql = Sql & " Left Outer Join PartMast d"
    Sql = Sql & " On a1.ItemCode=d.StageCode and a.Branchid=d.BranchID and a.Compid=d.CompID"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER b"
    Sql = Sql & " On a1.ItemCode=b.ItemCode and a.Branchid=b.Branchid and a.Compid=b.Compid"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER g"
    Sql = Sql & " On d.StageCode=g.ItemCode and a.Branchid=g.Branchid and a.Compid=g.Compid and g.status=1"
    Sql = Sql & " Left Outer Join INV_UOM_Master h"
    Sql = Sql & " On g.UOM=h.Code and a.Branchid=h.Branchid and a.Compid=h.CompId and h.Status=1"
  End If

  If Trim$(DocFlg) = "Job Work" Then
    Sql = Sql & " From  SAL_JobworkEnquiry_Hdr a"
    Sql = Sql & " Left Outer Join SAL_JobworkEnquiry_Dtl a1 on a.EnquiryType = a1.EnquiryType"
    Sql = Sql & " and a.EnquiryNo= a1.EnquiryNo and a.EnquiryDate=a1.EnquiryDate and a.Branchid=a1.BranchID and a.Compid=a1.CompID"
    Sql = Sql & " Left Outer Join Material_Master i"
    Sql = Sql & " On a1.ReceivedItemCode=i.ItemCode and a.Branchid=i.Branchid and a.Compid=i.Compid"
    Sql = Sql & " Left Outer Join INV_UOM_MASTER c"
    Sql = Sql & " On i.UOM=c.Code and a.Branchid=c.Branchid and a.Compid=c.Compid"
    Sql = Sql & " Left Outer Join Material_Master b"
    Sql = Sql & " On a1.SentItemCode=b.ItemCode and a.Branchid=b.Branchid and a.Compid=b.Compid"
    Sql = Sql & " Left Outer Join INV_UOM_MASTER d"
    Sql = Sql & " On b.UOM=d.Code and a.Branchid=d.Branchid and a.Compid=d.Compid"
    Sql = Sql & " Left Outer Join PartMast e"
    Sql = Sql & " On a1.SentItemCode=e.StageCode and a.Branchid=e.BranchID and a.Compid=e.CompID"
    Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER g"
    Sql = Sql & " On e.StageCode=g.ItemCode and a.Branchid=g.Branchid and a.Compid=g.Compid and g.status=1"
  End If

  Sql = Sql & " Where a.EnquiryType='" & Trim$(txtType) & "'"
  Sql = Sql & " and a.EnquiryNo='" & Trim$(txtNo) & "'"
  Sql = Sql & " and a.EnquiryDate='" & Format(txtDate, "yyyy-MM-dd") & "'"
  If Trim$(RptFlg) = "Obsolute" Then
    Sql = Sql & " and a.RevisionNo= " & Val(cmbRevNo) & ""  '<= " & Val(cmbRevNo) & ""
  Else
    Sql = Sql & " and a.RevisionNo=" & Val(cmbRevNo) & ""
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "'"
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
    
    Header
    j = j + 1
GridCount = GridCount + 1
RowProcess
    If Not snap.EOF() Then
        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Enquiry No & Date : " & snap("EnquiryType") & snap("EnquiryNo") & " - " & snap("EnquiryDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Customer Name     : " & snap("CustomerName")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Customer Type      : " & snap("CustomerType")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Item Type               : " & snap("ItemType")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Reference No        : " & snap("ReferenceNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

        frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Reference Date     : " & snap("ReferenceDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)
        If Trim$(DocFlg) = "Job Work" Then
          frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Revision No           : " & Val(cmbRevNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Tag = Val(frmPrint.sprPrint.Tag + 2)

          frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Space(5) & "Revision Date        : " & snap("RevisionDate")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If snap.RecordCount >= 1 Then
          snap.MoveLast
snap.MoveFirst
          If Trim$(DocFlg) = "Job Work" Then
            For i = 1 To snap.RecordCount
              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Itemgroupkey"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("RecdItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("SentItemCode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("PartNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("RecdUOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Qty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(snap("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
              j = j + 1
GridCount = GridCount + 1
RowProcess
    
              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("RecdItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("SentItemDesc"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("SentUom"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format((snap("AssRate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              snap.MoveNext
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
              j = j + 1
GridCount = GridCount + 1
RowProcess
            Next i
          Else
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
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("UOM"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Qty"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(snap("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("Remarks"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
              j = j + 1
GridCount = GridCount + 1
RowProcess
    
              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(snap("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

              frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format((snap("AssRate")), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              snap.MoveNext
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
              j = j + 1
GridCount = GridCount + 1
RowProcess
            Next i
          End If
        End If
        j = j + 3
GridCount = GridCount + 3
RowProcess
        If Trim$(DocFlg) = "Job Work" Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
        End If
        
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

        j = j + 1
GridCount = GridCount + 1
RowProcess
      
    End If
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 5
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
      sSendToPort CAlign(iPageWidth, "SALES ENQUIRY")
      sSendToPort ""
    ElseIf mDocSeq(1) = 5 Then
      Printer.Print CAlign(iPageWidth, "SALES ENQUIRY")
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

        sSendToPort Space(2) & LAlign(11, "Enquiry No.    :") & Space(1) & LAlign(10, snap("EnquiryType") & snap("EnquiryNo")) & Space(22) & LAlign(12, "Enquiry Date    :") & Space(1) & LAlign(10, snap("EnquiryDate"))
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

        Printer.Print Space(2) & LAlign(11, "Enquiry No.    :") & Space(1) & LAlign(10, snap("EnquiryType") & snap("EnquiryNo")) & Space(22) & LAlign(12, "Enquiry Date    :") & Space(1) & LAlign(10, snap("EnquiryDate"))
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

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptEnquiryDocument.frm", fpsList, "")

End Sub
