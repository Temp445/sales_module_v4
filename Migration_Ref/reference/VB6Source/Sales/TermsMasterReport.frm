VERSION 5.00
Begin VB.Form rptTermsMasterReport 
   Caption         =   "Terms Master"
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
   Begin VB.CommandButton CmdStep 
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
      Index           =   1
      Left            =   7110
      TabIndex        =   4
      Top             =   2295
      Width           =   345
   End
   Begin VB.CommandButton CmdStep 
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
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   1620
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   400
      Left            =   6585
      TabIndex        =   9
      Top             =   3840
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   400
      Left            =   5295
      TabIndex        =   8
      Top             =   3840
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Proceed"
      Height          =   400
      Left            =   4005
      TabIndex        =   7
      Top             =   3840
      Width           =   1300
   End
   Begin VB.Frame fraReportType 
      Caption         =   "Report Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   4575
      TabIndex        =   13
      Top             =   2790
      Width           =   2730
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "TermsMasterReport.frx":0000
         Left            =   90
         List            =   "TermsMasterReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   2550
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.Frame FraStep 
      Height          =   660
      Index           =   0
      Left            =   4260
      TabIndex        =   11
      Top             =   1470
      Visible         =   0   'False
      Width           =   3360
      Begin VB.ComboBox cmbTerms 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "TermsMasterReport.frx":0038
         Left            =   1155
         List            =   "TermsMasterReport.frx":0051
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Term"
         Top             =   225
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms Master"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   990
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Select"
      Height          =   600
      Index           =   1
      Left            =   4845
      TabIndex        =   12
      Top             =   2145
      Width           =   2190
      Begin VB.OptionButton optSelective 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1110
         TabIndex        =   3
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   225
         Left            =   135
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame fraCodeDet 
      Caption         =   "Terms Master"
      Height          =   1620
      Left            =   7605
      TabIndex        =   14
      Top             =   2115
      Width           =   4260
Begin AceSpread sprList
         Height          =   1365
         Left            =   75
         TabIndex        =   15
         Top             =   180
         Width           =   4095
         _Version        =   458752
         _ExtentX        =   7223
         _ExtentY        =   2408
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   3
         OperationMode   =   4
         RestrictRows    =   -1  'True
      End
   End
End
Attribute VB_Name = "rptTermsMasterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public rptName As String

Private sprow As Long
Private Erpno As String
Private TsNo As String

Private Enum eDesc
  mCode = 1
  mDescription = 2
  mType = 3
End Enum

Private Sub Sp_Help()
  Dim Sql As String
  Dim Cnt As Long
  Dim rCnt As Long

  Try
  
  rCnt = sprList.DataRowCnt + 1
  Select Case rptName
    Case "Terms Master"
      sprList.Col = 3
sprList.ColHidden = True
      Sql = " Select Code, TypeFlg, Description Name, Case Status When '1' Then 'Active' Else 'InActive' End as Status From TERMSMASTER a"
      Sql = Sql & " Where TypeFlg in ('" & Trim$(CmbTerms.Text) & "')"
      Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' "
      If FetchValue_New(True, db.ConnectionString, Sql, "0010", "List of Descriptions", "Description") = True Then
           sprList.ClearSelection
          sprList.MaxRows = sprList.DataRowCnt + UBound(getvalue, 1) + 1
          For Cnt = 1 To UBound(getvalue, 1)
            sprList.SetText eDesc.mCode, rCnt, getvalue(Cnt, 1)
            sprList.SetText eDesc.mDescription, rCnt, getvalue(Cnt, 3)
            Call Avoid_Duplicate(sprList, eDesc.mDescription)
            rCnt = rCnt + 1
          Next Cnt
        sprList.MaxRows = sprList.DataRowCnt + 1
      End If

    Case "Coutry/City Master"
      sprList.Col = 3
sprList.ColHidden = False
      Sql = " Select CountryName Name,CountryCode Code,"
      Sql = Sql & " Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
      'Sql = " Select CountryName Name, (Select CountryName From SAL_COUNTRYCITY_MASTER Where ParentCode = 1 and BranchID = a.BranchID"
      'Sql = Sql & " and CompID = a.CompID)as CityName, CountryCode Code, Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
      Sql = Sql & " From SAL_COUNTRYCITY_MASTER a    "
      Sql = Sql & " Where ParentCode = 0 and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"

    Case "Formula Master"
      sprList.Col = 3
sprList.ColHidden = True
      Sql = " Select FrmID Code, FrmLName Name, FrmDesc, IsActive from Com_FormulaMast_Hdr"
      Sql = Sql & " Where ModuleCode=4 and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"

  End Select
  
  
  If sprList.DataRowCnt >= 1 Then
      CmdStep(1).Enabled = True
    Else
      CmdStep(1).Enabled = False
  End If
  
  fraCodeDet.Visible = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sSpread_FormulaMasterDisplay() As Boolean
  Dim Sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rsExpr As AceADODB.Recordset
  Dim setstr As String
  Dim i As Long
  Dim FrmID As Integer
  Dim Status As String
  Dim RoundOFf As String
  Dim GrandTotalFlag As String
  Dim PostingAcc As String
  Dim retVal As Variant
  Dim Spr As AceSpread
  
  Try
  
  sSpread_FormulaMasterDisplay = False
  
  Sql = " Select FrmID, FrmSName, FrmLName, FrmDesc, IsActive from Com_FormulaMast_Hdr"
  Sql = Sql & " Where ModuleCode=4 and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then sShowMessage "Please select atlease one item"
GoTo ExitHere
    sprList.GetText 1, sprList.ActiveRow, retVal
    Sql = Sql & " and FrmID in ('" & Trim$(retVal) & "')"
  End If
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_FormulaMasterDisplay = False
    GoTo ExitHere
  End If
  
  Sql = " Select a.FrmID, "
  Sql = Sql & " Case When Convert(int, a.ExprRoute) < 0 then 'SUBTOTAL' Else  b.ExprGroupName End as ExprName,"
  Sql = Sql & " Isnull(dbo.fn_FrontEndExpression(4, a.FrmID, a.ExprSeq, '" & gBranchID & "', '" & gCompID & "'), '') Expr, a.RoundOff, a.RoundOffValue,"
  Sql = Sql & " Case a.GrandTotalFlag When 1 then '+' When -1 then '-' Else '' End as GTFlag,"
  Sql = Sql & " Case a.PostingAc When -1 then 'AS PER ORDER' When 0 Then '' Else (Select LedgerName from"
  Sql = Sql & " Fas_LedgerMaster Where LedgerCode=a.PostingAc and CompCode=a.CompID + a.BranchID) End as LedgerName, a.ExprSeq"
  Sql = Sql & " from Com_FormulaMast_Dtl a"
  Sql = Sql & " Left Outer Join COM_FormulaExpr_Group b"
  Sql = Sql & " On (b.Route+b.Width)=a.ExprRoute and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID"
  Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  Sql = Sql & " Order by a.FrmID, a.ExprSeq "
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  Sql = " Select a.FrmID,"
  Sql = Sql & " Case When Convert(int, a.ExprRoute) < 0 then 'SUBTOTAL' Else  b.ExprGroupName End as ExprName,"
  Sql = Sql & " Isnull(dbo.fn_FrontEndExpression(4, a.FrmID, a.ExprSeq, '" & gBranchID & "', '" & gCompID & "'), '') Expr"
  Sql = Sql & " from Com_FormulaMast_Dtl a"
  Sql = Sql & " Left Outer Join COM_FormulaExpr_Group b"
  Sql = Sql & " On (b.Route+b.Width)=a.ExprRoute and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID"
  Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Group by a.FrmID, a.ExprSeq, a.ExprRoute, b.ExprGroupName"
  Set rsExpr = New AceADODB.Recordset
  If RSOpen(rsExpr, Sql) = False Then
    GoTo ExitHere
  End If
  
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Formula Master Report"
  
  Set Spr = frmPrint.sprPrint
  
  Call spSetColWidth(Spr, 1, 20, 2, 40, 3, 20)
  
  Call Header
 
  i = 1
  Do While Not rsDetails.EOF
    If rsDetails("IsActive") = "True" Then
      Status = "Active"
    Else
      Status = "InActive"
    End If
    
    If FrmID <> rsDetails("FrmID") Then
      Call spSetCellBorder(Spr, 1, sprow, 3, , mBorderAllSides)
      setstr = rsDetails("FrmSName") & " [ " & rsDetails("FrmLName") & " ] "
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, True, "Formula Name ", Trim$(setstr), "Status : " & Trim$(Status))

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, True, "Description  ", rsDetails("FrmDesc"))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
    
    rs.Filter = adFilterNone
    If Not rs.EOF Then
      rs.Filter = " FrmID= " & rsDetails("FrmID") & ""
      Do While Not rs.EOF
        Select Case rs("RoundOff")
          Case "N"
            RoundOFf = "None"
          Case "B"
            RoundOFf = "Nearest"
          Case "U"
            RoundOFf = "UpWard/DownWard"
        End Select
        
        If rs("GTFlag") <> "" Then
          GrandTotalFlag = LAlign(14, "GrandTotalFlag ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(15, rs("GTFlag"))
          PostingAcc = LAlign(14, "Posting A/C  ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(15, rs("LedgerName"))
        Else
          GrandTotalFlag = ""
          PostingAcc = ""
        End If
        
        Call spSetCellBorder(Spr, 1, sprow, 3, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, "Expression Name ", rs("ExprName"), Trim$(GrandTotalFlag))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        
        Call spSetCellBorder(Spr, 1, sprow, 3, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, "Round Off      ", Trim$(RoundOFf), Trim$(PostingAcc))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        
        Call spSetCellBorder(Spr, 1, sprow, 3, , mBorderAllSides)
        Call spSetText(Spr, 1, sprow, "RoundOff Value ", mAlignLeftCenter, False)
        Call spSetText(Spr, 2, sprow, rs("RoundOffValue"), mAlignLeftCenter, False, 2)
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow

        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        rs.MoveNext
      Loop
    End If
    rsExpr.Filter = adFilterNone
    If Not rsExpr.EOF Then
      rsExpr.Filter = " FrmID= " & rsDetails("FrmID") & ""
      Call spSetText(Spr, 1, sprow, "Expression", mAlignLeftCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, 3, sprow, mBorderBottom)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Do While Not rsExpr.EOF
        If Trim$(rsExpr("Expr")) <> "" Then
          Call spSetCellBorder(Spr, 1, sprow, 3, sprow, mBorderAllSides)
          Call spSetText(Spr, 1, sprow, Left$(Trim$(rsExpr("ExprName")), 28), mAlignLeftCenter, False)
          sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        End If
        rsExpr.MoveNext
      Loop
    End If
  
    If i <> rsDetails.RecordCount Then
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
    
    FrmID = rsDetails("FrmID")
    i = i + 1
    rsDetails.MoveNext
  Loop
  
  totalpages = iPageNo
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  frmPrint.Show
  sSpread_FormulaMasterDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDetails = Nothing
  Set rsExpr = Nothing
End Try
End Function

Private Function sSpread_CountryCityDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim CountryCode As Integer
  Dim i As Long
  Dim retVal As Variant
  
  Try
  
  sSpread_CountryCityDisplay = False
  
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then
 sShowMessage " Please select atlease one item "
GoTo ExitHere
End If
  End If
  
  Sql = " Select CountryName,CountryCode CountryCode,"
  Sql = Sql & " Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
  Sql = Sql & " From SAL_COUNTRYCITY_MASTER a    "
  Sql = Sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If optSelective.Value = True Then
    sprList.GetText 1, sprList.ActiveRow, retVal
    Sql = Sql & " and ParentCode in ( 0, '" & Trim$(retVal) & "')      "
  End If
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_CountryCityDisplay = False
    GoTo ExitHere
  End If
 
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Country/City Master Report"
  
  Call Header
  
  i = 1
  Do While Not rsDetails.EOF
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
  
    If i = 1 And CountryCode <> rsDetails("CountryCode") Then
      spLAlign frmPrint.sprPrint, 1, sprow, 20, CheckNull(rsDetails("CountryName"))
      spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("Type"))
    ElseIf CountryCode <> rsDetails("CountryCode") Then
      spLAlign frmPrint.sprPrint, 2, sprow, 20, CheckNull(rsDetails("CountryName"))
      spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("Type"))
    Else
      spLAlign frmPrint.sprPrint, 2, sprow, 20, CheckNull(rsDetails("CountryName"))
      spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("Type"))
    End If
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  
    i = i + 1
    CountryCode = rsDetails("CountryCode")
    rsDetails.MoveNext
  Loop
  
  totalpages = iPageNo
  frmPrint.Show
  
  sSpread_CountryCityDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
 
  Try
  
  sSpread_sPrintcurrent = False
  
  Select Case rptName
    Case "Terms Master"
      If sSpread_TermsMasterDisplay = False Then
        GoTo ExitHere
      End If
    Case "Coutry/City Master"
      If sSpread_CountryCityDisplay = False Then
        GoTo ExitHere
      End If
    Case "Formula Master"
      If sSpread_FormulaMasterDisplay = False Then
        GoTo ExitHere
      End If
  End Select

  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub Header()
  Try
  
  Select Case rptName
    Case "Terms Master"
      Call RptStandard_Header(3, sprow, Erpno, TsNo, chkPrintdate)
      GridCount = GridCount + 6
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      spCAlign frmPrint.sprPrint, 1, sprow, 20, "Type Flag", True
      spCAlign frmPrint.sprPrint, 2, sprow, 40, "Description ", True
      spCAlign frmPrint.sprPrint, 3, sprow, 10, "Status", True
      sprow = sprow + 1
GridCount = GridCount + 1
    Case "Formula Master"
      Call RptStandard_Header(3, sprow, Erpno, TsNo, chkPrintdate.Value)
      GridCount = GridCount + 6
      
      Call spSetCellBorder(frmPrint.sprPrint, 1, sprow, 3, , mBorderTop)
    
'        sSendToPort String(iPageWidth, "-")
'        sSendToPort LAlign(30, "ParameterName") & Space(1) & LAlign(20, "Parameter1") & Space(1) & LAlign(4, "Operator") & Space(1) & LAlign(20, "Parameter2") & Space(1) & LAlign(5, "RoundingOff") & Space(1) & LAlign(8, "RoundingOfValue") & Space(1) & LAlign(25, "PostingName") & Space(1) & LAlign(7, "IsPrint") & Space(1) & LAlign(3, "Seq")
'        sSendToPort Space(78) & LAlign(5, "Off") & Space(1) & LAlign(8, "OfValue")
    Case "Coutry/City Master"
      Call RptStandard_Header(3, sprow, Erpno, TsNo, chkPrintdate)
      GridCount = GridCount + 6
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
      spCAlign frmPrint.sprPrint, 1, sprow, 20, "CountryName", True
      spCAlign frmPrint.sprPrint, 2, sprow, 20, "CityName ", True
      spCAlign frmPrint.sprPrint, 3, sprow, 10, "Type", True
      sprow = sprow + 1
GridCount = GridCount + 1
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sSpread_TermsMasterDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim TypeFlg As String
  Dim i As Long
  Dim retVal As Variant
  
  Try
  
  sSpread_TermsMasterDisplay = False
  
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then
 sShowMessage " Please select atlease one item "
GoTo ExitHere
End If
  End If
  
  Sql = " Select Code, TypeFlg, Description, Case Status When '1' Then 'Active' Else 'InActive' End as Status From TERMSMASTER a"
  Sql = Sql & " Where TypeFlg in ('" & Trim$(CmbTerms.Text) & "') "
  Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' "
  If optSelective.Value = True Then
'    If sprList.SelectionCount < 1 Then: sShowMessage " Please select atlease one item ": GoTo ExitHere
      Sql = Sql & " and Code in "
      Sql = Sql & Spread_Get_SelectionList(sprList, 1)
'    sprList.GetText 1, sprList.ActiveRow, retVal
'    Sql = Sql & " and Code in ('" & Trim$(retVal) & "')      "
  End If
  Sql = Sql & " Order by TypeFlg"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_TermsMasterDisplay = False
    GoTo ExitHere
  End If
 
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Terms Master Report "
 
  Call Header
  
  
  Do While Not rsDetails.EOF
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, sprow, frmPrint.sprPrint.MaxCols, sprow
    If TypeFlg <> rsDetails("TypeFlg") Then
        spLAlign frmPrint.sprPrint, 1, sprow, 20, CheckNull(rsDetails("TypeFlg"))
        spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(rsDetails("Description"))
        spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("Status"))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
     Else
        spLAlign frmPrint.sprPrint, 2, sprow, 40, CheckNull(rsDetails("Description"))
        spLAlign frmPrint.sprPrint, 3, sprow, 10, CheckNull(rsDetails("Status"))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
     End If
     i = i + 1
     TypeFlg = rsDetails("TypeFlg")
     rsDetails.MoveNext
  Loop
  
  sSpread_TermsMasterDisplay = True
  
  totalpages = iPageNo
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  frmPrint.Show
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function CountryCityDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim CountryCode As Integer
  Dim i As Long
  Dim retVal As Variant
  
  Try
  
  CountryCityDisplay = False
  
  '-----Modify Somuk
    'Sql = Sql & " Case ParentCode When '1' Then CountryName Else 'Country' End as Cityname"
    'Sql = " Select CountryName, (Select CountryName From SAL_COUNTRYCITY_MASTER Where ParentCode = 1 and BranchID = a.BranchID"
    'Sql = Sql & " and CompID = a.CompID)as CityName, CountryCode, Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
  '------
  
  Sql = " Select CountryName ,CountryCode CountryCode,"
  Sql = Sql & " Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
  Sql = Sql & " From SAL_COUNTRYCITY_MASTER a    "
  Sql = Sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then
 sShowMessage " Please select atlease one item "
GoTo ExitHere
End If
    sprList.GetText 1, sprList.ActiveRow, retVal
    Sql = Sql & " and ParentCode in ( 0, '" & Trim$(retVal) & "')      "
  End If
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    CountryCityDisplay = False
    GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 1 Then
    i = 0
    iPageWidth = 80
    
    iLineNo = 1
    pheight = 75
    pPaper = vbPRPSA4
    Orient = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    
    sSendToPort ""
    gStrReportHeading = "Country/City Master Report"
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    
    i = 1
    CountryCode = 0
    Do While Not rsDetails.EOF
      If i = 1 And CountryCode <> rsDetails("CountryCode") Then
        sSendToPort LAlign(35, rsDetails("CountryName")) & Space(37) & LAlign(15, rsDetails("Type"))
      ElseIf CountryCode <> rsDetails("CountryCode") Then
        sSendToPort Space(36) & LAlign(35, rsDetails("CountryName")) & Space(1) & LAlign(15, rsDetails("Type"))
      Else
        sSendToPort Space(36) & LAlign(35, rsDetails("CountryName")) & Space(1) & LAlign(15, rsDetails("Type"))
      End If
      If i <> rsDetails.RecordCount Then sSendToPort ""
      i = i + 1
      CountryCode = rsDetails("CountryCode")
      rsDetails.MoveNext
    Loop
    
    sSendToPort ""
    rsDetails.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
'    With spView
'      i = 5: CountryCode = 0
'      .ClearRange 1, i, .MaxCols, .MaxRows - 1, True
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexLeft, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexRight, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexTop, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexBottom, 0, CellBorderStyleBlank
'      .SetText 1, 1, gCompName
'      .SetText 1, 2, "Terms Master Report"
'      If chkPrintDate.Value = 1 Then: .SetText 2, 3, "Printed. Date : " & Format(ServerDate, "dd/MM/yyyy")
'      .SetText 1, 4, "COUNTRY NAME"
'      .SetText 2, 4, "CITY NAME"
'      .SetText 3, 4, "TYPE"
'
'      Do While Not rsDetails.EOF
'        If i = 5 And CountryCode <> rsDetails("CountryCode") Then
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexRight, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexBottom, 0, CellBorderStyleSolid
'          .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True: .FontSize = 9
'
'          .SetText 1, i, rsDetails("CountryName")
'          .SetText 3, i, rsDetails("Type")
'        Else
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexRight, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexBottom, 0, CellBorderStyleSolid
'          .Col = 2: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = False
'          .SetText 2, i, rsDetails("CountryName")
'          .Col = 3: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = False
'          .SetText 3, i, rsDetails("Type")
'        End If
'        i = i + 1
'        CountryCode = rsDetails("CountryCode")
'        rsDetails.MoveNext
'      Loop
'      .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'      .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = True: .FontSize = 10
'      .SetText 1, .MaxRows, gCompName & Space(10)
'      .Col = 3: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = True

'      .SetText 3, .MaxRows, "Page : 1"
'    End With
'    frmsp.Visible = True
'    frmsp.Height = 7455
'    frmsp.Width = 10950
'    spView.Height = 6690
'    spView.Width = 10260
  End If
  CountryCityDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function TermsMasterDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim TypeFlg As String
  Dim i As Long
  Dim retVal As Variant
  
  Try
  
  TermsMasterDisplay = False
  
  Sql = " Select Code, TypeFlg, Description, Case Status When '1' Then 'Active' Else 'InActive' End as Status From TERMSMASTER a"
  Sql = Sql & " Where TypeFlg in ('" & Trim$(CmbTerms.Text) & "') "
  Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' "
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then
 sShowMessage " Please select atleast one item "
GoTo ExitHere
End If
    sprList.GetText 1, sprList.ActiveRow, retVal
    Sql = Sql & " and Code in ('" & Trim$(retVal) & "')      "
  End If
  Sql = Sql & " Order by TypeFlg"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    TermsMasterDisplay = False
    GoTo ExitHere
  End If
  If cmbReportType.ListIndex = 1 Then
    i = 0
    iPageWidth = 80
    iLineNo = 1
    pheight = 75
    pPaper = vbPRPSA4
    Orient = 1
    
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    
    sSendToPort ""
    gStrReportHeading = "Terms Master Report"
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
   
    i = 1
    TypeFlg = ""
    Do While Not rsDetails.EOF
      If TypeFlg <> rsDetails("TypeFlg") Then
        sSendToPort LAlign(31, rsDetails("TypeFlg")) & Space(1) & LAlign(37, rsDetails("Description")) & Space(1) & LAlign(8, rsDetails("Status"))
      Else
        sSendToPort Space(32) & LAlign(37, rsDetails("Description")) & Space(1) & LAlign(8, rsDetails("Status"))
      End If
      If i <> rsDetails.RecordCount Then sSendToPort ""
      i = i + 1
      TypeFlg = rsDetails("TypeFlg")
      rsDetails.MoveNext
    Loop
  
    rsDetails.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
'    With spView
'      i = 5
'      TypeFlg = ""
'      .ClearRange 1, i, .MaxCols, .MaxRows - 1, True
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexLeft, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexRight, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexTop, 0, CellBorderStyleBlank
'      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexBottom, 0, CellBorderStyleBlank
'
'      Do While Not rsDetails.EOF
'        .SetText 1, 1, gCompName
'        .SetText 1, 2, "Terms Master Report"
'        If chkPrintDate.Value = 1 Then: .SetText 2, 3, "Printed. Date : " & Format(ServerDate, "dd/MM/yyyy")
'        .SetText 1, 4, "TYPE FLAG"
'        .SetText 2, 4, "DESCRIPTION"
'        .SetText 3, 4, "STATUS"
'
'        If TypeFlg <> rsDetails("TypeFlg") Then
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexRight, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexBottom, 0, CellBorderStyleSolid
'          .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True: .FontSize = 9
'          .SetText 1, i, rsDetails("TypeFlg")
'          .SetText 2, i, rsDetails("Description")
'          .SetText 3, i, rsDetails("Status")
'        Else
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexRight, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'          .SetCellBorder 2, i, .MaxCols, i, CellBorderIndexBottom, 0, CellBorderStyleSolid
'          .Col = 2: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = False
'          .SetText 2, i, rsDetails("Description")
'          .Col = 3: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = False
'          .SetText 3, i, rsDetails("Status")
'        End If
'        i = i + 1
'        TypeFlg = rsDetails("TypeFlg")
'        rsDetails.MoveNext
'      Loop
'      .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'      .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = True: .FontSize = 10
'      .SetText 1, .MaxRows, gCompName & Space(10)
'      .Col = 3: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter: .Font = "Arial": .FontBold = True

'      .SetText 3, .MaxRows, "Page : 1"
'    End With
'    frmsp.Visible = True
'    frmsp.Height = 7455
'    frmsp.Width = 10950
'    spView.Height = 6690
'    spView.Width = 10260
  End If
  
  TermsMasterDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sPageHeader()
  Try

  sSendToPort ""
  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  
  Select Case rptName
    Case "Terms Master"
      sSendToPort LAlign(31, "TypeFlag") & Space(1) & LAlign(37, "Description") & Space(1) & LAlign(8, "Status")
    Case "Formula Master"
'        sSendToPort String(iPageWidth, "-")
'        sSendToPort LAlign(30, "ParameterName") & Space(1) & LAlign(20, "Parameter1") & Space(1) & LAlign(4, "Operator") & Space(1) & LAlign(20, "Parameter2") & Space(1) & LAlign(5, "RoundingOff") & Space(1) & LAlign(8, "RoundingOfValue") & Space(1) & LAlign(25, "PostingName") & Space(1) & LAlign(7, "IsPrint") & Space(1) & LAlign(3, "Seq")
'        sSendToPort Space(78) & LAlign(5, "Off") & Space(1) & LAlign(8, "OfValue")
    Case "Coutry/City Master"
      sSendToPort LAlign(35, "CountryName") & Space(1) & LAlign(35, "CityName") & Space(1) & LAlign(15, "Type")
  End Select
    
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintCurrent() As Boolean
 
  Try
  sPrintCurrent = False
  
  Select Case rptName
    Case "Terms Master"
      If TermsMasterDisplay = False Then
        GoTo ExitHere
      End If
    Case "Coutry/City Master"
      If CountryCityDisplay = False Then
        GoTo ExitHere
      End If
    Case "Formula Master"
      If FormulaMasterDisplay = False Then
        GoTo ExitHere
      End If
  End Select

  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function FormulaMasterDisplay() As Boolean
  Dim Sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rsExpr As AceADODB.Recordset
  
  Dim i As Long
  Dim FrmID As Integer
  Dim Status As String
  Dim RoundOFf As String
  Dim GrandTotalFlag As String
  Dim PostingAcc As String
  Dim retVal As Variant
  
  Try
  
  FormulaMasterDisplay = False
  
  Sql = " Select FrmID, FrmSName, FrmLName, FrmDesc, IsActive from Com_FormulaMast_Hdr"
  Sql = Sql & " Where ModuleCode=4 and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If optSelective.Value = True Then
    If sprList.SelectionCount < 1 Then sShowMessage "Please select atlease one item"
GoTo ExitHere
    sprList.GetText 1, sprList.ActiveRow, retVal
    Sql = Sql & " and FrmID in ('" & Trim$(retVal) & "')"
  End If
  
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    FormulaMasterDisplay = False
    GoTo ExitHere
  End If
  
  Sql = " Select a.FrmID, "
  Sql = Sql & " Case When Convert(int, a.ExprRoute) < 0 then 'SUBTOTAL' Else  b.ExprGroupName End as ExprName,"
  Sql = Sql & " Isnull(dbo.fn_FrontEndExpression(4, a.FrmID, a.ExprSeq, '" & gBranchID & "', '" & gCompID & "'), '') Expr, a.RoundOff, a.RoundOffValue,"
  Sql = Sql & " Case a.GrandTotalFlag When 1 then '+' When -1 then '-' Else '' End as GTFlag,"
  Sql = Sql & " Case a.PostingAc When -1 then 'AS PER ORDER' When 0 Then '' Else (Select LedgerName from"
  Sql = Sql & " Fas_LedgerMaster Where LedgerCode=a.PostingAc and CompCode=a.CompID + a.BranchID) End as LedgerName, a.ExprSeq"
  Sql = Sql & " from Com_FormulaMast_Dtl a"
  Sql = Sql & " Left Outer Join COM_FormulaExpr_Group b"
  Sql = Sql & " On (b.Route+b.Width)=a.ExprRoute and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID"
  Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  Sql = Sql & " Order by a.FrmID, a.ExprSeq "
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  Sql = " Select a.FrmID,"
  Sql = Sql & " Case When Convert(int, a.ExprRoute) < 0 then 'SUBTOTAL' Else  b.ExprGroupName End as ExprName,"
  Sql = Sql & " Isnull(dbo.fn_FrontEndExpression(4, a.FrmID, a.ExprSeq, '" & gBranchID & "', '" & gCompID & "'), '') Expr"
  Sql = Sql & " from Com_FormulaMast_Dtl a"
  Sql = Sql & " Left Outer Join COM_FormulaExpr_Group b"
  Sql = Sql & " On (b.Route+b.Width)=a.ExprRoute and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID"
  Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Group by a.FrmID, a.ExprSeq, a.ExprRoute, b.ExprGroupName"
  Set rsExpr = New AceADODB.Recordset
  If RSOpen(rsExpr, Sql) = False Then
    GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  
  sSendToPort ""
  gStrReportHeading = "Formula Master Report"
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
'  sPageHeader
  sSendToPort String(iPageWidth, "-")
  i = 1
  Do While Not rsDetails.EOF
    If rsDetails("IsActive") = "True" Then
      Status = "Active"
    Else
      Status = "InActive"
    End If
    If FrmID <> rsDetails("FrmID") Then
      sSendToPort LAlign(15, "Formula Name") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(42, rsDetails("FrmSName") & " [ " & rsDetails("FrmLName") & " ] ") & Space(1) & LAlign(7, "Status") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(7, Trim$(Status))
      sSendToPort LAlign(15, "Description") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(50, rsDetails("FrmDesc"))
      sSendToPort ""
    End If
    rs.Filter = adFilterNone
    If Not rs.EOF Then
      rs.Filter = " FrmID= " & rsDetails("FrmID") & ""
      Do While Not rs.EOF
        Select Case rs("RoundOff")
          Case "N"
            RoundOFf = "None"
          Case "B"
            RoundOFf = "Nearest"
          Case "U"
            RoundOFf = "UpWard/DownWard"
        End Select
        
        If rs("GTFlag") <> "" Then
          GrandTotalFlag = LAlign(14, "GrandTotalFlag ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(15, rs("GTFlag"))
          PostingAcc = LAlign(14, "Posting A/C  ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(15, rs("LedgerName"))
        Else
          GrandTotalFlag = ""
          PostingAcc = ""
        End If
        
        sSendToPort LAlign(15, "Expression Name") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(30, rs("ExprName")) & Space(1) & Trim$(GrandTotalFlag)
        sSendToPort LAlign(15, "Round Off      ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(30, Trim$(RoundOFf)) & Space(1) & Trim$(PostingAcc)
        sSendToPort LAlign(15, "RoundOff Value ") & Space(1) & RAlign(1, ":") & Space(1) & LAlign(30, rs("RoundOffValue"))
        sSendToPort ""
        rs.MoveNext
      Loop
    End If
    rsExpr.Filter = adFilterNone
    If Not rsExpr.EOF Then
      rsExpr.Filter = " FrmID= " & rsDetails("FrmID") & ""
      sSendToPort LAlign(80, "Expression")
      sSendToPort LAlign(80, "----------")
      Do While Not rsExpr.EOF
        If Trim$(rsExpr("Expr")) <> "" Then
          sSendToPort LAlign(28, Left$(Trim$(rsExpr("ExprName")), 28)) & ": " & LAlign(60, Trim$(rsExpr("Expr")))
        End If
        rsExpr.MoveNext
      Loop
    End If
    
    If i <> rsDetails.RecordCount Then
      sSendToPort ""
      sSendToPort String(iPageWidth, "-")
    End If
    
    FrmID = rsDetails("FrmID")
    i = i + 1
    rsDetails.MoveNext
  Loop
  
  sSendToPort ""
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  
  FormulaMasterDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDetails = Nothing
  Set rsExpr = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  Form_Clear Me
  chkPrintdate.Value = 1
  cmbReportType.ListIndex = 0
  StepClear Me
  StepEnable Me
  CmbTerms.ListIndex = 0
  fraCodeDet.Visible = False
  optAll.Value = True
  If rptName = "Terms Master" Then
    CmdStep(1).Enabled = False
  ElseIf rptName = "Formula Master" Then
    CmdStep(0).Enabled = True
  End If
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdprint_Click()
    
  Try
  Me.MousePointer = vbHourglass
  
'  With spView
'    .PrintOrientation = PrintOrientationPortrait
'    .PrintPageOrder = PageOrderDownThenOver
'    .PrintFirstPageNumber = 1
'    .PrintPageStart = 1
'    .PrintScalingMethod = PrintScalingMethodSmartPrint
'    .BorderStyle = BorderStyleFixedSingle
'    .PrintBorder = True
'    .PrintCenterOnPageH = True
'    .PrintCenterOnPageV = True
'    .PrintMarginTop = 500
'    .PrintMarginLeft = 50
'    .PrintMarginRight = 1.25
'    .Action = ActionPrint
'  End With
  
  sShowMessage "Print job over"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next

  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdView.Enabled = True
  End If
  fraCodeDet.Enabled = False
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If cmbReportType <> "Spread Format Report" Then
    If sPrintCurrent = False Then
 GoTo ExitHere
End If
    Call FinaliseAndShowPrintPreview
  Else
    Erpno = Get_ERPandTSNo(Me.Tag, "SAL", mErp)
    TsNo = Get_ERPandTSNo(Me.Tag, "SAL", mTs)
    If sSpread_sPrintcurrent = False Then
 GoTo ExitHere
End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
    
  Select Case rptName
    Case "Terms Master"
      fraStep(0).Visible = True
      CmdStep(0).Enabled = True
      rptTermsMasterReport.Caption = "Terms Master Report"
      fraStep(1).Caption = "Terms Master"
'      fraDetails.Caption = "Terms Master"
    Case "Formula Master"
      fraStep(0).Visible = False
      CmdStep(0).Visible = False
      CmdStep(1).Enabled = True
'      fraTermsType.Visible = False
      rptTermsMasterReport.Caption = "Formula Master Report"
      fraStep(1).Caption = "Formula Master"
'      fraDetails.Caption = "Formula Master"
    Case "Coutry/City Master"
      fraStep(0).Visible = False
      CmdStep(0).Visible = False
      CmdStep(1).Enabled = True
'      fraTermsType.Visible = False
      rptTermsMasterReport.Caption = "Coutry/City Master Report"
      fraStep(1).Caption = "Coutry/City Master"
'      fraDetails.Caption = "Coutry/City Master"
  End Select
  
  CmbTerms.ListIndex = 0
  fraCodeDet.Visible = False
  cmbReportType.ListIndex = 0
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
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  
  Call Spread_Clear(sprList)
  sprList.MaxRows = 1
  fraCodeDet.Enabled = False
  fraCodeDet.Visible = False
  CmdStep(1).Enabled = True
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
'  Call Sp_Help
  
  CmdStep(1).Enabled = False
  fraCodeDet.Enabled = True
  fraCodeDet.Visible = True
  sprList.SetFocus
  
  Select Case rptName
    Case "Terms Master"
      sprList.Col = 3
sprList.ColHidden = True
      Sql = " Select Code, TypeFlg, Description Name, Case Status When '1' Then 'Active' Else 'InActive' End as Status From TERMSMASTER a"
      Sql = Sql & " Where TypeFlg in ('" & Trim$(CmbTerms.Text) & "')"
      Sql = Sql & " and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' "
    Case "Coutry/City Master"
      sprList.Col = 3
sprList.ColHidden = False
      Sql = " Select CountryName Name,CountryCode Code,"
      Sql = Sql & " Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
      'Sql = " Select CountryName Name, (Select CountryName From SAL_COUNTRYCITY_MASTER Where ParentCode = 1 and BranchID = a.BranchID"
      'Sql = Sql & " and CompID = a.CompID)as CityName, CountryCode Code, Case ParentCode When '0' Then 'Country' Else 'City' End as Type"
      Sql = Sql & " From SAL_COUNTRYCITY_MASTER a    "
      Sql = Sql & " Where ParentCode = 0 and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
    Case "Formula Master"
      sprList.Col = 3
sprList.ColHidden = True
      Sql = " Select FrmID Code, FrmLName Name, FrmDesc, IsActive from Com_FormulaMast_Hdr"
      Sql = Sql & " Where ModuleCode=4 and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  End Select

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
    If Not rs.EOF Then
      sprList.MaxRows = 1
      sprList.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprList.SetText 1, rCnt, rs("Code")
        sprList.SetText 2, rCnt, rs("Name")
        If Trim$(rptName) = "Coutry/City Master" Then
 sprList.SetText 3, rCnt, rs("Type")
End If
        rs.MoveNext
      Next rCnt
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If sprList.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub sprList_DblClick(sender as object, e as DblClickEventArgs)
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Call Sp_Help
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

Private Sub sprList_KeyDown(sender as object, e as KeyDownEventArgs)
'  On Error Resume Next
'  If KeyCode = vbKeyF2 Then: Call Sp_Help
End Sub

Private Sub sprList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If sprList.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub Design_sprList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprList


sprList.Click +=  new EventHandler(sprList_Click)
sprList.DblClick +=  new EventHandler(sprList_DblClick)
sprList.KeyDown +=  new EventHandler(sprList_KeyDown)
sprList.KeyUp +=  new EventHandler(sprList_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\TermsMasterReport.frm", sprList, "")

End Sub
