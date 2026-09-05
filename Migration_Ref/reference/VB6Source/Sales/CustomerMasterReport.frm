VERSION 5.00
Begin VB.Form rptCustomerMasterReport 
   Caption         =   "Customer Master Report"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   10980
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
      Index           =   2
      Left            =   7095
      TabIndex        =   13
      Top             =   2220
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
      Index           =   1
      Left            =   7380
      TabIndex        =   9
      Top             =   1635
      Width           =   345
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
      Left            =   7395
      TabIndex        =   4
      Top             =   1050
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6330
      TabIndex        =   19
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   3705
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   405
      Left            =   5175
      TabIndex        =   18
      Top             =   3705
      Width           =   1170
   End
   Begin VB.Frame fraStep 
      Caption         =   "Select Status"
      Height          =   615
      Index           =   1
      Left            =   4185
      TabIndex        =   8
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   1485
      Width           =   3150
      Begin VB.OptionButton optInActive 
         Caption         =   "Inactive"
         Height          =   255
         Left            =   2145
         TabIndex        =   7
         Top             =   255
         Width           =   945
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active"
         Height          =   255
         Left            =   1050
         TabIndex        =   6
         Top             =   255
         Width           =   960
      End
      Begin VB.OptionButton optBoth2 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Value           =   -1  'True
         Width           =   660
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Select A/cType"
      Height          =   615
      Index           =   0
      Left            =   4185
      TabIndex        =   3
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   870
      Width           =   3150
      Begin VB.OptionButton optBoth1 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton optCustomer 
         Caption         =   "Customer"
         Height          =   255
         Left            =   1050
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optParty 
         Caption         =   "Party"
         Height          =   255
         Left            =   2145
         TabIndex        =   2
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Selection"
      Height          =   570
      Index           =   2
      Left            =   4485
      TabIndex        =   12
      ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
      Top             =   2085
      Width           =   2550
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Select to choose all Customers (Alt + A)"
         Top             =   225
         Value           =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton optSelective 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1515
         TabIndex        =   11
         Top             =   225
         Width           =   960
      End
   End
   Begin VB.Frame fraReportType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4470
      TabIndex        =   16
      Top             =   2670
      Width           =   2580
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print.Date"
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "CustomerMasterReport.frx":0000
         Left            =   90
         List            =   "CustomerMasterReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   2415
      End
   End
   Begin VB.Frame fraselective 
      Height          =   1755
      Left            =   7710
      TabIndex        =   21
      Top             =   2355
      Visible         =   0   'False
      Width           =   4170
Begin AceSpread spList
         Height          =   1530
         Left            =   60
         TabIndex        =   20
         Top             =   165
         Width           =   4035
         _Version        =   458752
         _ExtentX        =   7117
         _ExtentY        =   2699
         _StockProps     =   64
         DAutoSizeCols   =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         SelectBlockOptions=   0
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   3885
      TabIndex        =   17
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   3705
      Width           =   1300
   End
End
Attribute VB_Name = "rptCustomerMasterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sprow As Long

Private Erpno As String
Private TsNo As String

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Public Sub Header()
  Try

  Call RptStandard_Header(4, sprow, Erpno, TsNo, chkPrintDate.Value)
  GridCount = GridCount + 6
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
  Dim Sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim sRow As Long
  Dim i As Long
  Dim PartyCode As String
  Dim AcType As String
  Dim Status As String
  Dim sRtHdAtInAt As String
  Dim Spr As AceSpread
  Dim setstr As String
  
  Try
  
  sSpread_sPrintcurrent = False
  
  If OptSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
  End If
  
  Sql = " select a.PartyName, a.PartyCode, a.VendorCode, a.AcType, a.FormCode,(select b.Description from TERMSMASTER b where  b.TypeFlg='Form' and b.Code=a.FormCode"
  Sql = Sql & " and b.Branchid=a.Branchid  and b.Compid=a.Compid) FormDesc, a.PackingCode,"
  Sql = Sql & " (select b1.Description  from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding'  and b1.Code=a.PackingCode"
  Sql = Sql & " and b1.Branchid=a.Branchid  and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,"
  Sql = Sql & " (select b2.Description  from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction'  and b2.Code=a.DeliveryCode"
  Sql = Sql & " and b2.Branchid=a.Branchid  and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,"
  Sql = Sql & " (select b3.Description  from TERMSMASTER b3 where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode"
  Sql = Sql & " and b3.Branchid=a.Branchid  and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,"
  Sql = Sql & " (select  b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode"
  Sql = Sql & " and b4.Branchid=a.Branchid  and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,"
  Sql = Sql & " (select  b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid"
  Sql = Sql & " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,"
  Sql = Sql & " (select b6.Description  from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode"
  Sql = Sql & " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc, a.IsSubcontractor, a.AliasName, a.EffStatus,"
  Sql = Sql & " a.EffDate, a.CreditDays, a.CustomerFormulaCode, (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7"
  Sql = Sql & " where b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.CustomerFormulaCode and b7.Branchid=a.Branchid  and b7.Compid=a.Compid) CustomerFormulaName, a.VendorFormulaCode,"
  Sql = Sql & " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where isActive = '1' and b7.ModuleCode = 9 and b7.FrmID=a.VendorFormulaCode"
  Sql = Sql & " and b7.Branchid=a.Branchid and b7.Compid=a.Compid) VendorFormulaName"
  Sql = Sql & " From FAS_PARTY_MASTER a"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  If optCustomer Then
 Sql = Sql & " and a.AcType='D'"
End If
  If OptParty Then
 Sql = Sql & " and a.AcType='Y'"
End If
  If optBoth1 Then
 Sql = Sql & " and a.AcType In('D','Y')"
End If
  If optActive Then
 Sql = Sql & " and a.EffStatus = 1"
End If
  If optInActive Then
 Sql = Sql & " and a.EffStatus = 0"
End If
  
  If OptSelective Then
    Sql = Sql & " and a.PartyCode in "
    Sql = Sql & Spread_Get_SelectionList(spList, 1)
'    With spList
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = 1
'        Sql = Sql & "'" & Trim$(.Text) & "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)
'    End With
'    Sql = Sql & ")"
  End If
  
  Sql = Sql & " Order by a.PartyCode, a.PartyName"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sSpread_sPrintcurrent = False
    GoTo ExitHere
  End If
  
  If OptSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
  End If
  
  Sql = " Select a.PartyCode, b.PartyName, b.AcType, a.Add1, a.Add2, a.Add3, a.City, a.PinCode, a.Phone, a.Fax,"
  Sql = Sql & " a.Cellular, a.Email, a.CST, a.CSTDate, a.lst, a.LSTDate, a.GSTRegNo, a.GSTRegDate,a.ProvisionalID,"
  Sql = Sql & " a.AreaCode, a.ECCNo, a.EDRegNo, a.EDRegDate, a.ContactPerson,"
  Sql = Sql & " IsNull(a.UdyamNo,'') As UdyamNo, a.UdyamDate"
  Sql = Sql & " From FAS_PARTY_ADDRESS a"
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On b.PartyCode = a.PartyCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  
  If optCustomer Then
 Sql = Sql & " and b.AcType='D'"
End If
  If OptParty Then
 Sql = Sql & " and b.AcType='Y'"
End If
  If optBoth1 Then
 Sql = Sql & " and b.AcType In('D','Y')"
End If
  If optActive Then
 Sql = Sql & " and b.EffStatus = 1"
End If
  If optInActive Then
 Sql = Sql & " and b.EffStatus = 0"
End If
  
  If OptSelective Then
    Sql = Sql & " and a.PartyCode in "
    Sql = Sql & Spread_Get_SelectionList(spList, 1)
  End If
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  If optBoth2 Then
    If OptSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "All Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "All Party Details"
End If
    End If
  Else
    If OptSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Party Details"
End If
    End If
  End If
  If optBoth1 Then
    If OptSelective And optBoth2 <> True Then
      If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - Active"
End If
      If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - InActive"
End If
    ElseIf optBoth2 <> True Then
      If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - Active"
End If
      If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - InActive"
End If
    End If
  ElseIf optBoth2 <> True Then
    If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & " - Active"
End If
    If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & " - InActive"
End If
  End If
  
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 1
  
  gStrReportHeading = "Customer Master List Report"
  
  Set Spr = frmPrint.sprPrint
  
  Call spSetColWidth(Spr, 1, 17, 2, 30, 3, 13, 4, 20)
  
  Call Header
  
  Do While Not rs.EOF
    If rs("AcType") = "P" Then
      AcType = "Vendor"
    ElseIf rs("AcType") = "D" Then
      AcType = "Customer"
    ElseIf rs("AcType") = "Y" Then
      AcType = "Party"
    End If
    If PartyCode <> rs("PartyCode") Then
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      setstr = "Customer Name : " & rs("PartyCode") & " - " & rs("PartyName")
      Call spSetText(Spr, 1, sprow, setstr, mAlignleftTop, True, 4, 1)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      setstr = "A/c Type : " & Space(2) & Trim$(AcType)
      Call spSetText(Spr, 1, sprow, setstr, mAlignleftTop, True, 4, 1)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
        
    If IsGSTRequired = True Then
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Address", rs("Add1"), "CST", CheckNull(rs("CST")) & IIf(CheckNull(rs("CSTDate")) = "", "", "  dt." & CheckNull(rs("CSTDate"))))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 2, sprow, 1, mAlignleftTop, False, rs("Add2"), "LST", CheckNull(rs("LST")) & IIf(CheckNull(rs("LSTDate")) = "", "", "   dt." & CheckNull(rs("LSTDate"))))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 2, sprow, 1, mAlignleftTop, False, CheckNull(rs("Add3")), "GST Reg No", CheckNull(rs("GSTRegNo")) & IIf(CheckNull(rs("GSTRegDate")) = "", "", "  dt." & CheckNull(rs("GSTRegDate"))))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "City", CheckNull(rs("City")), "Provisional ID", CheckNull(rs("ProvisionalID")))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    Else
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Address ", CheckNull(rs("Add1")), "CST ", CheckNull(rs("CST")))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, " ", CheckNull(rs("Add2")), "CST Date ", CheckNull(rs("CSTdate")))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, " ", CheckNull(rs("Add3")), "LST  ", CheckNull(rs("LST")))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      
      Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "City", CheckNull(rs("City")), "LST Date ", CheckNull(rs("LSTdate")))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    End If
    
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Pincode", CheckNull(rs("Pincode")), "Area Code  ", CheckNull(rs("AreaCode")))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Phone No", CheckNull(rs("Phone")), "ECC No", CheckNull(rs("ECCNo")))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Fax No", CheckNull(rs("Fax")), "ED Reg No", CheckNull(rs("EDRegNo")))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Cell No", CheckNull(rs("Cellular")), "ED Reg Date ", CheckNull(rs("EDRegDate")))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
     
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "E-MailID ", CheckNull(rs("Email")), "Contact Name ", CheckNull(CheckNull(rs("ContactPerson"))))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Udyam No. ", CheckNull(rs("UdyamNo")), "Udyam Date ", CheckNull(rs("UdyamDate")))
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
     
    rsDetails.Filter = adFilterNone
    If Not rsDetails.EOF Then
      rsDetails.Filter = " PartyCode= " & rs("PartyCode")
      Do While Not rsDetails.EOF
        If rsDetails("EffStatus") = "True" Then
          Status = "Active"
        Else
          Status = "InActive"
        End If
            
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Packing & Forwarding ", CheckNull(rsDetails("PackingDesc")), "Form", CheckNull(rsDetails("FormDesc")))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
         
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Insurance", CheckNull(rsDetails("InsuranceDesc")), "Freight ", CheckNull(rsDetails("FreightDesc")))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
            
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Delivery Instruct", CheckNull(rsDetails("DeliveryDesc")), "Mode", CheckNull(rsDetails("ModeDesc")))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
            
        
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Term Of Payment", CheckNull(rsDetails("PaymentDesc")), "Credit Days", CheckNull(rsDetails("CreditDays")))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Cust Formula Name", CheckNull(rsDetails("CustomerFormulaName")), "EffDate", CheckNull(rsDetails("EffDate")))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        
        Call spSetCellBorder(Spr, 1, sprow, 4, , mBorderAllSides)
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignleftTop, False, "Vendor Formula Name", CheckNull(rsDetails("VendorFormulaName")), "Status", CheckNull(Status))
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        rsDetails.MoveNext
      Loop
    End If
    PartyCode = rs("PartyCode")
    rs.MoveNext
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  Loop
  
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  totalpages = iPageNo
  frmPrint.Show
  
  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
  Set rs = Nothing
End Try
End Function

Public Sub sPageHeader()
  Try

'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrentold() As Boolean
  Dim Sql As String
  Dim rsDetails As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim sRow As Long
  Dim i As Long
  Dim PartyCode As String
  Dim AcType As String
  Dim Status As String
  Dim sRtHdAtInAt As String
  
  Try
  
  Sql = " select a.PartyName, a.PartyCode, a.VendorCode, a.AcType, a.FormCode,(select b.Description from TERMSMASTER b where  b.TypeFlg='Form' and b.Code=a.FormCode"
  Sql = Sql & " and b.Branchid=a.Branchid  and b.Compid=a.Compid) FormDesc, a.PackingCode,"
  Sql = Sql & " (select b1.Description  from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding'  and b1.Code=a.PackingCode"
  Sql = Sql & " and b1.Branchid=a.Branchid  and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,"
  Sql = Sql & " (select b2.Description  from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction'  and b2.Code=a.DeliveryCode"
  Sql = Sql & " and b2.Branchid=a.Branchid  and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,"
  Sql = Sql & " (select b3.Description  from TERMSMASTER b3 where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode"
  Sql = Sql & " and b3.Branchid=a.Branchid  and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,"
  Sql = Sql & " (select  b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode"
  Sql = Sql & " and b4.Branchid=a.Branchid  and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,"
  Sql = Sql & " (select  b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid"
  Sql = Sql & " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,"
  Sql = Sql & " (select b6.Description  from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode"
  Sql = Sql & " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc, a.IsSubcontractor, a.AliasName, a.EffStatus,"
  Sql = Sql & " a.EffDate, a.CreditDays, a.CustomerFormulaCode, (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7"
  Sql = Sql & " where b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.CustomerFormulaCode and b7.Branchid=a.Branchid  and b7.Compid=a.Compid) CustomerFormulaName, a.VendorFormulaCode,"
  Sql = Sql & " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where isActive = '1' and b7.ModuleCode = 9 and b7.FrmID=a.VendorFormulaCode"
  Sql = Sql & " and b7.Branchid=a.Branchid  and b7.Compid=a.Compid) VendorFormulaName"
  Sql = Sql & " From FAS_PARTY_MASTER a"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  If optCustomer Then
 Sql = Sql & " and a.AcType='D'"
End If
  If OptParty Then
 Sql = Sql & " and a.AcType='Y'"
End If
  If optBoth1 Then
 Sql = Sql & " and a.AcType In('D','Y')"
End If
  If optActive Then
 Sql = Sql & " and a.EffStatus = 1"
End If
  If optInActive Then
 Sql = Sql & " and a.EffStatus = 0"
End If
  
  If OptSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.PartyCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql & "'" & Trim$(spList.Text) & "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql & ")"
  End If
  
  Sql = Sql & " Order by a.PartyCode, a.PartyName"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sPrintcurrentold = False
    GoTo ExitHere
  End If
  Sql = " Select a.PartyCode, b.PartyName, b.AcType, a.Add1, a.Add2, a.Add3, a.City, a.PinCode, a.Phone, a.Fax,"
  Sql = Sql & " A.Cellular, A.Email, A.CST, A.CSTDate, A.lst, A.LSTDate, A.AreaCode, A.ECCNo, A.EDRegNo, A.EDRegDate, A.ContactPerson"
  Sql = Sql & " from FAS_PARTY_ADDRESS a"
  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
  Sql = Sql & " On b.PartyCode = a.PartyCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  If optCustomer Then
 Sql = Sql & " and b.AcType='D'"
End If
  If OptParty Then
 Sql = Sql & " and b.AcType='Y'"
End If
  If optBoth1 Then
 Sql = Sql & " and b.AcType In('D','Y')"
End If
  If optActive Then
 Sql = Sql & " and b.EffStatus = 1"
End If
  If optInActive Then
 Sql = Sql & " and b.EffStatus = 0"
End If
  
  If OptSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.PartyCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql & "'" & Trim$(spList.Text) & "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql & ")"
  End If
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  If optBoth2 Then
    If OptSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "All Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "All Party Details"
End If
    End If
  Else
    If OptSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Customer Details"
End If
      If OptParty Then
 sRtHdAtInAt = "Party Details"
End If
    End If
  End If
  If optBoth1 Then
    If OptSelective And optBoth2 <> True Then
      If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - Active"
End If
      If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - InActive"
End If
    ElseIf optBoth2 <> True Then
      If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - Active"
End If
      If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - InActive"
End If
    End If
  ElseIf optBoth2 <> True Then
    If optActive Then
 sRtHdAtInAt = sRtHdAtInAt & " - Active"
End If
    If optInActive Then
 sRtHdAtInAt = sRtHdAtInAt & " - InActive"
End If
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
  gStrReportHeading = "Customer Master Report"
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort CAlign(iPageWidth, sRtHdAtInAt)
  sSendToPort ""
  sPageHeader
  
  sRow = rsDetails.RecordCount
  PartyCode = ""
  i = 1
  Do While Not rs.EOF
    If rs("AcType") = "P" Then
      AcType = "Vendor"
    ElseIf rs("AcType") = "D" Then
      AcType = "Customer"
    ElseIf rs("AcType") = "Y" Then
      AcType = "Party"
    End If
        
    If PartyCode <> rs("PartyCode") Then
      sSendToPort LAlign(18, "Customer Name") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(36, rs("PartyCode") & Space(1) & LAlign(1, "-") & Space(1) & rs("PartyName")) & Space(1) & LAlign(8, "A/c Type") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(8, Trim$(AcType))
      sSendToPort ""
    End If
   
    sSendToPort LAlign(18, "Address ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add1")) & Space(1) & LAlign(11, "CST") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("CST")))
    sSendToPort Space(18) & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add2")) & Space(1) & LAlign(11, "CSTdate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("CSTDate")))
    sSendToPort Space(18) & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add3")) & Space(1) & LAlign(11, "LST") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("LST")))
    sSendToPort LAlign(18, "City    ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("City"))) & Space(1) & LAlign(11, "LSTDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("LSTDate")))
    sSendToPort LAlign(18, "Pincode ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Pincode"))) & Space(1) & LAlign(11, "AreaCode") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("AreaCode")))
    sSendToPort LAlign(18, "PhoneNo ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Phone"))) & Space(1) & LAlign(11, "ECCNo") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EccNo")))
    sSendToPort LAlign(18, "Fax No  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Fax"))) & Space(1) & LAlign(11, "EDRegNo") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EDRegNo")))
    sSendToPort LAlign(18, "Cell No ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Cellular"))) & Space(1) & LAlign(11, "EDRegDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EDRegDate")))
    sSendToPort LAlign(18, "E-MailID") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Email"))) & Space(1) & LAlign(11, "ContactName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("ContactPerson")))
    sSendToPort ""
    
    rsDetails.Filter = adFilterNone
    If Not rsDetails.EOF Then
      rsDetails.Filter = " PartyCode= " & rs("PartyCode")
      Do While Not rsDetails.EOF
        If rsDetails("EffStatus") = "True" Then
          Status = "Active"
        Else
          Status = "InActive"
        End If
        
        sSendToPort LAlign(18, "Packing&Forwarding") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("PackingDesc")) & Space(1) & LAlign(11, "Form") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("FormDesc")))
        sSendToPort LAlign(18, "Insurance          ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("InsuranceDesc")) & Space(1) & LAlign(11, "Freight") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("FreightDesc")))
        sSendToPort LAlign(18, "Delivery Instruct ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("DeliveryDesc")) & Space(1) & LAlign(11, "Mode  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ModeDesc")))
        sSendToPort LAlign(18, "Term Of Payment  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("PaymentDesc"))) & Space(1) & LAlign(11, "CreditDays") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("CreditDays")))
        sSendToPort LAlign(18, "CustFormulaName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("CustomerFormulaName"))) & Space(1) & LAlign(11, "EffDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("EffDate")))
        sSendToPort LAlign(18, "VendorFormulaName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("VendorFormulaName"))) & Space(1) & LAlign(11, "Status") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, Trim$(Status))
        sSendToPort ""
        rsDetails.MoveNext
      Loop
    End If
    PartyCode = rs("PartyCode")
    sSendToPort ""
    i = i + 1
    rs.MoveNext
  Loop
  
  sSendToPort ""
  rsDetails.Close
  rs.Close
  
  Call EndofReport(iPageWidth, sRow)
  
  sPrintcurrentold = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
  Set rs = Nothing
End Try
End Function

Private Function sPrintCurrent() As Boolean
'  Dim Sql As String
'  Dim rsDetails As ADODB.Recordset
'  Dim rs As ADODB.Recordset
'  Dim SRow As Long
'  Dim i As Long
'  Dim PartyCode As String
'  Dim AcType As String
'  Dim Status As String
'  Dim sRtHdAtInAt As String
'
'  On Error GoTo ErrHandler
'
'  Sql = " select a.PartyName, a.PartyCode, a.VendorCode, a.AcType, a.FormCode,(select b.Description from TERMSMASTER b where  b.TypeFlg='Form' and b.Code=a.FormCode"
'  Sql = Sql & " and b.Branchid=a.Branchid  and b.Compid=a.Compid) FormDesc, a.PackingCode,"
'  Sql = Sql & " (select b1.Description  from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding'  and b1.Code=a.PackingCode"
'  Sql = Sql & " and b1.Branchid=a.Branchid  and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,"
'  Sql = Sql & " (select b2.Description  from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction'  and b2.Code=a.DeliveryCode"
'  Sql = Sql & " and b2.Branchid=a.Branchid  and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,"
'  Sql = Sql & " (select b3.Description  from TERMSMASTER b3 where b3.TypeFlg='Payment Terms' and b3.Code=a.PaymentCode"
'  Sql = Sql & " and b3.Branchid=a.Branchid  and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,"
'  Sql = Sql & " (select  b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance' and b4.Code=a.InsuranceCode"
'  Sql = Sql & " and b4.Branchid=a.Branchid  and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,"
'  Sql = Sql & " (select  b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode' and b5.Code=a.ModeCode and b5.Branchid=a.Branchid"
'  Sql = Sql & " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,"
'  Sql = Sql & " (select b6.Description  from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode"
'  Sql = Sql & " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc, a.IsSubcontractor, a.AliasName, a.EffStatus,"
'  Sql = Sql & " a.EffDate, a.CreditDays, a.CustomerFormulaCode, (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7"
'  Sql = Sql & " where b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.CustomerFormulaCode and b7.Branchid=a.Branchid  and b7.Compid=a.Compid) CustomerFormulaName, a.VendorFormulaCode,"
'  Sql = Sql & " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where isActive = '1' and b7.ModuleCode = 9 and b7.FrmID=a.VendorFormulaCode"
'  Sql = Sql & " and b7.Branchid=a.Branchid  and b7.Compid=a.Compid) VendorFormulaName"
'  Sql = Sql & " From FAS_PARTY_MASTER a"
'  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
'  If optCustomer Then: Sql = Sql & " and a.AcType='D'"

'  If optParty Then: Sql = Sql & " and a.AcType='Y'"

'  If optBoth1 Then: Sql = Sql & " and a.AcType In('D','Y')"

'  If optActive Then: Sql = Sql & " and a.EffStatus = 1"

'  If optInActive Then: Sql = Sql & " and a.EffStatus = 0"
'
'  If optSelective Then
'    If spList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
'    Sql = Sql & " and a.PartyCode in("
'    With spList
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = 1
'        Sql = Sql & "'" & Trim$(.Text) & "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)
'    End With
'    Sql = Sql & ")"
'  End If
'
'  Sql = Sql & " Order by a.PartyCode, a.PartyName"
'  Set rsDetails = New ADODB.Recordset
'  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
'  If rsDetails.RecordCount = 0 Then
'    MsgBox "No Record(s) Found"
'    sPrintcurrent = False
'    GoTo ExitHere
'  End If
'  Sql = " Select a.PartyCode, b.PartyName, b.AcType, a.Add1, a.Add2, a.Add3, a.City, a.PinCode, a.Phone, a.Fax,"
'  Sql = Sql & " A.Cellular, A.Email, A.CST, A.CSTDate, A.lst, A.LSTDate, A.AreaCode, A.ECCNo, A.EDRegNo, A.EDRegDate, A.ContactPerson"
'  Sql = Sql & " from FAS_PARTY_ADDRESS a"
'  Sql = Sql & " Left Outer Join FAS_PARTY_MASTER b"
'  Sql = Sql & " On b.PartyCode = a.PartyCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
'  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
'  If optCustomer Then: Sql = Sql & " and b.AcType='D'"

'  If optParty Then: Sql = Sql & " and b.AcType='Y'"

'  If optBoth1 Then: Sql = Sql & " and b.AcType In('D','Y')"

'  If optActive Then: Sql = Sql & " and b.EffStatus = 1"

'  If optInActive Then: Sql = Sql & " and b.EffStatus = 0"
'
'  If optSelective Then
'    If spList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
'    Sql = Sql & " and a.PartyCode in("
'    With spList
'      SRow = 0
'      For i = 1 To .SelectionCount
'        SRow = .GetMultiSelItem(SRow)
'        .Row = SRow: .Col = 1
'        Sql = Sql & "'" & Trim$(.Text) & "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)
'    End With
'    Sql = Sql & ")"
'  End If
'
'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then GoTo ExitHere
'
'  If optBoth2 Then
'    If optSelective And optBoth1 <> True Then
'      If optCustomer Then: sRtHdAtInAt = "Selective Customer Details"

'      If optParty Then: sRtHdAtInAt = "Selective Party Details"
'    ElseIf optBoth1 <> True Then
'      If optCustomer Then: sRtHdAtInAt = "All Customer Details"

'      If optParty Then: sRtHdAtInAt = "All Party Details"
'    End If
'  Else
'    If optSelective And optBoth1 <> True Then
'      If optCustomer Then: sRtHdAtInAt = "Selective Customer Details"

'      If optParty Then: sRtHdAtInAt = "Selective Party Details"
'    ElseIf optBoth1 <> True Then
'      If optCustomer Then: sRtHdAtInAt = "Customer Details"

'      If optParty Then: sRtHdAtInAt = "Party Details"
'    End If
'  End If
'  If optBoth1 Then
'    If optSelective And optBoth2 <> True Then
'      If optActive Then: sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - Active"

'      If optInActive Then: sRtHdAtInAt = sRtHdAtInAt & "Selective Customer & Party Details - InActive"
'    ElseIf optBoth2 <> True Then
'      If optActive Then: sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - Active"

'      If optInActive Then: sRtHdAtInAt = sRtHdAtInAt & "All Customer & Party Details - InActive"
'    End If
'  ElseIf optBoth2 <> True Then
'    If optActive Then: sRtHdAtInAt = sRtHdAtInAt & " - Active"

'    If optInActive Then: sRtHdAtInAt = sRtHdAtInAt & " - InActive"
'  End If
'
'  If cmbReportType.ListIndex = 1 Then
'    pheight = 75
'    pPaper = vbPRPSA4
'    Orient = 1
'    iPageWidth = 80
'    iLineNo = 1
'    fOpenPort Me
'    sSendToPort ""
'    sSendToPort CAlign(iPageWidth, gCompName)
'    sSendToPort ""
'    gStrReportHeading = "Customer Master Report"
'    sSendToPort CAlign(iPageWidth, gStrReportHeading)
'    sSendToPort CAlign(iPageWidth, sRtHdAtInAt)
'    sSendToPort ""
'    sPageHeader
'
'    SRow = rsDetails.RecordCount
'    PartyCode = ""
'    i = 1
'    Do While Not rs.EOF
'      If rs("AcType") = "P" Then
'        AcType = "Vendor"
'      ElseIf rs("AcType") = "D" Then
'        AcType = "Customer"
'      ElseIf rs("AcType") = "Y" Then
'        AcType = "Party"
'      End If
'
'      If PartyCode <> rs("PartyCode") Then
'        sSendToPort LAlign(18, "Customer Name") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(36, rs("PartyCode") & Space(1) & LAlign(1, "-") & Space(1) & rs("PartyName")) & Space(1) & LAlign(8, "A/c Type") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(8, Trim$(AcType))
'        sSendToPort ""
'      End If
'
'      sSendToPort LAlign(18, "Address ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add1")) & Space(1) & LAlign(11, "CST") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("CST")))
'      sSendToPort Space(18) & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add2")) & Space(1) & LAlign(11, "CSTdate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("CSTDate")))
'      sSendToPort Space(18) & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rs("Add3")) & Space(1) & LAlign(11, "LST") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("LST")))
'      sSendToPort LAlign(18, "City    ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("City"))) & Space(1) & LAlign(11, "LSTDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("LSTDate")))
'      sSendToPort LAlign(18, "Pincode ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Pincode"))) & Space(1) & LAlign(11, "AreaCode") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("AreaCode")))
'      sSendToPort LAlign(18, "PhoneNo ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Phone"))) & Space(1) & LAlign(11, "ECCNo") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EccNo")))
'      sSendToPort LAlign(18, "Fax No  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Fax"))) & Space(1) & LAlign(11, "EDRegNo") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EDRegNo")))
'      sSendToPort LAlign(18, "Cell No ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Cellular"))) & Space(1) & LAlign(11, "EDRegDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("EDRegDate")))
'      sSendToPort LAlign(18, "E-MailID") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rs("Email"))) & Space(1) & LAlign(11, "ContactName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(30, CheckNull(rs("ContactPerson")))
'      sSendToPort ""
'
'      rsDetails.Filter = adFilterNone
'      If Not rsDetails.EOF Then
'        rsDetails.Filter = " PartyCode= " & rs("PartyCode")
'        Do While Not rsDetails.EOF
'          If rsDetails("EffStatus") = "True" Then
'            Status = "Active"
'          Else
'            Status = "InActive"
'          End If
'
'          sSendToPort LAlign(18, "Packing&Forwarding") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("PackingDesc")) & Space(1) & LAlign(11, "Form") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("FormDesc")))
'          sSendToPort LAlign(18, "Insurance          ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("InsuranceDesc")) & Space(1) & LAlign(11, "Freight") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("FreightDesc")))
'          sSendToPort LAlign(18, "Delivery Instruct ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("DeliveryDesc")) & Space(1) & LAlign(11, "Mode  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ModeDesc")))
'          sSendToPort LAlign(18, "Term Of Payment  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("PaymentDesc"))) & Space(1) & LAlign(11, "CreditDays") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("CreditDays")))
'          sSendToPort LAlign(18, "CustFormulaName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("CustomerFormulaName"))) & Space(1) & LAlign(11, "EffDate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("EffDate")))
'          sSendToPort LAlign(18, "VendorFormulaName") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("VendorFormulaName"))) & Space(1) & LAlign(11, "Status") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, Trim$(Status))
'          sSendToPort ""
'          rsDetails.MoveNext
'        Loop
'      End If
'      PartyCode = rs("PartyCode")
'      sSendToPort ""
'      i = i + 1
'      rs.MoveNext
'    Loop
'
'    sSendToPort ""
'    rsDetails.Close
'    rs.Close
'
'    Call EndofReport(iPageWidth, SRow)
'  Else
''    With spview
''      i = 4
''      .ClearRange 1, i, .MaxCols, .MaxRows - 1, True
''      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexLeft, 0, CellBorderStyleBlank
''      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexRight, 0, CellBorderStyleBlank
''      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexTop, 0, CellBorderStyleBlank
''      .SetCellBorder 1, i, .MaxCols, .MaxRows - 1, CellBorderIndexBottom, 0, CellBorderStyleBlank
''      For i = 4 To .MaxRows
''        .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = False: .FontSize = 8
''        .RemoveCellSpan 1, i
''      Next i
''      i = 4
''      .SetText 1, 1, gCompName
''      .SetText 1, 2, "Terms Master Report"
''      If chkPrintDate.Value = 1 Then: .SetText 2, 3, "Printed. Date : " & Format(ServerDate, "dd/MM/yyyy")
'      gStrReportHeading = "Terms Master Report"
'      Do While Not rs.EOF
'        If rs("AcType") = "P" Then
'          AcType = "Vendor"
'        ElseIf rs("AcType") = "D" Then
'          AcType = "Customer"
'        ElseIf rs("AcType") = "Y" Then
'          AcType = "Party"
'        End If
'        If PartyCode <> rs("PartyCode") Then
'          .AddCellSpan 1, i, 3, 1
'          .SetCellBorder 1, i, .MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
'          .Col = 1: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True: .FontSize = 9

'          .SetText 1, i, "Customer Name : " & Space(2) & rs("PartyCode") & " - " & rs("PartyName")

'          .Col = 4: .Row = i: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .Font = "Arial": .FontBold = True: .FontSize = 9

'          .SetText 4, i, "A/c Type : " & Space(2) & Trim$(AcType)
'          i = i + 1
'        End If
'        .SetCellBorder 1, i, .MaxCols, i + 9, CellBorderIndexLeft, 0, CellBorderStyleSolid
'        .SetCellBorder 1, i, .MaxCols, i + 9, CellBorderIndexRight, 0, CellBorderStyleSolid
'        .SetCellBorder 1, i, .MaxCols, i + 9, CellBorderIndexTop, 0, CellBorderStyleSolid
'        .SetCellBorder 1, i, .MaxCols, i + 9, CellBorderIndexBottom, 0, CellBorderStyleSolid
'
'        .SetText 1, i, "Address : "

'        .SetText 1, i + 3, "City : "

'        .SetText 1, i + 4, "Pincode : "

'        .SetText 1, i + 5, "PhoneNo : "

'        .SetText 1, i + 6, "Fax No : "

'        .SetText 1, i + 7, "Cell No : "

'        .SetText 1, i + 8, "E-MailID : "
'
'        .SetText 2, i, Space(2) & CheckNull(rs("Add1"))
'        .SetText 2, i + 1, Space(2) & CheckNull(rs("Add2"))
'        .SetText 2, i + 2, Space(2) & CheckNull(rs("Add3"))
'        .SetText 2, i + 3, Space(2) & CheckNull(rs("City"))
'        .SetText 2, i + 4, Space(2) & CheckNull(rs("Pincode"))
'        .SetText 2, i + 5, Space(2) & CheckNull(rs("Phone"))
'        .SetText 2, i + 6, Space(2) & CheckNull(rs("Fax"))
'        .SetText 2, i + 7, Space(2) & CheckNull(rs("Cellular"))
'        .SetText 2, i + 8, Space(2) & CheckNull(rs("Email"))
'
'        .SetText 3, i, "CST : "

'        .SetText 3, i + 1, "CSTdate : "

'        .SetText 3, i + 2, "LST : "

'        .SetText 3, i + 3, "LSTdate : "

'        .SetText 3, i + 4, "AreaCode : "

'        .SetText 3, i + 5, "ECCNo : "

'        .SetText 3, i + 6, "EDRegNo : "

'        .SetText 3, i + 7, "EDRegDate : "

'        .SetText 3, i + 8, "ContactName : "
'
'        .SetText 4, i, Space(2) & CheckNull(rs("CST"))
'        .SetText 4, i + 1, Space(2) & CheckNull(rs("CSTdate"))
'        .SetText 4, i + 2, Space(2) & CheckNull(rs("LST"))
'        .SetText 4, i + 3, Space(2) & CheckNull(rs("LSTdate"))
'        .SetText 4, i + 4, Space(2) & CheckNull(rs("AreaCode"))
'        .SetText 4, i + 5, Space(2) & CheckNull(rs("ECCNo"))
'        .SetText 4, i + 6, Space(2) & CheckNull(rs("EDRegNo"))
'        .SetText 4, i + 7, Space(2) & CheckNull(rs("EDRegDate"))
'        .SetText 4, i + 8, Space(2) & CheckNull(CheckNull(rs("ContactPerson")))
'        i = i + 9
'        rsDetails.Filter = adFilterNone
'        If Not rsDetails.EOF Then
'          rsDetails.Filter = " PartyCode= " & rs("PartyCode")
'          Do While Not rsDetails.EOF
'            If rsDetails("EffStatus") = "True" Then
'              Status = "Active"
'            Else
'              Status = "InActive"
'            End If
'            .AddCellSpan 1, i, .MaxCols, 1
'            i = i + 1
'            .SetCellBorder 1, i, .MaxCols, i + 5, CellBorderIndexLeft, 0, CellBorderStyleSolid
'            .SetCellBorder 1, i, .MaxCols, i + 5, CellBorderIndexRight, 0, CellBorderStyleSolid
'            .SetCellBorder 1, i, .MaxCols, i + 5, CellBorderIndexTop, 0, CellBorderStyleSolid
'            .SetCellBorder 1, i, .MaxCols, i + 5, CellBorderIndexBottom, 0, CellBorderStyleSolid
'
'            .SetText 1, i, "Packing&Forwarding : "

'            .SetText 1, i + 1, "Insurance : "

'            .SetText 1, i + 2, "Delivery Instruct : "

'            .SetText 1, i + 3, "Term Of Payment : "

'            .SetText 1, i + 4, "CustFormulaName : "

'            .SetText 1, i + 5, "VendorFormulaName : "
'
'            .SetText 2, i, Space(2) & CheckNull(rsDetails("PackingDesc"))
'            .SetText 2, i + 1, Space(2) & CheckNull(rsDetails("InsuranceDesc"))
'            .SetText 2, i + 2, Space(2) & CheckNull(rsDetails("DeliveryDesc"))
'            .SetText 2, i + 3, Space(2) & CheckNull(rsDetails("PaymentDesc"))
'            .SetText 2, i + 4, Space(2) & CheckNull(rsDetails("CustomerFormulaName"))
'            .SetText 2, i + 5, Space(2) & CheckNull(rsDetails("VendorFormulaName"))
'
'            .SetText 3, i, "Form : "

'            .SetText 3, i + 1, "Freight : "

'            .SetText 3, i + 2, "Mode : "

'            .SetText 3, i + 3, "CreditDays : "

'            .SetText 3, i + 4, "EffDate : "

'            .SetText 3, i + 5, "Status : "
'
'            .SetText 4, i, Space(2) & CheckNull(rsDetails("FormDesc"))
'            .SetText 4, i + 1, Space(2) & CheckNull(rsDetails("FreightDesc"))
'            .SetText 4, i + 2, Space(2) & CheckNull(rsDetails("ModeDesc"))
'            .SetText 4, i + 3, Space(2) & CheckNull(rsDetails("CreditDays"))
'            .SetText 4, i + 4, Space(2) & CheckNull(rsDetails("EffDate"))
'            .SetText 4, i + 5, Space(2) & CheckNull(Status)
'            rsDetails.MoveNext
'            .SetCellBorder 1, i, .MaxCols, i + 6, CellBorderIndexTop, 0, CellBorderStyleSolid
'            .AddCellSpan 1, i + 6, .MaxCols, 1
'            i = i + 7
'          Loop
'        End If
'        PartyCode = rs("PartyCode")
'        rs.MoveNext
'      Loop
'    End With
'
'    frmsp.Visible = True
'    frmsp.Height = 7470
'    frmsp.Width = 11310
'    spview.Height = 6660
'    spview.Width = 10290
'  End If
'  sPrintcurrent = True
'
'ExitHere:
'  On Error GoTo 0
'  Set rsDetails = Nothing
'  Set rs = Nothing
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call Form_Clear(Me)
  chkPrintDate.Value = 1
  
  Call StepClear(Me)
  Call StepEnable(Me)
  
  OptAll.Value = True
  cmbReportType.ListIndex = 0
  fraselective.Visible = False
  CmdStep(2).Enabled = False
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  If cmbReportType <> "Spread Format Report" Then
    If sPrintcurrentold = False Then
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

Private Sub cmdprint_Click()
    
  Try
  Me.MousePointer = vbHourglass

'  Call Spread_PagePrint(spview, 4, 54, spview.MaxCols, 3, cdlPortrait)
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
    cmdview.Enabled = True
  End If
  fraselective.Enabled = False
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try
  
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

Private Sub optAll_Click()

  Try
  Me.MousePointer = vbHourglass

  Call Spread_Clear(spList)
  spList.MaxRows = 1
  fraselective.Visible = False
  CmdStep(2).Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optCustomer_Click()

  Try
  Me.MousePointer = vbHourglass

  If optActive Or optInActive Then
    optBoth2.Value = True
  End If
  
  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optParty_Click()

  Try
  Me.MousePointer = vbHourglass

  If optActive Or optInActive Then
    optBoth2.Value = True
  End If

  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optBoth1_Click()

  Try
  Me.MousePointer = vbHourglass
  
  If optActive Or optInActive Then
    optBoth2.Value = True
  End If
  
  If OptSelective Then
    OptAll.Value = True
  End If

  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optActive_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If OptAll Then
    '
  ElseIf OptSelective Then
    OptAll.Value = True
  End If

  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optInActive_Click()

  Try
  Me.MousePointer = vbHourglass

  If OptAll Then
    '
  ElseIf OptSelective Then
    OptAll.Value = True
  End If

  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optBoth2_Click()
  Try
  Me.MousePointer = vbHourglass
  
  If OptSelective Then
    OptAll.Value = True
  End If

  If OptAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf OptSelective Then
    OptAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OptSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Select PartyCode [Customer Code],PartyName [Customer Name] from FAS_PARTY_MASTER where "
  Sql = Sql & " Branchid = '" & gBranchID & "' and compid='" & gCompID & "'"
  If optCustomer Then
 Sql = Sql & " and AcType='D'"
End If
  If OptParty Then
 Sql = Sql & " and AcType='Y'"
End If
  If optBoth1 Then
 Sql = Sql & " and AcType In('D','Y')"
End If
  If optActive Then
 Sql = Sql & " and EffStatus = 1"
End If
  If optInActive Then
 Sql = Sql & " and EffStatus = 0"
End If
'  If optBoth2 Then: Sql = Sql & " and EffStatus In(0,1)"
  
'  CmdStep(2).Enabled = False
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.MaxRows = spList.DataRowCnt
'  fraselective.Enabled = True
'  fraselective.Visible = True
'  spList.SetFocus

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    OptAll.Value = True
    GoTo ExitHere
  Else
    CmdStep(2).Enabled = False
    fraselective.Enabled = True
    fraselective.Visible = True
    spList.Enabled = True
 spList.Visible = True
    spList.ClearSelection
    Call ADORS_Spread_DataFill(rs, spList, True)
    
    spList.MaxRows = spList.DataRowCnt
    Call Spread_Set_Focus(spList, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set rs = Nothing
End Try
End Sub

Private Sub spList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList


spList.Click +=  new EventHandler(spList_Click)
spList.KeyUp +=  new EventHandler(spList_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerMasterReport.frm", spList, "")

End Sub
