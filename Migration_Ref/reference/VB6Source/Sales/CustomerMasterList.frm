VERSION 5.00
Begin VB.Form rptCustomerMasterList 
   Caption         =   "Customer Master"
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
   Begin VB.Frame frmsp 
      Height          =   645
      Left            =   300
      TabIndex        =   17
      Top             =   255
      Visible         =   0   'False
      Width           =   1530
      Begin VB.CommandButton cmdexit1 
         Caption         =   "E&xit"
         Height          =   420
         Left            =   9585
         TabIndex        =   18
         Top             =   6945
         Width           =   1215
      End
      Begin VB.CommandButton CMDPRINT 
         Caption         =   "P&rint"
         Height          =   420
         Left            =   8385
         TabIndex        =   19
         Top             =   6945
         Width           =   1215
      End
Begin AceSpread spview
         Height          =   375
         Left            =   495
         TabIndex        =   20
         Top             =   180
         Width           =   825
         _Version        =   458752
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         MaxCols         =   4
      End
   End
   Begin VB.Frame frmmain 
      Height          =   3705
      Left            =   4155
      TabIndex        =   2
      Top             =   1230
      Width           =   3345
      Begin VB.Frame frastep 
         Caption         =   "Report Type"
         Height          =   930
         Index           =   4
         Left            =   135
         TabIndex        =   21
         Top             =   2055
         Width           =   3060
         Begin VB.CheckBox chkPrintdate 
            Caption         =   "Print.Date"
            Height          =   255
            Left            =   90
            TabIndex        =   23
            Top             =   615
            Width           =   1080
         End
         Begin VB.ComboBox cmbReportType 
            Height          =   330
            ItemData        =   "CustomerMasterList.frx":1F8BA
            Left            =   90
            List            =   "CustomerMasterList.frx":1F8C4
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   210
            Width           =   2880
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   1650
         TabIndex        =   12
         ToolTipText     =   "Click to Exit (Alt + x)"
         Top             =   3135
         Width           =   1300
      End
      Begin VB.Frame FraAccount 
         Caption         =   "Select"
         ForeColor       =   &H00404040&
         Height          =   585
         Left            =   135
         TabIndex        =   14
         Top             =   1440
         Width           =   1890
         Begin VB.OptionButton optAll 
            Caption         =   "&All"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   525
         End
         Begin VB.OptionButton optSelective 
            Caption         =   "Single"
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   870
            TabIndex        =   15
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   400
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Click to View Report(Alt + R)"
         Top             =   3135
         Width           =   1300
      End
      Begin VB.Frame fraSelect2 
         Caption         =   "Select Status"
         Height          =   615
         Left            =   135
         TabIndex        =   8
         ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
         Top             =   795
         Width           =   3060
         Begin VB.OptionButton optInActive 
            Caption         =   "I&n Active"
            Height          =   255
            Left            =   2010
            TabIndex        =   11
            Top             =   285
            Width           =   945
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Activ&e"
            Height          =   255
            Left            =   870
            TabIndex        =   10
            Top             =   285
            Width           =   840
         End
         Begin VB.OptionButton optBoth2 
            Caption         =   "&All"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   255
            Value           =   -1  'True
            Width           =   660
         End
      End
      Begin VB.Frame fraSelect1 
         Caption         =   "Select A/cType"
         Height          =   615
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   "Select to choose Selective Customers (Alt + S)"
         Top             =   150
         Width           =   3060
         Begin VB.OptionButton optBoth1 
            Caption         =   "&All"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   255
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton optCustomer 
            Caption         =   "&Customer"
            Height          =   255
            Left            =   870
            TabIndex        =   6
            Top             =   270
            Width           =   990
         End
         Begin VB.OptionButton optParty 
            Caption         =   "&Party"
            Height          =   255
            Left            =   2010
            TabIndex        =   5
            Top             =   270
            Width           =   705
         End
      End
      Begin VB.CheckBox chkAddress 
         Caption         =   "Address"
         Height          =   210
         Left            =   2265
         TabIndex        =   3
         Top             =   1830
         Width           =   945
      End
   End
   Begin VB.Frame fraselective 
      Caption         =   "Customer Master"
      Height          =   1845
      Left            =   3150
      TabIndex        =   0
      Top             =   4965
      Visible         =   0   'False
      Width           =   5550
Begin AceSpread spList
         Height          =   1530
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   5325
         _Version        =   458752
         _ExtentX        =   9393
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
         MaxRows         =   1
         OperationMode   =   3
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SelectBlockOptions=   0
      End
   End
End
Attribute VB_Name = "rptCustomerMasterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public MListName As String

Private Function JobWorkItemMasterListDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim CustomerCode As String
  Dim retVal As Variant
  Dim i As Long
  Dim j As Integer

  Try
  JobWorkItemMasterListDisplay = False
  
  Sql = " Select a.ItemGroupKey, a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, a.ReceivedItemCode,"
  Sql = Sql & " c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM, a.RelationType, a.PartNo, c.PurchaseUOMDecimal as NoOfDec,"
  Sql = Sql & " a.RatePer, a.rate, a.ValidFrom, a.ValidTo, Case a.QcCheck When '0' Then 'No' Else 'Yes' End as QcCheck,"
  Sql = Sql & " Case a.isActive When '0' Then 'No' Else 'Yes' End as Status, a.FormulaCode,"
  Sql = Sql & " (select top 1 d.FrmLName from COM_FormulaMast_Hdr d where d.ModuleCode = " & gModuleCode & " and d.FrmID = a.FormulaCode  and d.Branchid = a.Branchid"
  Sql = Sql & " and d.Compid = a.Compid ) FrmLName, a.CustomerCode, (Select PartyName From Fas_Party_Master Where PartyCode = a.CustomerCode "
  Sql = Sql & " and BranchID = a.BranchID and CompID = a.CompID) PartyName"
  Sql = Sql & " From SAL_Jobwork_ItemMaster a"
  Sql = Sql & " Left Outer Join Material_Master b"
  Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
  Sql = Sql & " Left Outer Join Material_Master c"
  Sql = Sql & " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
  Sql = Sql & " where a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and b.Status = 1 "
  If optSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    spList.GetText 1, spList.ActiveRow, retVal
    Sql = Sql & " and a.CustomerCode in('" & retVal & "')  "
  End If
  Sql = Sql & " Order by a.CustomerCode, PartyName"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    JobWorkItemMasterListDisplay = False
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
    gStrReportHeading = "JobWork Item Master Report"
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
   
    i = 1
    CustomerCode = ""
    Do While Not rsDetails.EOF
      If CustomerCode <> rsDetails("CustomerCode") Then
        sSendToPort ""
        sSendToPort "Customer : " & "[" & rsDetails("CustomerCode") & "] " & rsDetails("PartyName")
        sSendToPort ""
      End If
  
      sSendToPort LAlign(12, "SentItemCode") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("SentItemCode")) & Space(1) & LAlign(12, "RecdItemCode") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ReceivedItemCode")))
      sSendToPort LAlign(12, "Description ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("SentDesc"))) & Space(1) & LAlign(12, "Description ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("RecdDesc")))
      sSendToPort LAlign(12, "Sent UOM    ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("SentUOM"))) & Space(1) & LAlign(12, "Recd UOM    ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("RecdUOM")))
      sSendToPort LAlign(12, "Rate Per    ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("RatePer")) & Space(1) & LAlign(12, "ValidFrom") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ValidFrom")))
      sSendToPort LAlign(12, "Rate        ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("Rate"))) & Space(1) & LAlign(12, "ValidTo  ") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ValidTo")))
      sSendToPort LAlign(12, "NoOf Decimal") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("NoOfDec"))) & Space(1) & LAlign(12, "QcCheck") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("QcCheck"))
      sSendToPort LAlign(12, "Formula Name") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("FrmLName"))) & Space(1) & LAlign(12, "Status") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("Status"))
      If i <> rsDetails.RecordCount Then sSendToPort ""
      i = i + 1
      CustomerCode = rsDetails("CustomerCode")
      rsDetails.MoveNext
    Loop
    
    rsDetails.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
    frmsp.Visible = True
    frmsp.Height = 7470
    frmsp.Width = 11310
    spview.Height = 6660
    spview.Width = 10290

      i = 4
      spview.ClearRange 1, i, spview.MaxCols, spview.MaxRows - 1, True
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexLeft, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexRight, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexTop, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexBottom, 0, CellBorderStyleBlank
      For i = 4 To spview.MaxRows
        For j = 1 To spview.MaxCols
          spview.Col = j
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = False
spview.FontSize = 8
        Next j
        spview.RemoveCellSpan 1, i
      Next i
      i = 4
      spview.SetText 1, 1, gCompName
      spview.SetText 1, 2, "Customer Item Master Report"
      If chkPrintDate.Value = 1 Then
 spview.SetText 2, 3, "Printed. Date : " & Format(ServerDate, "dd/MM/yyyy")
End If
      CustomerCode = ""
      Do While Not rsDetails.EOF
        If i > 494 Then
          spview.MaxRows = i + 8
        End If
        If CustomerCode <> rsDetails("CustomerCode") Then
          spview.SetCellBorder 1, i, spview.MaxCols, 1, CellBorderIndexTop, 0, CellBorderStyleSolid
          spview.AddCellSpan 1, i, spview.MaxCols, 1
          spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = True

          spview.SetText 1, i, "Customer : " & "[" & rsDetails("CustomerCode") & "] " & rsDetails("PartyName")
          i = i + 1
        End If
        
        spview.SetCellBorder 1, i, spview.MaxCols, i + 7, CellBorderIndexLeft, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 7, CellBorderIndexRight, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 7, CellBorderIndexTop, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 7, CellBorderIndexBottom, 0, CellBorderStyleSolid

        spview.SetText 1, i, "SentItemCode : "

        spview.SetText 1, i + 1, "ItemDescription : "

        spview.SetText 1, i + 2, "Sent UOM : "

        spview.SetText 1, i + 3, "RatePer : "

        spview.SetText 1, i + 4, "Rate : "

        spview.SetText 1, i + 5, "NoOfDecimal : "

        spview.SetText 1, i + 6, "FormulaName : "
        
        spview.SetText 2, i, rsDetails("SentItemCode")
        spview.SetText 2, i + 1, CheckNull(rsDetails("SentDesc"))
        spview.SetText 2, i + 2, rsDetails("SentUOM")
        spview.SetText 2, i + 3, CheckNull(rsDetails("Rateper"))
        spview.SetText 2, i + 4, CheckNull(rsDetails("Rate"))
        spview.SetText 2, i + 5, CheckNull(rsDetails("NoOfDec"))
        spview.SetText 2, i + 6, CheckNull(rsDetails("FrmLName"))
        
        spview.SetText 3, i, "RecdItemCode : "

        spview.SetText 3, i + 1, "ItemDescription : "

        spview.SetText 3, i + 2, "Recd UOm : "

        spview.SetText 3, i + 3, "ValidFrom : "

        spview.SetText 3, i + 4, "Valid To : "

        spview.SetText 3, i + 5, "QCcheck : "

        spview.SetText 3, i + 6, "Status : "
        
        spview.SetText 4, i, CheckNull(rsDetails("ReceivedItemCode"))
        spview.SetText 4, i + 1, CheckNull(rsDetails("RecdDesc"))
        spview.SetText 4, i + 2, CheckNull(rsDetails("RecdUOM"))
        spview.SetText 4, i + 3, rsDetails("ValidFrom")
        spview.SetText 4, i + 4, CheckNull(rsDetails("ValidTo"))
        spview.SetText 4, i + 5, rsDetails("QcCheck")
        spview.SetText 4, i + 6, CheckNull(rsDetails("Status"))
 
        CustomerCode = rsDetails("CustomerCode")
        i = i + 7
        rsDetails.MoveNext
        spview.AddCellSpan 1, i, spview.MaxCols, 1
        i = i + 1
      Loop
      rsDetails.Close
    frmsp.Visible = True
    frmsp.Height = 7470
    frmsp.Width = 11310
    spview.Height = 6660
    spview.Width = 10290
  End If
  
  JobWorkItemMasterListDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function CustomerItemMasterListDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  
  Dim CustomerCode As String
  Dim retVal As Variant
  Dim i As Long

  Try
  CustomerItemMasterListDisplay = False
  
  Sql = " select a.CustomerCode,(select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode"
  Sql = Sql & " and b.Branchid = '" & gBranchID & "' and Compid='" & gCompID & "') PartyName, a.ItemCode,a.PartNo,(select c.ItemDescription"
  Sql = Sql & " from INV_MATERIAL_MASTER c where c.ItemCode=a.ItemCode and c.Branchid = '" & gBranchID & "' and c.Compid='" & gCompID & "' and c.status = 1) ItemDescription,"
  Sql = Sql & " a.Rate,a.AssessableRate, ValidFrom, ValidTo, FormulaCode, Case QcCheck When '0' Then 'No' Else 'Yes' end as QCCheck,"
  Sql = Sql & " Case Status When '1' Then 'Active' Else 'InActive' End as Status,"
  Sql = Sql & " (select Top 1 FrmLName from COM_FormulaMast_Hdr d where d.ModuleCode = " & gModuleCode & " and d.FrmID=a.FormulaCode"
  Sql = Sql & " and d.Branchid = '" & gBranchID & "' and d.Compid ='" & gCompID & "') FrmLName"
  Sql = Sql & " from SAL_Customer_ItemMaster a where a.Branchid = '" & gBranchID & "' and a.Compid ='" & gCompID & "'"
  If optSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    spList.GetText 1, spList.ActiveRow, retVal
    Sql = Sql & " and a.CustomerCode in('" & retVal & "')  "
  End If
  Sql = Sql & " Order by CustomerCode,PartNo"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    CustomerItemMasterListDisplay = False
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
    gStrReportHeading = "Customer Item Master Report"
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
   
    i = 1
    CustomerCode = ""
    Do While Not rsDetails.EOF
      If CustomerCode <> rsDetails("CustomerCode") Then
        If i <> 1 Then sSendToPort ""
        sSendToPort "Customer : " & "[" & rsDetails("CustomerCode") & "] " & rsDetails("PartyName")
        sSendToPort ""
      End If
       
      sSendToPort LAlign(13, "Item Code") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(35, rsDetails("ItemCode")) & Space(1) & LAlign(11, "Rate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("Rate")))
      sSendToPort LAlign(13, "Item Descript") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(35, CheckNull(rsDetails("ItemDescription"))) & Space(1) & LAlign(11, "Assessable Rate") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("AssessableRate")))
      sSendToPort LAlign(13, "PartNo") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(35, rsDetails("PartNo")) & Space(1) & LAlign(11, "Valid From") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, CheckNull(rsDetails("ValidFrom")))
      sSendToPort LAlign(13, "Formula Name") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(35, CheckNull(rsDetails("FrmLName"))) & Space(1) & LAlign(11, "Valid To") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("ValidTo"))
      sSendToPort LAlign(13, "QCCheck") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(35, CheckNull(rsDetails("QCCheck"))) & Space(1) & LAlign(11, "Status") & Space(1) & LAlign(1, ":") & Space(1) & LAlign(28, rsDetails("Status"))
     
      If i <> rsDetails.RecordCount Then sSendToPort ""
      CustomerCode = rsDetails("CustomerCode")
      i = i + 1
      rsDetails.MoveNext
    Loop
  
    rsDetails.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
      i = 4
      spview.ClearRange 1, i, spview.MaxCols, spview.MaxRows - 1, True
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexLeft, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexRight, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexTop, 0, CellBorderStyleBlank
      spview.SetCellBorder 1, i, spview.MaxCols, spview.MaxRows - 1, CellBorderIndexBottom, 0, CellBorderStyleBlank
      For i = 4 To spview.MaxRows
        spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = False
spview.FontSize = 8
        spview.RemoveCellSpan 1, i
      Next i
      i = 4
      spview.SetText 1, 1, gCompName
      spview.SetText 1, 2, "Customer Item Master Report"
      If chkPrintDate.Value = 1 Then
 spview.SetText 2, 3, "Printed. Date : " & Format(ServerDate, "dd/MM/yyyy")
End If
      CustomerCode = ""
      Do While Not rsDetails.EOF
        If i > 494 Then
          spview.MaxRows = i + 6
        End If
        If CustomerCode <> rsDetails("CustomerCode") Then
          spview.SetCellBorder 1, i, spview.MaxCols, 1, CellBorderIndexTop, 0, CellBorderStyleSolid
          spview.AddCellSpan 1, i, spview.MaxCols, 1
          spview.Col = 1
spview.Row = i
spview.TypeVAlign = TypeVAlignCenter
spview.TypeHAlign = TypeHAlignLeft
spview.Font = "Arial"
spview.FontBold = True

          spview.SetText 1, i, "Customer : " & "[" & rsDetails("CustomerCode") & "] " & rsDetails("PartyName")
          i = i + 1
        End If
        
        spview.SetCellBorder 1, i, spview.MaxCols, i + 5, CellBorderIndexLeft, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 5, CellBorderIndexRight, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 5, CellBorderIndexTop, 0, CellBorderStyleSolid
        spview.SetCellBorder 1, i, spview.MaxCols, i + 5, CellBorderIndexBottom, 0, CellBorderStyleSolid

        spview.SetText 1, i, "Item Code : "

        spview.SetText 1, i + 1, "Item Description : "

        spview.SetText 1, i + 2, "PartNo : "

        spview.SetText 1, i + 3, "Formula Name : "

        spview.SetText 1, i + 4, "QCCheck : "
        
        spview.SetText 2, i, rsDetails("ItemCode")
        spview.SetText 2, i + 1, CheckNull(rsDetails("ItemDescription"))
        spview.SetText 2, i + 2, rsDetails("PartNo")
        spview.SetText 2, i + 3, CheckNull(rsDetails("FrmLName"))
        spview.SetText 2, i + 4, CheckNull(rsDetails("QCCheck"))
        
        spview.SetText 3, i, "Rate : "

        spview.SetText 3, i + 1, "Assessable Rate : "

        spview.SetText 3, i + 2, "Valid From : "

        spview.SetText 3, i + 3, "Valid To : "

        spview.SetText 3, i + 4, "Status : "
        
        spview.SetText 4, i, CheckNull(rsDetails("Rate"))
        spview.SetText 4, i + 1, CheckNull(rsDetails("AssessableRate"))
        spview.SetText 4, i + 2, CheckNull(rsDetails("ValidFrom"))
        spview.SetText 4, i + 3, rsDetails("ValidTo")
        spview.SetText 4, i + 4, CheckNull(rsDetails("Status"))
 
        CustomerCode = rsDetails("CustomerCode")
        i = i + 5
        rsDetails.MoveNext
        spview.AddCellSpan 1, i, spview.MaxCols, 1
        i = i + 1
      Loop
      rsDetails.Close
    frmsp.Visible = True
    frmsp.Height = 7470
    frmsp.Width = 11310
    spview.Height = 6660
    spview.Width = 10290

  End If
  
  CustomerItemMasterListDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function CustomerMasterListDisplay() As Boolean
  Dim Sql As String
'  Dim rsDetails As ADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim SRow As Long
  Dim i As Long
  Dim PartyCode As String
  Dim AcType As String
  Dim retVal As Variant
  Dim sRtHdAtInAt As String
  
  Try
  CustomerMasterListDisplay = False
   
  Sql = " select a.PartyName, a.PartyCode, a.VendorCode, a.AcType,"
  If chkAddress Then
    Sql = Sql & " a.EffStatus, b.Add1, b.Add2, b.Add3, b.City, b.PinCode"
  Else
    Sql = Sql & " a.EffStatus"
  End If
  Sql = Sql & " From FAS_PARTY_MASTER a"
  Sql = Sql & " Left Outer Join FAS_PARTY_ADDRESS b On b.PartyCode = a.PartyCode"
  Sql = Sql & " and b.BranchID = a.BranchID and b.CompID = a.CompID"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "'"
  
  If optCustomer Then
 Sql = Sql & " and a.AcType='D'"
End If
  If optParty Then
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
  
  If optSelective Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    spList.GetText 1, spList.ActiveRow, retVal
    Sql = Sql & " and a.PartyCode in('" & retVal & "')  "
  End If
  Sql = Sql & " Order by a.PartyCode, a.PartyName"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  If rs.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    CustomerMasterListDisplay = False
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
  
  If optBoth2 Then
    If optSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If optParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "All Customer Details"
End If
      If optParty Then
 sRtHdAtInAt = "All Party Details"
End If
    End If
  Else
    If optSelective And optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Selective Customer Details"
End If
      If optParty Then
 sRtHdAtInAt = "Selective Party Details"
End If
    ElseIf optBoth1 <> True Then
      If optCustomer Then
 sRtHdAtInAt = "Customer Details"
End If
      If optParty Then
 sRtHdAtInAt = "Party Details"
End If
    End If
  End If
  If optBoth1 Then
    If optSelective And optBoth2 <> True Then
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
  gStrReportHeading = "Customer Master Report"
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort CAlign(iPageWidth, sRtHdAtInAt)
  sSendToPort ""
  sPageHeader
  
  SRow = rs.RecordCount
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
      sSendToPort LAlign(10, rs("PartyCode")) & Space(2) & LAlign(40, rs("PartyName")) & Space(1) & LAlign(12, rs("VendorCode"))
    End If
    If chkAddress Then
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("Add1")))
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("Add2")))
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("Add3")))
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("City")))
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("Pincode")))
      sSendToPort Space(12) & LAlign(50, CheckNull(rs("Add1")))
    End If
    PartyCode = rs("PartyCode")
    If i <> rs.RecordCount Then sSendToPort ""
    i = i + 1
    rs.MoveNext
  Loop
    
  rs.Close
  
  Call EndofReport(iPageWidth, SRow)
  CustomerMasterListDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Sub sPageHeader()
  Try

'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")

  sSendToPort LAlign(10, "Party Code") & Space(2) & LAlign(40, "Party Name") & Space(1) & LAlign(12, "Vendor Code")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function sPrintcurrent() As Boolean
  
  Try
  
  sPrintcurrent = False
  
  Select Case MListName
    Case "Customer Master"
      If CustomerMasterListDisplay = False Then
        GoTo ExitHere
      End If
    Case "Customer Item Master"
      If CustomerItemMasterListDisplay = False Then
        GoTo ExitHere
      End If
    Case "JobWork Item Master"
      If JobWorkItemMasterListDisplay = False Then
        GoTo ExitHere
      End If
  End Select
  
  sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdExit1_Click()
' On Error Resume Next
  
  frmsp.Visible = False
  frmsp.Height = 705
  frmsp.Width = 1530
  spview.Height = 375
  spview.Width = 825
End Sub

Private Sub cmdPrint_Click()
  
  Try
  Me.MousePointer = vbHourglass

'  Call Spread_PagePrint(spview, 4, 54, spview.MaxCols, 3, cdlPortrait)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdview_Click()
  
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me

'  iLineNo = 1
'  iPageNo = 1
  
  If sPrintcurrent = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
  
'  sClosePort
'  Set frmOutput = New frmReportViewer
'  frmOutput.Caption = gStrReportHeading
'  frmOutput.txtTotalPages.Text = CStr(iPageNo)
'  frmOutput.wbrView.Navigate strFileName
'  frmOutput.Tag = strFileName
'  gStrReportHeading = ""
'  frmOutput.Show vbModal
'  DoEvents
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
  
  Select Case MListName
    Case "Customer Master"
      fraselective.Caption = "Customer Detail"
      rptCustomerMasterList.Caption = "Customer Master"
    Case "Customer Item Master"
      fraselective.Caption = "Customer Item Detail"
      rptCustomerMasterList.Caption = "Customer Item Master"
    Case "Jobwork Item Master"
      fraselective.Caption = "Jobwork Item Detail"
      rptCustomerMasterList.Caption = "Jobwork Item Master"
  End Select
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
    MDISales.Toolbar1.Visible = False
    RptFlg = ""
  End If
End Sub

Private Sub optAll_Click()

  Try
  Me.MousePointer = vbHourglass

  Call Spread_Clear(spList)
  spList.MaxRows = 1
  chkAddress.Value = 0
  fraselective.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optCustomer_Click()

  Try
  Me.MousePointer = vbHourglass
  
  chkAddress.Value = 0
  If optActive Or optInActive Then
    optBoth2.Value = True
  End If
  
  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
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

  chkAddress.Value = 0
  If optActive Or optInActive Then
    optBoth2.Value = True
  End If
  
  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
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
  
  chkAddress.Value = 0
  If optActive Or optInActive Then
    optBoth2.Value = True
  End If
  
  If optSelective Then
    optAll.Value = True
  End If

  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
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

  chkAddress.Value = 0
  If optAll Then
    '
  ElseIf optSelective Then
    optAll.Value = True
  End If

  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
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

  chkAddress.Value = 0
  If optAll Then
    '
  ElseIf optSelective Then
    optAll.Value = True
  End If

  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
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
  
  chkAddress.Value = 0
  If optSelective Then
    optAll.Value = True
  End If

  If optAll Then
    Call Spread_Clear(spList)
    spList.MaxRows = 1
    fraselective.Visible = False
  ElseIf optSelective Then
    optAll.Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optSelective_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Integer
    
  Try
  Me.MousePointer = vbHourglass
  
  chkAddress.Value = 0
  spList.ClearSelection
  fraselective.Visible = True

  Select Case MListName
    Case "Customer Master"
      Sql = "Select PartyCode Code, PartyName Name from FAS_PARTY_MASTER where "
      Sql = Sql & " AcType IN ('D','Y') and Branchid = '" & gBranchID & "' and compid='" & gCompID & "'"
    Case "Customer Item Master"
      Sql = "select distinct a.CustomerCode Code,(select PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode "
      Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) Name from SAL_Customer_ItemMaster a where"
      Sql = Sql & " a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    Case "JobWork Item Master"
      Sql = "select distinct a.CustomerCode Code,(select PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode "
      Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) Name from SAL_Jobwork_ItemMaster a where"
      Sql = Sql & " a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  End Select
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
    If Not rs.EOF Then
      spList.MaxRows = 1
      spList.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        spList.SetText 1, rCnt, rs("Code")
        spList.SetText 2, rCnt, rs("Name")
        rs.MoveNext
      Next rCnt
    End If
  
  rs.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spview(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spview

End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList

End Sub

Private Sub InitializeSpreads()

	Call Design_spview("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerMasterList.frm", spview, "")

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerMasterList.frm", spList, "")

End Sub
