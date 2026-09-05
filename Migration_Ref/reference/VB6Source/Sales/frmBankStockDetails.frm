VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmBankStockDetails 
   Caption         =   "Bank Stock Details"
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
   Begin VB.Frame Frame3 
      Caption         =   "Help"
      Height          =   480
      Left            =   10950
      TabIndex        =   13
      Top             =   2550
      Width           =   1935
      Begin VB.Label Label1 
         Caption         =   "Ctrl+Del = Delete Row"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   45
         TabIndex        =   14
         Top             =   180
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   5040
      TabIndex        =   9
      Top             =   9270
      Width           =   4575
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   405
         Left            =   3015
         TabIndex        =   10
         ToolTipText     =   "Click to Exit from the Current Screen"
         Top             =   165
         Width           =   1500
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   405
         Left            =   1530
         TabIndex        =   11
         ToolTipText     =   "Click to Exit (Alt + x)"
         Top             =   165
         Width           =   1500
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   405
         Left            =   45
         TabIndex        =   12
         ToolTipText     =   "Click to View the Report"
         Top             =   165
         Width           =   1500
      End
   End
   Begin VB.Frame fraProgress 
      Height          =   1260
      Left            =   4777
      TabIndex        =   5
      Top             =   1605
      Visible         =   0   'False
      Width           =   5100
      Begin MSComctlLib.ProgressBar pbrProcess 
         Height          =   420
         Left            =   105
         TabIndex        =   6
         Top             =   465
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblEmployee 
         AutoSize        =   -1  'True
         Caption         =   "Report Heading"
         Height          =   210
         Left            =   105
         TabIndex        =   8
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         Caption         =   "Pls wait. Process may take few minutes to complete..."
         Height          =   210
         Left            =   105
         TabIndex        =   7
         Top             =   225
         Width           =   3900
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8070
      TabIndex        =   4
      Top             =   1125
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Process Month"
      Height          =   615
      Left            =   6615
      TabIndex        =   2
      Top             =   990
      Width           =   1425
      Begin MSComCtl2.DTPicker dtpMonth 
         Height          =   330
         Left            =   75
         TabIndex        =   3
         ToolTipText     =   "Enter the From Date"
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM-yyyy"
         Format          =   64487427
         CurrentDate     =   37272
      End
   End
   Begin VB.Frame fraView 
      Caption         =   "Stock Details"
      Height          =   6240
      Left            =   1747
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   11160
Begin AceSpread spView
         Height          =   5910
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   10920
         _Version        =   458752
         _ExtentX        =   19262
         _ExtentY        =   10425
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         MaxCols         =   5
      End
   End
End
Attribute VB_Name = "frmBankStockDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim sprow As Long
Dim sRow(4) As String
Dim sShe As Integer

Private Sub sProceedTotal(ByVal sheet As Integer)
  Dim i As Integer
  Dim tmp As Variant
  
  Try
  
  Select Case sheet
    Case Is = 1
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 2, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
    
    Case Is = 2
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 3, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
    
    Case Is = 3
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 4, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
    
    Case Is = 4
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 5, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
    
    Case Is = 5
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 6, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
    
    Case Is = 6
      spView.sheet = sheet
      spView.GetText 5, spView.MaxRows, tmp
      spView.sheet = 7
      Call spSetText(spView, 4, 7, Format(Val(tmp), "0.000"), mAlignRightCenter, False, spView.MaxCols, 1)
  End Select
  
  spView.sheet = 7
  tmp = Spread_Cell_Total(spView, 4, 2, 4, (spView.MaxRows - 1))
  Call spSetText(spView, 4, 8, Format(Val(tmp), "0.000"), mAlignRightCenter, True, spView.MaxCols, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  dtpMonth.Value = Format(CurDate, "MMM-yyyy")
  
  Call Spread_Clear(spView)
  spView.MaxRows = 1
  cmdSave.Enabled = False
  cmdOK.Enabled = True
  spView.Enabled = False
  Frame1.Visible = True
  fraView.Visible = False
  fraProgress.Visible = False
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

Private Sub CmdOk_Click()
  Dim i As Long
  Dim j As Integer
  Dim Sql As String
  Dim Spr As AceSpread
  Dim genrec As AceADODB.Recordset
  Dim CurDate As Date
  Dim str As String
  Dim sShe As Integer
  Dim sValue As Double
  Dim Heading As String
  Dim TotQty(6) As String
  Dim Process As Boolean
  
  Try
  Me.MousePointer = vbHourglass

  Set Spr = spView
  
  sShe = 1
Process = False

  sprow = 1
GridCount = 1
  Spr.SheetCount = 7
  i = 1
  str = "TCS"
  CurDate = Format(dtpMonth, "yyyy-MM" & "-01")
  CurDate = DateAdd("d", -1, (DateAdd("M", 1, Format(CurDate, "yyyy-MM-dd"))))
  TotQty(0) = 0
TotQty(1) = 0
TotQty(2) = 0
TotQty(3) = 0
TotQty(4) = 0
TotQty(5) = 0
  fraView.Visible = True
  
    For i = 1 To 4
      If i = 1 Then
        Heading = "Tools & Consumables "
      ElseIf i = 2 Then
        str = "FGS"
        Heading = "Finished Goods "
      ElseIf i = 3 Then
        str = "CHI"
        Heading = "Checking Instruments "
      ElseIf i = 4 Then
        str = "RMS"
        Heading = "Raw Material "
      End If
      
      Sql = " Execute SAL_BankInformation_Result "
      Sql = Sql & " @WarehoseCode = '" & Trim$(str) & "',"
      Sql = Sql & " @AsOnDate = '" & Format(CurDate, "yyyy-MM-dd" & " 23:59:59") & "', "
      Sql = Sql & " @RptDate = '" & Format(CurDate, "yyyy-MM-dd" & " 23:59:59") & "', "
      Sql = Sql & " @BranchID = '" & gBranchID & "', "
      Sql = Sql & " @CompID = '" & gCompID & "',"
      Sql = Sql & " @Finyear = '" & gFinyear & "'"
      Set genrec = New AceADODB.Recordset
      If RSOpen(genrec, Sql) = False Then Exit Sub
      
      If Not genrec.EOF Then
        Process = True
        Spr.sheet = i
Spr.SheetName = str
        lblEmployee.Caption = Trim$(Heading)
        lblEmployee.Visible = True
        Spr.MaxRows = genrec.RecordCount
        
        'Progress Bar
        pbrProcess.Max = genrec.RecordCount
        pbrProcess.Min = 0
        fraProgress.Visible = True
  
        gStrReportHeading = Heading & "As On Stock " & Format(dtpMonth, "MMM-yy")
        sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
        
        'Header
        Call spSetColWidth(Spr, 1, 15, 2, 40, 3, 10, 4, 10, 5, 13)
        Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Item Code", "Description", "Quantity", "Rate", "Value")
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        Call Spread_Lock(spView, True, 1, sprow, 3, sprow, i)
        sprow = sprow + 1
GridCount = GridCount + 1

        genrec.MoveLast
genrec.MoveFirst
        j = 1
        Do While Not genrec.EOF
          Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
          Call spSetTextColArray(Spr, 3, sprow, 1, mAlignRightCenter, False, Format(Val(genrec("CBQty")), "0.00"), Format(genrec("Rate"), "0.00"), Format((genrec("CBQty") * genrec("Rate")), "0.00"))
          Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
          
          'Spr Lock
          Call Spread_Lock(spView, True, 1, sprow, spView.MaxCols, sprow, i)

          sprow = sprow + 1
GridCount = GridCount + 1
          TotQty(i - 1) = Val(genrec("CBQty") * genrec("Rate")) + Val(TotQty(i - 1))
          pbrProcess.Value = j
DoEvents
j = j + 1
          genrec.MoveNext
        Loop
        'Spr Lock
        Call Spread_Lock(spView, True, 1, sprow, 3, sprow, i)
        Call Spread_Lock(spView, True, spView.MaxCols, sprow, spView.MaxCols, sprow, i)

        Call spSetText(Spr, 4, sprow, "TOTAL", mAlignRightCenter, True)
        Call spSetText(Spr, 5, sprow, Format(Val(TotQty(i - 1)), "0.00"), mAlignRightCenter, True)
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        
'      Else
'        sShowMessage "Record(s) Not Found"
'        fraProgress.Visible = False
'        GoTo ExitHere
      End If
    Next i
    genrec.Close
    
    'SubContractor
    
    Sql = " Execute dbo.SAL_BankProcess_SubContract '" & Format(CurDate, "yyyy-MM-dd" & " 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "','" & gFinyear & "'"
    Set genrec = New AceADODB.Recordset
    If RSOpen(genrec, Sql) = False Then Exit Sub
''    If genrec.RecordCount = 0 Then
''      sShowMessage "No Record's Found"
''      GoTo ExitHere
''    End If
    If Not genrec.EOF Then
      Process = True
      lblEmployee.Caption = "Sub Contractor"
      lblEmployee.Visible = True
      
      'Progress Bar
      pbrProcess.Max = genrec.RecordCount
      pbrProcess.Min = 0
      fraProgress.Visible = True

      Spr.sheet = 5
Spr.SheetName = "SUB"
      gStrReportHeading = "Sub Contractor As On Stock " & Format(dtpMonth, "MMM-yy")
      sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
      Spr.MaxRows = genrec.RecordCount

      'Header
      Call spSetColWidth(Spr, 1, 15, 2, 40, 3, 10, 4, 10, 5, 13, 6, 10)
      Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Item Code", "Description", "Quantity", "Rate", "Value", "PartyCode")
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      sprow = sprow + 1
GridCount = GridCount + 1

      j = 1
      Do While Not genrec.EOF
        sValue = Val(genrec("CBQty") * CheckNull(genrec("Rate"), 0))
  
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, sprow, 1, mAlignRightCenter, False, Format(Val(genrec("CBQty")), "0.00"), Format(genrec("Rate"), "0.00"), Format(sValue, "0.00")) ', genrec("Partycode"))
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        
        'Spr Lock
        Call Spread_Lock(spView, True, 1, sprow, spView.MaxCols, sprow, Spr.sheet)

        sprow = sprow + 1
GridCount = GridCount + 1
        TotQty(4) = Val(sValue) + Val(TotQty(4))
        pbrProcess.Value = j
DoEvents
j = j + 1
        genrec.MoveNext
      Loop
      Call Spread_HideCol(spView, True, 6)
      'Spr Lock
      Call Spread_Lock(spView, True, 1, sprow, 3, sprow, Spr.sheet)
      Call Spread_Lock(spView, True, spView.MaxCols, sprow, spView.MaxCols, sprow, Spr.sheet)

      Call spSetText(Spr, 4, sprow, "TOTAL", mAlignRightCenter, True)
      Call spSetText(Spr, 5, sprow, Format(Val(TotQty(4)), "0.00"), mAlignRightCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
''    Else
''      sShowMessage "Record(s) Not Found"
''      fraProgress.Visible = False
''      GoTo ExitHere
    End If
    genrec.Close
    
    'WIP
    Sql = " Execute dbo.SAL_BankProcess_WIP 'WIP','" & Format(CurDate, "yyyy-MM-dd" & " 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "','" & gFinyear & "'"
'    Sql = "  Execute dbo.SAL_BankProcess_WIP 'WIP','2009-08-01 16:00:59', '2009-08-01 16:00:59', '00', '01','04-2009 03-2010'"
    Set genrec = New AceADODB.Recordset
    If RSOpen(genrec, Sql) = False Then Exit Sub
'    If genrec.RecordCount = 0 Then
'      sShowMessage "No Record's Found"
'      GoTo ExitHere
'    End If
    If Not genrec.EOF Then
      Process = True
      lblEmployee.Caption = "WIP"
      lblEmployee.Visible = True
      
      'Progress Bar
      pbrProcess.Max = genrec.RecordCount
      pbrProcess.Min = 0
      fraProgress.Visible = True
      
      Spr.sheet = 6
Spr.SheetName = "WIP"
      gStrReportHeading = "WIP As On Stock " & Format(dtpMonth, "MMM-yy")
      sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
      Spr.MaxRows = genrec.RecordCount

      'Header
      Call spSetColWidth(Spr, 1, 15, 2, 40, 3, 10, 4, 10, 5, 13)
      Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
      Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Item Code", "Description", "Quantity", "Rate", "Value")
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      sprow = sprow + 1
GridCount = GridCount + 1

      j = 1
      Do While Not genrec.EOF
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, sprow, 1, mAlignRightCenter, False, Format(Val(genrec("Qty")), "0.00"), Format(genrec("Rate"), "0.00"), Format((genrec("Qty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        
        'Spr Lock
        Call Spread_Lock(spView, True, 1, sprow, spView.MaxCols, sprow, Spr.sheet)

        sprow = sprow + 1
GridCount = GridCount + 1
        TotQty(5) = Val(genrec("Qty") * CheckNull(genrec("Rate"), 0)) + Val(TotQty(5))
        pbrProcess.Value = j
DoEvents
j = j + 1
        genrec.MoveNext
      Loop
      'Spr Lock
      Call Spread_Lock(spView, True, 1, sprow, 3, sprow, Spr.sheet)
      Call Spread_Lock(spView, True, spView.MaxCols, sprow, spView.MaxCols, sprow, Spr.sheet)

      Call spSetText(Spr, 4, sprow, "TOTAL", mAlignRightCenter, True)
      Call spSetText(Spr, 5, sprow, Format(Val(TotQty(5)), "0.00"), mAlignRightCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)

'    Else
'      sShowMessage "Record(s) Not Found"
'      fraProgress.Visible = False
'      GoTo ExitHere
    End If
    
    'GrantTotal
    sShe = 0
    Spr.sheet = 7
Spr.SheetName = "Grant Total"
    gStrReportHeading = "CLOSING BALANCE As On Stock " & Format(dtpMonth, "MMM-yy")
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
    Spr.MaxRows = 10

    'Header
    Call spSetColWidth(Spr, 1, 15, 2, 40, 3, 10, 4, 10, 5, 13)
    Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
    Call spSetText(Spr, 1, sprow, "Stock Type", mAlignCenterCenter, True, 3, 1)
    Spr.AddCellSpan 4, sprow, 5, 1
    Call spSetText(Spr, 4, sprow, "Value", mAlignCenterCenter, True)
    Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "TOOLS & CONSUMABLES", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(0)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow + 6, mBorderAllSides)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "FINISHED GOODS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(1)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "INSTRUMENTS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(2)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "RAW MATERIAL", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(3)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "SUB CONTRACTOR", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(4)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "WORK IN  PROGRESS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(5)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    Call spSetText(Spr, 1, sprow, "GRANT TOTAL", mAlignLeftCenter, True, 3, 1)
    TotQty(5) = (Val(TotQty(0)) + Val(TotQty(1)) + Val(TotQty(2)) + Val(TotQty(3)) + Val(TotQty(4)) + Val(TotQty(5)))
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(5)), "0.000"), mAlignRightCenter, True, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
    
    'Spr Lock
    Call Spread_Lock(spView, True, 1, 1, spView.MaxCols, spView.MaxRows, Spr.sheet)

    Spr.MaxRows = Spr.DataRowCnt + 2
    Spr.MaxCols = Spr.DataColCnt + 1
    totalpages = iPageNo
    Spr.SheetCount = 7
    For i = 1 To 7
      Spr.sheet = i
      Spr.Lock = True
      Spr.RowHeadersShow = False
      Spr.ColHeadersShow = False
      Spr.MaxRows = Spr.DataRowCnt
    Next i
  
  If Process = True Then
    fraView.Visible = True
    cmdOK.Enabled = False
    spView.Enabled = True
    Frame1.Visible = True
    cmdSave.Enabled = True
    sShowMessage "Process Over!"
  Else
    sShowMessage "Record(s) Not Found"
    fraProgress.Visible = False
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim Sql As String
  Dim genrec As AceADODB.Recordset
  Dim retVal As Variant
  Dim revNo As Integer
  Dim i As Integer
  Dim sCnt As Integer
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
  Sql = " Delete from SAL_BankStockProcess Where ProcessMonth = '" & Format(dtpMonth, "yyyy-MM") & "' "
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
  For sCnt = 1 To 6
    spView.sheet = sCnt
    For i = 2 To (spView.DataRowCnt - 1)
      Sql = "Execute Sal_Ins_BankStockProcess"
      spView.GetText 6, i, retVal
      Sql = Sql & " @Partycode = '" & Val(retVal) & "',"
      retVal = spView.SheetName
      Sql = Sql & " @WHCode = '" & Trim$(retVal) & "',"
      Sql = Sql & " @ProcessMonth = '" & Format(dtpMonth, "yyyy-MM") & "',"
      spView.GetText 1, i, retVal
      Sql = Sql & " @ItemCode = '" & Trim$(retVal) & "',"
      spView.GetText 3, i, retVal
      Sql = Sql & " @Qty = " & Format(retVal, "0.00") & ","
      spView.GetText 4, i, retVal
      Sql = Sql & " @Rate = " & Format(Val(retVal), "0.00") & ","
      Sql = Sql & " @UserID = '" & gUserID & "',"
      Sql = Sql & " @BranchID = '" & gBranchID & "',"
      Sql = Sql & " @CompID = '" & gCompID & "',"
      Sql = Sql & " @TranStamp = '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    Next i
  Next sCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record(s) Saved"
  Call cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  dtpMonth.Value = Format(CurDate, "MMM-yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spView_EditChange(sender as object, e as EditChangeEventArgs)
'  Dim retval As Variant
'  Dim retVal1 As Variant
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  spView.Sheet = spView.ActiveSheet
'  If Col = 4 And spView.Sheet < 7 Then
'    spView.GetText 3, Row, retval
'    spView.GetText 4, Row, retVal1
'    retval = (Val(retval) * Val(retVal1))
'
'    spView.SetText 5, Row, Format(retval, "0.000")
'
'    retval = Spread_Col_Total(spView, 5, 2, (spView.MaxRows - 1))
'    spView.SetText 5, spView.MaxRows, Format(retval, "0.000")
'  End If
'
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

Private Sub spView_EditMode(sender as object, e as EditModeEventArgs)
'  Dim retval As Variant
'  Dim retVal1 As Variant
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  spView.Sheet = spView.ActiveSheet
'  If Col = 4 Then
'    spView.GetText 3, Row, retval
'    spView.GetText 4, Row, retVal1
'    retval = (Val(retval) * Val(retVal1))
'
'    spView.SetText 5, Row, Format(retval, "0.000")
'
'    retval = Spread_Col_Total(spView, 5, 2, (spView.MaxRows - 1))
'    spView.SetText 5, spView.MaxRows, Format(retval, "0.000")
'  End If
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

Private Sub spView_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim tqty As Double
  
  Try
  Me.MousePointer = vbHourglass
  
    spView.sheet = spView.ActiveSheet
    If KeyCode = vbKeyDelete Then
      spView.DeleteRows spView.ActiveRow, 1
      spView.MaxRows = spView.MaxRows - 1
      tqty = Spread_Cell_Total(spView, 5, 2, 5, spView.MaxRows - 1)
      spView.SetText 5, spView.MaxRows, Format(tqty, "0.00")
      Call sProceedTotal(spView.sheet)
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spView(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spView


spView.EditChange +=  new EventHandler(spView_EditChange)
spView.EditMode +=  new EventHandler(spView_EditMode)
spView.KeyDown +=  new EventHandler(spView_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spView("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmBankStockDetails.frm", spView, "")

End Sub
