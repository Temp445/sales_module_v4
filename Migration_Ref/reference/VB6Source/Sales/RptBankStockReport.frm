VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RptBankStockReport 
   Caption         =   "Bank Stock Report"
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
   Begin VB.Frame FraPlaceHolder 
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   5362
      TabIndex        =   0
      Top             =   3532
      Width           =   3885
      Begin VB.CommandButton cbExit 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   2580
         TabIndex        =   1
         Top             =   1740
         Width           =   1300
      End
      Begin VB.Frame fratype 
         Height          =   945
         Index           =   3
         Left            =   375
         TabIndex        =   6
         Top             =   645
         Width           =   3135
         Begin VB.ComboBox cmbType 
            Height          =   330
            ItemData        =   "RptBankStockReport.frx":0000
            Left            =   1035
            List            =   "RptBankStockReport.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   210
            Width           =   1995
         End
         Begin VB.CheckBox chkPrintDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Print Date"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   585
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "Report Type"
            Height          =   210
            Left            =   105
            TabIndex        =   9
            Top             =   255
            Width           =   885
         End
      End
      Begin VB.Frame FraStep 
         Caption         =   "As On Date"
         Height          =   645
         Index           =   0
         Left            =   1222
         TabIndex        =   4
         Top             =   0
         Width           =   1440
         Begin MSComCtl2.DTPicker dtpMonth 
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   225
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
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
            Format          =   22282241
            CurrentDate     =   37273
         End
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
         Left            =   2745
         TabIndex        =   3
         Top             =   165
         Width           =   345
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   1292
         TabIndex        =   2
         Top             =   1740
         Width           =   1300
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&Proceed"
         Enabled         =   0   'False
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   1740
         Width           =   1300
      End
   End
End
Attribute VB_Name = "RptBankStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim sprow As Long
Dim sShe As Integer

Private Erpno As String
Private TsNo As String

Private Sub cmdStep_Click(ByVal Index As Integer)
' On Error Resume Next

  Call StepProceedNext(Me, Index)
  CmdView.Enabled = True
  
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call ClearForm(Me)
  
  Call StepClear(Me)
  Call StepEnable(Me)
  
  cmbType.ListIndex = 0
  chkPrintDate.Value = 1
  chkPrintDate.Value = 1
  dtpMonth.Value = Format(ServerDate, "dd/MM/yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  Erpno = Get_ERPandTSNo("mnuBankStock", "SAL", mErp)
  TsNo = Get_ERPandTSNo("mnuBankStock", "SAL", mTs)
  If sSpread_sPrintcurrent = False Then
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
  
  CmdView.Enabled = False
  
  cmbType.ListIndex = 0
  
  dtpMonth.Value = Format(ServerDate, "MM-yyyy")
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

Public Sub Header()
  
  Dim Spr As AceSpread
  
  Try
  
  Set Spr = frmPrint.sprPrint
  
  Spr.sheet = Spr.sheet
  Call spSetColWidth(Spr, 1, 15, 2, 30, 3, 8, 4, 8, 5, 10)
  
  Call RptStandard_Header(5, sprow, Erpno, TsNo, IIf(chkPrintDate, True, False))
  GridCount = GridCount + 6
  
  Call Spread_BkColour(Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow)
  
  If sShe = 1 Then
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Item Code", "Description", "Quantity", "Rate", "Value")
  ElseIf sShe = 0 Then
    Call spSetText(Spr, 1, sprow, "Stock Type", mAlignCenterCenter, True, 3, 1)
    Spr.AddCellSpan 4, sprow, 5, 1
    Call spSetText(Spr, 4, sprow, "Value", mAlignCenterCenter, True)
  End If
  
  Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)

  sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow, Val(Spr.sheet)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sprint()
' On Error Resume Next
  Call cmdClear_Click
End Sub

Public Sub sView()
' On Error Resume Next
  Call cmdClear_Click
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
  Dim Sql As String
  Dim genrec As AceADODB.Recordset
  
  Dim str As String
  Dim Heading As String
  Dim TotQty(6) As String
  
  Dim i As Integer
  Dim Spr As AceSpread
  Dim CurDate As Date
  
  Try
  
  sSpread_sPrintcurrent = False
  
  Set Spr = frmPrint.sprPrint
  
  sShe = 1
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
  
    For i = 1 To 6
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
      ElseIf i = 5 Then
        str = "SUB"
        Heading = "Sub Contractor "
      ElseIf i = 6 Then
        str = "WIP"
        Heading = "WIP "
      End If
      
      Sql = " Select * From dbo.fn_SAL_BankProcess('" & Format(CurDate, "yyyy-MM") & "','" & Format(CurDate, "yyyy-MM-dd " & "23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd " & "23:59:59") & "','" & Trim$(str) & "','" & gBranchID & "', '" & gCompID & "','" & gFinyear & "')"
      Set genrec = New AceADODB.Recordset
      If RSOpen(genrec, Sql) = False Then Exit Function
'      If genrec.RecordCount = 0 Then
'        sSpread_sPrintcurrent = False
'        sShowMessage "No Record's Found"
'        GoTo ExitHere
'      End If
      
      Spr.sheet = i
Spr.SheetName = str
      gStrReportHeading = Heading & "As On Stock " & Format(dtpMonth, "MMM-yy")
      sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
      
      Call Header
      
      Do While Not genrec.EOF
        Call spSetTextColArray(Spr, 1, sprow, 1, mAlignLeftCenter, False, genrec("ItemCode"), genrec("ItemDescription"))
        Call spSetTextColArray(Spr, 3, sprow, 1, mAlignRightCenter, False, Val(genrec("CBQty")), Format(genrec("Rate"), "0.00"), Format((genrec("CBQty") * genrec("Rate")), "0.00"))
        Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        TotQty(i - 1) = Val(genrec("CBQty") * genrec("Rate")) + Val(TotQty(i - 1))
        genrec.MoveNext
      Loop
      Call spSetText(Spr, 4, sprow, "TOTAL", mAlignRightCenter, True)
      Call spSetText(Spr, 5, sprow, Format(Val(TotQty(i - 1)), "0.00"), mAlignRightCenter, True)
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
    Next i
    genrec.Close
    
    'Grand Total
    sShe = 0
    Spr.sheet = 7
Spr.SheetName = "Grand Total"
    gStrReportHeading = "CLOSING BALANCE As On Stock " & Format(dtpMonth, "MMM-yy")
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 54
Orient = 2
    
    Call Header
    
    Call spSetText(Spr, 1, sprow, "TOOLS & CONSUMABLES", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(0)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow + 6, mBorderAllSides)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "FINISHED GOODS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(1)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "INSTRUMENTS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(2)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "RAW MATERIAL", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(3)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "SUB CONTRACTOR", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(4)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "WORK IN  PROGRESS", mAlignLeftCenter, False, 3, 1)
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(5)), "0.000"), mAlignRightCenter, False, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
    
    Call spSetText(Spr, 1, sprow, "GRAND TOTAL", mAlignLeftCenter, True, 3, 1)
    TotQty(5) = (Val(TotQty(0)) + Val(TotQty(1)) + Val(TotQty(2)) + Val(TotQty(3)) + Val(TotQty(4)) + Val(TotQty(5)))
    Call spSetText(Spr, 4, sprow, Format(Val(TotQty(5)), "0.000"), mAlignRightCenter, True, Spr.MaxCols, 1)
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
  
    Spr.MaxRows = Spr.DataRowCnt + 2
    Spr.MaxCols = Spr.DataColCnt + 1
    totalpages = iPageNo
    Spr.SheetCount = 7
    
    For i = 1 To 7
      Spr.sheet = i
      Spr.Lock = True
      Spr.MaxRows = Spr.DataRowCnt + 2
    Next i
    
    
  frmPrint.Show
  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set genrec = Nothing
End Try
End Function
