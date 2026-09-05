VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptAgeingAnalysisreport 
   Caption         =   "Ageing Analysis"
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
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6525
      TabIndex        =   11
      Top             =   4065
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5235
      TabIndex        =   12
      Top             =   4065
      Width           =   1300
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3945
      TabIndex        =   13
      Top             =   4065
      Width           =   1300
   End
   Begin VB.Frame FraReportType 
      Caption         =   "Report Type"
      Height          =   945
      Left            =   4625
      TabIndex        =   8
      Top             =   2985
      Width           =   2520
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "rptAgeingAnalysisreport.frx":0000
         Left            =   90
         List            =   "rptAgeingAnalysisreport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2340
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   5150
      TabIndex        =   5
      Top             =   1665
      Width           =   1470
      Begin MSComCtl2.DTPicker dtpAsondate 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57737217
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "As On Date"
         Height          =   210
         Left            =   300
         TabIndex        =   7
         Top             =   195
         Width           =   840
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "A/C Type"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   4625
      TabIndex        =   2
      Top             =   2430
      Width           =   2520
      Begin VB.OptionButton optCreditor 
         Caption         =   "Creditor"
         Height          =   225
         Left            =   1455
         TabIndex        =   4
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton optDeptor 
         Caption         =   "Deptor"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   810
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
      Left            =   6705
      TabIndex        =   1
      Top             =   1890
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
      Left            =   7230
      TabIndex        =   0
      Top             =   2550
      Width           =   345
   End
End
Attribute VB_Name = "rptAgeingAnalysisreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sprow As Long

Private Erpno As String
Private TsNo As String

Public Sub sView()
  'No Change
End Sub

Public Sub sPrint()
  'No Change
End Sub
Private Function sSpread_sPrintcurrent() As Boolean
  Dim i As Integer
  Dim Sql As String
  Dim Spr As AceSpread
  Dim rs As AceADODB.Recordset
  Dim TotAmt As Double
  
  Try
  
  TotAmt = 0
  If optDeptor.Value = True Then
    Sql = "Select LrCode, Particulars, abs(Less30) As Less1 ,"
    Sql = Sql & " Case When Less30<0 then 'Dr' When Less30>0 then 'Cr' Else '' End As Type0,"
    Sql = Sql & " 0 As Amount1to30, '' As Type1,"
'    Sql = Sql & " abs(Amount1to30) As Amount1to30, Case When Amount1to30<0 then 'Dr' when Amount1to30>0 then 'Cr' else '' end Type1,"
    Sql = Sql & " abs(Amount31to60) As Amount31to60, Case When Amount31to60<0 then 'Dr' when Amount31to60>0 then 'Cr' else '' end Type2,"
    Sql = Sql & " abs(Amount61to90) As Amount61to90, Case When Amount61to90<0 then 'Dr' when Amount61to90>0 then 'Cr' else '' end Type3,"
    Sql = Sql & " abs(Amount91to180) As Amount91to180, Case When Amount91to180<0 then 'Dr' when Amount91to180>0 then 'Cr' else '' end Type4,"
    Sql = Sql & " abs(Amount180) As Amount180, Case When Amount180<0 then 'Dr' when Amount180>0 then 'Cr' else '' end Type5,"
    Sql = Sql & " abs(PendingAmt) As PendingAmt, Case When PendingAmt<0 then 'Dr' when PendingAmt>0 then 'Cr' else '' end Type6,"
    Sql = Sql & " abs(AdvAmt) As AdvAmt, Case When AdvAmt<0 then 'Dr' when AdvAmt>0 then 'Cr' else '' end Type7,"
    Sql = Sql & " abs(TotAmt) As TotAmt, Case When TotAmt<0 then 'Dr' when TotAmt>0 then 'Cr' else '' end Type8"
    Sql = Sql & " From dbo.fn_FAS_Dis_AgeingAnalysis_tmp ('0206', '0' , 'B', '" & Format(dtpAsondate, "yyyy-MM-dd") & "', '" & Format((Left(gFinyear, 7)), "yyyy-MM-dd") & " ', '" & Format((Right(gFinyear, 7)), "yyyy-MM") & "-31" & " ', '" & gFasCompCode & "')"
    Sql = Sql & " Where PendingAmt<>0 or AdvAmt <> 0 or Type = 'Group' and TotAmt > 0 "
    Sql = Sql & " Order By LrRoute, Type, Particulars"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then Exit Function
    If rs.RecordCount <= 0 Then
      sShowMessage "Record(s) Not Found"
     GoTo ExitHere
    End If
    
    Set Spr = frmPrint.sprPrint
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 1

      Call Header
      Do While Not rs.EOF
        If Val(rs("TotAmt")) > 0 Then
          Call spSetText(Spr, 1, sprow, rs("Particulars"), mAlignLeftCenter)
    
          Call spSetTextColArray(Spr, 2, sprow, 1, mAlignRightCenter, False, Format(rs("Less1"), "0.00"), rs("Type0"), Format(rs("Amount1to30"), "0.00"), rs("Type1"), Format(rs("Amount31to60"), "0.00"), rs("Type2"), Format(rs("Amount61to90"), "0.00"), rs("Type3"), Format(rs("Amount91to180"), "0.00"), rs("Type4"), Format(rs("Amount180"), "0.00"), rs("Type5"), Format(rs("PendingAmt"), "0.00"), rs("Type6"), Format(rs("AdvAmt"), "0.00"), rs("Type7"), Format(rs("TotAmt"), "0.00"), rs("Type8"))
        
          Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
    
          sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
          If Trim$(rs("Type8")) = "Cr" Then
            TotAmt = Val(TotAmt) + Val(rs("TotAmt"))
          End If
        End If
        rs.MoveNext
      Loop
      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      Call spSetText(Spr, 1, sprow, "Total", mAlignRightCenter, True)
      Call spSetText(Spr, 18, sprow, Format(Val(TotAmt), "0.00"), mAlignRightCenter, True)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow

      Spr.MaxRows = Spr.DataRowCnt + 2
      totalpages = iPageNo  Else
  
    Sql = "Select LrCode, Particulars, abs(Less1) As Less1,"
    Sql = Sql & " case when Less1<0 then 'Cr' when Less1>0 then 'Dr' else '' end Type0,"
    Sql = Sql & " abs(Amount1to30) as Amount1to30, case when Amount1to30<0 then 'Cr' when Amount1to30>0 then 'Dr' else '' end Type1,"
    Sql = Sql & " abs(Amount31to60) as Amount31to60, case when Amount31to60<0 then 'Cr' when Amount31to60>0 then 'Dr' else '' end Type2,"
    Sql = Sql & " abs(Amount61to90) as Amount61to90, case when Amount61to90<0 then 'Cr' when Amount61to90>0 then 'Dr' else '' end Type3,"
    Sql = Sql & " abs(Amount90) as Amount90, case when Amount90<0 then 'Cr' when Amount90>0 then 'Dr' else '' end Type4,"
    Sql = Sql & " abs(PendingAmt) as PendingAmt, case when PendingAmt<0 then 'Cr' when PendingAmt>0 then 'Dr' else '' end Type5,"
    Sql = Sql & " abs(AdvAmt) as AdvAmt, case when AdvAmt<0 then 'Cr' when AdvAmt>0 then 'Dr' else '' end Type6,"
    Sql = Sql & " abs(TotAmt) as TotAmt, case when TotAmt<0 then 'Cr' when TotAmt>0 then 'Dr' else '' end Type7"
    Sql = Sql & " From dbo.fn_FAS_Dis_AgeingAnalysis ('0303', 'B', '" & Format(dtpAsondate, "yyyy-MM-dd") & "', '" & Format((Left(gFinyear, 7) & "-1"), "yyyy-MM-dd") & " ', '" & Format((Right(gFinyear, 7) & "-31"), "yyyy-MM-dd") & " ', '" & gCompID & gBranchID & "')"
    Sql = Sql & " Where PendingAmt<>0 or AdvAmt <> 0 or Type = 'Group' and TotAmt > 0 "
    Sql = Sql & " Order By LrRoute, Type, Particulars"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then Exit Function
    If rs.RecordCount <= 0 Then
      sShowMessage "Record(s) Not Found"
     GoTo ExitHere
    End If
    
    Set Spr = frmPrint.sprPrint
    sprow = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 1

      Call Header
      Do While Not rs.EOF
        If Val(rs("TotAmt")) > 0 Then
          Call spSetText(Spr, 1, sprow, rs("Particulars"), mAlignLeftCenter)
    
          Call spSetTextColArray(Spr, 2, sprow, 1, mAlignRightCenter, False, Format(rs("Less1"), "0.00"), rs("Type0"), Format(rs("Amount1to30"), "0.00"), rs("Type1"), Format(rs("Amount31to60"), "0.00"), rs("Type2"), Format(rs("Amount61to90"), "0.00"), rs("Type3"), Format(rs("Amount90"), "0.00"), rs("Type4"), Format(rs("PendingAmt"), "0.00"), rs("Type5"), Format(rs("AdvAmt"), "0.00"), rs("Type6"), Format(rs("TotAmt"), "0.00"), rs("Type7"))
        
          Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
        End If
        
        If Trim$(rs("Type7")) = "Dr" Then
          TotAmt = Val(TotAmt) + Val(rs("TotAmt"))
        End If
        
        sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
        rs.MoveNext
      Loop
      
'      If Trim$(rs("Type7")) = "Dr" Then
'        TotAmt = Val(TotAmt) + Val(rs("TotAmt"))
'      End If

      Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)
      Call spSetText(Spr, 1, sprow, "Total", mAlignRightCenter, True)
      Call spSetText(Spr, 18, sprow, Format(TotAmt, "0.00"), mAlignRightCenter, True)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
      Spr.MaxRows = Spr.DataRowCnt + 2
      totalpages = iPageNo
  End If
  
  frmPrint.Show
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Sub Header()
  
  Dim Spr As AceSpread
  
  Try
  
  Set Spr = frmPrint.sprPrint

  If optDeptor = True Then
    Call spSetColWidth(Spr, 1, 30, 2, 10, 3, 4, 4, 10, 5, 4, 6, 10, 7, 4, 8, 10, 9, 4, 10, 10, 11, 4, 12, 10, 13, 4, 14, 10, 15, 4, 16, 10, 17, 4, 18, 10, 19, 4)
    gStrReportHeading = "Ageing Analysis for Sundry Deptors " & Format(dtpAsondate, "MMM-yy")
    
    Call RptStandard_Header(19, sprow, Erpno, TsNo, IIf(chkPrintdate, True, False))
    GridCount = GridCount + 6
    
    Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Sundry Deptors", "<1", "Type", "<30 Days", "Type", ">30<60 Days", "Type", ">61<90 Days", "Type", ">91<180 Days", "Type", ">180 Days", "Type", "Pending Amt", "Type", "On Account", "Type", "Total Amt", "Type")
  Else
    Call spSetColWidth(Spr, 1, 30, 2, 10, 3, 4, 4, 10, 5, 4, 6, 10, 7, 4, 8, 10, 9, 4, 10, 10, 11, 4, 12, 10, 13, 4, 14, 10, 15, 4, 16, 10, 17, 4)
    gStrReportHeading = "Ageing Analysis for Sundry Creditors " & Format(dtpAsondate, "MMM-yy")
    
    Call RptStandard_Header(17, sprow, Erpno, TsNo, IIf(chkPrintdate, True, False))
    GridCount = GridCount + 6
    
    Spread_BkColour Spr, &H8000000F, 1, sprow, Spr.MaxCols, sprow
    Call spSetTextColArray(Spr, 1, sprow, 1, mAlignCenterCenter, True, "Sundry Creditors", "<1", "Type", "<30 Days", "Type", ">30<60 Days", "Type", ">61<90 Days", "Type", ">91 Days", "Type", "Pending Amt", "Type", "On Account", "Type", "Total Amt", "Type")
  End If
  
  Call spSetCellBorder(Spr, 1, sprow, Spr.MaxCols, sprow, mBorderAllSides)

  sprow = sprow + 1
GridCount = GridCount + 1
RowProcess Me, sprow
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdClear_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  dtpAsondate.Value = Format(CurDate, "dd/MM/yyyy")
  cmbReportType.ListIndex = 0
  optDeptor.Value = True
  Call StepClear(Me)
  Call StepEnable(Me)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Call StepProceedNext(Me, Index)
  
End Sub

Private Sub cmdview_Click()
  
  
  Try
  Me.MousePointer = vbHourglass
  
  If cmbReportType.ListIndex = 1 Then
'    If sPrintcurrent = False Then: GoTo ExitHere
'    Call FinaliseAndShowPrintPreview
  Else
    Erpno = Get_ERPandTSNo("mnuBaseMasterReport", "SAL", mErp)
    TsNo = Get_ERPandTSNo("mnuBaseMasterReport", "SAL", mTs)
    If sSpread_sPrintcurrent = False Then
 GoTo ExitHere
End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
  Dim CurDate As Date
  
  Try

  CurDate = ServerDate
  dtpAsondate.Value = Format(CurDate, "dd/MM/yyyy")
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

