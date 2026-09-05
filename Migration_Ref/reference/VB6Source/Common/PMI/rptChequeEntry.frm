VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptChequeEntry 
   Caption         =   "Cheque Entry"
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
      Left            =   7875
      TabIndex        =   2
      Top             =   1965
      Width           =   345
   End
   Begin VB.Frame FraReport 
      Caption         =   "Report"
      ForeColor       =   &H00404040&
      Height          =   570
      Index           =   3
      Left            =   4820
      TabIndex        =   16
      Top             =   3030
      Width           =   3000
      Begin VB.OptionButton OptDoc 
         Caption         =   "DocumentWise"
         Height          =   225
         Left            =   135
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton OptParty 
         Caption         =   "Partywise"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6960
      TabIndex        =   11
      Top             =   3720
      Width           =   1300
   End
   Begin VB.Frame FraStep 
      Caption         =   "Period"
      Height          =   675
      Index           =   0
      Left            =   4820
      TabIndex        =   14
      Top             =   1770
      Width           =   3000
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   22544385
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   22544385
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To "
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
         Left            =   1380
         TabIndex        =   15
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.Frame FraStep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   600
      Index           =   1
      Left            =   5083
      TabIndex        =   13
      Top             =   2430
      Width           =   2475
      Begin VB.OptionButton OptSelective 
         Caption         =   "Selective"
         Height          =   255
         Left            =   1275
         TabIndex        =   4
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
         Height          =   225
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Value           =   -1  'True
         Width           =   480
      End
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
      Left            =   7620
      TabIndex        =   6
      Top             =   2580
      Width           =   345
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5670
      TabIndex        =   10
      Top             =   3720
      Width           =   1300
   End
   Begin VB.Frame FraSpr 
      Caption         =   "Customer"
      Height          =   1800
      Left            =   8295
      TabIndex        =   12
      Top             =   1800
      Width           =   3555
Begin AceSpread SprCust
         Height          =   1545
         Left            =   75
         TabIndex        =   5
         Top             =   180
         Width           =   3390
         _Version        =   458752
         _ExtentX        =   5980
         _ExtentY        =   2725
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
         MaxCols         =   2
         MaxRows         =   5
         OperationMode   =   4
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4380
      TabIndex        =   9
      Top             =   3720
      Width           =   1300
   End
End
Attribute VB_Name = "rptChequeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eCustomer
  mCustomerCode = 1
  mCustomerName = 2
End Enum

Private Sub CmdClear_Click()
' On Error Resume Next
  Call StepClear(Me)
  FraSpr.Visible = False
  SprCust.ClearRange 1, 1, -1, -1, True
  SprCust.ClearSelection
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If FraStep.UBound > Index Then
    If optSelective = True Then
      Call OptSelective_Click
    Else
      Call optAll_Click
    End If
  Else
    If SprCust.SelectionCount >= 1 Then
      FraSpr.Enabled = False
'      CmdStep(Index).Enabled = True
    Else
      cmdStep(Index).Enabled = False
    End If
  End If
End Sub

Private Sub cmdView_Click()
  Try
  Me.MousePointer = vbHourglass
  
  If dtpToDate < dtpFromDate Then
    sShowMessage "To Date Should be greater than From Date"
    dtpToDate.SetFocus
    GoTo ExitHere
  End If
  
  If sPrintCurrent = False Then
    GoTo ExitHere
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date

  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  
  dtpFromDate = Format(CurDate, "dd-MM-yyyy")
  dtpToDate = Format(CurDate, "dd-MM-yyyy")

  FraSpr.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  cmdStep(1).Enabled = True
  FraSpr.Visible = False
  SprCust.ClearRange 1, 1, -1, -1, True
  SprCust.ClearSelection
  SprCust.MaxRows = 1
  If optSelective Then
    Call OptSelective_Click
  End If
End Sub

Private Sub OptSelective_Click()
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim Cnt As Integer

  Try
  Me.MousePointer = vbHourglass
  
  SprCust.ClearRange 1, 1, -1, -1, True
  SprCust.ClearSelection
  
  cmdStep(1).Enabled = False
  
  FraSpr.Enabled = True

  sql = " Select Distinct a.VendorCode, b.PartyName"
  sql = sql & " From PMI_Invoice_Ref a"
  sql = sql & " Left Outer Join FAS_Party_Master b"
  sql = sql & " On a.VendorCode = b.PartyCode and a.BranchID = b.BranchID and a.CompID = b.CompID"
  sql = sql & " Where a.ModuleCode = '" & gModuleCode & "'  and a.DocDate Between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "' "
  sql = sql & " and a.BranchID = '" & gBranchID & " ' and a.CompID = '" & gCompID & "'"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    FraSpr.Visible = True
    SprCust.Visible = True
      SprCust.MaxRows = rs.RecordCount
      For Cnt = 1 To rs.RecordCount
        SprCust.SetText eCustomer.mCustomerCode, Cnt, rs("VendorCode")
        SprCust.SetText eCustomer.mCustomerName, Cnt, rs("PartyName")
        rs.MoveNext
      Next Cnt
  Else
    SprCust.Visible = False
    sShowMessage "No Records Found."
    optAll.Value = True
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SprCust_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If SprCust.SelectionCount >= 1 Then
    cmdStep(1).Enabled = True
  Else
    cmdStep(1).Enabled = False
  End If
End Sub

Private Sub SprCust_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If SprCust.SelectionCount >= 1 Then
    cmdStep(1).Enabled = True
  Else
    cmdStep(1).Enabled = False
  End If
End Sub

Public Sub sPageHeader()
  Try
  
  sSendToPort String$(iPageWidth, "-")
  If OptDoc = True Then
    sSendToPort LAlign(10, "DocNo") & Space(1) & LAlign(45, "Customer Name") & Space(1) & LAlign(10, "RefNo") & Space(2) & RAlign(11, "Amount") '& Space(2) & LAlign(11, "Status")
    sSendToPort LAlign(10, "DocDate") & Space(47) & LAlign(10, "RefDate")
  Else
    sSendToPort LAlign(26, "DocNo") & Space(1) & LAlign(26, "RefNo") & Space(4) & RAlign(23, "Amount") '& Space(5) & LAlign(11, "Status")
    sSendToPort LAlign(26, "DocDate") & Space(1) & LAlign(26, "RefDate")
  End If
  sSendToPort String$(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim Cnt As Integer
  Dim PartyCode As String
  
  Try
  
  sPrintCurrent = False
  
  If optSelective Then
    If SprCust.SelectionCount < 1 Then
      sSendToPort "Please Select Atleast one Vendor"
      GoTo ExitHere
    End If
  End If

  sql = " Select Distinct a.DocType, a.DocNo, a.DocDate, a.RefType, a.RefNo, a.RefDate, abs(a.InvAmount) as InvAmount, a.VendorCode, b.PartyName,"
  sql = sql & " Case when DocType = 'CHQ' then 'Completed' else 'Pending' end as status"
  sql = sql & " From PMI_Invoice_Ref a"
  sql = sql & " Left Outer Join FAS_Party_Master b"
  sql = sql & " On a.VendorCode = b.PartyCode and a.BranchID = b.BranchID and a.CompID = b.CompID"
  sql = sql & " Where a.Modulecode = '" & gModuleCode & "' and a.DocDate Between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "' "
  If optSelective = True Then
    sql = sql & "  and a.VendorCode in "
    sql = sql & Spread_Get_SelectionList(SprCust, eCustomer.mCustomerCode)
  End If
'  If optPending = True Then
'    Sql = Sql & " and a.DocType <> 'CHQ'"
'  End If
'  If optCompleted = True Then
'    Sql = Sql & " and a.DocType = 'CHQ'"
'  End If
  sql = sql & " and a.BranchID = '" & gBranchID & " ' and a.CompID = '" & gCompID & "'"
  sql = sql & " Order By b.PartyName"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If rs.EOF Then
    MsgBox " No Records Found."
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
  sSendToPort CAlign(iPageWidth, IIf(optSelective, "Cheque Entry Report For Selective Vendors", "Cheque Entry Report"))
  sSendToPort CAlign(iPageWidth, "(" & Format(dtpFromDate, "DD/MM/YYYY") & " To " & Format(dtpToDate, "DD/MM/YYYY") & ")")
  Call sPageHeader
  
  Cnt = 0
  Do While Not rs.EOF
    If OptDoc = True Then
      If PartyCode <> rs("VendorCode") Then
        If Cnt >= 1 Then
 sSendToPort String(iPageWidth, "-")
End If
        sSendToPort LAlign(10, rs("DocType") & "-" & rs("DocNo")) & Space(1) & LAlign(45, rs("PartyName")) & Space(1) & LAlign(10, IIf(Trim$(rs("RefType")) = "", rs("RefNo"), rs("RefType") & "-" & rs("RefNo"))) & Space(2) & RAlign(11, Format(rs("InvAmount"), "0.00")) '& Space(2) & LAlign(11, rs("Status"))
        sSendToPort LAlign(10, Format(rs("DocDate"), "dd/MM/yyyy")) & Space(47) & LAlign(10, Format(rs("RefDate"), "dd/MM/yyyy"))
'        sSendToPort
      Else
        sSendToPort LAlign(10, rs("DocType") & "-" & rs("DocNo")) & Space(46) & Space(1) & LAlign(10, IIf(Trim$(rs("RefType")) = "", rs("RefNo"), rs("RefType") & "-" & rs("RefNo"))) & Space(2) & RAlign(11, Format(rs("InvAmount"), "0.00")) '& Space(2) & LAlign(11, rs("Status"))
        sSendToPort LAlign(10, Format(rs("DocDate"), "dd/MM/yyyy")) & Space(47) & LAlign(10, Format(rs("RefDate"), "dd/MM/yyyy"))
'        sSendToPort
      End If
    Else
      If PartyCode <> rs("VendorCode") Then
        If Cnt >= 1 Then
 sSendToPort String(iPageWidth, "-")
End If
        sSendToPort LAlign(10, "Customer :") & Space(1) & LAlign(50, CheckNull(rs("PartyName")))
        sSendToPort ""
      End If
      sSendToPort LAlign(26, rs("DocType") & "-" & rs("DocNo")) & Space(1) & LAlign(26, IIf(Trim$(rs("RefType")) = "", rs("RefNo"), rs("RefType") & "-" & rs("RefNo"))) & Space(4) & RAlign(23, Format(rs("InvAmount"), "0.00")) '& Space(5) & LAlign(11, rs("Status"))
      sSendToPort LAlign(26, Format(rs("DocDate"), "dd/MM/yyyy")) & Space(1) & LAlign(26, Format(rs("RefDate"), "dd/MM/yyyy"))
    End If
    PartyCode = rs("VendorCode")
    rs.MoveNext
    Cnt = Cnt + 1
  Loop
  Call EndofReport(iPageWidth, Val(Cnt))
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Function

Private Sub Design_SprCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprCust


SprCust.Click +=  new EventHandler(SprCust_Click)
SprCust.KeyUp +=  new EventHandler(SprCust_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_SprCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\PMI\rptChequeEntry.frm", SprCust, "")

End Sub
