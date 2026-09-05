VERSION 5.00
Begin VB.Form rptFormulaParameterReport 
   Caption         =   "Formula Parameter Report"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   5940
      TabIndex        =   0
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   4065
      Width           =   1300
   End
   Begin VB.Frame fraselective 
      Caption         =   "Formula Parameter Master"
      Height          =   1845
      Left            =   3180
      TabIndex        =   5
      Top             =   4545
      Width           =   5550
Begin AceSpread spList
         Height          =   1530
         Left            =   105
         TabIndex        =   6
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
   Begin VB.Frame FraAccount 
      Caption         =   "Select"
      ForeColor       =   &H00404040&
      Height          =   585
      Left            =   5010
      TabIndex        =   2
      Top             =   3420
      Width           =   1890
      Begin VB.OptionButton OptAll 
         Caption         =   "&All"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   195
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton OptSelective 
         Caption         =   "Single"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   855
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   405
      Left            =   4650
      TabIndex        =   1
      ToolTipText     =   "Click to View Report(Alt + R)"
      Top             =   4065
      Width           =   1300
   End
End
Attribute VB_Name = "rptFormulaParameterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Function ParameterMasterListDisplay() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim sql As String
  Dim i As Long
  Dim retVal As Variant
  
  Try
  ParameterMasterListDisplay = False
  
  sql = " Select Code, Description, Case Tag When 'C' then 'Calculated Parameter'"
  sql = sql & " When '%' then 'Percentage' When 'V' Then 'Value' End as Tag, ValPer,"
  sql = sql & " Case GVTag When '1' Then 'GoodsValue' Else 'No GoodsValue' End as GVTag,"
  sql = sql & " Case AGVTag When '1' Then 'Assessable Value' Else 'No Assessable Value'  End as AGVTag"
  sql = sql & " From COM_FORMULA_PARAMETER"
  sql = sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If OptSelective.Value = True Then
    If spList.SelectionCount < 1 Then
 sShowMessage " Please select atlease one item "
Exit Function
End If
      spList.GetText 1, spList.ActiveRow, retVal
      sql = sql & " and Code in ('" & Trim$(retVal) & "')      "
  End If
  sql = sql & " Order by Code"
  Set rsDetails = New AceADODB.Recordset
  If RSOpen(rsDetails, sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    ParameterMasterListDisplay = False
    GoTo ExitHere
  End If
  
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
  gStrReportHeading = "Parameter Master Report"
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
  sPageHeader
  
  i = 1
  Do While Not rsDetails.EOF
    sSendToPort LAlign(17, rsDetails("Description")) & Space(1) & LAlign(20, rsDetails("Tag")) & Space(1) & LAlign(6, Format(rsDetails("ValPer"), "0.00")) & Space(1) & LAlign(13, rsDetails("GVTag")) & Space(1) & LAlign(19, rsDetails("AGVTag"))
    
    If i <> rsDetails.RecordCount Then sSendToPort ""
    i = i + 1
    rsDetails.MoveNext
  Loop
 
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  
  ParameterMasterListDisplay = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sPageHeader()
  Try

  sSendToPort ""
'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  
  sSendToPort LAlign(17, "Parameter Name") & Space(1) & LAlign(20, "Tag") & Space(1) & LAlign(6, "Value") & Space(1) & LAlign(13, "GVTag") & Space(1) & LAlign(19, "AGVTag")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrent() As Boolean
 
  Try
  
  sPrintcurrent = False
 
  If ParameterMasterListDisplay = False Then
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

End Try
End Function
Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdView_Click()
  Dim frmOutput As frmReportViewer
  
  Try
  Me.MousePointer = vbHourglass

  Set frmCurrent = Me

  iLineNo = 1
  iPageNo = 1

  If sPrintcurrent = False Then
 GoTo ExitHere
End If
  sClosePort
  Set frmOutput = New frmReportViewer
  frmOutput.Caption = gStrReportHeading
  frmOutput.txtTotalPages.Text = CStr(iPageNo)
  frmOutput.wbrView.Navigate strFileName
  frmOutput.Tag = strFileName
  gStrReportHeading = ""
  frmOutput.Show
  Me.MousePointer = Default
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  
  Call Spread_Clear(spList)
  spList.MaxRows = 1
  fraselective.Enabled = False
End Sub

Private Sub optSelective_Click()
  Dim sql As String
  Dim rs As AceADODB.Recordset
  
  Dim RCnt As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
  spList.ClearSelection
  fraselective.Enabled = True
  
  sql = " Select Code, Description, Case Tag When 'C' then 'Calculated Parameter'"
  sql = sql & " When '%' then 'Percentage' When 'V' Then 'Value' End as Tag, ValPer,"
  sql = sql & " Case GVTag When '1' Then 'GoodsValue' Else 'No GoodsValue' End as GVTag,"
  sql = sql & " Case AGVTag When '1' Then 'Assessable Value' Else 'No Assessable Value'  End as AGVTag"
  sql = sql & " From COM_FORMULA_PARAMETER"
  sql = sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
    If Not rs.EOF Then
      spList.MaxRows = 1
      spList.MaxRows = rs.RecordCount
      For RCnt = 1 To rs.RecordCount
        spList.SetText 1, RCnt, rs("Code")
        spList.SetText 2, RCnt, rs("Description")
        rs.MoveNext
      Next RCnt
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\FormulaParameterReport.frm", spList, "")

End Sub
