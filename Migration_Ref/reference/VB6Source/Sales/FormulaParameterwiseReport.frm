VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptFormulaParameterwiseReport 
   Caption         =   "Formula Parameterwise Report"
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame fraList 
      Height          =   630
      Left            =   3390
      TabIndex        =   8
      Top             =   1740
      Width           =   4935
      Begin VB.TextBox txtFormulaParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   210
         Width           =   2805
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Formula Parameter Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   165
         TabIndex        =   9
         Top             =   255
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   4597
      TabIndex        =   5
      Top             =   945
      Width           =   2520
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   196214785
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   0
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   196214785
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1380
         TabIndex        =   7
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   6
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   5730
      TabIndex        =   4
      Top             =   2445
      Width           =   1300
   End
   Begin VB.CommandButton cbView 
      Caption         =   "&View"
      Height          =   400
      Left            =   4440
      TabIndex        =   3
      Top             =   2445
      Width           =   1300
   End
End
Attribute VB_Name = "rptFormulaParameterwiseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cbView_Click()

  Try
  Me.MousePointer = vbHourglass
  
  If Trim(txtFormulaParameter) = "" Then
    sShowMessage "Please select the Formula Parameter"
    txtFormulaParameter.SetFocus
    GoTo ExitHere
  End If
  Set frmCurrent = Me
  DoEvents
  iLineNo = 1
  iPageNo = 1
  If sPrintCurrent = False Then
 GoTo ExitHere
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Public Sub sPageHeader()
  Try

'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(10, "Invoice No") & Space(1) & LAlign(12, "Invoice Date") & Space(1) & RAlign(10, "Amount")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim Sql As String
  Dim i As Long
  Dim tamt As Double
  Dim rsDetails As AceADODB.Recordset
  
  Try
  
  Set rsDetails = New AceADODB.Recordset

  Sql = " select a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.Code,a.Amount from"
  Sql = Sql + " SAL_INVOICE_CALCULATION a where a.Code=" & Val(txtFormulaParameter.Tag) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.finyear = '" & gFinyear & "' and a.InvoiceDate between"
  Sql = Sql + " convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " order by a.InvoiceDate,a.InvoiceType,a.InvoiceNo"
  If RSOpen(rsDetails, Sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sPrintCurrent = False
    GoTo ExitHere
  End If
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  fOpenPort Me
  
  tamt = 0
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  
  sSendToPort ""
  gStrReportHeading = "Formula Parameterwise Report between " & dpfromdt & " and " & dpTodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort CAlign(iPageWidth, txtFormulaParameter)
  sSendToPort ""
  sPageHeader
  i = 1
  Do While Not rsDetails.EOF
    sSendToPort LAlign(10, rsDetails("InvoiceType") & rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(10, Format(rsDetails("Amount"), "0.00"))
    tamt = tamt + rsDetails("Amount")
    i = i + 1
    rsDetails.MoveNext
  Loop
  sSendToPort ""
  sSendToPort LAlign(10, "Total") & Space(14) & RAlign(10, Format(tamt, "0.00"))
  sSendToPort ""
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub Form_Load()
  
  Try

  CurrDate Me
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub txtFormulaParameter_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select distinct (select b.Description from COM_FORMULA_PARAMETER b"
  Sql = Sql + " where b.Code=a.Code and b.Branchid=a.Branchid and b.Compid=a.Compid) Description,"
  Sql = Sql + " a.Code from SAL_INVOICE_CALCULATION a where a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula Parameter", "Parameter Name") = True Then
    txtFormulaParameter.Text = getvalue(1)
    txtFormulaParameter.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormulaParameter_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFormulaParameter_DblClick
End If
End Sub

