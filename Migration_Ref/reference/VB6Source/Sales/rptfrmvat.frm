VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptfrmvat 
   Caption         =   "VAT Report"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   6180
      Width           =   1260
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4500
      TabIndex        =   6
      Top             =   6165
      Width           =   1260
   End
   Begin VB.Frame fraRptType 
      Caption         =   "Select"
      Height          =   765
      Left            =   4575
      TabIndex        =   4
      Top             =   1755
      Width           =   2385
      Begin VB.ComboBox cmbRptType 
         Height          =   315
         ItemData        =   "rptfrmvat.frx":0000
         Left            =   135
         List            =   "rptfrmvat.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   2130
      End
   End
   Begin VB.Frame fraRpt 
      Caption         =   "Restrict"
      Height          =   765
      Left            =   4125
      TabIndex        =   0
      Top             =   2520
      Width           =   3270
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   195887107
         CurrentDate     =   38052
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   345
         Left            =   1830
         TabIndex        =   2
         Top             =   285
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   195887107
         CurrentDate     =   38052
      End
      Begin VB.Label lblResDayTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   240
         Left            =   1515
         TabIndex        =   3
         Top             =   315
         Width           =   210
      End
   End
Begin AceSpread sprLedgerDet
      Height          =   2850
      Left            =   3165
      TabIndex        =   8
      Top             =   3285
      Width           =   4935
      _Version        =   458752
      _ExtentX        =   8705
      _ExtentY        =   5027
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   1
      OperationMode   =   4
      ScrollBars      =   2
   End
End
Attribute VB_Name = "rptfrmvat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum eSpr
  anLrCode = 1
  anParticulars = 2
  anVatType = 3
  anCompCode = 4
  anCompName = 5
End Enum
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
 
  Try
  Me.MousePointer = vbHourglass

  GoTo UnloadHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Exit Sub
UnloadHere:
  Unload Me
End Try
End Sub

Private Sub cmdView_Click()
Try
  Me.MousePointer = vbHourglass
  
  If Format(dtpFrom, "yyyy-MM-dd") > Format(dtpTo, "yyyy-MM-dd") Then
    MsgBox "From Date Should be <= To date.", vbOKOnly, App.Title
    GoTo ExitHere
  End If

  
  If ExtractRegister = False Then GoTo ExitHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Function ExtractRegister() As Boolean
''  Dim Sql As String
''  Dim rs As ADODB.Recordset
''  Dim rs1 As ADODB.Recordset
''  Dim lcnt As Integer
''  Dim k As Long
''  Dim j As Long
''  Dim sallocal As Variant
''  Dim salinter As Variant
''  Dim vat As Variant
''  Dim tot As Long
''  Dim TotAmt As Double
''  Dim rptType As String
''
''  On Error GoTo ErrHandler
''  ExtractRegister = False
''
''  Select Case cmbRptType.ListIndex
''    Case 0 ' Sales
''      rptType = "0K"
''    Case 1 'Purchase
''      rptType = "0B"
''  End Select
''
''  Sql = " Select code,Particulars,round(closingbalance,0)closingbalance From dbo.fn_FAS_Rpt_ExtractRegister('0K', '" & Format(dtpFrom, "yyyy-MM-dd") & "', '" & Format(dtpFrom, "yyyy-MM-dd") & "', '" & Format(dtpTo, "yyyy-MM-dd") & "', '0200')"
''  Sql = Sql & "  Order by SortOrder, Path, Type, Particulars"
''  Set rs = New ADODB.Recordset
''  If RSOpen(rs, Sql) = False Then
''    GoTo ExitHere
''  End If
''
''   If rs.RecordCount = 0 Then
''    sShowMessage "No records Found."
''    GoTo ExitHere
''  End If
''
''  Sql = " Select code,Particulars,round(closingbalance,0)closingbalance From dbo.fn_FAS_Rpt_ExtractRegister('0B', '" & Format(dtpFrom, "yyyy-MM-dd") & "', '" & Format(dtpFrom, "yyyy-MM-dd") & "', '" & Format(dtpTo, "yyyy-MM-dd") & "', '0200')"
''  Sql = Sql & "  Order by SortOrder, Path, Type, Particulars"
''  Set rs1 = New ADODB.Recordset
''  If RSOpen(rs1, Sql) = False Then
''    GoTo ExitHere
''  End If
''
''   If rs1.RecordCount = 0 Then
''    sShowMessage "No records Found."
''    GoTo ExitHere
''  End If
''
''
''
''  rptfrmvat.Show
''
''  rs.MoveLast: rs.MoveFirst
''
''  k = 0
''    For k = 1 To rs.RecordCount
''    With Frmprintvat.sprPrintvat
''
''    .Row = 11: .Col = 4: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 4, 11, Mid(dtpFrom.Value, 1, 1)
''    .Row = 11: .Col = 5: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 5, 11, Mid(dtpFrom.Value, 2, 1)
''    .Row = 11: .Col = 6: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 6, 11, Mid(dtpFrom.Value, 4, 1)
''    .Row = 11: .Col = 7: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 7, 11, Mid(dtpFrom.Value, 5, 1)
''    .Row = 11: .Col = 8: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 8, 11, Mid(dtpFrom.Value, 7, 1)
''    .Row = 11: .Col = 9: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 9, 11, Mid(dtpFrom.Value, 8, 1)
''    .Row = 11: .Col = 10: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 10, 11, Mid(dtpFrom.Value, 9, 1)
''    .Row = 11: .Col = 11: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 11, 11, Mid(dtpFrom.Value, 10, 1)
''   .Row = 11: .Col = 14: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 14, 11, Mid(dtpTo.Value, 1, 1)
''    .Row = 11: .Col = 15: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 15, 11, Mid(dtpTo.Value, 2, 1)
''    .Row = 11: .Col = 16: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 16, 11, Mid(dtpTo.Value, 4, 1)
''     .Row = 11: .Col = 17: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 17, 11, Mid(dtpTo.Value, 5, 1)
''     .Row = 11: .Col = 18: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 18, 11, Mid(dtpTo.Value, 7, 1)
''    .Row = 11: .Col = 19: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 19, 11, Mid(dtpTo.Value, 8, 1)
''    .Row = 11: .Col = 20: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 20, 11, Mid(dtpTo.Value, 9, 1)
''    .Row = 11: .Col = 21: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 21, 11, Mid(dtpTo.Value, 10, 1)
''
''    If rs("code") = "5" Then
''    sallocal = rs("closingbalance")
''     .Row = 19: .Col = 4: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 4, 19, Mid((StrReverse(rs("closingbalance"))), 11, 1)
''     .Row = 19: .Col = 5: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 5, 19, Mid(StrReverse(rs("closingbalance")), 10, 1)
''     .Row = 19: .Col = 6: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 6, 19, Mid(StrReverse(rs("closingbalance")), 9, 1)
''     .Row = 19: .Col = 7: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 7, 19, Mid(StrReverse(rs("closingbalance")), 8, 1)
''     .Row = 19: .Col = 8: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 8, 19, Mid(StrReverse(rs("closingbalance")), 7, 1)
''     .Row = 19: .Col = 9: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 9, 19, Mid(StrReverse(rs("closingbalance")), 6, 1)
''     .Row = 19: .Col = 10: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 10, 19, Mid(StrReverse(rs("closingbalance")), 5, 1)
''     .Row = 19: .Col = 11: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 11, 19, Mid(StrReverse(rs("closingbalance")), 4, 1)
''     .Row = 19: .Col = 12: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 12, 19, Mid(StrReverse(rs("closingbalance")), 3, 1)
''     .Row = 19: .Col = 13: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 13, 19, Mid(StrReverse(rs("closingbalance")), 2, 1)
''     .Row = 19: .Col = 14: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 14, 19, Mid(StrReverse(rs("closingbalance")), 1, 1)
''    End If
''    If rs("code") = "353" Then
''    vat = rs("closingbalance")
''     .Row = 19: .Col = 15: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 15, 19, Mid(StrReverse(rs("closingbalance")), 10, 1)
''     .Row = 19: .Col = 16: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 16, 19, Mid(StrReverse(rs("closingbalance")), 9, 1)
''     .Row = 19: .Col = 17: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 17, 19, Mid(StrReverse(rs("closingbalance")), 8, 1)
''     .Row = 19: .Col = 18: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 18, 19, Mid(StrReverse(rs("closingbalance")), 7, 1)
''     .Row = 19: .Col = 19: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 19, 19, Mid(StrReverse(rs("closingbalance")), 6, 1)
''   .Row = 19: .Col = 20: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 20, 19, Mid(StrReverse(rs("closingbalance")), 5, 1)
''    .Row = 19: .Col = 21: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 21, 19, Mid(StrReverse(rs("closingbalance")), 4, 1)
''   .Row = 19: .Col = 22: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 22, 19, Mid(StrReverse(rs("closingbalance")), 3, 1)
''   .Row = 19: .Col = 23: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 23, 19, Mid(StrReverse(rs("closingbalance")), 2, 1)
''   .Row = 19: .Col = 24: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 24, 19, Mid(StrReverse(rs("closingbalance")), 1, 1)
''End If
''   If rs("code") = "6" Then
''   salinter = rs("closingbalance")
''     .Row = 27: .Col = 4: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 4, 27, Mid(StrReverse(rs("closingbalance")), 11, 1)
''     .Row = 27: .Col = 5: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 5, 27, Mid(StrReverse(rs("closingbalance")), 10, 1)
''     .Row = 27: .Col = 6: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 6, 27, Mid(StrReverse(rs("closingbalance")), 9, 1)
''     .Row = 27: .Col = 7: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 7, 27, Mid(StrReverse(rs("closingbalance")), 8, 1)
''     .Row = 27: .Col = 8: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 8, 27, Mid(StrReverse(rs("closingbalance")), 7, 1)
''     .Row = 27: .Col = 9: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 9, 27, Mid(StrReverse(rs("closingbalance")), 6, 1)
''     .Row = 27: .Col = 10: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 10, 27, Mid(StrReverse(rs("closingbalance")), 5, 1)
''     .Row = 27: .Col = 11: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 11, 27, Mid(StrReverse(rs("closingbalance")), 4, 1)
''     .Row = 27: .Col = 12: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 12, 27, Mid(StrReverse(rs("closingbalance")), 3, 1)
''     .Row = 27: .Col = 13: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 13, 27, Mid(StrReverse(rs("closingbalance")), 2, 1)
''     .Row = 27: .Col = 14: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 14, 27, Mid(StrReverse(rs("closingbalance")), 1, 1)
''    End If
''   tot = sallocal + salinter
''   ''''sal total
''    .Row = 34: .Col = 4: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 4, 34, Mid(StrReverse(tot), 11, 1)
''     .Row = 34: .Col = 5: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 5, 34, Mid(StrReverse(tot), 10, 1)
''     .Row = 34: .Col = 6: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 6, 34, Mid(StrReverse(tot), 9, 1)
''     .Row = 34: .Col = 7: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 7, 34, Mid(StrReverse(tot), 8, 1)
''     .Row = 34: .Col = 8: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 8, 34, Mid(StrReverse(tot), 7, 1)
''     .Row = 34: .Col = 9: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 9, 34, Mid(StrReverse(tot), 6, 1)
''     .Row = 34: .Col = 10: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 10, 34, Mid(StrReverse(tot), 5, 1)
''     .Row = 34: .Col = 11: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 11, 34, Mid(StrReverse(tot), 4, 1)
''     .Row = 34: .Col = 12: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 12, 34, Mid(StrReverse(tot), 3, 1)
''     .Row = 34: .Col = 13: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 13, 34, Mid(StrReverse(tot), 2, 1)
''     .Row = 34: .Col = 14: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 14, 34, Mid(StrReverse(tot), 1, 1)
''    '''''''''''
''    ' vat total
''    .Row = 34: .Col = 15: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 15, 34, Mid(StrReverse(vat), 10, 1)
''     .Row = 34: .Col = 16: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 16, 34, Mid(StrReverse(vat), 9, 1)
''     .Row = 34: .Col = 17: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 17, 34, Mid(StrReverse(vat), 8, 1)
''     .Row = 34: .Col = 18: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 18, 34, Mid(StrReverse(vat), 7, 1)
''     .Row = 34: .Col = 19: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 19, 34, Mid(StrReverse(vat), 6, 1)
''   .Row = 34: .Col = 20: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 20, 34, Mid(StrReverse(vat), 5, 1)
''    .Row = 34: .Col = 21: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 21, 34, Mid(StrReverse(vat), 4, 1)
''   .Row = 34: .Col = 22: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 22, 34, Mid(StrReverse(vat), 3, 1)
''   .Row = 34: .Col = 23: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 23, 34, Mid(StrReverse(vat), 2, 1)
''   .Row = 34: .Col = 24: .TypeHAlign = TypeHAlignLeft
''    Frmprintvat.sprPrintvat.SetText 24, 34, Mid(StrReverse(vat), 1, 1)
''   ''''''''''''''
''
''  rs.MoveNext
''
''
''  End With
''  Next k
''
''  k = 0
''    For k = 1 To rs1.RecordCount
''    With Frmprintvat.sprPrintvat
''
''
''
''
''
''
''
''
''
''    rs1.MoveNext
''
''
''  End With
''  Next k
''
''  Frmprintvat.sprPrintvat.MaxRows = Frmprintvat.sprPrintvat.DataRowCnt + 1
''  Screen.MousePointer = 1
''  totalpages = pageno
''  ExtractRegister = True
''
''ExitHere:
''  On Error GoTo 0
''  Set rs = Nothing
''  Exit Function
''ErrHandler:
''  Call ErrMsgShow
''  Resume ExitHere

  Dim fromdate As Date 'Start Date
  Dim toDate As Date 'End Date
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rs1 As AceADODB.Recordset
  Dim sRow As Long
  Dim StartCol As Long
  Dim CapitalGoods(2) As Double
  Dim Total(5) As Double
  Dim RptType As String
  Dim ListStr As String
Dim    vat(2) As Double '0=1%
Dim 1=4% As 
Dim 2=12.5%

Dim    Goods(2) As Double '0=1%
Dim 1=4% As 
Dim 2=12.5%

  
  Try
  Me.MousePointer = vbHourglass
  
'  cmdView.Enabled = False
  vat(0) = 0
vat(1) = 0
vat(2) = 0

  Goods(0) = 0
Goods(1) = 0
Goods(2) = 0:
    '.SetText 3, 4, CAlign(70, txtdealer)
    Frmprintvat.sprPrintvat.SetText 3, 6, txtAuthority
    Frmprintvat.sprPrintvat.SetText 7, 6, txtTinNo
    Frmprintvat.sprPrintvat.SetText 2, 66, txtPlace
    Frmprintvat.sprPrintvat.SetText 2, 68, Format(DtprptDate, "dd.MM.yyyy")
    Frmprintvat.sprPrintvat.SetText 8, 66, "For " & TxtCompName
    Frmprintvat.sprPrintvat.SetText 8, 70, txtName  
  ListStr = Spread_Get_ListAll(sprLedgerDet, eSpr.anLrCode)
  RptType = "0B"
  
  If Len(ListStr) > 1 Then
    Sql = "select Isnull(sum(dbo.fn_FAS_LedgerOpBalance (a.LrCode,'" & Format(DtpMonthYear, "yyyy-MM") & "-01'" & ",0,a.CompCode)),0)as TotalAmt from"
    Sql = Sql & " FAS_VatReport_Option a where CompCode like '" & gCompID & "%' AND Type in ('IT','IG')"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
      Frmprintvat.sprPrintvat.SetText 5, 8, Format(rs("TotalAmt"), "0.00")
    Else
      Frmprintvat.sprPrintvat.SetText 5, 8, Format(0, "0.00")
    End If
  End If
  
  RptType = "0B"

  ListStr = Spread_Get_ListAll(sprLedgerDet, eSpr.anLrCode)
  
  If Len(ListStr) > 1 Then
    Sql = " Select X.VatType, sum(X.Amount)Amount,sum(x.TotalAmt)TotalAmt,x.Branchid,x.Compid"
    Sql = Sql & " FROM (Select a.CompCode, a.DocNo, a.VrRoute, b.LrCode, c.LedgerName, (b.DrAmt+b.CrAmt) Amount,"
    Sql = Sql & " isnull((select sum(DrAmt-CrAmt) from FAS_VoucherTranSub h where Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where h.DocNo = m.DocNo And h.DocDate = m.DocDate And h.VrRoute = m.VrRoute"
    Sql = Sql & " and h.CompCode=m.CompCode and m.LrCode in " & ListStr & ")"
    Sql = Sql & " and (Exists (Select * from Fas_VATReport_Option n where n.LrCode=h.LrCode and n.CompCode=h.CompCode and n.Type='VG')or h.VTSNo=2)"
    Sql = Sql & " and a.DocNo=h.DocNo and a.DocDate=h.DocDate and a.VrRoute=h.VrRoute and a.CompCode=h.CompCode),0) as TotalAmt,"
    Sql = Sql & " Right(a.CompCode, 2) as Branchid, Left(a.CompCode, 2) As CompID,e.VatValue as VatType From FAS_VoucherTranMast a"
    Sql = Sql & " Left Outer join FAS_VoucherTranSub b On a.DocNo = b.DocNo and a.docDate = b.DocDate and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode"
    Sql = Sql & " Left Outer Join FAS_VoucherTranRef b1 On a.DocNo = b1.DocNo and a.docDate = b1.DocDate and a.VrRoute = b1.VrRoute and a.CompCode = b1.CompCode"
    Sql = Sql & " Left Outer Join FAS_BillDetails b2 On a.DocNo = b2.DocNo and a.docDate = b2.DocDate and a.VrRoute = b2.VrRoute and a.CompCode = b2.CompCode and b2.Type='New Ref'"
    Sql = Sql & " Left Outer join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode"
    Sql = Sql & " Left Outer join FAS_LrAddress d On d.LrCode = c.LedgerCode and d.CompCode = c.CompCode"
    Sql = Sql & " Left Outer Join Fas_VatReport_Option e On e.Lrcode=b.LrCode and e.CompCode=b.CompCode"
    Sql = Sql & " Where Month(a.DocDate)='" & Format(dtpFrom, "MM") & "' and Year(a.DocDate)='" & Format(dtpFrom, "yyyy") & "'"
    Sql = Sql & " and a.VrRoute Like '" & Trim$(RptType) & "%' and left(A.CompCode, 2)= '" & gCompID & "'"
    Sql = Sql & " and Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where a.DocNo=m.DocNo  and a.DocDate=m.DocDate and a.VrRoute=m.VrRoute and a.CompCode=m.CompCode"
    Sql = Sql & " and b.LrCode in " & ListStr & ")"
    Sql = Sql & " )X  group by X.VatType,x.Branchid,x.Compid Order by X.VatType,x.Branchid,x.Compid"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    sRow = 3
    Frmprintvat.sprPrintvat.SetText 8, 6, "Month : " & Format(DtpMonthYear, "MMM")

    Frmprintvat.sprPrintvat.SetText 9, 6, "Year  : " & Format(DtpMonthYear, "yyyy")
    Do While Not rs.EOF
        Select Case rs("BranchID")
          Case "01"
            Select Case rs("VatType")
              Case "Vat 1%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 11, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(0) = rs("Amount")
Goods(0) = rs("TotalAmt")
              Case "Vat 4%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 14, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(1) = rs("Amount")
Goods(1) = rs("TotalAmt")
              Case "Vat 12.5%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 17, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(2) = rs("Amount")
Goods(2) = rs("TotalAmt")
            End Select
          Case "02"
            Select Case rs("VatType")
              Case "Vat 1%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 12, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(0) = vat(0) + rs("Amount")
Goods(0) = Goods(0) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mPurval, 12, Format(Goods(0), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mInputTax, 12, Format(vat(0), "0.00")
              Case "Vat 4%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 15, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(1) = vat(1) + rs("Amount")
Goods(1) = Goods(1) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mPurval, 15, Format(Goods(1), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mInputTax, 15, Format(vat(1), "0.00")
              Case "Vat 12.5%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mHeading, 18, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(2) = vat(2) + rs("Amount")
Goods(2) = Goods(2) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mPurval, 18, Format(Goods(2), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mInputTax, 18, Format(vat(2), "0.00")
            End Select
        End Select      rs.MoveNext
    Loop
    
   ' Frmprintvat.sprPrintvat.SetText eVATDetail.mPurval, 20, Format(Val(Val(SpVal(Frmprintvat.sprPrintvat, 12, 3)) + Val(SpVal(Frmprintvat.sprPrintvat, 15, 3)) + Val(SpVal(Frmprintvat.sprPrintvat, 18, 3)) + Val(SpVal(Frmprintvat.sprPrintvat, 19, 3))), "0.00")
    'Frmprintvat.sprPrintvat.SetText eVATDetail.mPurval, 22, Format(Val(SpVal(Frmprintvat.sprPrintvat, 20, 3)) - Val(SpVal(Frmprintvat.sprPrintvat, 21, 3)), "0.00")
    
    'Frmprintvat.sprPrintvat.SetText eVATDetail.mInputTax, 20, Format(Val(Val(SpVal(Frmprintvat.sprPrintvat, 12, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 15, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 18, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 19, 5))), "0.00")
    'Frmprintvat.sprPrintvat.SetText eVATDetail.mInputTax, 22, Format(Val(SpVal(Frmprintvat.sprPrintvat, 20, 5)) - Val(SpVal(Frmprintvat.sprPrintvat, 21, 5)), "0.00")
  
  ' Salse
    RptType = "0K"
    
    Sql = " SELECT X.VatType,sum(X.Amount)Amount,sum(x.TotalAmt)TotalAmt,x.Branchid,x.Compid FROM (Select b.LrCode, c.LedgerName, (b.DrAmt+b.CrAmt) Amount," 'a.Amount-(b.DrAmt+b.CrAmt) as TotalAmt,"
    Sql = Sql & " isnull((select sum(DrAmt+CrAmt) from FAS_VoucherTranSub h where Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where h.DocNo = m.DocNo And h.DocDate = m.DocDate And h.VrRoute = m.VrRoute"
    Sql = Sql & " and h.CompCode=m.CompCode and m.LrCode in " & ListStr & ")"
    Sql = Sql & " and (Exists (Select * from Fas_VATReport_Option n where n.LrCode=h.LrCode and n.CompCode=h.CompCode and n.Type='VG')or h.VTSNo=2)"
    Sql = Sql & " and a.DocNo=h.DocNo and a.DocDate=h.DocDate and a.VrRoute=h.VrRoute and a.CompCode=h.CompCode),0) as TotalAmt,"
    Sql = Sql & " Right(a.CompCode,2) as Branchid,left(A.CompCode, 2) As Compid,"
    Sql = Sql & " (Select VatValue from Fas_VatReport_Option where LrCode=b.LrCode and CompCode=b.CompCode and Type='VT')VatType From FAS_VoucherTranMast a"
    Sql = Sql & " Left Outer join FAS_VoucherTranSub b On a.DocNo = b.DocNo and a.docDate = b.DocDate and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode"
    Sql = Sql & " Left Outer Join FAS_VoucherTranRef b1 On a.DocNo = b1.DocNo and a.docDate = b1.DocDate and a.VrRoute = b1.VrRoute and a.CompCode = b1.CompCode"
    Sql = Sql & " Left Outer Join FAS_BillDetails b2 On a.DocNo = b2.DocNo and a.docDate = b2.DocDate and a.VrRoute = b2.VrRoute and a.CompCode = b2.CompCode and b2.Type='New Ref'"
    Sql = Sql & " Left Outer join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode"
    Sql = Sql & " Left Outer join FAS_LrAddress d On d.LrCode = c.LedgerCode and d.CompCode = c.CompCode"
    Sql = Sql & " Where Month(a.DocDate)='" & Format(DtpMonthYear, "MM") & "' and Year(a.DocDate)='" & Format(DtpMonthYear, "yyyy") & "'"
    Sql = Sql & " and a.VrRoute Like '" & Trim$(RptType) & "%' and left(A.CompCode, 2)= '" & gCompID & "'"
    Sql = Sql & " and Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where a.DocNo=m.DocNo  and a.DocDate=m.DocDate and a.VrRoute=m.VrRoute and a.CompCode=m.CompCode"
    Sql = Sql & " and b.LrCode in " & ListStr & ")"
    Sql = Sql & " )X   group by X.VatType,x.Branchid,x.Compid Order by X.VatType,x.Branchid,x.Compid"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    sRow = 3
    Do While Not rs.EOF
        Select Case rs("BranchID")
          Case "01"
            Select Case rs("VatType")
              Case "Vat 1%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 11, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(0) = rs("Amount")
Goods(0) = rs("TotalAmt")
              Case "Vat 4%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 14, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(1) = rs("Amount")
Goods(1) = rs("TotalAmt")
              Case "Vat 12.5%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 17, "       Unit I  = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit I
                vat(2) = rs("Amount")
Goods(2) = rs("TotalAmt")
            End Select
          Case "02"
            Select Case rs("VatType")
              Case "Vat 1%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 12, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(0) = vat(0) + rs("Amount")
Goods(0) = Goods(0) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalTurnOver, 12, Format(Goods(0), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mVATDue, 12, Format(vat(0), "0.00")
              Case "Vat 4%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 15, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(1) = vat(1) + rs("Amount")
Goods(1) = Goods(1) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalTurnOver, 15, Format(Goods(1), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mVATDue, 15, Format(vat(1), "0.00")
              Case "Vat 12.5%"
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalHeading, 18, "       Unit II = " & Format(rs("TotalAmt"), "0.00") 'Goods Taxable at 1% for Unit II
                vat(2) = vat(2) + rs("Amount")
Goods(2) = Goods(2) + rs("TotalAmt")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mSalTurnOver, 18, Format(Goods(2), "0.00")
                Frmprintvat.sprPrintvat.SetText eVATDetail.mVATDue, 18, Format(vat(2), "0.00")
            End Select
        End Select
      rs.MoveNext
    Loop
  
   
  End If
  
  ListStr = Spread_Get_ListAll(sprLedgerDet, eSpr.anLrCode)
  RptType = "0B"
  
  If Len(ListStr) > 1 Then
    Sql = " Select sum(X.Amount)Amount,sum(x.TotalAmt)TotalAmt,x.Compid"
    Sql = Sql & " FROM (Select a.CompCode, a.DocNo, a.VrRoute, b.LrCode, c.LedgerName, (b.DrAmt+b.CrAmt) Amount,"
    Sql = Sql & " isnull((select sum(DrAmt-CrAmt) from FAS_VoucherTranSub h where Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where h.DocNo = m.DocNo And h.DocDate = m.DocDate And h.VrRoute = m.VrRoute"
    Sql = Sql & " and h.CompCode=m.CompCode and m.LrCode in " & ListStr & ")"
    Sql = Sql & " and (Exists (Select * from Fas_VATReport_Option n where n.LrCode=h.LrCode and n.CompCode=h.CompCode and n.Type='CG')or h.VTSNo=2)"
    Sql = Sql & " and a.DocNo=h.DocNo and a.DocDate=h.DocDate and a.VrRoute=h.VrRoute and a.CompCode=h.CompCode),0) as TotalAmt,"
    Sql = Sql & " Right(a.CompCode, 2) as Branchid, Left(a.CompCode, 2) As CompID,e.VatValue as VatType From FAS_VoucherTranMast a"
    Sql = Sql & " Left Outer join FAS_VoucherTranSub b On a.DocNo = b.DocNo and a.docDate = b.DocDate and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode"
    Sql = Sql & " Left Outer Join FAS_VoucherTranRef b1 On a.DocNo = b1.DocNo and a.docDate = b1.DocDate and a.VrRoute = b1.VrRoute and a.CompCode = b1.CompCode"
    Sql = Sql & " Left Outer Join FAS_BillDetails b2 On a.DocNo = b2.DocNo and a.docDate = b2.DocDate and a.VrRoute = b2.VrRoute and a.CompCode = b2.CompCode and b2.Type='New Ref'"
    Sql = Sql & " Left Outer join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode"
    Sql = Sql & " Left Outer join FAS_LrAddress d On d.LrCode = c.LedgerCode and d.CompCode = c.CompCode"
    Sql = Sql & " Left Outer Join Fas_VatReport_Option e On e.Lrcode=b.LrCode and e.CompCode=b.CompCode"
    Sql = Sql & " Where Month(a.DocDate)='" & Format(DtpMonthYear, "MM") & "' and Year(a.DocDate)='" & Format(DtpMonthYear, "yyyy") & "'"
    Sql = Sql & " and a.VrRoute Like '" & Trim$(RptType) & "%' and left(A.CompCode, 2)= '" & gCompID & "'"
    Sql = Sql & " and Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where a.DocNo=m.DocNo  and a.DocDate=m.DocDate and a.VrRoute=m.VrRoute and a.CompCode=m.CompCode"
    Sql = Sql & " and b.LrCode in " & ListStr & ")"
    Sql = Sql & " )X  group by x.Compid Order by x.Compid"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
      Frmprintvat.sprPrintvat.SetText 5, 30, Format(rs("Amount"), "0.00")
      Frmprintvat.sprPrintvat.SetText 3, 30, Format(rs("TotalAmt"), "0.00")
      Frmprintvat.sprPrintvat.SetText 3, 31, Val(SpVal(Frmprintvat.sprPrintvat, 29, 3)) + Val(SpVal(Frmprintvat.sprPrintvat, 30, 3))
      Frmprintvat.sprPrintvat.SetText 3, 33, Val(SpVal(Frmprintvat.sprPrintvat, 31, 3)) - Val(SpVal(Frmprintvat.sprPrintvat, 32, 3))
      Frmprintvat.sprPrintvat.SetText 5, 31, Val(SpVal(Frmprintvat.sprPrintvat, 29, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 30, 5))
      Frmprintvat.sprPrintvat.SetText 5, 33, Val(SpVal(Frmprintvat.sprPrintvat, 31, 5)) - Val(SpVal(Frmprintvat.sprPrintvat, 32, 5))
    End If
      
    Frmprintvat.sprPrintvat.SetText 5, 36, Val(SpVal(Frmprintvat.sprPrintvat, 8, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 22, 5)) + Val(SpVal(Frmprintvat.sprPrintvat, 33, 5))
  End If
  
  ListStr = Spread_Get_ListAll(sprLedgerDet, eSpr.anLrCode)
  RptType = "0K"
  
  If Len(ListStr) > 1 Then
    Sql = " Select sum(X.Amount)Amount,sum(x.TotalAmt)TotalAmt,x.Compid"
    Sql = Sql & " FROM (Select a.CompCode, a.DocNo, a.VrRoute, b.LrCode, c.LedgerName, (b.DrAmt+b.CrAmt) Amount,"
    Sql = Sql & " isnull((select sum(DrAmt-CrAmt) from FAS_VoucherTranSub h where Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where h.DocNo = m.DocNo And h.DocDate = m.DocDate And h.VrRoute = m.VrRoute"
    Sql = Sql & " and h.CompCode=m.CompCode and m.LrCode in " & ListStr & ")"
    Sql = Sql & " and (Exists (Select * from Fas_VATReport_Option n where n.LrCode=h.LrCode and n.CompCode=h.CompCode and n.Type='TG')or h.VTSNo=2)"
    Sql = Sql & " and a.DocNo=h.DocNo and a.DocDate=h.DocDate and a.VrRoute=h.VrRoute and a.CompCode=h.CompCode),0) as TotalAmt,"
    Sql = Sql & " Right(a.CompCode, 2) as Branchid, Left(a.CompCode, 2) As CompID,e.VatValue as VatType From FAS_VoucherTranMast a"
    Sql = Sql & " Left Outer join FAS_VoucherTranSub b On a.DocNo = b.DocNo and a.docDate = b.DocDate and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode"
    Sql = Sql & " Left Outer Join FAS_VoucherTranRef b1 On a.DocNo = b1.DocNo and a.docDate = b1.DocDate and a.VrRoute = b1.VrRoute and a.CompCode = b1.CompCode"
    Sql = Sql & " Left Outer Join FAS_BillDetails b2 On a.DocNo = b2.DocNo and a.docDate = b2.DocDate and a.VrRoute = b2.VrRoute and a.CompCode = b2.CompCode and b2.Type='New Ref'"
    Sql = Sql & " Left Outer join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode"
    Sql = Sql & " Left Outer join FAS_LrAddress d On d.LrCode = c.LedgerCode and d.CompCode = c.CompCode"
    Sql = Sql & " Left Outer Join Fas_VatReport_Option e On e.Lrcode=b.LrCode and e.CompCode=b.CompCode and e.Type='CS'"
    Sql = Sql & " Where Month(a.DocDate)='" & Format(DtpMonthYear, "MM") & "' and Year(a.DocDate)='" & Format(DtpMonthYear, "yyyy") & "'"
    Sql = Sql & " and a.VrRoute Like '" & Trim$(RptType) & "%' and left(A.CompCode, 2)= '" & gCompID & "'"
    Sql = Sql & " and Exists (Select * from FAS_VoucherTranSub m"
    Sql = Sql & " Where a.DocNo=m.DocNo  and a.DocDate=m.DocDate and a.VrRoute=m.VrRoute and a.CompCode=m.CompCode"
    Sql = Sql & " and b.LrCode in " & ListStr & ")"
    Sql = Sql & " )X group by x.Compid Order by x.Compid"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
  End If
  
  If Not rs.EOF Then
    Frmprintvat.sprPrintvat.SetText 6, 57, Format(rs("Amount"), "0.00")
  End If
  
  Frmprintvat.sprPrintvat.SetText 9, 52, Val(SpVal(Frmprintvat.sprPrintvat, 36, 5)) - Val(SpVal(Frmprintvat.sprPrintvat, 38, 10))
  Frmprintvat.sprPrintvat.SetText 9, 59, Val(SpVal(Frmprintvat.sprPrintvat, 55, 6)) + Val(SpVal(Frmprintvat.sprPrintvat, 57, 6))
  Frmprintvat.sprPrintvat.SetText 9, 60, Val(SpVal(Frmprintvat.sprPrintvat, 52, 9)) - Val(SpVal(Frmprintvat.sprPrintvat, 59, 9))
  

 
Catch ex As Exception
  Call ErrMsgShow(ex)
'  Resume ExitHere
Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
  
End Function

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
Try
  
  

  cmbRptType.ListIndex = 0
  dtpFrom.Value = Format(ServerDate, "dd-MM-yyyy")
  dtpTo.Value = dtpFrom.Value
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprLedgerDet_DblClick(sender as object, e as DblClickEventArgs)
Dim Sql As String
  Dim rCnt As Integer
    
  Dim RptType As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If Col = 2 Then
      Select Case cmbRptType.ListIndex
        Case 0 ' Sales
          RptType = "0K%"
        Case 1 'Purchase
          RptType = "0B%"
      End Select
    
      Sql = " Select distinct a.LrCode, b.LedgerName, a.VrRoute,b.CompCode "
      Sql = Sql & " From FAS_VoucherTransub a"
      Sql = Sql & " Left Outer join FAS_LedgerMaster b"
      Sql = Sql & " On a.LrCode = b.LedgerCode and a.CompCode = b.CompCode"
      Sql = Sql & " Where a.VrRoute like '" & RptType & "'"
      Sql = Sql & " and Not (b.LrRoute like '0206%' or LrRoute like '0303%' or LrRoute like '04%')"
      Sql = Sql & " and a.CompCode = '0200' "
      Sql = Sql & " and a.DocDate between '" & Format(dtpFrom, "yyyy-MM-dd") & "' and '" & Format(dtpTo, "yyyy-MM-dd") & "'"
      If FetchValue_New(False, db.ConnectionString, Sql, "0101", "List of Ledgers.", "Ledger Name", "CompCode") = True Then
        sprLedgerDet.SetText eSpr.anLrCode, sprLedgerDet.MaxRows, getvalue(1)
        sprLedgerDet.SetText eSpr.anCompCode, sprLedgerDet.MaxRows, getvalue(4)
        If Spread_Clear_Duplicate(sprLedgerDet, sprLedgerDet.MaxRows, eSpr.anLrCode, eSpr.anCompCode) = False Then
          sprLedgerDet.SetText eSpr.anParticulars, sprLedgerDet.MaxRows, getvalue(2)
          sprLedgerDet.SetText eSpr.anCompCode, sprLedgerDet.MaxRows, getvalue(4)
        End If
        sprLedgerDet.MaxRows = sprLedgerDet.DataRowCnt + 1
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprLedgerDet(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprLedgerDet


sprLedgerDet.DblClick +=  new EventHandler(sprLedgerDet_DblClick)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprLedgerDet("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmvat.frm", sprLedgerDet, "")

End Sub
