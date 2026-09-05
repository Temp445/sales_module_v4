VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesEnquiryModify 
   Caption         =   "Sales Enquiry Modify"
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4335
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5145
      TabIndex        =   19
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4335
      Width           =   1300
   End
   Begin VB.Frame FraEnq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   1365
      TabIndex        =   21
      Top             =   855
      Width           =   9270
      Begin VB.TextBox txtEnqtype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   1
         ToolTipText     =   "Sales Enquiry No"
         Top             =   330
         Width           =   930
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         MaxLength       =   15
         TabIndex        =   7
         ToolTipText     =   "Enter the Ref. No"
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox txtEnqNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   3
         ToolTipText     =   "Sales Enquiry No"
         Top             =   330
         Width           =   795
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   885
         Width           =   4605
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4725
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Enter the Remarks"
         Top             =   885
         Width           =   3915
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   8670
         TabIndex        =   16
         ToolTipText     =   "Click to OK"
         Top             =   885
         Width           =   495
      End
      Begin VB.ComboBox cmbCustomerType 
         Height          =   330
         ItemData        =   "frmSalesEnquiryModify.frx":0000
         Left            =   5835
         List            =   "frmSalesEnquiryModify.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Select the Customer Type"
         Top             =   315
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtpEnqDate 
         Height          =   315
         Left            =   1905
         TabIndex        =   5
         ToolTipText     =   "Sales Enquiry Date"
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   55902209
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker dtpRefDate 
         Height          =   315
         Left            =   4590
         TabIndex        =   9
         ToolTipText     =   "Enter the Ref. Date"
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55902209
         CurrentDate     =   37258
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   12
         Top             =   705
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   105
         TabIndex        =   0
         Top             =   135
         Width           =   945
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   4725
         TabIndex        =   14
         Top             =   705
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.No."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1095
         TabIndex        =   2
         Top             =   135
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1905
         TabIndex        =   4
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. No"
         Height          =   210
         Left            =   3150
         TabIndex        =   6
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Date"
         Height          =   210
         Left            =   4590
         TabIndex        =   8
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
         Height          =   210
         Left            =   5835
         TabIndex        =   10
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Click to Save the Data"
      Top             =   4335
      Width           =   1300
   End
Begin AceSpread spItem
      Height          =   1995
      Left            =   525
      TabIndex        =   17
      ToolTipText     =   "Enter the Sales Item Details"
      Top             =   2220
      Width           =   10890
      _Version        =   458752
      _ExtentX        =   19209
      _ExtentY        =   3519
      _StockProps     =   64
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   1
      SelectBlockOptions=   0
   End
End
Attribute VB_Name = "frmSalesEnquiryModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim CurrDate As Date

Private Enum eEnq
  mItmType = 1
  mItmCode = 2
  mItmDesc = 3
  mUOM = 4
  mTarVolme = 5
  mVolmeper = 6
  mQty = 7
  mRate = 8
  mAssRate = 9
  mTargtDate = 10
  mSpeci = 11
End Enum
Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim tmp As Variant
  Dim CurDate As Date
  
  Try
  
  If spItem.DataRowCnt = 0 Then
    sShowMessage "Enter atleast One Item"
    Spread_Set_Focus spItem, 1, 1
    Chkbefsave = False
    GoTo ExitHere
  End If
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
    For i = 1 To spItem.DataRowCnt
      spItem.GetText eEnq.mItmCode, i, tmp
      If Trim(tmp) = "" Then
        sShowMessage "Enter the Item Code"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mItmCode
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mItmDesc, i, tmp
      If Trim(tmp) = "" Then
        sShowMessage "Enter the Item Description"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mItmDesc
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mUOM, i, tmp
      If Trim(tmp) = "" Then
        sShowMessage "Select The UOM"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mUOM
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mQty, i, tmp
      If Val(tmp) = 0 Then
        sShowMessage "Quantity cannot be zero"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mQty
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mTargtDate, i, tmp
      If IsDate(tmp) = False Then
        sShowMessage "Enter the Target Date"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mTargtDate
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mVolmeper, i, tmp
      If Trim$(tmp) = "" Then
        sShowMessage "Select the Volume Per"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mVolmeper
        GoTo ExitHere
      End If
      
      spItem.GetText eEnq.mTargtDate, i, tmp
      If Format(tmp, "yyyyMMdd") < Format(CurDate, "yyyyMMdd") Then
        sShowMessage "Target Date shoule be greater than current date"
        spItem.SetFocus
        Spread_Set_Focus spItem, i, eEnq.mTargtDate
        GoTo ExitHere
      End If
    Next i
  
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmbCustomerType_Click()
  Try
  
  If cmbCustomerType = "Existing Customer" Then
    txtCustomer.Locked = True
    Label2.ForeColor = &HFF0000
    txtCustomer = ""
txtCustomer.Tag = ""
  ElseIf cmbCustomerType = "New Customer" Then
    txtCustomer.Locked = False
    Label2.ForeColor = &H80000012
    txtCustomer = ""
txtCustomer.Tag = ""
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  FraEnq.Enabled = True
  spItem.Enabled = False
  cmdSave.Enabled = False
  spItem.MaxRows = 1
  txtEnqNo.SetFocus
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpEnqDate = CurrDate
  dtpRefDate = CurrDate
  cmbCustomerType.ListIndex = 0
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

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim i As Long
  Dim rs As AceADODB.Recordset
  
  Try
  
  If Trim$(txtEnqtype) = "" Then
    sShowMessage "Select the Enquiry type"
    txtEnqtype.SetFocus
    GoTo ExitHere
  End If

  If contrl_valid(txtRefNo, "Enter the Customer Ref. No") = False Then
 GoTo ExitHere
End If
  If contrl_valid(txtCustomer, "Select the Customer") = False Then
 GoTo ExitHere
End If
  txtRefNo = UCase(txtRefNo)
  txtCustomer = UCase(txtCustomer)
  
  Sql = "Select *,(select NoofDecimals from INV_UOM_MASTER where SDesc=a.UOM and Branchid =a.Branchid"
  Sql = Sql & " and compid=a.Compid)NoofDec from SAL_VW_Enquiry_Dtl a where EnquiryType='" & Trim$(txtEnqtype) & "' and EnquiryNo=" & Val(txtEnqNo) & " and EnquiryDate='" & Format(dtpEnqDate, "yyyy-MM-dd") & "' "
  Sql = Sql & " and a.QuotationType='' and BRanchid='" & Trim$(gBranchID) & "' and Compid='" & Trim$(gCompID) & "' "
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If Not rs.EOF Then
      For i = 1 To rs.RecordCount
        spItem.MaxRows = rs.RecordCount + 1
        spItem.SetText eEnq.mItmType, i, IIf(Trim$(rs("ItemType")) = "N", "New Item", "Existing Item")
        Call spItem_ComboSelChange(eEnq.mItmType, i)
        Call Spread_Set_Decimal(spItem, Val(rs("NoofDec")), eEnq.mQty, i)
        Call Spread_Set_Decimal(spItem, Val(rs("NoofDec")), eEnq.mTarVolme, i)
        Call Spread_Lock(spItem, False, eEnq.mQty, i)
        Call Spread_Lock(spItem, False, eEnq.mTarVolme, i)
        spItem.SetText eEnq.mItmCode, i, rs("ItemCode")
        spItem.SetText eEnq.mItmDesc, i, rs("ItemDescription")
        spItem.SetText eEnq.mUOM, i, rs("UOM")
        spItem.SetText eEnq.mQty, i, Val(rs("Qty"))
        spItem.SetText eEnq.mTarVolme, i, Val(rs("Volume"))
        spItem.SetText eEnq.mVolmeper, i, Trim$(rs("VolumePer"))
        spItem.SetText eEnq.mRate, i, rs("Rate")
        spItem.SetText eEnq.mAssRate, i, rs("AssessableRate")
        spItem.SetText eEnq.mTargtDate, i, rs("TargetDate")
        spItem.SetText eEnq.mSpeci, i, Trim$(rs("Specification"))
        rs.MoveNext
      Next i
  End If
  
  FraEnq.Enabled = False
  spItem.Enabled = True
  spItem.SetText 1, 1, "Existing Item"
  Call Spread_Lock(spItem, True, eEnq.mItmCode, 1, eEnq.mUOM)
  spItem.SetFocus
  Spread_Set_Focus spItem, 1, 1
  cmdSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

    Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim retVal
  Dim enqnorec As AceADODB.Recordset
  Dim cmd As AceADODB.Command
  Dim i As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  Set enqnorec = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  If Chkbefsave = False Then GoTo ExitHere
  CurDate = ServerDate
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_INS_ENQUIRY_HDR"
  cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adVarChar, adParamInput, 3, Trim$(txtEnqtype))
  cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(Trim$(txtEnqNo)))
  cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtEnqNo.Tag))
  cmd.Parameters.Append cmd.CreateParameter("RevisionDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ReferenceNo", adVarChar, adParamInput, 30, Trim$(txtRefNo))
  cmd.Parameters.Append cmd.CreateParameter("ReferenceDate", adDBTimeStamp, adParamInput, , Format(dtpRefDate, "yyyy-MM-dd"))
  If Trim(cmbCustomerType) = "Existing Customer" Then
    cmd.Parameters.Append cmd.CreateParameter("CustomerType", adChar, adParamInput, 1, "E")
  Else
    cmd.Parameters.Append cmd.CreateParameter("CustomerType", adChar, adParamInput, 1, "N")
  End If
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtCustomer.Tag))
  cmd.Parameters.Append cmd.CreateParameter("CustomerName", adVarChar, adParamInput, 50, Trim$(txtCustomer.Text))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtRemarks))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute

    cmd.CommandText = "SAL_INS_ENQUIRY_DTL"
    For i = 1 To spItem.DataRowCnt
      spItem.GetText eEnq.mItmCode, i, retVal
      If Trim(retVal) <> "" Then
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adVarChar, adParamInput, 3, Trim$(txtEnqtype))
        cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(Trim$(txtEnqNo)))
        cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
        cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtEnqNo.Tag))
        spItem.GetText eEnq.mItmType, i, retVal
        If Trim(retVal) = "New Item" Then
          cmd.Parameters.Append cmd.CreateParameter("ItemType", adChar, adParamInput, 1, "N")
        Else
          cmd.Parameters.Append cmd.CreateParameter("ItemType", adChar, adParamInput, 1, "E")
        End If
        spItem.GetText eEnq.mItmCode, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        spItem.GetText eEnq.mItmDesc, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemDescription", adVarChar, adParamInput, 50, Trim$(retVal))
        spItem.GetText eEnq.mUOM, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("UOM", adVarChar, adParamInput, 3, Trim$(retVal))
        spItem.GetText eEnq.mQty, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("Qty", adDouble, adParamInput, , Val(retVal))
        spItem.GetText eEnq.mTarVolme, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("Volume", adDouble, adParamInput, , Val(retVal))
        spItem.GetText eEnq.mVolmeper, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("VolumePer", adVarChar, adParamInput, 15, Trim$(retVal))
        spItem.GetText eEnq.mRate, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(retVal))
        spItem.GetText eEnq.mAssRate, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(retVal))
        spItem.GetText eEnq.mTargtDate, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("TargetDate", adDBTimeStamp, adParamInput, , Format(Trim$(retVal), "yyyy-MM-dd"))
        spItem.GetText eEnq.mSpeci, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("Specification", adVarChar, adParamInput, 30, Trim$(retVal))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
    Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Enquiry Number :" & Trim$(txtEnqtype) & Trim$(txtEnqNo) & "  Is Updated"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set enqnorec = Nothing
  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpenqdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)

  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtprefdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  If TypeOf Me.ActiveControl Is fpSpread Then
 GoTo ExitHere
End If
  If KeyAscii = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try

'  Loadenqtype '''Loading all the enquiry types to combo
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpEnqDate = CurrDate
  dtpRefDate = CurrDate
  cmbCustomerType.ListIndex = 0
  spItem.Enabled = False
  cmdSave.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub spItem_ComboSelChange(sender as object, e as ComboSelChangeEvent)
  Dim tmp As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Col = eEnq.mItmType Then
    spItem.Row = Row
spItem.Col = Col
    If Trim(spItem.Text) = "New Item" Then
      Call Spread_Lock(spItem, False, eEnq.mItmCode, Row, eEnq.mUOM, Row)
    Else
      Call Spread_Lock(spItem, True, eEnq.mItmCode, Row, eEnq.mUOM, Row)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

    spItem.Row = Row
spItem.Col = eEnq.mItmType
    If Col = eEnq.mItmCode And Trim(spItem.Text) = "Existing Item" Then
      If cmbCustomerType = "Existing Customer" Then
        Sql = "select a.ItemCode,a.ItemDescription,a.UOMDesc,a.UOMDecimal from MATERIAL_MASTER a"
        Sql = Sql & " where a.Branchid ='" & gBranchID & "' and a.compid='" & gCompID & "' and a.Status = 1 order by a.ItemCode"
      Else
        Sql = "select a.ItemCode,a.ItemDescription,a.UOMDesc,a.UOMDecimal from MATERIAL_MASTER a"
        Sql = Sql & " where a.Branchid ='" & gBranchID & "' and a.compid='" & gCompID & "' and a.Status = 1 order by ItemCode"
      End If
      If FetchValue_New(False, db.ConnectionString, Sql, "1110", "List of Items for the Customer", "Item Code", "Item Description", "UOM") = True Then
        spItem.SetText eEnq.mItmCode, Row, IIf(IsNull(getvalue(1)), "", getvalue(1))
        spItem.SetText eEnq.mItmDesc, Row, IIf(IsNull(getvalue(2)), "", getvalue(2))
        spItem.SetText eEnq.mUOM, Row, IIf(IsNull(getvalue(3)), "", getvalue(3))
        Call Spread_Lock(spItem, False, eEnq.mQty, Row)
        Call Spread_Set_Decimal(spItem, Val(getvalue(4)), eEnq.mQty, Row)
      End If
      If Trim$(spItem.Text) <> "" Then spItem.MaxRows = spItem.DataRowCnt + 1
    ElseIf Col = eEnq.mUOM And Trim(spItem.Text) = "New Item" Then
      Sql = "select SDesc,NoofDecimals from INV_UOM_MASTER where Branchid ='" & gBranchID & "'"
      Sql = Sql & " and compid='" & gCompID & "'"
      If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of UOM's", "UOM") = True Then
        spItem.SetText eEnq.mUOM, Row, IIf(IsNull(getvalue(1)), "", getvalue(1))
        Call Spread_Set_Decimal(spItem, Val(getvalue(2)), eEnq.mQty, Row)
        Call Spread_Set_Decimal(spItem, Val(getvalue(2)), eEnq.mTarVolme, Row)
        Call Spread_Lock(spItem, False, eEnq.mQty, Row)
        Call Spread_Lock(spItem, False, eEnq.mTarVolme, Row)
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_EditMode(sender as object, e as EditModeEventArgs)
  If Col = eEnq.mItmCode Then
    Spread_Clear_Duplicate spItem, spItem.ActiveRow, eEnq.mItmCode
  End If
End Sub

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then spItem_DblClick spItem.ActiveCol, spItem.ActiveRow
  If spItem.ActiveCol = spItem.MaxCols And KeyCode = 13 Then
    spItem.MaxRows = spItem.DataRowCnt + 1
    spItem.SetText eEnq.mItmType, spItem.MaxRows, "Existing Item"
    Call Spread_Lock(spItem, True, 2, spItem.MaxRows, 3)
  ElseIf KeyCode = 46 And Shift = 1 And spItem.ActiveCol = eEnq.mItmType Then
    Spread_Clear_Row spItem, eEnq.mItmType, spItem.MaxCols, spItem.ActiveRow, spItem.ActiveRow
    spItem.DeleteRows spItem.ActiveRow, 1
    spItem.MaxRows = spItem.DataRowCnt + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim tmp As Variant
  
    spItem.GetText eEnq.mItmType, NewRow, tmp
    If Trim$(tmp) = "Existing item" Then
      If NewCol = eEnq.mItmCode Then
        Spread_Set_HeaderColor spItem, anSelectOnly, anCol, NewCol, , NewCol
        Spread_Lock spItem, True, NewCol, NewRow, NewCol, NewRow
      ElseIf NewCol = eEnq.mUOM Then
        Spread_Set_HeaderColor spItem, anEditable, anCol, NewCol, , NewCol
        Spread_Lock spItem, True, NewCol, NewRow, NewCol, NewRow
      End If
    ElseIf Trim$(tmp) = "NEW ITEM" Then
      If NewCol = eEnq.mUOM Then
        Spread_Set_HeaderColor spItem, anSelectOnly, anCol, NewCol, , NewCol
        Spread_Lock spItem, True, NewCol, NewRow, NewCol, NewRow
      ElseIf NewCol = eEnq.mItmCode Then
        Spread_Set_HeaderColor spItem, anEditable, anCol, NewCol, , NewCol
        Spread_Lock spItem, False, NewCol, NewRow, NewCol, NewRow
      End If
    End If
End Sub

Private Sub txtCustomer_Change()
' On Error Resume Next
  If txtCustomer = "" Then
 txtCustomer.Tag = ""
End If
End Sub

Private Sub txtcustomer_DblClick()
  Dim Sql As String

  Try
  
  If cmbCustomerType = "New Customer" Then
 GoTo ExitHere
End If
  
  Me.MousePointer = vbHourglass
  Sql = " select distinct (select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode"
  Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.CustomerCode"
  Sql = Sql + " from SAL_Customer_ItemMaster a where a.Branchid='" & gBranchID & "' "
  Sql = Sql + " and a.Compid='" & gCompID & "'"
  Sql = Sql & " order by PartyName"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer Name") = True Then
     txtCustomer.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
     txtCustomer = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustomer_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
     txtcustomer_DblClick
  End If
End Sub

Private Sub txtCustomer_KeyPress(ByVal KeyAscii As Integer)
    KeyAscii = ToUpper(KeyAscii)
End Sub

Private Sub txtEnqNo_DblClick()
  Dim Sql As String
  

  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct a.EnquiryType,a.EnquiryNo,a.EnquiryDate,a.Revisionno,ReferenceNo,ReferenceDate,CustomerType,CustomerCode,"
  Sql = Sql & " CustomerName,Remarks from SAL_VW_Enquiry_Hdr a Left outer join SAL_VW_Enquiry_Dtl b"
  Sql = Sql & " on a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo and a.EnquiryDate=b.EnquiryDate"
  Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
  Sql = Sql & " where b.QuotationType='' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "1111000010", "List of Enquirys", "Enquiry Type", "Enquiry No", "Enquiry Date", "Revisionno", "Customer Name") = True Then
    txtEnqtype = getvalue(1)
    txtEnqNo = getvalue(2)
    dtpEnqDate = getvalue(3)
    txtEnqNo.Tag = Val(getvalue(4)) + 1 'Revno
    txtRefNo = getvalue(5)
    dtpRefDate = getvalue(6)
    cmbCustomerType.ListIndex = IIf(Trim$(getvalue(7)) = "E", 0, 1)
    txtCustomer.Tag = getvalue(8)
    txtCustomer = getvalue(9)
    txtRemarks = getvalue(10)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtEnqNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtEnqNo_DblClick
End If
  End Sub

Private Sub txtrefno_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = ToUpper(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtremarks_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.ComboSelChange +=  new EventHandler(spItem_ComboSelChange)
spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.EditMode +=  new EventHandler(spItem_EditMode)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)
spItem.LeaveCell +=  new EventHandler(spItem_LeaveCell)

End Sub

Private Sub InitializeSpreads()

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesEnquiryModify.frm", spItem, "")

End Sub
