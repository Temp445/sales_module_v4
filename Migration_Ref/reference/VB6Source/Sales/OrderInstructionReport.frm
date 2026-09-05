VERSION 5.00
Begin VB.Form rptOrderInstructionReport 
   Caption         =   "Order Instruction"
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3870
      TabIndex        =   1
      Top             =   1500
      Width           =   3300
      Begin VB.TextBox txtDcsrno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   345
         Width           =   990
      End
      Begin VB.TextBox txtDcsrDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   345
         Width           =   1215
      End
      Begin VB.TextBox txtDCSRType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   345
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
         Height          =   210
         Left            =   975
         TabIndex        =   7
         Top             =   165
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Order Type"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         Height          =   210
         Left            =   1980
         TabIndex        =   5
         Top             =   165
         Width           =   795
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   420
      Left            =   5550
      TabIndex        =   0
      Top             =   2310
      Width           =   1365
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   420
      Left            =   4200
      TabIndex        =   8
      Top             =   2310
      Width           =   1365
   End
End
Attribute VB_Name = "rptOrderInstructionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Sub sPageHeader()
  Try

  sSendToPort String(80, "-")
  sSendToPort LAlign(45, "Item Code") & Space(1) & LAlign(3, "UOM") & Space(1) & RAlign(8, "Qty") & Space(1) & RAlign(10, "Rate") & Space(1) & RAlign(10, "Value")
  sSendToPort LAlign(45, "Item Description")
  sSendToPort String(80, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub PREVIEW()
  Dim recditem As String
  Dim rs As AceADODB.Recordset
  Dim RQty As String
  Dim Sql As String
  Dim Value As Double
  Dim CurDate As Date
  Dim i As Integer
  
  Try
  
  Set rs = New AceADODB.Recordset
  CurDate = ServerDate
  
  If FrmFlg = "S" Then
    If Pen_Enquiry = eFormList.mSalOrder Then
      Sql = "select a.OrderType, a.OrderNo,d.PartyName,a.OrderDate,a.ValidFrom,a.ValidTo, a.CustomerCode, "
      Sql = Sql + " b.PackingDesc , b.InsuranceDesc, b.DeliveryDesc,b.PaymentDesc , "
      Sql = Sql + " b.FreightDesc, b.ModeDesc, b.formDesc, b.FormulaName, c.Add1, c.Add2, c.City, c.PinCode"
      Sql = Sql + " from dbo.fn_SAL_Order_Hdr ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') a"
      Sql = Sql + " left outer join dbo.fn_SAL_SO_Terms('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') b"
      Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate and a.Branchid = b.Branchid"
      Sql = Sql + " and a.Compid = b.Compid and a.Finyear = b.Finyear left outer join FAS_Party_Address c"
      Sql = Sql + " on a.CustomerCode = c.PartyCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
      Sql = Sql + " left outer Join PartyMaster d on  a.CustomerCode = d.PartyCode"
      Sql = Sql + " and a.Branchid = d.Branchid and a.Compid = d.Compid"
      Sql = Sql + " where a.ordertype = '" & Trim(txtDCSRType.Text) & "'"
      Sql = Sql + " and a.orderNo = '" & Trim(txtDcsrno.Text) & "'"
      Sql = Sql + " and a.compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
    Else
      Sql = "select a.OrderType, a.OrderNo,d.PartyName,a.OrderDate,a.ValidFrom,a.ValidTo, a.CustomerCode, "
      Sql = Sql + " b.PackingDesc , b.InsuranceDesc, b.DeliveryDesc,b.PaymentDesc , "
      Sql = Sql + " b.FreightDesc, b.ModeDesc, b.formDesc, b.FormulaName, c.Add1, c.Add2, c.City, c.PinCode"
      Sql = Sql + " from dbo.fn_SAL_JobWorkOrder_Hdr ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') a"
      Sql = Sql + " left outer join dbo.fn_SAL_SO_Terms('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') b"
      Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate and a.Branchid = b.Branchid"
      Sql = Sql + " and a.Compid = b.Compid and a.Finyear = b.Finyear left outer join FAS_Party_Address c"
      Sql = Sql + " on a.CustomerCode = c.PartyCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
      Sql = Sql + " left outer Join PartyMaster d on  a.CustomerCode = d.PartyCode"
      Sql = Sql + " and a.Branchid = d.Branchid and a.Compid = d.Compid"
      Sql = Sql + " where a.ordertype = '" & Trim(txtDCSRType.Text) & "'"
      Sql = Sql + " and a.orderNo = '" & Trim(txtDcsrno.Text) & "'"
      Sql = Sql + " and a.compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
    End If
    
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    If Pen_Enquiry = eFormList.mSalOrder Then
      sSendToPort CAlign(iPageWidth, "SALES ORDER INSTRUCTION")
    Else
      sSendToPort CAlign(iPageWidth, "JOBWORK ORDER INSTRUCTION")
    End If
    sSendToPort ""
    sSendToPort "Printed Date : " & Format(CurDate, "dd/mm/yyyy") & " " & Format(CurDate, "HH:MM AMPM")
    sSendToPort ""
    If Pen_Enquiry = eFormList.mSalOrder Then
      sSendToPort "S.O No                :" & Space(1) & rs("Ordertype") & rs("OrderNo") & Space(30) & "Date :" & rs("Orderdate")
    Else
      sSendToPort "J.O No                :" & Space(1) & rs("Ordertype") & rs("OrderNo") & Space(30) & "Date :" & rs("Orderdate")
    End If
    sSendToPort ""
    sSendToPort "Customer              :" & Space(1) & rs("PartyName")
    sSendToPort Space(24) & rs("add1")
    If Trim$(rs("add2")) <> "" Then
 sSendToPort Space(24) & rs("add2")
End If
    sSendToPort Space(24) & rs("city")
    sSendToPort Space(24) & rs("pincode")
    sSendToPort ""
    sSendToPort "Validity              :" & Space(1) & rs("ValidFrom") & Space(1) & "To " & rs("ValidTo")
    sSendToPort ""
    sSendToPort "Mode                  :" & Space(1) & rs("ModeDesc")
    sSendToPort ""
    sSendToPort "Freight               :" & Space(1) & rs("FreightDesc")
    sSendToPort ""
    sSendToPort "Payment Terms         :" & Space(1) & rs("PaymentDesc")
    sSendToPort ""
    sSendToPort "Delivery Terms        :" & Space(1) & rs("DeliveryDesc")
    sSendToPort ""
    sSendToPort "Insurance             :" & Space(1) & rs("InsuranceDesc")
    sSendToPort ""
    sSendToPort "Packing & Forwarding  :" & Space(1) & rs("PackingDesc")
    sSendToPort ""
    sSendToPort "Against Form          :" & Space(1) & rs("FormDesc")
    sSendToPort ""
    
    sPageHeader
    
    If Pen_Enquiry = eFormList.mSalOrder Then
      Sql = "select a.ItemCode, b.ItemDescription , b.UOMDesc UOM,"
      Sql = Sql + " c.ItemDescription RecdDesc, c.UOMDesc RecdUOM, c.UOMDecimal RecdDecimal, a.OrderQty,"
      Sql = Sql + " a.Rate from dbo.fn_SAL_Order_Dtl ('" + Format(CurDate, "yyyy-MM-dd HH:mm:ss") + "','" + Format(CurDate, "yyyy-MM-dd") + "', '" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') a"
      Sql = Sql + " left outer join Material_Master b on a.ItemCode = b.ItemCode and a.Branchid = b.Branchid"
      Sql = Sql + " and a.Compid = b.Compid left outer join Material_Master c on a.ItemCode = c.ItemCode"
      Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid where a.ordertype = '" & Trim(txtDCSRType.Text) & "'"
      Sql = Sql + " and a.orderNo = '" & Trim(txtDcsrno.Text) & "' and b.Status = 1 and c.Status = 1"
      Sql = Sql + " and a.compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
    Else
      Sql = "select a.SentItemCode ItemCode, b.ItemDescription , b.UOMDesc UOM,"
      Sql = Sql + " c.ItemDescription RecdDesc, c.UOMDesc RecdUOM, c.UOMDecimal RecdDecimal, a.OrderQty,"
      Sql = Sql + " a.Rate from dbo.fn_SAL_JobWorkOrder_Dtl ('" + Format(CurDate, "yyyy-MM-dd HH:mm:ss") + "','" + Format(CurDate, "yyyy-MM-dd") + "', '" & gBranchID & "','" & gCompID & "', '" & gFinyear & "') a"
      Sql = Sql + " left outer join Material_Master b on a.SentItemCode = b.ItemCode and a.Branchid = b.Branchid"
      Sql = Sql + " and a.Compid = b.Compid left outer join Material_Master c on a.ReceivedItemCode = c.ItemCode"
      Sql = Sql + " and a.Branchid = c.Branchid and a.Compid = c.Compid where a.ordertype = '" & Trim(txtDCSRType.Text) & "'"
      Sql = Sql + " and a.orderNo = '" & Trim(txtDcsrno.Text) & "' and b.Status = 1 and c.Status = 1"
      Sql = Sql + " and a.compid='" & gCompID & "' and a.finyear='" & gFinyear & "'"
    End If
    If rs.State = adStateOpen Then
 rs.Close
End If
    rs.Open Sql, db, adOpenStatic, adLockReadOnly
    rs.ActiveConnection = Nothing
    For i = 1 To rs.RecordCount
      If recditem <> rs("ItemCode") Then
        'RQty = Format(rs("ToBeRecdQty"), "0") & IIf(Val(rs("RecdDecimal")) > 0, "." & String(Val(rs("RecdDecimal")), "0"), "")
        RQty = IIf(Val(rs("OrderQty")) = "0", "Open", rs("OrderQty"))
        Value = Val(RQty) * Val(rs("Rate"))
        sSendToPort LAlign(45, rs("ItemCode")) & Space(1) & RAlign(3, rs("UOM")) & Space(1) & RAlign(8, CStr(RQty)) & Space(1) & RAlign(10, Format(rs("Rate"), "0.00")) & Space(1) & RAlign(10, CDbl(Format(Value, "0.00")))
        sSendToPort LAlign(45, rs("ItemDescription"))
        sSendToPort ""
'      Else
'        sSendToPort LAlign(28, rs("SentItemCode")) & Space(1) & RAlign(9, CStr(SQty))
'        sSendToPort LAlign(28, rs("SentDesc"))
'        sSendToPort ""
      End If
      recditem = rs("ItemCode")
      rs.MoveNext
    Next i
    sSendToPort String(80, "-")
    sSendToPort ""
    sSendToPort ""
    sSendToPort ""
    sSendToPort ""
    sSendToPort ""
    sSendToPort ""
    sSendToPort ""
    sSendToPort " Prepared by" & Space(20) & "Manager" & Space(20) & "Authorised Signatory"
    sSendToPort ""
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdView_Click()
  
  Dim frmOutput As frmReportViewer

  
  Try
  Me.MousePointer = vbHourglass

  If Trim(txtDcsrno.Text) = "" Then
    MsgBox "Select the order number.", vbInformation, App.Title
    txtDcsrno.SetFocus
    GoTo ExitHere
  End If
  
  Set frmCurrent = Me
  cmdView.Enabled = False
  iPageNo = 1
  iLineNo = 1
  pheight = 75
  pPaper = vbPRPSFanfoldUS
  Orient = 1
  iPageWidth = 80
  If fOpenPort(Me) Then
    PREVIEW
    sClosePort
    Set frmOutput = New frmReportViewer
    frmOutput.Caption = gStrReportHeading
    frmOutput.txtTotalPages.Text = CStr(iPageNo)
    frmOutput.wbrView.Navigate strFileName
    frmOutput.Tag = strFileName
    gStrReportHeading = ""
    frmOutput.Show
  End If
  
  cmdView.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
  If Pen_Enquiry = eFormList.mSalOrder Then
    rptOrderInstructionReport.Caption = "Sales Order InStruction"
  Else
    rptOrderInstructionReport.Caption = "JobWork Order InStruction"
  End If

End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub txtDCSRType_DblClick()
  Dim CurDate As Date
  Dim Sql As String
  
    
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  If Pen_Enquiry = eFormList.mSalOrder Then
    If FrmFlg = "S" Then
      Sql = "select distinct a.OrderType,a.OrderNo,a.OrderDate,b.PartyName from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + "left outer join Partymaster b on  a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid " 'where a.Status='A'"
      Sql = Sql + " order by OrderDate desc,OrderNo desc,OrderType desc"
    End If
  Else
    If FrmFlg = "S" Then
      Sql = "select distinct a.OrderType,a.OrderNo,a.OrderDate,b.PartyName from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a "
      Sql = Sql + "left outer join Partymaster b on  a.CustomerCode=b.PartyCode and a.Branchid=b.Branchid and a.Compid=b.Compid where a.Status='A'"
      Sql = Sql + " order by OrderDate desc,OrderNo desc,OrderType desc"
    End If
  End If
  If FetchValue_New(False, db.ConnectionString, Sql, "1111", "List of Order's", "Order Type", "Order No", "Order Date", "VendorName") = True Then
    If getvalue(1) = "" Then
      GoTo ExitHere
    End If
    txtDCSRType.Text = getvalue(1)
    txtDcsrno.Text = getvalue(2)
    txtDcsrDate.Text = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

  Me.MousePointer = vbDefault
End Try
End Sub
