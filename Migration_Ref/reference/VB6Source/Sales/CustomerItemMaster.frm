VERSION 5.00
Begin VB.Form frmCustomerItemMaster 
   Caption         =   "Customer Item Master"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4770
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5207
      TabIndex        =   5
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4770
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3915
      TabIndex        =   4
      ToolTipText     =   "Click to Save the Data"
      Top             =   4770
      Width           =   1300
   End
   Begin VB.Frame fracust 
      Height          =   690
      Left            =   2542
      TabIndex        =   7
      Top             =   675
      Width           =   6630
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   330
         Left            =   5970
         TabIndex        =   1
         ToolTipText     =   "Click to OK"
         Top             =   285
         Width           =   585
      End
      Begin VB.TextBox txtcustcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   300
         Width           =   5865
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   120
         Width           =   690
      End
   End
Begin AceSpread spItem
      Height          =   2610
      Left            =   345
      TabIndex        =   3
      ToolTipText     =   "Enter the Item Details"
      Top             =   2100
      Width           =   11025
      _Version        =   458752
      _ExtentX        =   19447
      _ExtentY        =   4604
      _StockProps     =   64
      Enabled         =   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   1
      SelectBlockOptions=   0
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   4177
      TabIndex        =   9
      Top             =   1335
      Width           =   3360
      Begin VB.TextBox txtDefaultFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Select the Defalut Formula. Press F2 or Double Click for Help."
         Top             =   300
         Width           =   3225
      End
      Begin VB.Label lblVendCode 
         AutoSize        =   -1  'True
         Caption         =   "Default Formula"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   120
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCustomerItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub CmdClear_Click()
  Try
  Me.MousePointer = vbHourglass
  
  Form_Clear Me
  FraCust.Enabled = True
  spItem.Enabled = False
  cmdSave.Enabled = False
  spItem.MaxRows = 1
  txtcustcode.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim snap As AceADODB.Recordset
  Dim sql As String
  Dim i As Integer
  
  Try
  
  If Trim(txtcustcode) = "" Then
    sShowMessage "Select the Customer Code"
    txtcustcode.SetFocus
    GoTo ExitHere
  End If
  
  Me.MousePointer = vbHourglass
  
  sql = "Select distinct a.Frmid, a.FrmLName from COM_FormulaMast_Hdr a Right Outer Join"
  sql = sql & " FAS_Party_Master b On a.Frmid = b.CustomerFormulaCode and a.BranchID = b.BranchID "
  sql = sql + " and a.CompID = b.CompID where a.ModuleCode = '" & gModuleCode & "' and b.PartyCode=" & Val(txtcustcode.Tag) & ""
  sql = sql + " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then
    GoTo ExitHere
  End If
  If Not snap.EOF Then
    txtDefaultFormula.Tag = snap("Frmid")
    txtDefaultFormula.Text = snap("FrmLName")
  Else
    sShowMessage "No Default Formula Found. So select the Default Formula & also open and update it in Party Master."
    Call txtDefaultFormula_DblClick
  End If

  sql = "select a.ItemCode,(select b.ItemDescription from INV_MATERIAL_MASTER b where"
  sql = sql + " b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid and b.status = 1) ItemDescription,"
  sql = sql + " a.Model,a.IssuedNo,a.PartNo,a.Rate,a.AssessableRate,a.ValidFrom,a.ValidTo,a.FormulaCode,(select top 1 b.FrmLName "
  sql = sql + " from COM_FormulaMast_Hdr b where b.ModuleCode = '" & gModuleCode & "' and b.FrmId=a.FormulaCode and b.Branchid=a.Branchid "
  sql = sql + " and b.Compid=a.Compid) FrmLName,QCCheck, Status from SAL_Customer_ItemMaster a"
  sql = sql + " where a.CustomerCode=" & Val(txtcustcode.Tag) & " and a.Branchid='" & gBranchID & "'"
  sql = sql + " and a.Compid='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spItem.MaxRows = snap.RecordCount + 1
    For i = 1 To snap.RecordCount
      spItem.SetText 1, i, snap("ItemCode")
      spItem.SetText 2, i, IIf(IsNull(snap("ItemDescription")), "", snap("ItemDescription"))
      spItem.SetText 3, i, IIf(IsNull(snap("Model")), "", snap("Model"))
      spItem.SetText 4, i, IIf(IsNull(snap("IssuedNo")), "", snap("IssuedNo"))
      spItem.SetText 5, i, IIf(IsNull(snap("PartNo")), "", snap("PartNo"))
      spItem.SetText 6, i, IIf(IsNull(snap("Rate")), "", snap("Rate"))
      spItem.SetText 7, i, IIf(IsNull(snap("AssessableRate")), "", snap("AssessableRate"))
      spItem.SetText 8, i, IIf(IsNull(snap("ValidFrom")), "", Format(snap("ValidFrom"), "dd/mm/yyyy"))
      spItem.SetText 9, i, IIf(IsNull(snap("ValidTo")), "", Format(snap("ValidTo"), "dd/mm/yyyy"))
      If Val(txtDefaultFormula.Tag) <> snap("FormulaCode") Then
        If IsNull(snap("FrmLName")) Then
          sShowMessage "Formula Not found in the Master for the Item in Row " & i & ". So default formula applied."
        Else
          spItem.SetText 10, i, CheckNull(snap("FormulaCode"))
          spItem.SetText 11, i, CheckNull(snap("FrmLName"))
        End If
      End If
'      spItem.SetText 8, i, IIf(IsNull(snap("CustomerFormulaCode")), "", snap("CustomerFormulaCode"))
'      spItem.SetText 9, i, IIf(IsNull(snap("FrmLName")), "", snap("FrmLName"))
      spItem.SetText 12, i, IIf(snap("QCCheck") = True, "Yes", "No")
      spItem.SetText 13, i, IIf(snap("Status") = True, "Active", "Inactive")
      spItem.SetText 14, i, 0
      snap.MoveNext
    Next i
  Else
    spItem.MaxRows = 1
  End If
  snap.Close
  FraCust.Enabled = False
  spItem.Enabled = True
  cmdSave.Enabled = True
  spItem.SetFocus
  Spread_Set_Focus spItem, 1, 1
  Me.MousePointer = 0
  GoTo ExitHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdSave_Click()
  Dim retVal As Variant
  Dim cmd As AceADODB.Command
  Dim sql As String
  Dim i As Long

  
  Try
  Me.MousePointer = vbHourglass

  Set cmd = New AceADODB.Command
  If Chkbefsave = False Then GoTo ExitHere
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If

  sql = "delete from SAL_Customer_ItemMaster where CustomerCode = " & Val(txtcustcode.Tag) & ""
  sql = sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
    For i = 1 To spItem.DataRowCnt
      cmd.CommandText = "SAL_INS_CUSTOMER_ITEMMASTER"
      commClear cmd
      cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcode.Tag))
      spItem.GetText 1, i, retVal
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
      spItem.GetText 3, i, retVal
       cmd.Parameters.Append cmd.CreateParameter("Model", adVarChar, adParamInput, 20, Trim$(retVal))
      spItem.GetText 4, i, retVal
       cmd.Parameters.Append cmd.CreateParameter("IssuedNo", adVarChar, adParamInput, 10, Trim$(retVal))
      spItem.GetText 5, i, retVal
       cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retVal))
      spItem.GetText 6, i, retVal
       cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(retVal))
      spItem.GetText 7, i, retVal
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(retVal))
      spItem.GetText 8, i, retVal
      If Not retVal = "" Then
        cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(CDate(retVal), "yyyy-MM-dd"))
      Else
        cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Null)
      End If
      spItem.GetText 9, i, retVal
      If Not retVal = "" Then
        cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(CDate(retVal), "yyyy-MM-dd"))
      Else
        cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Null)
      End If
      spItem.GetText 10, i, retVal
      If Val(retVal) <= 0 Then
        'Default Formula
        cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtDefaultFormula.Tag))
      Else
        'Specified Formula
        cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(retVal))
      End If
      spItem.GetText 12, i, retVal
      cmd.Parameters.Append cmd.CreateParameter("QCCheck", adBoolean, adParamInput, , IIf(Trim(retVal) = "Yes", "1", "0"))
      spItem.GetText 13, i, retVal
      cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, , IIf(Trim(retVal) = "Active", "1", "0"))
      Call AddCommonCmdParameters(cmd, , False)
      cmd.Execute
    Next
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  sShowMessage "Customer Item Details Saved"
  CmdClear_Click
  GoTo ExitHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = 0
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

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If Col = 1 Then
    sql = "select a.ItemCode,a.ItemDescription,(select top 1 b.FrmLName from "
    sql = sql + " COM_FormulaMast_Hdr b where b.ModuleCode = '" & gModuleCode & "' and b.FrmId=" & Val(cmdOK.Tag) & ""
    sql = sql + " and b.Branchid='" & gBranchID & "' and b.Compid='" & gCompID & "') FrmLName "
    sql = sql + " from INV_MATERIAL_MASTER a where Status = 1 and Branchid = '" & gBranchID & "'"
    sql = sql & " and compid='" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, sql, "*", "List of Items", "Item Code", "Description") = True Then
      spItem.SetText 1, Row, getvalue(1)
      spItem.SetText 2, Row, getvalue(2)
      spItem.SetText 10, Row, Val(cmdOK.Tag)
      spItem.SetText 11, Row, getvalue(3)
      spItem.SetText 13, Row, "No"
      spItem.SetText 14, Row, "Active"
      spItem.SetText 15, Row, 1
      Spread_Set_Focus spItem, Row, 3
      Avoid_Duplicate spItem, 1
    End If
  ElseIf Col = 11 Then
    sql = " select distinct FrmLName,FrmID from COM_FormulaMast_Hdr where ModuleCode = " & gModuleCode & ""
    sql = sql + " and isActive = '1' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Formula's", "Formula Name") = True Then
      spItem.SetText 10, Row, getvalue(2)
      spItem.SetText 11, Row, getvalue(1)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)

  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
    spItem_DblClick spItem.ActiveCol, spItem.ActiveRow
  ElseIf KeyCode = 46 And spItem.ActiveCol = 1 And Val(SpVal(spItem, spItem.ActiveRow, 12)) = 1 Then
    Spread_Clear_Row spItem, 1, spItem.MaxCols, spItem.ActiveRow, spItem.ActiveRow
    spItem.DeleteRows spItem.ActiveRow, 1
    spItem.MaxRows = spItem.DataRowCnt + 1
  ElseIf KeyCode = 13 Then
    If spItem.ActiveCol = 9 Then
       spItem.MaxRows = spItem.DataRowCnt + 1
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtcustcode_DblClick()
  Dim sql As String

  Try
  Me.MousePointer = vbHourglass

  sql = "select PartyName,PartyCode,CustomerFormulaCode from FAS_PARTY_MASTER where AcType IN ('D','Y')"
  sql = sql & " and Branchid='" & gBranchID & "' and compid='" & gCompID & "' and EffStatus = 1"
  If FetchValue_New(False, db.ConnectionString, sql, "100", "List of Customers", "Customer Name") = True Then
    txtcustcode.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcode = IIf(IsNull(getvalue(1)), "", getvalue(1))
    cmdOK.Tag = IIf(IsNull(getvalue(3)), "", getvalue(3))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtcustcode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)

  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub sprHelp()
End Sub

Private Function Chkbefsave() As Boolean
  Dim retVal As Variant
  Dim chkval As Variant
  Dim i As Long
  
  Try
  
  If Trim(txtDefaultFormula) = "" Then
    sShowMessage "Select the Formula"
    txtDefaultFormula.SetFocus
    Chkbefsave = False
Exit Function
  End If

    For i = 1 To spItem.DataRowCnt
      spItem.GetText 5, i, retVal
      If Trim(retVal) = "" Then
        sShowMessage "Enter the Part No"
        spItem.Col = 5
spItem.Row = i
        spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If

      spItem.GetText 6, i, retVal
      If Val(retVal) = 0 Then
        sShowMessage "Enter the Rate"
        spItem.Col = 6
spItem.Row = i
        spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If
      spItem.GetText 7, i, retVal
      If Val(retVal) = 0 Then
        sShowMessage "Enter the Assessable Rate"
        spItem.Col = 7
spItem.Row = i
        spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If
      
      spItem.GetText 8, i, retVal
      If retVal = "" Then
        sShowMessage "Enter Valid From Date"
        spItem.Row = i
spItem.Col = 8
spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If
      spItem.GetText 9, i, chkval
      If chkval = "" Then
        sShowMessage "Enter Valid To Date"
        spItem.Row = i
spItem.Col = 9
spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If
      If Not retVal = "" Or Not chkval = "" Then
        If DateValue(retVal) > DateValue(chkval) Then
           sShowMessage "Valid To Date should be greater than Valid From Date"
           spItem.Row = i
spItem.Col = 9
           spItem.Action = ActionActiveCell
           spItem.SetFocus
           Chkbefsave = False
Exit Function
        End If
      End If
      spItem.GetText 12, i, retVal
      If retVal = "" Then
        sShowMessage "Select the QCCheck"
        spItem.Row = i
spItem.Col = 12
spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If
      
      spItem.GetText 13, i, retVal
      If retVal = "" Then
        sShowMessage "Select the Status"
        spItem.Row = i
spItem.Col = 13
spItem.Action = ActionActiveCell
        spItem.SetFocus
        Chkbefsave = False
Exit Function
      End If

    Next
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub txtDefaultFormula_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  sql = " select distinct FrmLName,FrmID from COM_FormulaMast_Hdr where ModuleCode = " & gModuleCode & ""
  sql = sql + " and isActive = '1' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Formula", "Formula Name") = True Then
    txtDefaultFormula.Text = getvalue(1)
    txtDefaultFormula.Tag = getvalue(2)
  Else
    If Val(txtDefaultFormula.Tag) <= 0 Then
      sShowMessage "Default Formula must be selected."
      Call txtDefaultFormula_DblClick
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtDefaultFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
    Call txtDefaultFormula_DblClick
  End If
End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerItemMaster.frm", spItem, "")

End Sub
