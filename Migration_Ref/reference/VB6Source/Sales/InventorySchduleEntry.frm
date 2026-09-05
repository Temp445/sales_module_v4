VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInventorySchduleEntry 
   Caption         =   "Inventory Schdule Entry"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
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
   ScaleHeight     =   6765
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMonth 
      Height          =   720
      Left            =   3365
      TabIndex        =   9
      Top             =   795
      Width           =   5100
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   4560
         TabIndex        =   4
         Top             =   315
         Width           =   435
      End
      Begin VB.TextBox txtscheno 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   330
         Width           =   900
      End
      Begin VB.ComboBox CmbSchtype 
         Height          =   330
         ItemData        =   "InventorySchduleEntry.frx":0000
         Left            =   90
         List            =   "InventorySchduleEntry.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Inventory Schedule Type"
         Top             =   330
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpScheduleMonth 
         Height          =   315
         Left            =   2970
         TabIndex        =   3
         ToolTipText     =   "Select the Inventory Schedule Month"
         Top             =   330
         Width           =   1575
         _ExtentX        =   2778
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
         OLEDropMode     =   1
         CustomFormat    =   "MMMM-yyyy"
         Format          =   24707075
         CurrentDate     =   37174
      End
      Begin MSComCtl2.DTPicker dtpScheduleDate 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         Format          =   24707075
         CurrentDate     =   37174
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Sch. No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1170
         TabIndex        =   13
         Top             =   150
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Sch. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1740
         TabIndex        =   12
         Top             =   150
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sch. Type"
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Schedule Month"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2970
         TabIndex        =   10
         Top             =   150
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6555
      TabIndex        =   8
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5265
      TabIndex        =   7
      ToolTipText     =   "Click To Clear the Data"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3975
      TabIndex        =   6
      ToolTipText     =   "Click to Save the Data"
      Top             =   4590
      Width           =   1300
   End
Begin AceSpread spSchedule
      Height          =   2805
      Left            =   1545
      TabIndex        =   5
      ToolTipText     =   "Enter the ItemCode Schedule Details"
      Top             =   1665
      Width           =   8730
      _Version        =   458752
      _ExtentX        =   15399
      _ExtentY        =   4948
      _StockProps     =   64
      Enabled         =   0   'False
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
      MaxCols         =   5
      ScrollBars      =   2
      SelectBlockOptions=   0
      UserResize      =   2
   End
End
Attribute VB_Name = "frmInventorySchduleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private flg As Boolean

'Private Sub TypeLoad()
'  Dim Sql As String
'  Dim adoAddItem As ADODB.Recordset
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Set adoAddItem = New ADODB.Recordset
'  Sql = "Select DocumentType from COM_DOCUMENTTYPE_MASTER where pertainingto='V'"
'  Sql = Sql & " and ModuleName='SAL' and Branchid='" & gBranchID & "' and compid = '" & gCompID & "'"
'  If RSOpen(adoAddItem, Sql) = False Then: GoTo ExitHere
'  If Not adoAddItem.EOF Then
'    CmbSchtype.Clear
'    Do While Not adoAddItem.EOF = True
'      CmbSchtype.AddItem Trim$(adoAddItem("DocumentType"))
'      adoAddItem.MoveNext
'    Loop
'    CmbSchtype.ListIndex = 0
'    txtscheno = GetNo("SAL_INVENTORYSCHEDULEENTRY", "InventoryScheduleType", "InventoryScheduleNo", Trim$(CmbSchtype), "V")
'  End If
'  adoAddItem.Close
'
'ExitHere:
'  On Error GoTo 0
'  Set adoAddItem = Nothing
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

Private Sub CmbSchtype_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Call SetDocNo
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetDocNo()
  Try

  txtscheno = GetDocNo("SAL_INVENTORYSCHEDULEENTRY", "InventoryScheduleType", "InventoryScheduleNo", Trim$(CmbSchtype), "V")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Private Sub cmdClear_Click()

  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  fraMonth.Enabled = True
  spSchedule.Enabled = False
  cmdSave.Enabled = False
'  TypeLoad
  Call LoadDocType(mInventScheEntry)
  
  dtpScheduleDate.Value = ServerDate
  dtpScheduleMonth.Value = dtpScheduleDate
  CmbSchtype.SetFocus
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
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  If Trim(CmbSchtype) = "" Then
    sShowMessage "Please select the Inventory Schedule Type"
    CmbSchtype.SetFocus
    GoTo ExitHere
  End If
  
  Set snap = New AceADODB.Recordset
  
  Sql = "select a.InventoryScheduleType,a.InventoryScheduleNo,a.InventoryScheduleDate,"
  Sql = Sql + " a.ItemCode,b.ItemDescription,b.UOMDesc,b.UOMDecimal,"
  Sql = Sql + " a.InventoryScheduleQty from SAL_INVENTORYSCHEDULEENTRY a "
  Sql = Sql + " left outer join MATERIAL_MASTER b on a.ItemCode=b.ItemCode "
  Sql = Sql + " and a.Branchid=b.Branchid and a.Compid=b.Compid where "
  Sql = Sql + " month(a.InventoryScheduleMonth)=" & Month(dtpScheduleMonth.Value) & ""
  Sql = Sql + " and year(a.InventoryScheduleMonth)=" & Year(dtpScheduleMonth.Value) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql + " and a.Finyear='" & gFinyear & "' and b.Status = 1 and a.ModifiedDate=(select "
  Sql = Sql + " max(e.ModifiedDate) from SAL_INVENTORYSCHEDULEENTRY e "
  Sql = Sql + " where e.ItemCode=a.ItemCode and e.InventoryScheduleType=a.InventoryScheduleType"
  Sql = Sql + " and e.InventoryScheduleNo=a.InventoryScheduleNo "
  Sql = Sql + " and e.InventoryScheduleDate=a.InventoryScheduleDate"
  Sql = Sql + " and e.Branchid=a.Branchid and e.Compid=a.Compid and e.Finyear=a.finyear)"
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spSchedule.MaxRows = snap.RecordCount + 1
    CmbSchtype = snap("InventoryScheduleType")
    txtscheno = snap("InventoryScheduleNo")
    dtpScheduleDate = snap("InventoryScheduleDate")
    flg = True
    For i = 1 To snap.RecordCount
      spSchedule.SetText 1, i, snap("ItemCode")
      spSchedule.SetText 2, i, snap("ItemDescription")
      spSchedule.SetText 3, i, snap("UOMDesc")
'      Spread_Decimal_Set spSchedule, 4, 4, i, i, snap("UOMDecimal")
      Call Spread_Set_Decimal(spSchedule, snap("UOMDecimal"), 4, i)
      spSchedule.SetText 4, i, snap("InventoryScheduleQty")
      Call Spread_Lock(spSchedule, False, 4, i)
      spSchedule.SetText 5, i, 0
      snap.MoveNext
    Next i
  End If
  fraMonth.Enabled = False
  spSchedule.Enabled = True
  cmdSave.Enabled = True
  spSchedule.SetFocus
  Spread_Set_Focus spSchedule, 1, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim tmp As Variant
  Dim CurDate As Date
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass

  If validate_data = False Then
 GoTo ExitHere
End If
  
  If flg = False Then
    Call SetDocNo
  End If
  
  Set cmd = New AceADODB.Command
  CurDate = ServerDate
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "SAL_INS_INVENTORYSCHEDULEENTRY"
  For i = 1 To spSchedule.DataRowCnt
    spSchedule.GetText 4, i, tmp
    If Val(tmp) > 0 Then
      commClear cmd
      cmd.Parameters.Append cmd.CreateParameter("InventoryScheduleType", adVarChar, adParamInput, 3, Trim$(CmbSchtype.Text))
      cmd.Parameters.Append cmd.CreateParameter("InventoryScheduleNo", adInteger, adParamInput, , Val(txtscheno.Text))
      cmd.Parameters.Append cmd.CreateParameter("dtpScheduleDate", adDBTimeStamp, adParamInput, , Format$(dtpScheduleDate.Value, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("ScheduleMonth", adDBTimeStamp, adParamInput, , Format$(DateAdd("d", -(Day(dtpScheduleMonth) - 1), dtpScheduleMonth), "yyyy-MM-dd"))
      spSchedule.GetText 1, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      spSchedule.GetText 4, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("EffDate", adDBTimeStamp, adParamInput, , Format$(CurDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("ModifiedDate", adDBTimeStamp, adParamInput, , Format$(CurDate, "yyyy-MM-dd HH:mm:ss"))
      cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, , True)
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
    End If
  Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record(s) Saved"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpScheduleMonth_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
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
 SendKeys "{TAB}"
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
   
  dtpScheduleDate.Value = ServerDate
  dtpScheduleMonth.Value = DateAdd("d", -(Day(dtpScheduleDate) - 1), dtpScheduleDate)
  dtpScheduleMonth.MinDate = DateAdd("d", -(Day(dtpScheduleDate) - 1), dtpScheduleDate)
'  TypeLoad  'to load the Doc type
  flg = False
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

Private Function validate_data() As Boolean
  Dim tmp As Variant
  Dim flg As Boolean
  Dim i As Long
  
  Try
  flg = False
  validate_data = False
  
  If spSchedule.DataRowCnt = 0 Then
    MsgBox "No Items Selected ", vbOKOnly, App.Title
    GoTo ExitHere
  End If
  For i = 1 To spSchedule.DataRowCnt
    spSchedule.GetText 4, i, tmp
    If Val(tmp) > 0 Then
 flg = True
End If
  Next i
  If flg = False Then
    sShowMessage "Please enter the Inventory Schedule Qty"
    spSchedule.SetFocus
    Spread_Set_Focus spSchedule, 1, 4
    GoTo ExitHere
  End If
  
  validate_data = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub spSchedule_DblClick(sender as object, e as DblClickEventArgs)
  Dim tmp As Variant
  Dim srow As Long
  Dim Sql As String
  Dim i As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate

  spSchedule.GetText 1, Row, tmp
  If Col = 1 And Trim$(tmp) = "" Then
    Sql = "select distinct a.ItemCode,b.ItemDescription,b.UOMDesc,"
    Sql = Sql + " b.UOMDecimal from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join "
    Sql = Sql + " MATERIAL_MASTER b on a.ItemCode=b.ItemCode and a.Branchid=b.Branchid "
    Sql = Sql + " and a.Compid=b.Compid where a.Status='A'"
    Sql = Sql + " and a.ValidTo>=convert(datetime,'" + Format(dtpScheduleDate, "yyyy-MM-dd") + "',111)"
    Sql = Sql + " and a.Status='A' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Sql = Sql + " and a.Finyear='" & gFinyear & "' and b.Status = 1 and b.ItemNature in ('F', 'S')"
    If FetchValue_New(True, db.ConnectionString, Sql, "1100", "List of Item's", "Item Code", "Item Description") = True Then
      spSchedule.MaxRows = spSchedule.DataRowCnt + UBound(getvalue, 1) + 1
      For i = 1 To UBound(getvalue, 1)
        srow = spSchedule.DataRowCnt + 1
        spSchedule.SetText 1, srow, Trim(getvalue(i, 1))
        spSchedule.SetText 2, srow, Trim(getvalue(i, 2))
        spSchedule.SetText 3, srow, Trim(getvalue(i, 3))
'        Spread_Decimal_Set spSchedule, 4, 4, SRow, SRow, Val(getvalue(i, 4))
        Call Spread_Set_Decimal(spSchedule, Val(getvalue(i, 4)), 4, srow)
        Call Spread_Lock(spSchedule, False, 4, srow)
        spSchedule.SetText 5, srow, 1
      Next i
      Avoid_Duplicate spSchedule, 1
      spSchedule.SetFocus
      Spread_Set_Focus spSchedule, Row, 4
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spSchedule_KeyDown(sender as object, e as KeyDownEventArgs)
  If KeyCode = 46 And spSchedule.ActiveCol = 1 And Val(SpVal(spSchedule, spSchedule.ActiveRow, 5)) = 1 Then
    Spread_Clear_Row spSchedule, 1, spSchedule.MaxCols, spSchedule.ActiveRow, spSchedule.ActiveRow
    spSchedule.DeleteRows spSchedule.ActiveRow, 1
    spSchedule.MaxRows = spSchedule.DataRowCnt + 1
  End If
End Sub

Private Sub Design_spSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spSchedule


spSchedule.DblClick +=  new EventHandler(spSchedule_DblClick)
spSchedule.KeyDown +=  new EventHandler(spSchedule_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\InventorySchduleEntry.frm", spSchedule, "")

End Sub
