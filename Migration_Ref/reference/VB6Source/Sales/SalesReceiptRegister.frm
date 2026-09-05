VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSalesReceiptRegister 
   Caption         =   "Sales Receipt Register"
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
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCustomer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1132
      TabIndex        =   10
      Top             =   720
      Width           =   9690
      Begin VB.Frame Frame2 
         Height          =   420
         Left            =   6045
         TabIndex        =   14
         ToolTipText     =   "Select the RR Type"
         Top             =   195
         Width           =   3015
         Begin VB.OptionButton optEntered 
            Caption         =   "RR Entered"
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
            Left            =   1785
            TabIndex        =   4
            Top             =   165
            Width           =   1140
         End
         Begin VB.OptionButton optNEntered 
            Caption         =   "RR Not Entered"
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
            Left            =   90
            TabIndex        =   3
            Top             =   165
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         TabIndex        =   0
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3435
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   9090
         TabIndex        =   5
         ToolTipText     =   "Click to OK"
         Top             =   300
         Width           =   465
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   315
         Left            =   3555
         TabIndex        =   1
         ToolTipText     =   "Select the Invoice From Date"
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   37643
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         ToolTipText     =   "Select the Invoice To Date"
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   37643
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date To"
         Height          =   210
         Left            =   4800
         TabIndex        =   13
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date From"
         Height          =   210
         Left            =   3540
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   105
         TabIndex        =   11
         Top             =   120
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6270
      TabIndex        =   9
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4545
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4980
      TabIndex        =   8
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4545
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3690
      TabIndex        =   7
      ToolTipText     =   "Click to Save the Data"
      Top             =   4545
      Width           =   1300
   End
Begin AceSpread SprRR
      Height          =   2955
      Left            =   75
      TabIndex        =   6
      ToolTipText     =   "Enter the RR Details"
      Top             =   1500
      Width           =   11805
      _Version        =   458752
      _ExtentX        =   20823
      _ExtentY        =   5212
      _StockProps     =   64
      Enabled         =   0   'False
      ColsFrozen      =   5
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
      MaxCols         =   13
      MaxRows         =   1
      ProcessTab      =   -1  'True
      UserResize      =   1
   End
End
Attribute VB_Name = "frmSalesReceiptRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub cmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  fraCustomer.Enabled = True
  SprRR.Enabled = False
  cmdSave.Enabled = False
  txtCustName.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdOk_Click()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
 
  Try
  Me.MousePointer = vbHourglass
    
  Set rs = New AceADODB.Recordset
  If contrl_valid(txtCustName, "Select the Custer Name") = False Then GoTo ExitHere
  If CDate(DTPFrom.Value) > CDate(DTPTo.Value) Then
    sShowMessage "From date can't be greater than to date"
    DTPFrom.SetFocus
    DTPFrom.Value = Format(DTPFrom, "dd/mm/yyyy")
    GoTo ExitHere
  End If

  Sql = "select a.InvoiceType,a.InvoiceNo,a.InvoiceDate,(Select b.PartNo from "
  Sql = Sql + " SAL_Customer_ItemMaster b where b.CustomerCode=a.CustomerCode "
  Sql = Sql + " and b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid) PartNo,"
  Sql = Sql + " a.ItemCode,c.ItemDescription,c.UOMDecimal,a.InvoiceQty,a.RRNo,a.RRDate,"
  Sql = Sql + " a.ReceivedQty,a.AcceptedQty,a.RRRemarks, a.ItemGroupKey from SalesInvoice a left outer join "
  Sql = Sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid "
  Sql = Sql + " and a.Compid=c.Compid where a.CustomerCode = " & Val(txtCustName.Tag) & ""
  Sql = Sql + " and a.InvoiceDate between convert(datetime,'" & Format(DTPFrom.Value, "yyyy-MM-dd") & "',111)"
  Sql = Sql + " and convert(datetime,'" & Format(DTPTo.Value, "yyyy-MM-dd") & "',111)"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'" 'and a.finyear = '" & gFinyear & "' "
  If optNEntered.Value = True Then
 Sql = Sql + " and a.RRNo=''"
End If
  If optEntered.Value = True Then
 Sql = Sql + " and a.RRNo <> ''"
End If
  Sql = Sql + " order by a.InvoiceDate,a.InvoiceType,a.InvoiceNo"
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  If rs.BOF = True Then
    MsgBox "Records not found", vbInformation, App.Title
    Spread_Clear SprRR
    txtCustName.SetFocus
    GoTo ExitHere
  End If
  SprRR.MaxRows = 0
  Do While Not rs.EOF
    SprRR.MaxRows = SprRR.MaxRows + 1
    SprRR.SetText 1, SprRR.MaxRows, Trim(rs("InvoiceType"))
    SprRR.SetText 2, SprRR.MaxRows, rs("InvoiceNo")
    SprRR.SetText 3, SprRR.MaxRows, Format(rs("InvoiceDate"), "dd-MM-yyyy")
    SprRR.SetText 4, SprRR.MaxRows, rs("PartNo")
    SprRR.SetText 5, SprRR.MaxRows, Trim(rs("ItemCode"))
    SprRR.SetText 6, SprRR.MaxRows, rs("ItemDescription")
    Call Spread_Set_Decimal(SprRR, rs("UOMDecimal"), 7, SprRR.MaxRows)
    SprRR.SetText 7, SprRR.MaxRows, rs("InvoiceQty")
    Call Spread_Set_Decimal(SprRR, rs("UOMDecimal"), 10, SprRR.MaxRows, 11)
    If rs("rrno") <> "" Then
      SprRR.SetText 8, SprRR.MaxRows, rs("RRNo")
      SprRR.SetText 9, SprRR.MaxRows, rs("RRDate")
      SprRR.SetText 10, SprRR.MaxRows, rs("ReceivedQty")
      SprRR.SetText 11, SprRR.MaxRows, rs("AcceptedQty")
      SprRR.SetText 12, SprRR.MaxRows, rs("RRRemarks")
    End If
    SprRR.SetText 13, SprRR.MaxRows, rs("ItemGroupKey")
    SprRR.Col = 9
SprRR.Row = SprRR.MaxRows
    SprRR.TypeDateMin = Format(rs("InvoiceDate"), "MMddyyyy")
    rs.MoveNext
  Loop
  fraCustomer.Enabled = False
  SprRR.Enabled = True
  cmdSave.Enabled = True
  SprRR.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdSave_Click()
  Dim i As Long
  Dim tmp As Variant
  Dim cmd As AceADODB.Command
  
  Try
  Me.MousePointer = vbHourglass

  If Validate = False Then GoTo ExitHere
  
  Set cmd = New AceADODB.Command
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc

  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  For i = 1 To SprRR.DataRowCnt
    SprRR.GetText 8, i, tmp
    If Trim$(tmp) <> "" Then
      commClear cmd
      cmd.CommandText = "SAL_UP_INVOICE_HDR"
      SprRR.GetText 1, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adVarChar, adParamInput, 3, Trim$(tmp))
      SprRR.GetText 2, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(tmp))
      SprRR.GetText 3, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(tmp, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtCustName.Tag))
      SprRR.GetText 8, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RRNo", adVarChar, adParamInput, 20, Trim$(tmp))
      SprRR.GetText 9, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RRDate", adDBTimeStamp, adParamInput, , Format(tmp, "yyyy-MM-dd"))
      SprRR.GetText 12, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RRRemarks", adVarChar, adParamInput, 50, Trim$(tmp))
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
      
      commClear cmd
      cmd.CommandText = "SAL_UP_INVOICE_DTL"
      SprRR.GetText 1, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adVarChar, adParamInput, 3, Trim$(tmp))
      SprRR.GetText 2, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(tmp))
      SprRR.GetText 3, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(tmp, "yyyy-MM-dd"))
      SprRR.GetText 13, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      SprRR.GetText 5, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      SprRR.GetText 10, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ReceivedQty", adInteger, adParamInput, , Val(tmp))
      SprRR.GetText 11, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("AcceptedQty", adInteger, adParamInput, , Val(tmp))
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
  Me.MousePointer = 0
End Try
End Sub

Private Sub DTPFrom_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub DTPTo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

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

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurrDate As Date
  
  Try
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  DTPFrom = CurrDate
  DTPTo = CurrDate
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

Private Sub SprRR_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 1 Then
 GoTo ExitHere
End If
  SprRR.GetText 10, Row, tmp
'  Spread_Max_Set SprRR, 11, 11, Row, Val(tmp)
  Call Spread_Set_Max(SprRR, Val(tmp), 11, Row)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustname_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select distinct (select b.PartyName from FAS_PARTY_MASTER b where "
  Sql = Sql + " b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.compid=a.Compid) PartyName,a.CustomerCode "
  Sql = Sql + " from SAL_Customer_ItemMaster a where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer Name") = True Then
    txtCustName.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtCustName.Text = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

  Me.MousePointer = 0
End Try
End Sub

Private Function Validate() As Boolean
  Dim i As Long
  Dim bool As Boolean
  Dim temp As Variant

  Try
  
  Validate = False
  
  If SprRR.DataRowCnt = 0 Then
    sShowMessage "Please enter the RR Details"
    GoTo ExitHere
  End If
    
  bool = False
  For i = 1 To SprRR.DataRowCnt
    SprRR.GetText 8, i, temp
    If Trim$(temp) <> "" Then
      bool = True
      SprRR.GetText 9, i, temp
      If Trim$(temp) = "" Then
        sShowMessage "Enter the RRDate"
        SprRR.SetFocus
        Spread_Set_Focus SprRR, i, 9
        Exit Function
      End If
      SprRR.GetText 10, i, temp
      If Trim$(temp) = "" Then
        sShowMessage "Enter the Recived Qty"
        SprRR.SetFocus
        Spread_Set_Focus SprRR, i, 10
        Exit Function
      End If
      SprRR.GetText 11, i, temp
      If Trim$(temp) = "" Then
        sShowMessage "Enter the Accepted Qty"
        SprRR.SetFocus
        Spread_Set_Focus SprRR, i, 11
        Exit Function
      End If
    End If
  Next i
  If bool = False Then
    sShowMessage "Enter atleast one RR Number"
    SprRR.SetFocus
    Spread_Set_Focus SprRR, 1, 8
    Exit Function
  End If

  Validate = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub txtCustName_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
     txtcustname_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Design_SprRR(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprRR


SprRR.EditMode +=  new EventHandler(SprRR_EditMode)

End Sub

Private Sub InitializeSpreads()

	Call Design_SprRR("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesReceiptRegister.frm", SprRR, "")

End Sub
