VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInvoiceGrouping 
   Caption         =   "Invoice Grouping Screen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame FraCust 
      Caption         =   "* Customer "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   3103
      TabIndex        =   1
      Top             =   660
      Width           =   4545
      Begin VB.TextBox txtcust 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   210
         Width           =   4335
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
      Left            =   7050
      TabIndex        =   13
      Top             =   2280
      Width           =   345
   End
   Begin VB.Frame FraInvoice 
      Caption         =   "Invoice "
      Height          =   3000
      Left            =   1858
      TabIndex        =   14
      Top             =   2805
      Width           =   7035
Begin AceSpread SprInvoice
         Height          =   2655
         Left            =   120
         TabIndex        =   15
         Top             =   225
         Width           =   6795
         _Version        =   458752
         _ExtentX        =   11986
         _ExtentY        =   4683
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   10
         OperationMode   =   1
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6015
      TabIndex        =   18
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5910
      Width           =   1300
   End
   Begin VB.Frame Fradate 
      Caption         =   "Period"
      Height          =   720
      Left            =   3793
      TabIndex        =   9
      ToolTipText     =   "Select to Choose the Date"
      Top             =   2070
      Width           =   3165
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         ToolTipText     =   "Select the From Date"
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         Format          =   20119553
         CurrentDate     =   38393
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   330
         Left            =   1740
         TabIndex        =   12
         ToolTipText     =   "Select the To Date"
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         Format          =   20119553
         CurrentDate     =   38393
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   1455
         TabIndex        =   11
         Top             =   330
         Width           =   195
      End
   End
   Begin VB.CommandButton cbClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4725
      TabIndex        =   17
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5910
      Width           =   1300
   End
   Begin VB.Frame fraRef 
      Height          =   720
      Left            =   3013
      TabIndex        =   2
      Top             =   1290
      Width           =   4725
      Begin VB.TextBox txtRefNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "Ref No"
         Top             =   315
         Width           =   2475
      End
      Begin VB.TextBox txtRefType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         MaxLength       =   3
         TabIndex        =   4
         ToolTipText     =   "Ref Type"
         Top             =   315
         Width           =   870
      End
      Begin MSComCtl2.DTPicker dtpRefdate 
         Height          =   315
         Left            =   3465
         TabIndex        =   8
         ToolTipText     =   "Ref Date"
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20119555
         CurrentDate     =   37250
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ref Type"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* Ref No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1035
         TabIndex        =   5
         Top             =   135
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ref Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   6
         Top             =   135
         Width           =   645
      End
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3435
      TabIndex        =   16
      ToolTipText     =   "Click to Save the Data"
      Top             =   5910
      Width           =   1300
   End
End
Attribute VB_Name = "FrmInvoiceGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eInv
  mInvType = 1
  mInvNo = 2
  mInvDate = 3
  mInvDate2 = 4
  mInvAmt = 5
  mVend = 6
End Enum

Private Function SaveData() As Boolean
  Dim sql As String
  Dim rCnt As Long
  Dim retVal As Variant
  
  Dim CurDate As Date
  
  Try
  SaveData = False
    
  CurDate = ServerDate
  
    For rCnt = 1 To SprInvoice.DataRowCnt
      sql = "Insert into PMI_Invoice_Ref Values ("
      sql = sql + " '" & gModuleCode & "',"
      
      SprInvoice.GetText eInv.mInvType, rCnt, retVal
      sql = sql + " '" & Trim$(retVal) & "',"
      
      SprInvoice.GetText eInv.mInvNo, rCnt, retVal
      sql = sql + " " & Trim$(retVal) & ","
      
      SprInvoice.GetText eInv.mInvDate2, rCnt, retVal
      sql = sql + " '" & Trim$(retVal) & "',"
      
      sql = sql + " '" & Trim$(txtcust.Tag) & "',"
      
      sql = sql + " '" & Trim$(txtRefType.Text) & "',"
      sql = sql + " '" & Trim$(txtRefNo.Text) & " ',"
      sql = sql + " '" & Format(dtpRefdate, "yyyy-MM-dd") & "',"
      
      SprInvoice.GetText eInv.mInvAmt, rCnt, retVal
      sql = sql + " " & Val(retVal) & ","
      sql = sql + " 0,"
      sql = sql + " '" & gUserID & "',"
      sql = sql + " '" & gBranchID & "',"
      sql = sql + " '" & gCompID & "',"
      sql = sql + "'" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
      sql = sql + " )"
      db.Execute sql
    Next rCnt
   
  SaveData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateData() As Boolean
  Dim rCnt As Long
  
' On Error Resume Next
  
  ValidateData = False
  
  If Trim$(txtcust.Text) = "" Then
    sShowMessage "Select the Customer"
    Exit Function
  End If
  
  If txtRefNo.Text = "" Then
    sShowMessage "Enter the RefNo"
    txtRefNo.SetFocus
    Exit Function
  End If
  
  If dtpfromDate > dtptoDate Then
    sShowMessage "To Date should be Greater than From Date"
    dtpfromDate.SetFocus
    Exit Function
  End If
  
  ValidateData = True
End Function

Private Sub cbclear_Click()
  
  Try
  Me.MousePointer = vbHourglass

  FraCust.Enabled = True
  fraRef.Enabled = True
  Fradate.Enabled = True
  CmdStep.Enabled = True
  FraInvoice.Enabled = False
  txtRefType = ""
  txtRefNo = ""
  
  SprInvoice.ClearRange 1, 1, -1, -1, True
  
  SprInvoice.MaxRows = 1
  txtcust.Text = ""
  txtcust.SetFocus
  
  Call Spread_Set_HeaderColor(SprInvoice, anSelectOnly, anCol, 1)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 2)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 3)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 5)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim isDBTranActive As Boolean
 
  Try
  Me.MousePointer = vbHourglass
  
  db.BeginTrans
  isDBTranActive = True
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
    
  If SaveData() = False Then
    GoTo ExitHere
  End If
  
  db.CommitTrans
  isDBTranActive = False
  
  sShowMessage "Record Saved Successfully."
  
  Call cbclear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  If isDBTranActive = True Then
    db.RollbackTrans
  End If
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  FraCust.Enabled = False
  FraInvoice.Enabled = True
  fraRef.Enabled = False
  Fradate.Enabled = False
  CmdStep.Enabled = False
  SprInvoice.MaxRows = SprInvoice.DataRowCnt + 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = Format(ServerDate, "dd/MM/yyyy")
  
  dtpfromDate = CurDate
dtpfromDate.MaxDate = CurDate

  dtptoDate = CurDate
dtptoDate.MaxDate = CurDate

  dtpRefdate = CurDate
dtpRefdate.MaxDate = CurDate
  FraInvoice.Enabled = False
  Call Spread_Set_HeaderColor(SprInvoice, anSelectOnly, anCol, 1)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 2)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 3)
  Call Spread_Set_HeaderColor(SprInvoice, &HC0&, anCol, 5)
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

Private Sub SprInvoice_DblClick(sender as object, e as DblClickEventArgs)
  Dim sql As String
  Dim rCnt As Long
  Dim Cnt As Long
  
  Try
  Me.MousePointer = vbHourglass
  
    If Row = SprInvoice.DataRowCnt + 1 Then
      If gModuleCode = 4 Then
        sql = "select InvoiceType, InvoiceNo, InvoiceDate, TotalAmount"
        sql = sql & " From Sal_Invoice_Hdr a"
        sql = sql & " Where InvoiceDate between '" & Format(dtpfromDate, "yyyy-MM-dd") & "' and '" & Format(dtptoDate, "yyyy-MM-dd") & "' "
        sql = sql & " and CustomerCode = " & Val(txtcust.Tag) & " and BranchID = '" & gBranchID & " ' and CompId = '" & gCompID & " '"
        sql = sql & " and Not Exists (Select * From PMI_Invoice_Ref b where a.InvoiceNo=b.DocNo and a.InvoiceDate=b.DocDate and a.InvoiceType=b.DocType and a.BranchID=b.BranchID and a.CompID=b.CompID)"
      Else
        sql = "select DocType, DocNo, DocDate, InvoiceAmount"
        sql = sql & " From PUR_PurchaseInvoice_Hdr a"
        sql = sql & " Where DocDate between '" & Format(dtpfromDate, "yyyy-MM-dd") & "' and '" & Format(dtptoDate, "yyyy-MM-dd") & "' "
        sql = sql & " and VendorCode = " & Val(txtcust.Tag) & " and BranchID = '" & gBranchID & " ' and CompId = '" & gCompID & " '"
        sql = sql & " and Not Exists (Select * From PMI_Invoice_Ref b where a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.DocType=b.DocType and a.BranchID=b.BranchID and a.CompID=b.CompID)"
      End If
      If FetchValue_New(True, db.ConnectionString, sql, "11110", "List of Invoices", "Inv Type", "Inv No", "Inv Date", "Inv Amount") = True Then
        SprInvoice.MaxRows = SprInvoice.DataRowCnt + UBound(getvalue, 1) + 1
        rCnt = SprInvoice.DataRowCnt + 1
        For Cnt = 1 To UBound(getvalue, 1)
          SprInvoice.SetText eInv.mInvType, rCnt, getvalue(Cnt, 1)
          SprInvoice.SetText eInv.mInvNo, rCnt, getvalue(Cnt, 2)
          SprInvoice.SetText eInv.mInvDate, rCnt, getvalue(Cnt, 3)
          SprInvoice.SetText eInv.mInvDate, rCnt, getvalue(Cnt, 3)
          SprInvoice.SetText eInv.mInvDate2, rCnt, Format(getvalue(Cnt, 3), "yyyy-MM-dd")
          SprInvoice.SetText eInv.mInvAmt, rCnt, getvalue(Cnt, 4)
          If Spread_Clear_Duplicate(SprInvoice, rCnt, eInv.mInvType, eInv.mInvNo, eInv.mInvDate) = False Then
            rCnt = rCnt + 1
          End If
        Next Cnt
        SprInvoice.MaxRows = SprInvoice.DataRowCnt + 1
      End If
    End If
  
  cmdSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcust_DblClick()
  Dim sql As String

  Try
  Me.MousePointer = vbHourglass
  
  If gModuleCode = 4 Then
    sql = " Select Distinct a.CustomerCode, b.PartyName "
    sql = sql & " From Sal_Invoice_Hdr a"
    sql = sql & " Left Outer Join FAS_Party_Master b"
    sql = sql & " On a.CustomerCode = b.PartyCode and a.BranchId = b.BranchID and a.CompID = b.CompID"
  Else
    sql = " Select Distinct a.VendorCode, b.PartyName "
    sql = sql & " From PUR_PurchaseInvoice_Hdr a"
    sql = sql & " Left Outer Join FAS_Party_Master b"
    sql = sql & " On a.VendorCode = b.PartyCode and a.BranchId = b.BranchID and a.CompID = b.CompID"
  End If
  sql = sql & " Where a.BranchID = '" & gBranchID & "' and a.CompId = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "01", "List of Customers", "Customer Name") = True Then
    txtcust.Tag = getvalue(1)
    txtcust.Text = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtrefno_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

'  KeyAscii = NumValChk(KeyAscii)
  KeyAscii = Asc(UCase$(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtRefType_KeyPress(ByVal KeyAscii As Integer)
 
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = Asc(UCase$(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_SprInvoice(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprInvoice


SprInvoice.DblClick +=  new EventHandler(SprInvoice_DblClick)

End Sub

Private Sub InitializeSpreads()

	Call Design_SprInvoice("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\PMI\FrmInvoiceGrouping.frm", SprInvoice, "")

End Sub
