VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBankInformation 
   Caption         =   "Bank Information"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtrate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10800
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2325
      Width           =   570
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11625
      TabIndex        =   9
      Top             =   2310
      Width           =   1155
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   8400
      TabIndex        =   5
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   7680
      Width           =   1500
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   405
      Left            =   6915
      TabIndex        =   6
      ToolTipText     =   "Click to Exit (Alt + x)"
      Top             =   7680
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   405
      Left            =   5430
      TabIndex        =   7
      ToolTipText     =   "Click to View the Report"
      Top             =   7680
      Width           =   1500
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9630
      TabIndex        =   8
      Top             =   2010
      Width           =   735
   End
   Begin VB.Frame fraSpread 
      Enabled         =   0   'False
      Height          =   4890
      Left            =   2543
      TabIndex        =   3
      Top             =   2580
      Width           =   10245
Begin AceSpread sprDetails
         Height          =   4605
         Left            =   105
         TabIndex        =   4
         Top             =   195
         Width           =   10035
         _Version        =   458752
         _ExtentX        =   17701
         _ExtentY        =   8123
         _StockProps     =   64
         DAutoSizeCols   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         ScrollBars      =   2
         SelectBlockOptions=   0
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "WareHouse Name"
      Height          =   855
      Index           =   0
      Left            =   5895
      TabIndex        =   2
      Top             =   1725
      Width           =   3645
Begin AceSpread SPWHS
         Height          =   540
         Left            =   105
         TabIndex        =   10
         Top             =   195
         Width           =   3420
         _Version        =   458752
         _ExtentX        =   6033
         _ExtentY        =   953
         _StockProps     =   64
         DAutoSizeCols   =   1
         DisplayColHeaders=   0   'False
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
         MaxCols         =   3
         MaxRows         =   6
         OperationMode   =   2
         ScrollBars      =   2
         SelectBlockOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Month"
      Height          =   615
      Left            =   6975
      TabIndex        =   0
      Top             =   1110
      Width           =   1425
      Begin MSComCtl2.DTPicker dtpMonth 
         Height          =   330
         Left            =   75
         TabIndex        =   1
         ToolTipText     =   "Enter the From Date"
         Top             =   210
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
         CustomFormat    =   "MM-yyyy"
         Format          =   22282241
         CurrentDate     =   37272
      End
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Left            =   11430
      TabIndex        =   12
      Top             =   2385
      Width           =   210
   End
End
Attribute VB_Name = "frmBankInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub cmdClear_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  dtpMonth.Value = Format(CurDate, "MM-yyyy")
  
  Call Spread_Clear(sprDetails)
  txtrate.Text = ""
  cmdSave.Enabled = False
  CmdOk.Enabled = True
  fraSpread.Enabled = False
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  Dim i As Integer
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  For i = 1 To sprDetails.DataRowCnt
    If Val(txtrate.Text) > 0 Then
      sprDetails.GetText 4, i, retVal
      If Val(retVal) = 0 Then
        sprDetails.GetText 3, i, retVal
        retVal = ((retVal * Val(txtrate.Text)) / 100)
        sprDetails.SetText 4, i, Format(retVal, "0.00")
      End If
    Else
      sprDetails.GetText 3, i, retVal
      sprDetails.SetText 4, i, Format(retVal, "0.00")
    End If
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdOk_Click()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim i As Integer
  Dim tmp As Variant
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  sprDetails.ClearSelection
  CurDate = ServerDate
  
  Sql = " Select distinct a.ItemCode, a.ItemDescription as ItemDesc,a.Rate,"
  SPWHS.GetText 1, SPWHS.ActiveRow, tmp
  If Trim$(tmp) = "CHI" Or Trim$(tmp) = "RMS" Or Trim$(tmp) = "TCS" Then
    Sql = Sql & " dbo.fn_PUR_Item_Rate(a.ItemCode,0, '',0,NULL,'" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & gBranchID & "', '" & gCompID & "','" & gFinyear & "') as Purrate"
  Else
    Sql = Sql & " dbo.fn_SAL_Item_Rate(a.ItemCode, 0, '', 0, NULL, '" & gBranchID & "', '" & gCompID & "','" & gFinyear & "') Salrate"
  End If
  Sql = Sql & " From Inv_Material_Master a"
  SPWHS.GetText 3, SPWHS.ActiveRow, tmp
  Sql = Sql & " Where a.Branchid = '" & gBranchID & "' and a.compid = '" & gCompID & "' and a.ItemNature in  " & Trim$(tmp) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If snap.RecordCount = 0 Then
    sShowMessage "No records exist"
    GoTo ExitHere
  End If
  If Not snap.EOF Then
    sprDetails.MaxRows = snap.RecordCount '
      For i = 1 To snap.RecordCount
        sprDetails.SetText 1, i, snap("Itemcode")
        sprDetails.SetText 2, i, snap("ItemDesc")
        SPWHS.GetText 1, SPWHS.ActiveRow, tmp
        If Trim$(tmp) = "CHI" Or Trim$(tmp) = "RMS" Or Trim$(tmp) = "TCS" Then
          tmp = CheckNull(snap("PurRate"), 0)
        Else
          tmp = CheckNull(snap("SalRate"), 0)
        End If
        If Val(tmp) > 0 Then
          sprDetails.SetText 3, i, Format(Val(tmp), "0.00")
        Else
          sprDetails.SetText 3, i, Format(snap("Rate"), "0.00")
        End If
        snap.MoveNext
      Next i
  End If
    
  CmdOk.Enabled = False
  fraSpread.Enabled = True
  cmdSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim Sql As String
  Dim retVal As Variant
  Dim i As Integer
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
    
  sprDetails.Tag = 0
  For i = 1 To sprDetails.DataRowCnt
    sprDetails.GetText 4, i, retVal
    If Val(retVal) > 0 Then
      Sql = "Execute Sal_Ins_BankInformation_Dtl"
      SPWHS.GetText 1, SPWHS.ActiveRow, retVal
      Sql = Sql & " @WHCode = '" & Trim$(retVal) & "',"
      sprDetails.GetText 1, i, retVal
      Sql = Sql & " @ItemCode = '" & Trim$(retVal) & "',"
      sprDetails.GetText 4, i, retVal
      Sql = Sql & " @Rate = " & Format(retVal, "0.00") & ","
      Sql = Sql & " @UserID = '" & gUserID & "',"
      Sql = Sql & " @BranchID = '" & gBranchID & "',"
      Sql = Sql & " @CompID = '" & gCompID & "',"
      Sql = Sql & " @TranStamp = '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      sprDetails.Tag = 1
    End If
  Next i
  
  If Val(sprDetails.Tag) = 0 Then
    sShowMessage "Atleast One Change To Press Save"
    Call cmdClear_Click
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record(s) Saved"
  Call cmdClear_Click
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
  
  CurDate = ServerDate
  dtpMonth.Value = Format(CurDate, "MM-yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprDetails(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprDetails

End Sub

Private Sub Design_SPWHS(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SPWHS

End Sub

Private Sub InitializeSpreads()

	Call Design_sprDetails("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmBankInformation.frm", sprDetails, "")

	Call Design_SPWHS("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmBankInformation.frm", SPWHS, "")

End Sub
