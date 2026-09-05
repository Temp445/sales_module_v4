VERSION 5.00
Begin VB.Form frmDocumentTypeMaster 
   Caption         =   "Document Type Master"
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
   Icon            =   "DocumentTypeMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6375
      TabIndex        =   6
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   2625
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5085
      TabIndex        =   5
      ToolTipText     =   "Click to Clear the Data"
      Top             =   2625
      Width           =   1300
   End
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
      Height          =   1770
      Left            =   3675
      TabIndex        =   7
      Top             =   765
      Width           =   4155
      Begin VB.ComboBox cmbPertaining 
         Height          =   330
         ItemData        =   "DocumentTypeMaster.frx":000C
         Left            =   1305
         List            =   "DocumentTypeMaster.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select the Pertaining To"
         Top             =   1290
         Width           =   2505
      End
      Begin VB.TextBox txttype 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   0
         ToolTipText     =   "Enter the Document Type"
         Top             =   180
         Width           =   525
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Enter the Document Description"
         Top             =   555
         Width           =   2775
      End
      Begin VB.TextBox txtstart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         MaxLength       =   5
         TabIndex        =   2
         ToolTipText     =   "Enter Starting No"
         Top             =   915
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Document Type"
         Height          =   210
         Left            =   30
         TabIndex        =   8
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "* Description"
         Height          =   210
         Left            =   345
         TabIndex        =   9
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "* Starting No"
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pertaining To"
         Height          =   210
         Left            =   330
         TabIndex        =   11
         Top             =   1350
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3795
      TabIndex        =   4
      ToolTipText     =   "Click to Save the Data"
      Top             =   2625
      Width           =   1300
   End
Begin AceSpread sprdoctype
      Height          =   1560
      Left            =   2040
      TabIndex        =   12
      Top             =   3105
      Width           =   7350
      _Version        =   458752
      _ExtentX        =   12965
      _ExtentY        =   2752
      _StockProps     =   64
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
      MaxRows         =   10
      ScrollBars      =   2
      SelectBlockOptions=   0
   End
End
Attribute VB_Name = "frmDocumentTypeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  Option Compare Text

Private Sub LoadSpr()
  Dim sql As String
  Dim adoRs As AceADODB.Recordset

  Try

  sql = "select a.DocumentType,a.Description,a.StartingNo,a.PertainingTo from"
  sql = sql + " COM_DOCUMENTTYPE_MASTER a where a.Branchid='" & gBranchID & "'"
  sql = sql + " and a.Compid='" & gCompID & "' and a.ModuleName='SAL'"
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, sql) = False Then GoTo ExitHere
  If Not adoRs.EOF Then
    Spread_Clear sprdoctype
    Do Until adoRs.EOF
        sprdoctype.MaxRows = sprdoctype.DataRowCnt + 1
        sprdoctype.Row = sprdoctype.MaxRows
        sprdoctype.Col = 1
sprdoctype.Text = IIf(IsNull(adoRs("DocumentType")), "", adoRs("DocumentType"))

        sprdoctype.Col = 2
sprdoctype.Text = IIf(IsNull(adoRs("Description")), "", adoRs("Description"))

        sprdoctype.Col = 3
sprdoctype.Text = IIf(IsNull(adoRs("StartingNo")), "", adoRs("StartingNo"))

        sprdoctype.Col = 4
sprdoctype.Text = PerToCode_Name(adoRs("PertainingTo"))
      adoRs.MoveNext
    Loop
  End If
  adoRs.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Public Function PerToCode_Name(ByVal PertainingCode As String) As String
  Dim sql As String
  
  Try
  
  PerToCode_Name = ""
  
  If Trim$(PertainingCode) <> "" Then
    Select Case UCase(Trim$(PertainingCode))
      Case "E"
        sql = "Sales Enquiry"
      Case "N"
        sql = "Jobwork Enquiry"
      Case "Q"
        sql = "Sales Quotation"
      Case "U"
        sql = "Jobwork Quotation"
      Case "C"
        sql = "Customer PO Entry"
      Case "O"
        sql = "Sales Order"
      Case "J"
        sql = "Jobwork Order"
      Case "S"
        sql = "Schedule Entry"
      Case "V"
        sql = "Inventory Schedule Entry"
      Case "W"
        sql = "Weekly Plan"
      Case "A"
        sql = "Sales Order Amendment"
      Case "M"
        sql = "Jobwork Order Amendment"
      Case "D"
        sql = "Delivery Challan"
      Case "I"
        sql = "Sales Invoice"
      Case "YP"
        sql = "Sales Plan Yearly"
      Case "MP"
        sql = "Sales Plan Monthly"
      Case "WP"
        sql = "Sales Plan Weekly"
      Case "DP"
        sql = "Sales Plan Daily"
      Case "CR"
        sql = "Contract Review"
      Case "DD"
        sql = "Downstream Document"
    End Select
  End If
  
  PerToCode_Name = sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function PerToName_Code(ByVal PertainingText As String) As String
  Dim sql As String
  
  Try
  
  PerToName_Code = ""
  
  If Trim$(PertainingText) <> "" Then
    Select Case UCase(Trim$(PertainingText))
      Case "Sales Enquiry"
        sql = "E"
      Case "Jobwork Enquiry"
        sql = "N"
      Case "Sales Quotation"
        sql = "Q"
      Case "Jobwork Quotation"
        sql = "U"
      Case "Customer PO Entry"
        sql = "C"
      Case "Sales Order"
        sql = "O"
      Case "Jobwork ORDER"
        sql = "J"
      Case "Schedule Entry"
        sql = "S"
      Case "Inventory Schedule Entry"
        sql = "V"
      Case "Weekly Plan"
        sql = "W"
      Case "Sales Order Amendment"
        sql = "A"
      Case "Jobwork Order Amendment"
        sql = "M"
      Case "Delivery Challan"
        sql = "D"
      Case "Sales Invoice"
        sql = "I"
      Case "Sales Plan Yearly"
        sql = "YP"
      Case "Sales Plan Monthly"
        sql = "MP"
      Case "Sales Plan Weekly"
        sql = "WP"
      Case "Sales Plan Daily"
        sql = "DP"
      Case "Contract Review"
        sql = "CR"
      Case "Downstream Document"
        sql = "DD"
    End Select
  End If
  
  PerToName_Code = sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateData() As Boolean
  Dim sql As String
  Dim adoRs As AceADODB.Recordset
  
  Try
  ValidateData = False
  
  If contrl_valid(txtType, "Enter the Document Type") = False Then GoTo ExitHere
  If contrl_valid(txtdesc, "Enter the  Description") = False Then GoTo ExitHere
  If contrl_valid(txtStart, "Enter the Starting Number") = False Then GoTo ExitHere
  
  sql = " select DocumentType from COM_DOCUMENTTYPE_MASTER where "
  sql = sql + " DocumentType = '" & Trim$(txtType.Text) & "' and Branchid='" & gBranchID & "'"
  sql = sql + " and Compid = '" & gCompID & "'"
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, sql) = False Then GoTo ExitHere
  If Not adoRs.BOF Then
    sShowMessage "Document Type already exists"
    txtType.SetFocus
    GoTo ExitHere
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Function

Private Sub CmdClear_Click()

  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  cmbPertaining.ListIndex = 0
  txtType.Enabled = True
  txtType.SetFocus
  cmbPertaining.Enabled = True
  txtStart.Enabled = True
  txtdesc.Enabled = True
  LoadSpr
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

Private Sub cmdSave_Click()
  Dim Str As String
  Dim cmd As New AceADODB.Command

  Try
  Me.MousePointer = vbHourglass

  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  Str = PerToName_Code(Trim(cmbPertaining))
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "COM_INS_DOCUMENTTYPE_MASTER"
  commClear cmd
  cmd.Parameters.Append cmd.CreateParameter("DocumentType", adChar, adParamInput, 3, Trim$(txtType))
  cmd.Parameters.Append cmd.CreateParameter("Description", adVarChar, adParamInput, 30, Trim$(txtdesc))
  cmd.Parameters.Append cmd.CreateParameter("StartingNo", adInteger, adParamInput, , Val(Trim$(txtStart)))
  cmd.Parameters.Append cmd.CreateParameter("PertainingTo", adVarChar, adParamInput, 3, Trim$(Str))
  cmd.Parameters.Append cmd.CreateParameter("Module", adVarChar, adParamInput, 3, "SAL")
  cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, 3, True)
  Call AddCommonCmdParameters(cmd, , False)
  cmd.Execute
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record(s) Saved"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  
  Try
  
  If KeyCode = 13 Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Sub
Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Try
  
  LoadSpr
  cmbPertaining.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Sub
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

Private Sub txtdesc_KeyPress(ByVal KeyAscii As Integer)
  Try
  
  KeyAscii = Asc(UCase$(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Sub
Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtstart_KeyPress(ByVal KeyAscii As Integer)
  Try
  
  KeyAscii = NumValChk(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Sub
Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txttype_KeyPress(ByVal KeyAscii As Integer)
  Try
  
  KeyAscii = Asc(UCase$(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Sub
Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprdoctype(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprdoctype

End Sub

Private Sub InitializeSpreads()

	Call Design_sprdoctype("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\DocumentTypeMaster.frm", sprdoctype, "")

End Sub
