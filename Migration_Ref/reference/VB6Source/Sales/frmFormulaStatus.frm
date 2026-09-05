VERSION 5.00
Begin VB.Form frmFormulaStatus 
   Caption         =   "Formula Activation / Deactivation Screen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   390
      Left            =   6000
      TabIndex        =   3
      Top             =   4110
      Width           =   1170
   End
   Begin VB.CommandButton cmdLoadFormula 
      Caption         =   "Load Formula"
      Height          =   390
      Left            =   5115
      TabIndex        =   2
      Top             =   300
      Width           =   1800
   End
   Begin VB.Frame fraFormula 
      Caption         =   "Formula"
      Enabled         =   0   'False
      Height          =   3105
      Left            =   1830
      TabIndex        =   0
      Top             =   825
      Width           =   8370
Begin AceSpread sprFormula
         Height          =   2715
         Left            =   150
         TabIndex        =   1
         Top             =   255
         Width           =   8040
         _Version        =   458752
         _ExtentX        =   14182
         _ExtentY        =   4789
         _StockProps     =   64
         ButtonDrawMode  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4845
      TabIndex        =   4
      Top             =   4110
      Width           =   1170
   End
End
Attribute VB_Name = "frmFormulaStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'For Formula Spread Column No
Private Enum eFrm
  mFrmID = 1
  mSName = 2
  mLName = 3
  mDesc = 4
  mStatus = 5
  mOldStatus = 6
End Enum

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdLoadFormula_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim tmpActiveStatus As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select a.FrmID, a.FrmSName, a.FrmLName, a.FrmDesc, a.isActive"
  Sql = Sql & " from Com_FormulaMast_Hdr a"
  Sql = Sql & " where a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Order by a.FrmSName"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If rs.EOF Then
    sShowMessage "No Record Found."
    GoTo ExitHere
  Else
    rs.MoveLast
rs.MoveFirst
      sprFormula.MaxRows = rs.RecordCount
      For rCnt = 1 To sprFormula.MaxRows
        sprFormula.SetText eFrm.mFrmID, rCnt, rs("FrmID")
        sprFormula.SetText eFrm.mSName, rCnt, rs("FrmSName")
        sprFormula.SetText eFrm.mLName, rCnt, rs("FrmLName")
        sprFormula.SetText eFrm.mDesc, rCnt, rs("FrmDesc")
        tmpActiveStatus = IIf(rs("isActive"), "Active", "Inactive")
        sprFormula.SetText eFrm.mStatus, rCnt, tmpActiveStatus
        sprFormula.SetText eFrm.mOldStatus, rCnt, tmpActiveStatus
        rs.MoveNext
      Next rCnt
    fraFormula.Enabled = True
    cmdSave.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim Sql As String
  Dim retVal As Variant
  Dim rCnt As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
    For rCnt = 1 To sprFormula.MaxRows
      sprFormula.GetText eFrm.mStatus, rCnt, retVal
      If Trim$(retVal) <> Trim$(GetSprText(sprFormula, 6, rCnt)) Then
        Sql = "Update Com_FormulaMast_Hdr"
        Sql = Sql & " Set isActive='" & IIf(Trim$(retVal) = "Active", "1", "0") & "'"
        Sql = Sql & " Where FrmID=" & GetSprText(sprFormula, 1, rCnt)
        Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      End If
    Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved Successfully."
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

Private Sub Design_sprFormula(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprFormula

End Sub

Private Sub InitializeSpreads()

	Call Design_sprFormula("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmFormulaStatus.frm", sprFormula, "")

End Sub
