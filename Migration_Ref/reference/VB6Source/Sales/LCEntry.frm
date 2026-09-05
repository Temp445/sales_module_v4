VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLCEntry 
   Caption         =   "LC Entry"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   14985
   WindowState     =   2  'Maximized
   Begin VB.Frame FraLC 
      Height          =   2535
      Left            =   3448
      TabIndex        =   0
      Top             =   2640
      Width           =   4845
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Select  the Customer. Press F2 or Double Click for Help."
         Top             =   225
         Width           =   3315
      End
      Begin VB.TextBox txtLCNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Enter the Cheque No"
         Top             =   595
         Width           =   2265
      End
      Begin VB.TextBox txtLCAmt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         TabIndex        =   6
         ToolTipText     =   "Enter the Cheque Amount"
         Top             =   1335
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpLCDate 
         Height          =   315
         Left            =   1365
         TabIndex        =   4
         ToolTipText     =   "Cheque Date"
         Top             =   965
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
         Format          =   22544387
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dtpValidFrom 
         Height          =   315
         Left            =   1365
         TabIndex        =   12
         ToolTipText     =   "Enter the Valid From Date"
         Top             =   1695
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   22544387
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dtpValidTo 
         Height          =   315
         Left            =   1365
         TabIndex        =   13
         ToolTipText     =   "Enter the Valid Till Date"
         Top             =   2070
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   22544387
         CurrentDate     =   37250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid From"
         Height          =   210
         Left            =   525
         TabIndex        =   15
         Top             =   1747
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         Height          =   210
         Left            =   705
         TabIndex        =   14
         Top             =   2122
         Width           =   585
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   277
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* LC No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   750
         TabIndex        =   1
         Top             =   647
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* LC Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   5
         Top             =   1387
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "LC Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   3
         Top             =   1017
         Width           =   570
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6510
      TabIndex        =   9
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5265
      Width           =   1300
   End
   Begin VB.CommandButton cbClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5220
      TabIndex        =   8
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5265
      Width           =   1300
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3930
      TabIndex        =   7
      ToolTipText     =   "Click to Save the Data"
      Top             =   5265
      Width           =   1300
   End
End
Attribute VB_Name = "FrmLCEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbclear_Click()
' On Error Resume Next
  
  Form_Clear Me
  CurrDate Me

End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdSave_Click()
 
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If

  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  If SaveData() = False Then
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved Successfully."
  
  Call cbclear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdStep2_Click()
  Try
  Me.MousePointer = vbHourglass
  
  If txtLCNo.Text = "" Then
    sShowMessage "Enter the Cheque No"
    txtLCNo.SetFocus
    GoTo ExitHere
  End If

  If txtLCAmt.Text = "" Then
    sShowMessage "Enter the Cheque Amount"
    txtLCAmt.SetFocus
    GoTo ExitHere
  End If

  FraLC.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
 
  Try
  Me.MousePointer = vbHourglass
  
  dtpLCDate = Format(ServerDate, "dd/MM/yyyy")
  dtpValidFrom = dtpLCDate
  dtpValidTo = dtpLCDate
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

Private Sub txtcustcd_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  sql = "select distinct a.PartyName, a.PartyCode "
  sql = sql + " from FAS_PARTY_MASTER a "
  sql = sql + " where a.Branchid = '" & gBranchID & "' and a.compid='" & gCompID & "'"
  sql = sql & " order by a.PartyName"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Customers", "Customer") = True Then
    txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtLCAmt_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumValChk(KeyAscii)
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function SaveData() As Boolean
  Dim sql As String
  Dim rCnt As Long
  Dim retval As Variant
  
  Dim CurDate As Date
  
  Try
  SaveData = False
  
  CurDate = ServerDate
  
  sql = "Insert into SAL_LC_Dtl Values ("
  sql = sql + " " & Val(txtcustcd.Tag) & ","
  sql = sql + " '" & Trim$(txtLCNo.Text) & "',"
  sql = sql + " '" & Format(dtpLCDate, "yyyy-MM-dd") & "',"
  sql = sql + " " & Val(txtLCAmt.Text) & ","
  sql = sql + " '" & Format(dtpValidFrom, "yyyy-MM-dd") & "',"
  sql = sql + " '" & Format(dtpValidTo, "yyyy-MM-dd") & "',"
  sql = sql + " '" & gUserID & "',"
  sql = sql + " '" & gBranchID & "',"
  sql = sql + " '" & gCompID & "',"
  sql = sql + " '" & Format(CurDate, "yyyy-MM-dd HH:MM:ss") & "'"
  sql = sql & " )"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  SaveData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateData() As Boolean
  Dim TotAmt As Double
  Dim CurrDate As Date
  
  Try

  ValidateData = False
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  
  If txtLCNo.Text = "" Then
    sShowMessage "Enter the LC No"
    txtLCNo.SetFocus
    Exit Function
  End If

  If txtLCAmt.Text = "" Then
    sShowMessage "Enter the LC Amount"
    txtLCAmt.SetFocus
    Exit Function
  End If
  
  If dtpValidFrom > Format(CurrDate, "dd/MM/yyyy") Then
    sShowMessage "Reference Date should be less the or equal to Current Date"
    dtpValidFrom.Enabled = True
    dtpValidFrom.SetFocus
    Exit Function
  End If
  
  If dtpValidTo <= dtpValidFrom Then
    sShowMessage "Valid Till Date should be greater than From Date"
    dtpValidTo.Enabled = True
    dtpValidTo.SetFocus
    Exit Function
  End If
  
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

