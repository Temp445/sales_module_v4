VERSION 5.00
Begin VB.Form frmDBConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Connection Parameters"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10935
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConnect 
      Height          =   5055
      Left            =   105
      TabIndex        =   4
      Top             =   60
      Width           =   10740
      Begin VB.CommandButton cmdOpen 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10095
         TabIndex        =   26
         Top             =   330
         Width           =   510
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7200
         TabIndex        =   25
         Top             =   345
         Width           =   510
      End
      Begin VB.ComboBox cmbDatabase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7755
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   345
         Width           =   2280
      End
      Begin VB.CheckBox chkLatestFinyear 
         Caption         =   "  Latest Finyear Only"
         Height          =   255
         Left            =   7950
         TabIndex        =   23
         Top             =   750
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkShowLogin 
         Caption         =   "Show Login Form"
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   22
         Top             =   3810
         Width           =   1545
      End
Begin AceSpread spList
         Height          =   2715
         Left            =   150
         TabIndex        =   21
         Top             =   1020
         Width           =   10485
         _Version        =   458752
         _ExtentX        =   18494
         _ExtentY        =   4789
         _StockProps     =   64
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         OperationMode   =   3
         ScrollBars      =   2
         UserResize      =   0
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         CausesValidation=   0   'False
         Height          =   480
         Left            =   4950
         TabIndex        =   3
         Top             =   4455
         Width           =   1800
      End
      Begin VB.CommandButton cmdConnect 
         Cancel          =   -1  'True
         Caption         =   "Co&nnect"
         CausesValidation=   0   'False
         Default         =   -1  'True
         Height          =   480
         Left            =   3180
         TabIndex        =   11
         Top             =   4455
         Width           =   1800
      End
      Begin VB.ComboBox cmbUserType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmDBConnect.frx":036A
         Left            =   3180
         List            =   "frmDBConnect.frx":036C
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4050
         Width           =   795
      End
      Begin VB.ComboBox cmbUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4050
         Width           =   2745
      End
      Begin VB.ComboBox cmbERPUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmDBConnect.frx":036E
         Left            =   5805
         List            =   "frmDBConnect.frx":0370
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3750
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.ComboBox cmbServerIP 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   345
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.ComboBox cmbServer 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   345
         Width           =   3570
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6390
         TabIndex        =   2
         Text            =   "1433"
         Top             =   345
         Width           =   780
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5055
         PasswordChar    =   "_"
         TabIndex        =   1
         Top             =   345
         Width           =   1320
      End
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3705
         TabIndex        =   0
         Top             =   345
         Width           =   1320
      End
Begin AceSpread sprUserData
         Height          =   270
         Left            =   7305
         TabIndex        =   13
         Top             =   4080
         Visible         =   0   'False
         Width           =   2655
         _Version        =   458752
         _ExtentX        =   4683
         _ExtentY        =   476
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         ScrollBars      =   0
      End
      Begin VB.Label lblUserType 
         AutoSize        =   -1  'True
         Caption         =   "UserType"
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
         Left            =   3180
         TabIndex        =   20
         Top             =   3855
         Width           =   690
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "UserName"
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
         Left            =   3990
         TabIndex        =   19
         Top             =   3840
         Width           =   750
      End
      Begin VB.Label lblERPUserID 
         AutoSize        =   -1  'True
         Caption         =   "UserID"
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
         Left            =   5280
         TabIndex        =   18
         Top             =   3795
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblSpr 
         Caption         =   "sprUserData"
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
         Left            =   8865
         TabIndex        =   14
         Top             =   3870
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblPort 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Port No"
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
         Left            =   6390
         TabIndex        =   9
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblDatabase 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Database"
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
         Left            =   7770
         TabIndex        =   8
         Top             =   150
         Width           =   690
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   5055
         TabIndex        =   7
         Top             =   165
         Width           =   690
      End
      Begin VB.Label lblUserID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UserID"
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
         Left            =   3705
         TabIndex        =   6
         Top             =   165
         Width           =   495
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Server"
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
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmDBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isSuccess As Boolean
Private db1 As AceADODB.Connection

Private Enum eINI
  mVersion = 1
  mServer = 2
  mIPAddr = 3
  mAltIPAddr = 4
  mUserID = 5
  mPWD = 6
  mDefaultDB = 7
End Enum

Private Sub chkLatestFinyear_Click()
  Try
  Me.MousePointer = vbHourglass

'  If GetCompIDBranchidFinyear = False Then
'    GoTo ExitHere
'  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbFinyear_Click()
  Try
  Me.MousePointer = vbHourglass

  If GetUserType = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbServer_Click()
  Dim retVal As Variant
  Dim SelRow As Long

  Try
  Me.MousePointer = vbHourglass

  cmbServerIP.ListIndex = cmbServer.ListIndex
  SelRow = cmbServer.ListIndex + 1
  
    sprUserData.GetText eINI.mUserID, SelRow, retVal
    txtUserID.Text = Trim$(retVal)
    
    sprUserData.GetText eINI.mPWD, SelRow, retVal
    txtPassword.Text = Trim$(retVal)
    
  
  CmdOk.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbUserType_Click()
  Try
  Me.MousePointer = vbHourglass

  If GetUserName = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClose_Click()

  Try
  Me.MousePointer = vbHourglass

  Unload Me
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdConnect_Click()
  Dim ConStr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
  isSuccess = False
  
  If spList.DataRowCnt = 0 Then GoTo ExitHere
  
  Sql = "Select * from COM_Company_Master a Left Outer Join COM_Company_Finyear  b On a.CompID=b.CompID and a.BranchID=b.BranchID"
  Sql = Sql & " Where a.BranchID='" & Trim$(SpVal(spList, spList.ActiveRow, 3)) & "' and a.CompID='" & Trim$(SpVal(spList, spList.ActiveRow, 4)) & "'"
  Sql = Sql & " and Convert(varchar(10), b.FromYear, 105)='" & "01-" & Left$(Trim$(SpVal(spList, spList.ActiveRow, 5)), 7) & "' and Convert(varchar(10), b.ToYear, 105)='" & "31-" & Right$(Trim$(SpVal(spList, spList.ActiveRow, 5)), 7) & "'"
  Set rs = New AceADODB.Recordset
  If LocalRSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "No Branch / Company is found for the given Financial Year."
    GoTo ExitHere
  Else
    If IsNull(rs(0)) Then
      sShowMessage "No Branch / Company is found for the given Financial Year."
      GoTo ExitHere
    Else
      gBranchID = Trim$(SpVal(spList, spList.ActiveRow, 3))
      gCompID = Trim$(SpVal(spList, spList.ActiveRow, 4))
      gFinyear = Trim$(SpVal(spList, spList.ActiveRow, 5))
      gUserID = Trim$(cmbERPUserID.List(cmbUserName.ListIndex))
      If chkShowLogin.Value = 1 Then
        gUserID = ""
      Else
        gUserID = Trim$(cmbERPUserID.List(cmbUserName.ListIndex))
      End If
    End If
  End If
  
  isSuccess = True
  
  Unload Me
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmdConnectOld_Click()
'  Dim ConStr As String
'  Dim Sql As String
'  Dim rs As ADODB.Recordset
'
'  On Error GoTo ErrHandler
'
'  isSuccess = False
'
'  If cmbDatabase.ListIndex < 0 Then
'    sShowMessage "Please select the Database"
'    cmbDatabase.SetFocus
'    GoTo ExitHere
'  End If
'
''  #If isWeb = False Then
''    ConStr = "Provider=SQLOLEDB.1;uid=" & Trim$(txtUserID.Text) & ";pwd=" & Trim$(txtPassword.Text) & ";database=" & Trim$(cmbDatabase.Text) & ";server=" & Trim$(cmbServer.Text) & ""
''  #Else
'    ConStr = "Driver={SQL Server};" & "Server=" & IIf(cmbServer.ListIndex = -1, '      Trim$(cmbServer.Text), '      Trim$(cmbServerIP.Text)) & ";" & "Address=" & IIf(cmbServer.ListIndex = -1, '      Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$( '      txtPort.Text) & ";" & "Network=DBMSSOCN;" & "Database=" & Trim$( '      cmbDatabase.Text) & ";" & "Pwd=" & Trim$( '      txtPassword.Text) & ";" & "Uid=" & Trim$(txtUserID.Text) & ""






''  #End If
'
'  Set db = New ADODB.Connection
'  db.CursorLocation = adUseClient
'  db.Open ConStr
'
'  Sql = '    "Select * from COM_Company_Master a Left Outer Join COM_Company_Finyear  b On a.CompID=b.CompID and a.BranchID=b.BranchID"

'  Sql = Sql & " where a.BranchID='" & Trim$( '    cmbBranchID.Text) & "' and a.CompID='" & Trim$(cmbCompID.Text) & "'"

'  Sql = Sql & " and Convert(varchar(10), b.FromYear, 105)='" & "01-" & Left$( '    Trim$(cmbFinYear.Text), 7) & "' and Convert(varchar(10), b.ToYear, 105)='" & '    "31-" & Right$(Trim$(cmbFinYear.Text), 7) & "'"


'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then
'    GoTo ExitHere
'  End If
'
'  If rs.EOF Then
'    sShowMessage "No Branch / Company is found for the given Financial Year."
'    GoTo ExitHere
'  Else
'    If IsNull(rs(0)) Then
'      sShowMessage "No Branch / Company is found for the given Financial Year."
'      GoTo ExitHere
'    Else
'      gBranchID = Trim$(cmbBranchID.Text)
'      gCompID = Trim$(cmbCompID.Text)
'      gFinyear = Trim$(cmbFinYear.Text)
'    End If
'  End If
'
''  gServer = Trim$(cmbServer.Text)
''  gDatabase = Trim$(cmbDatabase.Text)
'
'  isSuccess = True
'
'  Unload Me
'
'ExitHere:
'  On Error GoTo 0
'  Set rs = Nothing
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

Private Sub cmdOk_Click()
  Dim ConStr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim ConnectionType As String
  Dim retVal As Variant
  Dim DefaultDB As String
  Dim lstCnt As Integer
  
  Try
  
  'Set Connectiontype here
'  ConnectionType = "NamedPipes"
  ConnectionType = "TCP"
  
  Select Case ConnectionType
    Case "NamedPipes"
      'For Forced Named Pipes Connection use "Network=dbnmpntw"
      ConStr = "Driver={SQL Server};" & "Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & ";" & "Address=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$(txtPort.Text) & ";" & "Network=dbnmpntw;" & "Database=master;" & "Pwd=" & Trim$(txtPassword.Text) & ";" & "Uid=" & Trim$(txtUserID.Text) & ""





    Case Else
      'For Forced TCP/IP Connection use "Network=DBMSSOCN"
      ConStr = "Driver={SQL Server};" & "Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & ";" & "Address=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$(txtPort.Text) & ";" & "Network=DBMSSOCN;" & "Database=master;" & "Pwd=" & Trim$(txtPassword.Text) & ";" & "Uid=" & Trim$(txtUserID.Text) & ""





  End Select
  
  Set db1 = New AceADODB.Connection
  db1.CursorLocation = adUseClient
  db1.Open ConStr

  Sql = " Select [Name] as dbName from sysdatabases"
  Sql = Sql & " Where [Name] not in ('master', 'model', 'msdb', 'tempdb', 'northwind', 'pubs')"
  Sql = Sql & " Order by [Name]"
  Set rs = New AceADODB.Recordset
  If LocalRSOpen(rs, Sql, , , , db1) = False Then
    GoTo ExitHere
  End If
  
    cmbDatabase.Clear
    Do While Not rs.EOF
      cmbDatabase.AddItem rs("dbName")
      rs.MoveNext
    Loop
    
    If DefaultDB <> "" Then
      sprUserData.GetText eINI.mVersion, cmbServer.ListIndex + 1, retVal
      If App.Major = Val(retVal) Then
        For lstCnt = 0 To cmbDatabase.ListCount - 1
          If DefaultDB = UCase(cmbDatabase.List(lstCnt)) Then
            cmbDatabase.ListIndex = lstCnt
            Exit For
          End If
        Next lstCnt
      End If
    End If
  
  CmdOk.Enabled = False
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub cmdOpen_Click()

  Try
  Me.MousePointer = vbHourglass

  If GetCompIDBranchidFinyear = False Then
    GoTo ExitHere
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
' On Error Resume Next
  
  If ReadFromFile = False Then
    Exit Sub
  End If
  
  If gModuleCode = eMdlCode.mACE Then
    chkShowLogin.Enabled = True
  Else
    chkShowLogin.Enabled = False
  End If
  
  CmdOk.Enabled = True
  
  isSuccess = False
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  Set db1 = Nothing
End Sub

Private Sub spList_Click(sender as object, e as ClickEventArgs)
  Try
  Me.MousePointer = vbHourglass

  If spList.DataRowCnt = 0 Then GoTo ExitHere

  If GetUserType = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Function ConnectToDatabase() As Boolean

  Try
 
  Me.Show vbModal
  ConnectToDatabase = isSuccess
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Decrypt(ByVal dtext As String) As String
  ' procedure to decrypt the existing password
Dim    destr As String
Dim  dechar As String
Dim  dttext As String

Dim    Count As Integer
Dim  nval As Integer
Dim  slen As Integer
Dim  dd As Integer

  
' On Error Resume Next

  slen = Len(dtext)
  If slen > 2 Then
    dttext = Right(dtext, 2)
    dd = Val(Chr(Asc(Mid(dttext, 1, 1)) - 110) + Chr(Asc(Mid(dttext, 2, 1)) - 110))

  
    destr = ""
    For Count = 1 To slen - 2
        nval = Asc(Mid(dtext, Count, 1))
        dechar = Chr(nval - dd - Count)
        destr = destr + dechar
    Next Count
  End If
  Decrypt = destr
End Function

Private Function ConnectDatabase() As Boolean
  Dim ConStr As String
  
  Try
  ConnectDatabase = False
  
  If cmbDatabase.ListIndex < 0 Then
    sShowMessage "Please select the Database"
    cmbDatabase.SetFocus
    GoTo ExitHere
  End If
  
'  #If isWeb = False Then
'    ConStr = "Provider=SQLOLEDB.1;uid=" & Trim$(txtUserID.Text) & ";pwd=" & Trim$(txtPassword.Text) & ";database=" & Trim$(cmbDatabase.Text) & ";server=" & Trim$(cmbServer.Text) & ""
'  #Else
'    ConStr = "Driver={SQL Server};" & '    "Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & ";" & '    "Address=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$(txtPort.Text) & ";" & '    "Network=DBMSSOCN;" & '    "Database=" & Trim$(cmbDatabase.Text) & ";" & '    "Pwd=" & Trim$(txtPassword.Text) & ";" & '    "Uid=" & Trim$(txtUserID.Text) & ""






'  #End If
  
'  If gModuleCode = eMdlCode.mFAS Then
'
'    ConStr = "Driver={SQL Server Native Client 10.0};Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & '              "," & IIf(Trim$(txtPort.Text) = "", 1433, Val(txtPort.Text)) & ";Database=" & Trim$(cmbDatabase.Text) & ";Uid=" & Trim$(txtUserID.Text) & ";Pwd=" & Trim$(txtPassword.Text) & ";"

'  Else
'    ConStr = "Provider=SQLNCLI10;Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & '              "," & IIf(Trim$(txtPort.Text) = "", 1433, Val(txtPort.Text)) & ";Database=" & Trim$(cmbDatabase.Text) & ";Uid=" & Trim$(txtUserID.Text) & "; Pwd=" & Trim$(txtPassword.Text) & ";"

'  End If
  
  
  ConStr = "Provider=SQLNCLI10;Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & IIf(Trim$(txtPort.Text) = "", 1433, Val(txtPort.Text)) & ";Database=" & Trim$(cmbDatabase.Text) & ";Uid=" & Trim$(txtUserID.Text) & "; Pwd=" & Trim$(txtPassword.Text) & ";"

  
  
  'Commented by Anandh.s on 2011-12-01 (Unknown usage)
'  toServer = IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text))
'  toDB = Trim$(cmbDatabase.Text)
'  toUID = Trim$(txtUserID.Text)
'  toPWD = Trim$(txtPassword.Text)
  
  Set db = New AceADODB.Connection
  db.CursorLocation = adUseClient
  db.ConnectionTimeout = 100            'db.ConnectionTimeout = 600 - changed on 25/11/2025
  db.CommandTimeout = 300
  
  db.Open ConStr
  
'  db.Close
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function GetCompIDBranchidFinyear() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  Dim i As Integer
  
  Try
  GetCompIDBranchidFinyear = False
  
  CurDate = Now
  
  If Trim$(cmbDatabase) = "" Then Exit Function
  
  If chkLatestFinyear Then
    Sql = " Select a.CompanyName , isnull(b.CompanyName,'') as BranchName, isnull(b.Branchid, a.Branchid) as Branchid, a.Compid, "
    Sql = Sql & " (Select Top 1 right(convert(varchar(10),c.FromYear,105),7) + ' ' + right(convert(varchar(10),c.ToYear,105),7)"
    Sql = Sql & " from " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_finyear('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') c"
    Sql = Sql & " Where isnull(b.Compid,a.Compid)=c.Compid and isnull(b.Branchid,a.Branchid)=c.Branchid"
    Sql = Sql & " Order By c.ToYear desc) as Finyear"
    Sql = Sql & " From " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_master('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') a"
    Sql = Sql & " left outer join " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_master('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') b"
    Sql = Sql & " on a.Compid=b.Compid and b.BranchID <> '00'"
    Sql = Sql & " Where a.BranchID='00'"
    Sql = Sql & " Order by a.CompID, isnull(b.Branchid,a.Branchid)"
  Else
    Sql = " Select a.CompanyName , isnull(b.CompanyName,'') as BranchName, isnull(b.Branchid,a.Branchid) as Branchid, a.Compid,"
    Sql = Sql & " right(convert(varchar(10),c.FromYear,105),7) + ' ' + right(convert(varchar(10),c.ToYear,105),7) as Finyear"
    Sql = Sql & " From " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_master('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') a"
    Sql = Sql & " left outer join " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_master('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') b"
    Sql = Sql & " on a.Compid=b.Compid and b.BranchID <> '00'"
    Sql = Sql & " left outer join " & Trim$(cmbDatabase.Text) & ".dbo.fn_com_company_finyear('A', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "') c"
    Sql = Sql & " on isnull(b.Compid,a.Compid)=c.Compid and isnull(b.Branchid,a.Branchid)=c.Branchid"
    Sql = Sql & " Where a.BranchID='00'"
    Sql = Sql & " Order by  a.Compid, isnull(b.Branchid,a.Branchid), c.ToYear desc, c.FromYear"
  End If
  Set rs = New AceADODB.Recordset
  If LocalRSOpen(rs, Sql, , , , db1) = False Then
    GoTo ExitHere
  End If

'  rs.Open Sql, db1, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
      Spread_Clear spList
      rs.MoveLast
rs.MoveFirst
      spList.MaxRows = rs.RecordCount
      For i = 1 To rs.RecordCount
        spList.SetText 1, i, Trim$(rs("CompanyName"))
        spList.SetText 2, i, Trim$(rs("BranchName"))
        spList.SetText 3, i, Trim$(rs("Branchid"))
        spList.SetText 4, i, Trim$(rs("Compid"))
        spList.SetText 5, i, Trim$(rs("Finyear"))
        rs.MoveNext
      Next i
    End If

  GetCompIDBranchidFinyear = True
  
  Call spList_Click(1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function GetUserName() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  GetUserName = False
  
  If Trim$(cmbUserType.Text) <> "" Then
    CurDate = Now
    
    Sql = "Select UserId, UserName"
    Sql = Sql & " from dbo.fn_Sec_Users('" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Trim$(SpVal(spList, spList.ActiveRow, 3)) & "', '" & Trim$(SpVal(spList, spList.ActiveRow, 4)) & "')"
    Sql = Sql & " Where UserGroupID='" & Trim$(cmbUserType.Text) & "'"
    Sql = Sql & " Order by UserName"
    If LocalRSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
      cmbUserName.Clear
      cmbERPUserID.Clear
      
      If Trim$(cmbUserType.Text) = "D" Then
        cmbUserName.AddItem "Developer"
        cmbERPUserID.AddItem -1
      End If
      
      If Not rs.EOF Then
        Do While Not rs.EOF
          cmbUserName.AddItem rs("UserName")
          cmbERPUserID.AddItem rs("UserId")
          rs.MoveNext
        Loop
      End If
      cmbUserName.ListIndex = 0
      cmbERPUserID.ListIndex = 0
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function GetUserType() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  GetUserType = False
  
  If ConnectDatabase = False Then
    cmdConnect.Enabled = False
    GoTo ExitHere
  Else
    cmdConnect.Enabled = True
  End If
  
  CurDate = Now
  
  Sql = " Select Distinct UserGroupId "
  Sql = Sql & " from dbo.fn_sec_users('" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Trim$(SpVal(spList, spList.ActiveRow, 3)) & "', '" & Trim$(SpVal(spList, spList.ActiveRow, 4)) & "')"
  If LocalRSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
    cmbUserType.Clear
    cmbUserType.AddItem "D"
    If Not rs.EOF Then
      Do While Not rs.EOF
        cmbUserType.AddItem rs("UserGroupId")
        rs.MoveNext
      Loop
      cmbUserType.ListIndex = 0
    End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function LocalRSOpen(ByRef RecSet As AceADODB.Recordset, ByVal SqlStr As String,Optional ByVal RSLocation As CursorLocationEnum = adUseClient,Optional ByVal RecOpenType As AceADODB.CursorTypeEnum = adOpenStatic,Optional ByVal RecLockType As AceADODB.LockTypeEnum = adLockReadOnly,Optional ByVal dbCon As AceADODB.Connection) As Boolean





  Try
  LocalRSOpen = False
  
  If RecSet Is Nothing Then
    'RecSet not instantiated
    Set RecSet = New AceADODB.Recordset
  Else
    'Already RecSet Instantiated
  End If

  If RecSet.State = adStateOpen Then
    RecSet.Filter = adFilterNone
    RecSet.Close
  End If
  
  RecSet.CursorLocation = RSLocation
  
  If dbCon Is Nothing Then Set dbCon = db
  If dbCon.State = adStateClosed Then dbCon.Open
  
  RecSet.Open SqlStr, dbCon, RecOpenType, RecLockType
  
  If RecSet.State = adStateOpen Then
    If RSLocation = adUseClient Then RecSet.ActiveConnection = Nothing
    
    LocalRSOpen = True
  Else
    Call MsgBox("Error: Unable to Open Recordset.", vbOKOnly + vbCritical, "Error")
  End If
  
'  If gTransactionCount <= 0 Then dbCon.Close
Catch ex As Exception
  Select Case Err.Number
    Case 3246       'db Connection is in Middle of Transaction.
         'So no need to close
    Case Else
      Call ErrMsgShow("Error Occured while Opening RecordSet.")

  End Select
Finally

End Try
End Function

Private Function ReadFromFile() As Boolean
  Dim fso As FileSystemObject
  Dim fsoTxt As TextStream
  Dim cnt As Long
  Dim Cnt2 As Integer
  Dim txtstr As String
  Dim tmp As String
  
  Try
  ReadFromFile = False

  Set fso = New FileSystemObject
  
  tmp = fso.GetParentFolderName(App.Path)
  Set fsoTxt = fso.OpenTextFile(tmp & "\Common\Ace.ini", ForReading, False, TristateTrue)

  
    sprUserData.MaxRows = 1
    cmbServer.Clear
cmbServerIP.Clear
    
    Do While Not fsoTxt.AtEndOfStream
      For cnt = eINI.mVersion To eINI.mDefaultDB
        If Not fsoTxt.AtEndOfStream Then
          txtstr = fsoTxt.ReadLine
          tmp = Decrypt(txtstr)
          'To Skip Server Details which are not compatible with the version
          If cnt = eINI.mVersion Then
            If Val(tmp) <> Val(App.Major) Then
              For Cnt2 = eINI.mVersion + 1 To eINI.mDefaultDB
fsoTxt.SkipLine
Next Cnt2
              Exit For
            End If
          End If
          If cnt = eINI.mServer Then
            cmbServer.AddItem tmp
          ElseIf cnt = eINI.mIPAddr Then
            cmbServerIP.AddItem tmp
          End If
          sprUserData.SetText cnt, sprUserData.MaxRows, tmp
        End If
      Next cnt
      sprUserData.MaxRows = sprUserData.DataRowCnt + 1
    Loop
    
    If cmbServer.ListCount > 0 Then
      cmbServer.ListIndex = 0
    End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set fsoTxt = Nothing
  Set fso = Nothing
End Try
End Function

Private Function RefreshDatabases() As Boolean
  Dim ConStr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim ConnectionType As String
  Dim retVal As Variant
  Dim DefaultDB As String
  Dim lstCnt As Integer
  
  Try
  
  'Set Connectiontype here
'  ConnectionType = "NamedPipes"
  ConnectionType = "TCP"
  
  Select Case ConnectionType
    Case "NamedPipes"
      'For Forced Named Pipes Connection use "Network=dbnmpntw"
      ConStr = "Driver={SQL Server};" & "Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & ";" & "Address=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$(txtPort.Text) & ";" & "Network=dbnmpntw;" & "Database=master;" & "Pwd=" & Trim$(txtPassword.Text) & ";" & "Uid=" & Trim$(txtUserID.Text) & ""





    Case Else
      'For Forced TCP/IP Connection use "Network=DBMSSOCN"
      ConStr = "Driver={SQL Server};" & "Server=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & ";" & "Address=" & IIf(cmbServer.ListIndex = -1, Trim$(cmbServer.Text), Trim$(cmbServerIP.Text)) & "," & Trim$(txtPort.Text) & ";" & "Network=DBMSSOCN;" & "Database=master;" & "Pwd=" & Trim$(txtPassword.Text) & ";" & "Uid=" & Trim$(txtUserID.Text) & ""





  End Select
  
  Set db1 = New AceADODB.Connection
  db1.CursorLocation = adUseClient
  db1.Open ConStr

  Sql = " Select [Name] as dbName from sysdatabases"
  Sql = Sql & " Where [Name] not in ('master', 'model', 'msdb', 'tempdb', 'northwind', 'pubs')"
  Sql = Sql & " Order by [Name]"
  Set rs = New AceADODB.Recordset
  If LocalRSOpen(rs, Sql, , , , db1) = False Then
    GoTo ExitHere
  End If
  
    
  
  
    cmbDatabase.Clear
    Do While Not rs.EOF
      cmbDatabase.AddItem rs("dbName")
      rs.MoveNext
    Loop
    
    If DefaultDB <> "" Then
      sprUserData.GetText eINI.mVersion, cmbServer.ListIndex + 1, retVal
      If App.Major = Val(retVal) Then
        For lstCnt = 0 To cmbDatabase.ListCount - 1
          If DefaultDB = UCase(cmbDatabase.List(lstCnt)) Then
            cmbDatabase.ListIndex = lstCnt
            Exit For
          End If
        Next lstCnt
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList


spList.Click +=  new EventHandler(spList_Click)

End Sub

Private Sub Design_sprUserData(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprUserData

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\frmDBConnect.frm", spList, "")

	Call Design_sprUserData("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\frmDBConnect.frm", sprUserData, "")

End Sub
