VERSION 5.00
Begin VB.Form rptSalesorderRev 
   Caption         =   "Sales Order Revision Summary"
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
   Begin VB.Frame FraStep 
      Height          =   720
      Index           =   4
      Left            =   3825
      TabIndex        =   10
      Top             =   465
      Width           =   3555
      Begin VB.CommandButton cmdStep 
         Caption         =   ">"
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   3135
         TabIndex        =   14
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   1050
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Order Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   17
         Top             =   135
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Order No."
         Height          =   210
         Left            =   960
         TabIndex        =   16
         Top             =   135
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         Height          =   210
         Left            =   2010
         TabIndex        =   15
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4140
      TabIndex        =   5
      Top             =   1215
      Width           =   2925
      Begin VB.ComboBox cmbOrderby 
         Height          =   330
         ItemData        =   "frmSalesorderRev.frx":0000
         Left            =   810
         List            =   "frmSalesorderRev.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   615
         Width           =   1995
      End
      Begin VB.ComboBox cmbType 
         Height          =   330
         ItemData        =   "frmSalesorderRev.frx":0029
         Left            =   810
         List            =   "frmSalesorderRev.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   225
         Width           =   1995
      End
      Begin VB.Label lblOrderby 
         AutoSize        =   -1  'True
         Caption         =   "Order by"
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   645
         Width           =   645
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   210
         Left            =   405
         TabIndex        =   8
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   3607
      TabIndex        =   0
      Top             =   2295
      Width           =   3990
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   2625
         TabIndex        =   1
         Top             =   165
         Width           =   1300
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   400
         Left            =   1335
         TabIndex        =   2
         Top             =   165
         Width           =   1300
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Enabled         =   0   'False
         Height          =   400
         Left            =   45
         TabIndex        =   4
         Top             =   165
         Visible         =   0   'False
         Width           =   1300
      End
      Begin VB.CommandButton cmdPnt 
         Caption         =   "&Print"
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
         Left            =   45
         TabIndex        =   3
         Top             =   165
         Width           =   1300
      End
   End
End
Attribute VB_Name = "rptSalesorderRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
' On Error Resume Next
  Call StepClear(Me)
  cmdView.Enabled = False
  cmdStep(4).Enabled = True
  cmbType.ListIndex = 0
  cmbOrderby.ListIndex = 0
End Sub

Private Sub cmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  cmdView.Enabled = True
  cmdStep(Index).Enabled = False
  cmbType.ListIndex = 0
  cmbOrderby.ListIndex = 0
End Sub

Private Sub cmdView_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass
  
'  Select * from
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtType_DblClick()
  Dim Sql As String
    
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct OrderType, OrderNo, OrderDate"
  
  Sql = Sql & " From SAL_Order_Hdr"
'  If optInnerJobWork Then: Sql = Sql & " From SAL_JobworkOrder_Hdr"
  
  Sql = Sql & " Where Branchid='" & gBranchID & "'"
  Sql = Sql & " and Compid = '" & gCompID & "' and finyear = '" & gFinyear & "'"
  Sql = Sql & " Order by OrderDate desc, OrderNo desc"
  
  If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of Order No's", "Order Type", "Order No", "Order Date") = True Then
    txtType.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)
    If Trim$(txtType.Text) <> "" Then
 cmdStep(4).Enabled = True
End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = 0
End Try
End Sub

