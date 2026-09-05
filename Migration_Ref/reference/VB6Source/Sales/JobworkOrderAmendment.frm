VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobworkOrderAmendment 
   Caption         =   "Jobwork Order Amendment"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
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
   ScaleHeight     =   7485
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFinyear 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   330
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   56
      Top             =   6990
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Frame fraSch 
      Height          =   1875
      Left            =   315
      TabIndex        =   46
      Top             =   5070
      Visible         =   0   'False
      Width           =   7260
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Top             =   195
         Width           =   2175
      End
Begin AceSpread spSchedule
         Height          =   1185
         Left            =   60
         TabIndex        =   48
         Top             =   600
         Width           =   7125
         _Version        =   458752
         _ExtentX        =   12568
         _ExtentY        =   2090
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   1
         ScrollBars      =   2
      End
      Begin VB.Label lblItemCode 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Item Code"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   49
         Top             =   240
         Width           =   855
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5160
      TabIndex        =   18
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3870
      TabIndex        =   17
      ToolTipText     =   "Click to Save the Data"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.Frame fraAmd 
      Height          =   720
      Left            =   375
      TabIndex        =   24
      Top             =   195
      Width           =   3045
      Begin VB.ComboBox cmbamdtype 
         Height          =   330
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Amendment Type"
         Top             =   315
         Width           =   870
      End
      Begin VB.TextBox txtamdno 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Amendment No"
         Top             =   315
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dpamddate 
         Height          =   315
         Left            =   1725
         TabIndex        =   2
         ToolTipText     =   "Amendment Date"
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   22347777
         CurrentDate     =   37262
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Amd. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1725
         TabIndex        =   27
         Top             =   135
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Amd. No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1095
         TabIndex        =   26
         Top             =   135
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Amd. Type"
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   135
         Width           =   885
      End
   End
   Begin VB.Frame fraOrder 
      Height          =   720
      Left            =   3435
      TabIndex        =   20
      Top             =   195
      Width           =   8130
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   7515
         TabIndex        =   7
         ToolTipText     =   "Click to OK"
         Top             =   315
         Width           =   480
      End
      Begin VB.TextBox txtOrdType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Jobwork Order Type"
         Top             =   315
         Width           =   795
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Customer"
         Top             =   315
         Width           =   4590
      End
      Begin VB.TextBox txtordnum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Select the Jobwork Order. Press F2 or Double Click for Help"
         Top             =   315
         Width           =   870
      End
      Begin VB.TextBox txtorddt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Jobwork Order Date"
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2880
         TabIndex        =   28
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Order Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1800
         TabIndex        =   23
         Top             =   135
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Order No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   930
         TabIndex        =   22
         Top             =   135
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "OrderType"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   21
         Top             =   135
         Width           =   780
      End
   End
   Begin VB.Frame fraTerms 
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
      Height          =   1710
      Left            =   360
      TabIndex        =   29
      Top             =   885
      Width           =   11220
      Begin VB.TextBox txtordertyping 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9195
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   54
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   1230
      End
      Begin VB.TextBox txtFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   15
         ToolTipText     =   "Select the Formula. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   2760
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Select the Packing & Forwarding. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtForm 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Select the Form. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtTerms 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Select the Tearms of Payment. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3675
      End
      Begin VB.TextBox txtInsurance 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtFreight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtDelIns 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Select the Delivery Instruction. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   6555
         TabIndex        =   50
         ToolTipText     =   "Enter the Valid From Date"
         Top             =   1305
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
         Format          =   22347779
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   7935
         TabIndex        =   51
         ToolTipText     =   "Enter the Valid Till Date"
         Top             =   1305
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   22347779
         CurrentDate     =   37250
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OrderTyping"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   9195
         TabIndex        =   55
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valid Till"
         Height          =   210
         Left            =   7935
         TabIndex        =   53
         Top             =   1110
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
         Height          =   210
         Left            =   6555
         TabIndex        =   52
         Top             =   1110
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   37
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Instruction"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7455
         TabIndex        =   36
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   35
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Freight"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   34
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Packing && Forwarding"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   33
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Form"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   32
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Insurance"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   31
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Mode"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7440
         TabIndex        =   30
         Top             =   630
         Width           =   390
      End
   End
Begin AceSpread sprItem
      Height          =   1920
      Left            =   360
      TabIndex        =   16
      ToolTipText     =   "Enter the Jobwork Amendment Details"
      Top             =   2610
      Width           =   11205
      _Version        =   458752
      _ExtentX        =   19764
      _ExtentY        =   3387
      _StockProps     =   64
      Enabled         =   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7965
      TabIndex        =   38
      Top             =   4425
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   39
         Top             =   135
         Width           =   900
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   150
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   41
         Top             =   150
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   40
         Top             =   2580
         Visible         =   0   'False
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   2085
         Left            =   45
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   3570
         _Version        =   458752
         _ExtentX        =   6297
         _ExtentY        =   3678
         _StockProps     =   64
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
         MaxCols         =   2
         ScrollBars      =   2
      End
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   45
         Top             =   165
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   975
         TabIndex        =   44
         Top             =   2625
         Visible         =   0   'False
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmJobworkOrderAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim SKey As String

'Item Spread Column Numbers
Private Enum eItem
  mFirstRow = 1
  mrecdItem = 2
  mRecdDesc = 3
  mRecdUOM = 4
  msentItem = 5
  mSentDesc = 6
  mPartNo = 7
  mSentUOM = 8
  mSentQty = 9
  mRatePer = 10
  mRate = 11
  mAssRate = 12
  mGoodsValue = 13
  mAssGoodsValue = 14
  mStatus = 15
  mRelation = 16
  mItemGroupKey = 17
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Enum eSch
  mSchDate = 1
  mSchQty = 2
  mLocation = 3
End Enum

Private Function ManyToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim sCnt As Long
  Dim retVal As Variant
  Dim sNo As Integer
  Dim sParentCode As String
  Dim sStageCode As String
  Dim sChildCode As String
  Dim DubLicate As Boolean
  Dim srow As Long
  Dim i As Long
  Dim j As Long
  
  Try
  DubLicate = False
  
  
    If fRow > sprItem.DataRowCnt Then     'already saved itemcode should not be changed
      'Fill the Child Items for the selected Item Structure Details in SentItem
        
      Sql = "Select Distinct b.SentItemCode, c.ItemDescription as ChildDesc,b.ItemGroupKey,b.PartNo,b.Rate,c.UOMDesc as UOM,c.UOMDecimal from SAl_JobWork_ItemMaster b"
      Sql = Sql & " Left Outer Join MATERIAL_MASTER c  On b.SentItemCode=c.ItemCode and b.BranchID=c.BranchID"
      Sql = Sql & " and b.CompID=c.CompID where CustomerCode=" & txtcustcd.Tag & " and RelationType='N-1'"
      Sql = Sql & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "' Order By b.ItemGroupKey"
      If FetchValue_New(True, db.ConnectionString, Sql, "111000", "Select Received Item", "Item Code", "Description", "Set No") = True Then
        For i = 1 To UBound(getvalue, 1)
          sCnt = sprItem.DataRowCnt + 1
          sprItem.SetText eItem.msentItem, sCnt, getvalue(i, 1)
          sprItem.SetText eItem.mSentDesc, sCnt, CheckNull(getvalue(i, 2))
          sprItem.SetText eItem.mSentUOM, sCnt, CheckNull(getvalue(i, 6))
          sprItem.SetText eItem.mRatePer, sCnt, 1
          sprItem.SetText eItem.mFirstRow, sCnt, Val(sCnt)
          
          If txtordertyping.Text <> "Open" Then
            Spread_Lock sprItem, False, eItem.mSentQty, sCnt, eItem.mAssRate, sCnt
          Else
            Spread_Lock sprItem, False, eItem.mRate, sCnt, eItem.mAssRate, sCnt
          End If
          
          Call Spread_Set_Decimal(sprItem, Val(CheckNull(getvalue(i, 7), 0)), eItem.mRatePer, sCnt)
        
          Sql = "Select Distinct b.ReceivedItemCode, c.ItemDescription as ChildDesc,b.ItemGroupKey, c.UOMDesc as UOM from SAl_JobWork_ItemMaster b"
          Sql = Sql & " Left Outer Join MATERIAL_MASTER c  On b.ReceivedItemCode=c.ItemCode and b.BranchID=c.BranchID"
          Sql = Sql & " and b.CompID=c.CompID where ItemGroupKey =" & getvalue(i, 3) & " "
          Sql = Sql & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "' Order By b.ItemGroupKey"
          Set rs = New AceADODB.Recordset
          If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
          sprItem.MaxRows = sprItem.DataRowCnt + rs.RecordCount
          sChildCode = ""
          For j = 1 To rs.RecordCount
            sChildCode = sChildCode + rs("ReceivedItemCode")
            rs.MoveNext
          Next j
            
          srow = sprItem.SearchCol(eItem.msentItem, 0, sprItem.MaxRows, Trim$(getvalue(i, 1)), SearchFlagsValue)
          Do While srow <> -1
            If Group_Child(srow, Trim$(getvalue(i, 1))) = sChildCode Then
              DubLicate = True
              sprItem.DeleteRows sCnt, 1
              Exit Do
            End If
            srow = sprItem.SearchCol(eItem.msentItem, srow, sprItem.MaxRows, Trim$(getvalue(i, 1)), SearchFlagsValue)
          Loop
          If DubLicate = False Then
            rs.MoveLast
rs.MoveFirst
            rCnt = sprItem.DataRowCnt
            sprItem.SetText eItem.mRate, rCnt, getvalue(i, 5)
            sprItem.SetText eItem.mAssRate, rCnt, getvalue(i, 5)
            sprItem.SetText eItem.mAssRate, rCnt, getvalue(i, 5)
            sprItem.SetText eItem.mStatus, rCnt, "Active"
            For rCnt = sprItem.DataRowCnt To rs.RecordCount + sprItem.DataRowCnt - 1
              sprItem.SetText eItem.mFirstRow, rCnt, Val(sCnt)
              sprItem.SetText eItem.mrecdItem, rCnt, rs("ReceivedItemCode")
              sprItem.SetText eItem.mRecdDesc, rCnt, rs("ChildDesc")
              sprItem.SetText eItem.mRecdUOM, rCnt, rs("UOM")
              sprItem.SetText eItem.mRelation, rCnt, "N-1"
              sprItem.SetText eItem.mItemGroupKey, rCnt, getvalue(i, 3)
                 rs.MoveNext
            Next rCnt
          Else
            sprItem.MaxRows = sprItem.DataRowCnt + 1
          End If
        Next i
      End If
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function Group_Child(ByVal fRow As Long,ByVal  ParentCode As String) As String
  Dim i As Integer
  Dim retVal As Variant
  Dim Row As Long
  Dim srow As Long
  Dim ChildCode As String
  
  Try
  Me.MousePointer = vbHourglass
  
    srow = sprItem.SearchCol(eItem.mFirstRow, 0, sprItem.MaxRows, Val(fRow), SearchFlagsValue)
    ChildCode = ""
    Do While srow <> -1
      sprItem.GetText eItem.mrecdItem, srow, retVal
      ChildCode = ChildCode + retVal
      srow = sprItem.SearchCol(eItem.mFirstRow, srow, sprItem.MaxRows, Trim$(fRow), SearchFlagsValue)
    Loop
  
  Group_Child = ChildCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

Private Sub SetSheet(ByVal fRow As Long)
' On Error Resume Next
    If spSchedule.SheetCount < fRow Then
      spSchedule.SheetCount = fRow
      Call CopySpreadStructure(spSchedule, 1, , fRow)
    End If
    spSchedule.Sheet = fRow
    spSchedule.ActiveSheet = fRow
End Sub

Private Function ValidateData() As Boolean
  Dim retVal As Variant
  Dim flg As Boolean
  Dim i As Long
  Dim sCnt As Long
  Dim Sql As String
  Dim CurDate As Date
  Dim srow As Long
  Dim rRow As Long
  
  Try
 
  CurDate = ServerDate
  
  flg = False
  ValidateData = False
  
  For i = 1 To sprItem.DataRowCnt
    sprItem.GetText eItem.mFirstRow, i, retVal
    srow = retVal
    If cmdOk.Tag <> "Open" Then
    sprItem.GetText eItem.mSentQty, srow, retVal
      If Val(retVal) <= Val(Spread_Get_CellTag(sprItem, eItem.mSentUOM, srow)) Then
        sShowMessage "Amend Qty should be greater than Invoice Qty (" & Val(Spread_Get_CellTag(sprItem, eItem.mSentUOM, srow)) & ")"
        sprItem.SetFocus
        Spread_Set_Focus sprItem, srow, eItem.mSentQty
        GoTo ExitHere
      End If
    End If
  Next i
  
  For i = 1 To sprItem.DataRowCnt
    sprItem.GetText eItem.mFirstRow, i, retVal
    srow = retVal
    spSchedule.Sheet = srow
    If fraSch.Visible = True And rRow <> srow Then
      If spSchedule.DataRowCnt = 0 Then
        sShowMessage "Enter the Schedule Details"
        GoTo ExitHere
      End If
      For sCnt = 1 To spSchedule.DataRowCnt
        spSchedule.GetText eSch.mSchDate, sCnt, retVal
        If Len(retVal) = 0 Then
          sShowMessage "Enter the ScheduleDate"
          spSchedule.ActiveSheet = srow
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
          GoTo ExitHere
        End If
        spSchedule.GetText eSch.mSchDate, sCnt, retVal
        If IsDate(retVal) = 0 Then
          sShowMessage "invalid ScheduleDate"
          spSchedule.ActiveSheet = srow
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
          GoTo ExitHere
        End If
        spSchedule.GetText eSch.mSchDate, sCnt, retVal
        If sCnt > spSchedule.GetColItemData(eSch.mLocation) Then
          If Format(retVal, "yyyyMMdd") < Format(CurDate, "yyyyMMdd") Then
            sShowMessage "Schedule Date Should be Creater then Or Equal to CurentDate"
            spSchedule.ActiveSheet = srow
            Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
            GoTo ExitHere
          End If
        End If
        spSchedule.GetText eSch.mSchQty, sCnt, retVal
        If Val(retVal) = 0 Then
          sShowMessage "Enter the ScheduleQty"
          spSchedule.ActiveSheet = srow
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchQty
          GoTo ExitHere
        End If
      Next sCnt
      sprItem.GetText eItem.mSentQty, srow, retVal
      If Val(retVal) <> Spread_Cell_Total(spSchedule, eSch.mSchQty, 1, eSch.mSchQty, spSchedule.DataRowCnt) Then
        sShowMessage "Schedule TotalQty Should be OrderQty"
        spSchedule.ActiveSheet = srow
        Spread_Set_Focus spSchedule, 1, eSch.mSchQty
        GoTo ExitHere
      End If
    End If
    rRow = srow
  Next i
  flg = False
  For i = 1 To sprItem.DataRowCnt
    sprItem.GetText eItem.mFirstRow, i, retVal
    srow = retVal
    If Trim$(SpVal(sprItem, srow, eItem.mStatus)) = "Active" Then
      flg = True
      Exit For
    End If
  Next i

  If flg = False Then
    sShowMessage "Atlease one Item should be Active"
    sprItem.SetFocus
    Spread_Set_Focus sprItem, 1, eItem.mStatus
    GoTo ExitHere
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function DataChanged(ByVal fRow As Long) As Boolean
  Dim OldValue As Variant
  Dim NewValue As Variant
  Dim cCnt As Long
  
  Try
  DataChanged = False
  
    sprItem.Row = fRow
    For cCnt = 1 To eItem.mStatus
      If cCnt <= eItem.mAssRate Or cCnt = eItem.mStatus Then
        OldValue = Spread_Get_CellTag(sprItem, cCnt, fRow)
        sprItem.GetText cCnt, fRow, NewValue
        If Trim$(OldValue) <> Trim$(NewValue) Then
          DataChanged = True
          GoTo ExitHere
        End If
      End If
    Next cCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub OrderCalulation()
  Dim tot As Double
  Dim atot As Double
  Dim i As Long
  Dim Sql As String
  Dim snap As AceADODB.Recordset

  Try

  If sprItem.DataRowCnt = 0 Then
 GoTo ExitHere
End If
  
  tot = 0
atot = 0
  For i = 1 To sprItem.DataRowCnt
    If Trim$(SpVal(sprItem, i, eItem.mStatus)) = "Active" Then
      tot = tot + Val(SpVal(sprItem, i, eItem.mGoodsValue))
      atot = atot + Val(SpVal(sprItem, i, eItem.mAssGoodsValue))
    End If
  Next i
  
  If txtFormula <> "" Then
'    Calculation Tot, atot, Val(txtFormula.Tag), SKey  'To calculate the formula
'
'    Sql = " select a.Code,(select b.description from com_formula_parameter b "
'    Sql = Sql + " where b.code=a.Code and b.Branchid=a.Branchid and b.Compid=a.Compid) Description,"
'    Sql = Sql + " a.amt,a.Sequence,a.Printing from com_formula_tmp a where a.FormulaCode=" & Val(txtFormula.Tag) & ""
'    Sql = Sql + " and a.printing=1 and SKey = '" + SKey + "' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
'    Sql = Sql + " order by a.sequence"
'    Set snap = New ADODB.Recordset
'    If RSOpen(snap, Sql) = False Then GoTo ExitHere
'    If Not snap.EOF Then
'      snap.MoveLast: snap.MoveFirst
'      spCalc.MaxRows = snap.RecordCount
'      For i = 1 To snap.RecordCount
'        spCalc.SetText 1, i, snap("Code")
'        spCalc.SetText 2, i, snap("Description")
'        spCalc.SetText 3, i, snap("Amt")
'        spCalc.SetText 4, i, snap("Sequence")
'        spCalc.SetText 5, i, snap("Printing")
'        snap.MoveNext
'      Next i
'    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Sub

Private Sub TermsDetails(ByVal CurDate As Date)
  Dim snap As New AceADODB.Recordset
  Dim Sql As String

  Try
  
  Sql = " select Distinct a.FormCode,(select b.Description from TERMSMASTER b where "
  Sql = Sql + " b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid "
  Sql = Sql + " and b.Compid=a.Compid) FormDesc, a.PackingCode,(select b1.Description "
  Sql = Sql + " from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding' "
  Sql = Sql + " and b1.Code=a.PackingCode and b1.Branchid=a.Branchid "
  Sql = Sql + " and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,(select b2.Description "
  Sql = Sql + " from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction' "
  Sql = Sql + " and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid "
  Sql = Sql + " and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,(select b3.Description "
  Sql = Sql + " from TERMSMASTER b3 where b3.TypeFlg='Payment Terms'"
  Sql = Sql + " and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid "
  Sql = Sql + " and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,(select "
  Sql = Sql + " b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance'"
  Sql = Sql + " and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid "
  Sql = Sql + " and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,(select "
  Sql = Sql + " b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid "
  Sql = Sql + " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,(select b6.Description "
  Sql = Sql + " from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode "
  Sql = Sql + " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc,a.FormulaCode,"
  Sql = Sql + " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where "
  Sql = Sql + " b7.ModuleCode = " & gModuleCode & " and b7.FrmID=a.FormulaCode and b7.Branchid=a.Branchid and b7.Compid=a.Compid) FrmLName"
  Sql = Sql + " from dbo.fn_SAL_JobWorkOrder_HDR('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a where a.OrderType ='" + Trim$(txtOrdType) + "'"
  Sql = Sql + " and a.OrderNo = " & Val(txtordnum) & " and a.OrderDate = '" + Format(txtorddt, "yyyy-MM-dd") + "'"
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.CustomerCode = " & Val(txtcustcd.Tag) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    If Not IsNull(snap("FormulaCode")) = True Then txtFormula.Tag = snap("FormulaCode")
    If Not IsNull(snap("FrmLName")) = True Then txtFormula.Text = snap("FrmLName")

    If Not IsNull(snap("DeliveryCode")) = True Then
 txtDelIns.Tag = snap("DeliveryCode")
End If
    If Not IsNull(snap("DeliveryDesc")) = True Then
 txtDelIns.Text = snap("DeliveryDesc")
End If
    
    If Not IsNull(snap("FormCode")) = True Then
 txtForm.Tag = snap("FormCode")
End If
    If Not IsNull(snap("FormDesc")) = True Then
 txtForm.Text = snap("FormDesc")
End If
    
    If Not IsNull(snap("PackingCode")) = True Then
 txtPack.Tag = snap("PackingCode")
End If
    If Not IsNull(snap("PackingDesc")) = True Then
 txtPack.Text = snap("PackingDesc")
End If
    
    If Not IsNull(snap("InsuranceCode")) = True Then
 txtInsurance.Tag = snap("InsuranceCode")
End If
    If Not IsNull(snap("InsuranceDesc")) = True Then
 txtInsurance.Text = snap("InsuranceDesc")
End If

    If Not IsNull(snap("FreightCode")) = True Then
 txtFreight.Tag = snap("FreightCode")
End If
    If Not IsNull(snap("FreightDesc")) = True Then
 txtFreight.Text = snap("FreightDesc")
End If

    If Not IsNull(snap("PaymentCode")) = True Then
 txtTerms.Tag = snap("PaymentCode")
End If
    If Not IsNull(snap("PaymentDesc")) = True Then
 txtTerms.Text = snap("PaymentDesc")
End If

    If Not IsNull(snap("ModeCode")) = True Then
 txtMode.Tag = snap("ModeCode")
End If
    If Not IsNull(snap("ModeDesc")) = True Then
 txtMode.Text = snap("ModeDesc")
End If
  End If
  snap.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbamdtype_Click()
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
  
  txtamdno = GetDocNo("SAL_Order_Amendment", "AmendmentType", "AmendmentNo", Trim$(cmbamdtype), "M")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

'Private Sub LoadAmdtype()
'  Dim amdtyperec As ADODB.Recordset
'  Dim Sql As String
'
'  On Error GoTo ErrHandler
'
'  Set amdtyperec = New ADODB.Recordset
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='M'"
'  Sql = Sql + " and ModuleName='SAL' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
'  If RSOpen(amdtyperec, Sql) = False Then GoTo ExitHere
'  If Not amdtyperec.EOF Then
'     Do Until amdtyperec.EOF
'        cmbamdtype.AddItem amdtyperec("DocumentType")
'        amdtyperec.MoveNext
'     Loop
'     cmbamdtype.ListIndex = 0
'     txtamdno = GetNo("SAL_JOBWORKAMENDMENT_HDR", "AmendmentType", "AmendmentNo", Trim$(cmbamdtype), "M", "SAL")
'  End If
'  amdtyperec.Close
'
'ExitHere:
'  On Error GoTo 0
'    Set amdtyperec = Nothing
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

Private Sub CmdClear_Click()
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  Call LoadDocType(mJWOrderAmendt)
  fraAmd.Enabled = True
  fraOrder.Enabled = True
  fraTerms.Enabled = False
  sprItem.Enabled = False
  cmdsave.Enabled = False
  cmbamdtype.SetFocus
  sprItem.MaxRows = 1
  CurrDate Me
  Call SetDocNo
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

Private Sub cmdOk_Click()
  Dim snap As AceADODB.Recordset
  Dim Schsnap As AceADODB.Recordset
  Dim Sql As String
  Dim tot As Double
  Dim tot1 As Double
  Dim rCnt As Long
  Dim sCnt As Long
  Dim PrevItemGroup As Integer
  Dim ValueRow As Long
  Dim CurDate As Date

  Try
  
  If Trim$(cmbamdtype) = "" Then
    sShowMessage "Please select the Amendment Type"
    cmbamdtype.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtordnum) = "" Then
    sShowMessage "Please select the Order"
    txtordnum.SetFocus
    GoTo ExitHere
  End If

  CurDate = ServerDate
  Set snap = New AceADODB.Recordset
  
  TermsDetails (CurDate) 'To display th terms details
  
  tot = 0
tot1 = 0
  Sql = " Select Distinct a.ItemGroupKey, a.SentItemCode, b.ItemDescription as SentDesc, b.UOMDesc as SentUOM,b.UOMDecimal as NoOfDec,"
  Sql = Sql + " a.ReceivedItemCode,c.ItemDescription as RecdDesc, c.UOMDesc as RecdUOM, "
  Sql = Sql + " a.RatePer,a.rate,a.AssRate,a.OrderQty,(select isnull(Sum(InvoiceQty),0) from SAL_INVOICE_Dtl where OrderType=a.OrderType and OrderNo=a.OrderNo and Itemcode=a.SentItemCode"
  Sql = Sql + " and OrderDate=a.OrderDate and Branchid=a.Branchid and Compid=a.Compid) InvoiceQty,a.Status,(select top 1 d.RelationType from SAL_Jobwork_ItemMaster d where d.ItemGroupKey = a.ItemGroupKey"
  Sql = Sql + " and d.Branchid = a.Branchid and d.Compid = a.Compid and d.CustomerCode = " & Val(txtcustcd.Tag) & ") RelationType,(select top 1 d.PartNo from SAL_Jobwork_ItemMaster d where d.ItemGroupKey = a.ItemGroupKey and d.Branchid = a.Branchid "
  Sql = Sql + " and d.Compid = a.Compid and d.ReceivedItemCode=a.ReceivedItemCode and d.SentItemCode=a.SentItemCode)PartNo,a.OrderingType"
  Sql = Sql + " from dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a Left Outer Join Material_Master b On a.SentItemCode=b.ItemCode"
  Sql = Sql + " and a.BranchID=b.BranchID and a.CompID=b.CompID Left Outer Join Material_Master c"
  Sql = Sql + " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID where"
  Sql = Sql + " a.OrderType='" + Trim(txtOrdType) + "' and a.OrderNo = " & Trim(txtordnum) & ""
  Sql = Sql + " and a.OrderDate = '" + Format(Trim(txtorddt), "yyyy-MM-dd") + "' and b.Status = 1 and c.Status = 1"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
    If Not snap.EOF Then
      sprItem.Enabled = True
      cmdOk.Tag = snap("OrderingType")
      snap.MoveLast
snap.MoveFirst
      sprItem.MaxRows = snap.RecordCount + 1
      For rCnt = 1 To snap.RecordCount
        If PrevItemGroup <> snap("ItemGroupKey") Then
          PrevItemGroup = snap("ItemGroupKey")
          ValueRow = rCnt
        End If
        sprItem.SetText eItem.mFirstRow, rCnt, ValueRow
        
        sprItem.SetText eItem.mrecdItem, rCnt, snap("ReceivedItemCode")
        Spread_Set_CellTag sprItem, snap("ReceivedItemCode"), eItem.mrecdItem, rCnt
        
        sprItem.SetText eItem.mRecdDesc, rCnt, CheckNull(snap("RecdDesc"))
        sprItem.SetText eItem.mRecdUOM, rCnt, CheckNull(snap("RecdUOM"))
        
        sprItem.SetText eItem.mRelation, rCnt, CheckNull(snap("RelationType"))
        sprItem.SetText eItem.mItemGroupKey, rCnt, CheckNull(snap("ItemGroupKey"))
        Spread_Set_CellTag sprItem, CheckNull(snap("ItemGroupKey")), eItem.mItemGroupKey, rCnt
        
        sprItem.SetText eItem.mStatus, rCnt, IIf(snap("Status") <> "C", "Active", "InActive")
        Spread_Set_CellTag sprItem, IIf(snap("Status") <> "C", "Active", "InActive"), eItem.mStatus, rCnt
        
        If snap("RelationType") = "1-1" Or (snap("RelationType") = "N-1" And ValueRow = rCnt) Then
          sprItem.SetText eItem.msentItem, rCnt, snap("SentItemCode")
          Spread_Set_CellTag sprItem, CheckNull(snap("SentItemCode")), eItem.msentItem, rCnt
          
          If txtordertyping.Text <> "Open" Then
            Spread_Lock sprItem, False, eItem.mSentQty, rCnt, eItem.mAssRate, rCnt
            Spread_Lock sprItem, False, eItem.mStatus, rCnt, eItem.mStatus, rCnt
          Else
            Spread_Lock sprItem, False, eItem.mRate, rCnt, eItem.mAssRate, rCnt
            Spread_Lock sprItem, False, eItem.mStatus, rCnt, eItem.mStatus, rCnt
          End If
          
          sprItem.SetText eItem.mSentDesc, rCnt, CheckNull(snap("SentDesc"))
          sprItem.SetText eItem.mSentUOM, rCnt, CheckNull(snap("SentUOM"))
          
          sprItem.SetText eItem.mPartNo, rCnt, CheckNull(snap("PartNo"))
'          Spread_Decimal_Set sprItem, eItem.mExSentQty, eItem.mRatePer, rCnt, rCnt, CheckNull(snap("NoOfDec"), 0)
          Call Spread_Set_Decimal(sprItem, CheckNull(snap("NoOfDec"), 0), eItem.mSentQty, rCnt, eItem.mRatePer)
          sprItem.SetText eItem.mRatePer, rCnt, CheckNull(snap("RatePer"))
          Spread_Set_CellTag sprItem, CheckNull(snap("SentItemCode")), eItem.msentItem, rCnt
          
          sprItem.SetText eItem.mRate, rCnt, CheckNull(snap("Rate"))
          Spread_Set_CellTag sprItem, CheckNull(snap("Rate")), eItem.mRate, rCnt
          
          sprItem.SetText eItem.mAssRate, rCnt, CheckNull(snap("AssRate"))
          Spread_Set_CellTag sprItem, CheckNull(snap("AssRate")), eItem.mAssRate, rCnt
          
          sprItem.SetText eItem.mSentQty, rCnt, CheckNull(snap("OrderQty"))
          Spread_Set_CellTag sprItem, CheckNull(snap("OrderQty")), eItem.mSentQty, rCnt
          Spread_Set_CellTag sprItem, CheckNull(snap("InvoiceQty")), eItem.mSentUOM, rCnt
          
          sprItem.SetText eItem.mGoodsValue, rCnt, (CheckNull(snap("OrderQty")) * CheckNull(snap("Rate")))
          tot = tot + (CheckNull(snap("OrderQty")) * CheckNull(snap("Rate")))
          sprItem.SetText eItem.mAssGoodsValue, rCnt, (CheckNull(snap("OrderQty")) * CheckNull(snap("AssRate")))
          tot1 = tot1 + (CheckNull(snap("OrderQty")) * CheckNull(snap("AssRate")))
          
          If txtordertyping.Text <> "Open" Then
            SetSheet (ValueRow)
            spSchedule.Sheet = ValueRow
            Sql = "Select ScheduleDate,ScheduleQty,Location from SAL_JobWorkOrder_Schedule where "
            Sql = Sql & " SentItemCode='" & Trim$(snap("SentItemCode")) & "'And OrderType='" & txtOrdType.Text & "'"
            Sql = Sql & " and OrderNo=" & txtordnum.Text & " and OrderDate='" & Format(txtorddt, "yyyy-MM-dd") & "'"
            Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
            Set Schsnap = New AceADODB.Recordset
            If RSOpen(Schsnap, Sql) = False Then
 GoTo ExitHere
End If
            If Not Schsnap.EOF Then
              spSchedule.MaxRows = Schsnap.RecordCount
              For sCnt = 1 To Schsnap.RecordCount
                spSchedule.SetText eSch.mSchDate, sCnt, Schsnap("ScheduleDate")
                spSchedule.SetText eSch.mSchQty, sCnt, Schsnap("ScheduleQty")
                spSchedule.SetText eSch.mLocation, sCnt, Schsnap("Location")
                Schsnap.MoveNext
              Next sCnt
              spSchedule.SetColItemData eSch.mLocation, spSchedule.DataRowCnt
            End If
          End If
        Else
          Call Spread_Lock(sprItem, True, eItem.mSentUOM, rCnt, eItem.mItemGroupKey, rCnt)
        End If
        snap.MoveNext
      Next rCnt
      sprItem.SetColItemData eItem.mSentDesc, sprItem.DataRowCnt
      sprItem.Tag = snap.RecordCount
      fraAmd.Enabled = False
      fraOrder.Enabled = False
      fraTerms.Enabled = True
      sprItem.Enabled = True
      cmdsave.Enabled = True
      txtPack.SetFocus
    Else
      sprItem.MaxRows = 1
      sprItem.Tag = 0
    End If
    If txtordertyping.Text <> "Open" Then
      fraSch.Visible = True
    Else
      fraSch.Visible = False
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set snap = Nothing
  Set Schsnap = Nothing
End Try
End Sub

Private Sub cmdSave_Click()
  Dim isDBTranActive As Boolean
  Dim retVal As Variant
  Dim tmp As Variant
  Dim amdnorec As AceADODB.Recordset
  Dim cmd As AceADODB.Command
  Dim flg As Boolean
  Dim i As Long
  Dim sCnt As Long
  Dim Sql As String
  Dim CurDate As Date
  Dim sRelation As String
  Dim sFirstRow As Long
  Dim srow As Long

  Try
 
  Set amdnorec = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  isDBTranActive = False
  
  CurDate = ServerDate
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  db.BeginTrans
  isDBTranActive = True
  
  Call SetDocNo
  
  '****SAL_Ins_Order_Amendment
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_Ins_Order_Amendment"
  cmd.Parameters.Append cmd.CreateParameter("AmendmentType", adVarChar, adParamInput, 3, Trim$(cmbamdtype))
  cmd.Parameters.Append cmd.CreateParameter("AmendmentNo", adInteger, adParamInput, , Val(txtamdno))
  cmd.Parameters.Append cmd.CreateParameter("AmendmentDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType))
  cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum))
  cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt, "yyyy-MM-dd"))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute
  
  '*****
  '@@@@@@@SAL_INS_JobWorkORDER_HDR
  
  Sql = "insert into SAL_JobWorkOrder_Hdr"
  Sql = Sql & " Select Distinct OrderType,OrderNo,OrderDate,OrderSource,CustomerCode,ReferenceNo,ReferenceDate,'" & Format(dpvalidfrom.Value, "yyyy-MM-dd") & "','" & Format(dpvalidto.Value, "yyyy-MM-dd") & "',OrderingType,"
  Sql = Sql & " Remarks,'" & txtForm.Tag & "','" & txtPack.Tag & "','" & txtDelIns.Tag & "','" & txtTerms.Tag & "','" & txtInsurance.Tag & "','" & txtMode.Tag & "',"
  Sql = Sql & " '" & txtFreight.Tag & "','" & txtFormula.Tag & "','" & txtGrandTotal.Text & "',PostingAc,'M',"
  Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "' , Status, OrderAuthorizationDate ,OrderAuthorizationBy,'" & gUserID & "',"
  Sql = Sql & " Branchid,Compid,Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'  from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
  Sql = Sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo='" & Val(txtordnum.Text) & "' and OrderDate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
  db.Execute Sql
    
  'Update Selected Order Finyear for Generate DocNo to this Financial year
  
  Sql = " Update SAL_JobWorkOrder_Hdr Set Finyear = '" & Trim$(txtFinyear.Text) & "'"
  Sql = Sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo='" & Val(txtordnum.Text) & "' and OrderDate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and compid='" & gCompID & "' and Finyear = '" & gFinyear & "'"
  db.Execute Sql
    
  '@@@@@@@@
  
  '!!!!!!!!!!!SAL_INS_JobWorkOrder_DTL
    For i = 1 To sprItem.DataRowCnt
      If i <= sprItem.GetColItemData(eItem.mSentDesc) Then
        If DataChanged(i) = True Then
          sprItem.GetText eItem.mFirstRow, i, tmp
          sFirstRow = Val(tmp)
          commClear cmd
          cmd.CommandText = "SAL_INS_JOBWORKORDER_DTL"
          cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
          cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
          cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt, "yyyy-MM-dd"))
          sprItem.GetText eItem.mItemGroupKey, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
          sprItem.GetText eItem.mRelation, i, tmp
          sRelation = Trim$(tmp)
          If sRelation = "1-1" Or sRelation = "N-1" Then
            sprItem.GetText eItem.mrecdItem, i, tmp
            cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
            sprItem.GetText eItem.msentItem, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          Else
            sprItem.GetText eItem.mrecdItem, i, tmp
            cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
            sprItem.GetText eItem.msentItem, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          End If
          cmd.Parameters.Append cmd.CreateParameter("ScheduleType", adChar, adParamInput, 1, Left(txtordertyping, 1))
          If txtordertyping.Text = "Open" Then
            cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , "0")
          Else
            sprItem.GetText eItem.mSentQty, sFirstRow, tmp
            cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
          End If
          sprItem.GetText eItem.mRatePer, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("RatePer", adDouble, adParamInput, , Val(tmp))
          sprItem.GetText eItem.mRate, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(tmp))
          sprItem.GetText eItem.mAssRate, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(tmp))
          cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "M")
          cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
          Call AddCommonCmdParameters(cmd)
          cmd.Execute
        End If
      Else
        sprItem.GetText eItem.mFirstRow, i, tmp
        sFirstRow = Val(tmp)
        commClear cmd
        cmd.CommandText = "SAL_INS_JOBWORKORDER_DTL"
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt, "yyyy-MM-dd"))
        sprItem.GetText eItem.mItemGroupKey, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
        sprItem.GetText eItem.mRelation, i, tmp
        sRelation = Trim$(tmp)
        If sRelation = "1-1" Or sRelation = "N-1" Then
          sprItem.GetText eItem.mrecdItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          sprItem.GetText eItem.msentItem, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        Else
          sprItem.GetText eItem.mrecdItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          sprItem.GetText eItem.msentItem, sFirstRow, tmp
          cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        End If
        cmd.Parameters.Append cmd.CreateParameter("ScheduleType", adChar, adParamInput, 1, Left(txtordertyping, 1))
        sprItem.GetText eItem.mSentQty, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
        sprItem.GetText eItem.mRatePer, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("RatePer", adDouble, adParamInput, , Val(tmp))
        sprItem.GetText eItem.mRate, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(tmp))
        sprItem.GetText eItem.mAssRate, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(tmp))
        cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "M")
        cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
'        Call AddCommonCmdParameters(cmd, , Trim$(txtFinyear.Text))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
     '!!!!!!!!
  If cmdOk.Tag <> "Open" And i = sFirstRow Then
    With spSchedule
      sprItem.Sheet = sFirstRow
      Sql = "delete from SAL_JobWorkOrder_Schedule"
      Sql = Sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo=" & Val(txtordnum.Text) & " and orderdate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
      sprItem.GetText eItem.msentItem, sFirstRow, retVal
      Sql = Sql & " and SentItemCode='" & Trim$(retVal) & "' "
      sprItem.GetText eItem.mItemGroupKey, sFirstRow, retVal
      Sql = Sql & " and ItemGroupKey ='" & Val(retVal) & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
      db.Execute Sql
      
      For sCnt = 1 To sprItem.DataRowCnt
        cmd.CommandText = "Sal_Ins_JobWorkOrder_Schedule"
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
        sprItem.GetText eItem.mItemGroupKey, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(retVal))
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mPartNo, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eSch.mSchDate, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(retVal, "yyyy-MM-dd"))
        sprItem.GetText eSch.mSchQty, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(retVal))
        sprItem.GetText eSch.mLocation, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(retVal))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      Next sCnt
      spSchedule.Sheet = sFirstRow
      Sql = "delete from SAL_JobWorkOrder_Schedule"
      Sql = Sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo=" & Val(txtordnum.Text) & " and orderdate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
      sprItem.GetText eItem.msentItem, sFirstRow, retVal
      Sql = Sql & " and SentItemCode='" & Trim$(retVal) & "' "
      sprItem.GetText eItem.mItemGroupKey, sFirstRow, retVal
      Sql = Sql & " and ItemGroupKey ='" & Val(retVal) & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
      db.Execute Sql
      
      For sCnt = 1 To spSchedule.DataRowCnt
        cmd.CommandText = "Sal_Ins_JobWorkOrder_Schedule"
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
        sprItem.GetText eItem.mItemGroupKey, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(retVal))
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mPartNo, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retVal))
        spSchedule.GetText eSch.mSchDate, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(retVal, "yyyy-MM-dd"))
        spSchedule.GetText eSch.mSchQty, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(retVal))
        spSchedule.GetText eSch.mLocation, sCnt, retVal
        cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(retVal))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      Next sCnt
  End If
  Next i
  End With

'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If Val(txtGoodsValue.Text) <> 0 Then
    If SaveCalcDetails = False Then
      GoTo ExitHere
    End If
    
    If UpdateFinyearToCalcDetails(Trim$(txtUserKey.Text), gModuleCode, Val(txtFormula.Tag), Trim$(txtOrdType), Val(txtordnum), txtorddt, Trim$(txtFinyear.Text)) = False Then
      GoTo ExitHere
    End If
  End If
  
  db.CommitTrans
  isDBTranActive = False
  sShowMessage "Amendment Number :" & Trim$(cmbamdtype) & " " & Trim$(txtamdno) & "  Is Generated"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  If isDBTranActive = True Then
    db.RollbackTrans
  End If
  Set amdnorec = Nothing
  Set cmd = Nothing
End Try
End Sub

Private Function SaveCalcDetails() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    Sql = "delete from COM_FormulaCalc_Hdr where DocType = '" & Trim$(txtOrdType) & "' "
    Sql = Sql + " and DocNo = " & Val(txtordnum) & " and DocDate = '" & Format(txtorddt, "yyyy-MM-dd") & "'"
    Sql = Sql + " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    db.Execute Sql
  
    Sql = "delete from COM_FormulaCalc_Dtl where DocType = '" & Trim$(txtOrdType) & "' "
    Sql = Sql + " and DocNo = " & Val(txtordnum) & " and DocDate = '" & Format(txtorddt, "yyyy-MM-dd") & "'"
    Sql = Sql + " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    db.Execute Sql
  
    CurDate = ServerDate
    
    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(txtOrdType) & "' DocType," & Val(txtordnum) & " DocNo,"
    Sql = Sql & "'" & Format(txtorddt, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,"
    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    db.Execute Sql

    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(txtOrdType) & "', " & Val(txtordnum) & ", '" & Format(txtorddt, "yyyy-MM-dd") & "',"
    Sql = Sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, "
    Sql = Sql & " a.IsValueEditable, a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    db.Execute Sql
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub dpamddate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
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
 SendKeys "{tab}"
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
  
  dpamddate = Format(ServerDate, "dd/MM/yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub sprItem_DblClick(sender as object, e as DblClickEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  If Col = eItem.mrecdItem Then
    If Row = sprItem.MaxRows Then
      Call HelpDirectItem(Col, Row)
    End If
  ElseIf Col = eItem.msentItem Then
    ManyToOneItemHelp (Row)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function HelpDirectItem(ByVal fCol As Long,ByVal  fRow As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim sFirstRow As Long
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim tmp3 As Variant

  Try
  HelpDirectItem = False
  
  If fCol = eItem.mrecdItem Then
    'Show 1 to 1 Item for the Vendor
    Sql = " Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
    Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
    Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and a.RelationType='1-1' and a.isActive='1'"
    Sql = Sql & " and a.ValidFrom <= '" & Format(dpamddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dpamddate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and b.Status = 1 and c.Status = 1 and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "1111", "List of Item", "Recd Item", "Description", "Sent Item", "Description") = True Then
        If fRow <= sprItem.DataRowCnt Then
          rCnt = fRow
        Else
          sprItem.MaxRows = sprItem.DataRowCnt + 1
          rCnt = sprItem.DataRowCnt + 1
        End If
        
        'Key Values First Set in Spread
        sprItem.SetText eItem.mItemGroupKey, rCnt, getvalue(10)
        sprItem.SetText eItem.mrecdItem, rCnt, getvalue(1)
        sprItem.SetText eItem.msentItem, rCnt, getvalue(4)
        
        If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.mrecdItem, eItem.msentItem) = False Then
          sprItem.SetText eItem.mFirstRow, rCnt, rCnt
          
          sprItem.SetText eItem.mRecdDesc, rCnt, getvalue(2)
          sprItem.SetText eItem.mRecdUOM, rCnt, getvalue(3)
          
          sprItem.SetText eItem.mSentDesc, rCnt, getvalue(5)
          sprItem.SetText eItem.mSentUOM, rCnt, getvalue(6)
          Call Spread_Set_Decimal(sprItem, Val(getvalue(7)), eItem.mSentQty, rCnt)
          If Trim$(cmdOk.Tag) = "Open" Then
            Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          Else
            Call Spread_Lock(sprItem, False, eItem.mRatePer, rCnt, eItem.mSentQty, rCnt)
          End If
          Call Spread_Lock(sprItem, True, eItem.mRate, rCnt, eItem.mAssRate, rCnt)
          sprItem.SetText eItem.mRatePer, rCnt, getvalue(8)
          sprItem.SetText eItem.mRate, rCnt, getvalue(9)
          sprItem.SetText eItem.mAssRate, rCnt, getvalue(9)
          sprItem.SetText eItem.mRelation, rCnt, "1-1"
          sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
          sprItem.GetText eItem.mRate, rCnt, tmp1  'Rate
          sprItem.GetText eItem.mRatePer, rCnt, tmp3   'Rateper
          sprItem.SetText eItem.mStatus, rCnt, "Active"
          Call Spread_Lock(sprItem, True, eItem.mStatus, rCnt, eItem.mStatus, rCnt)
          If Val(tmp3) > 0 Then
            sprItem.SetText eItem.mGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
            sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
            sprItem.GetText eItem.mAssRate, rCnt, tmp1 'ARate
            sprItem.SetText eItem.mAssGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
          End If
        End If
        sprItem.MaxRows = sprItem.DataRowCnt + 1    Else
      GoTo ExitHere
    End If
  Else
    'Show N to 1 Item for the Vendor
    
    Sql = "Select a.SentItemCode, b.ItemDescription as SentDesc, a.ReceivedItemCode, c.ItemDescription as RecdDesc, a.ItemGroupKey"
    Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
    Sql = Sql & " and (a.ValidFrom <= '" & Format(dpamddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dpamddate, "yyyy-MM-dd") & "')"
    Sql = Sql & " and b.Status = 1 and c.Status = 1 and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    Sql = Sql & " Order by a.ReceivedItemCode, a.SentItemCode, a.ItemGroupKey"
    If FetchValue_New(False, db.ConnectionString, Sql, "11110", "List of Item", "Sent Item", "Description", "Recd Item", "Description") = True Then
    
      Sql = "Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
      Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
      Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey"
      Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
      Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
      Sql = Sql & " Left Outer Join Material_Master c"
      Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
      Sql = Sql & " where a.CustomerCode=" & Val(txtcustcd.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
      Sql = Sql & " and a.ItemGroupKey=" & getvalue(5)
      Sql = Sql & " and a.ValidFrom <= '" & Format(dpamddate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dpamddate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and b.Status = 1 and c.Status = 1 and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If rs.EOF Then
 GoTo ExitHere
End If
        rs.MoveLast
rs.MoveFirst
        If fRow <= sprItem.DataRowCnt Then
          sFirstRow = fRow
          sprItem.MaxRows = rs.RecordCount + 1
        Else
          sFirstRow = sprItem.DataRowCnt + 1
          sprItem.MaxRows = sprItem.DataRowCnt + rs.RecordCount + 1
        End If
        For rCnt = sFirstRow To sprItem.MaxRows - 1
          'Key Values First Set
          sprItem.SetText eItem.mItemGroupKey, rCnt, rs("ItemGroupKey")
          sprItem.SetText eItem.mrecdItem, rCnt, rs("ReceivedItemCode")
          If sFirstRow = rCnt Then
            sprItem.SetText eItem.msentItem, rCnt, rs("SentItemCode")
          End If
          
          If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.msentItem, eItem.mrecdItem) = True Then
            sprItem.MaxRows = sprItem.DataRowCnt + 1
            Exit For
          End If
        
          sprItem.SetText eItem.mFirstRow, rCnt, sFirstRow
          
          sprItem.SetText eItem.mSentDesc, rCnt, rs("SentDesc")
          sprItem.SetText eItem.mSentUOM, rCnt, rs("SentUOM")
          
          If sFirstRow = rCnt Then
            sprItem.SetText eItem.mRecdDesc, rCnt, rs("RecdDesc")
            sprItem.SetText eItem.mRecdUOM, rCnt, rs("RecdUOM")
            sprItem.SetText eItem.mRatePer, rCnt, rs("RatePer")
            sprItem.SetText eItem.mRate, rCnt, rs("Rate")
            sprItem.SetText eItem.mAssRate, rCnt, rs("Rate")
            If Trim$(cmdOk.Tag) = "Y" Then
              Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            Else
              Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
            End If
            Call Spread_Lock(sprItem, True, eItem.mRate, rCnt, eItem.mAssRate, rCnt)
            sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
            sprItem.GetText eItem.mRate, rCnt, tmp1  'Rate
            sprItem.GetText eItem.mRatePer, rCnt, tmp3   'Rateper
            If Val(tmp3) > 0 Then
              sprItem.SetText eItem.mGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
              sprItem.GetText eItem.mSentQty, rCnt, tmp  'OrderQty
              sprItem.GetText eItem.mAssRate, rCnt, tmp1 'ARate
              sprItem.SetText eItem.mAssGoodsValue, rCnt, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
            End If
          Else
            Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          End If
          sprItem.SetText eItem.mRelation, rCnt, "N-1"
          sprItem.SetText eItem.mStatus, rCnt, "Active"
          rs.MoveNext
        Next rCnt
    Else
      GoTo ExitHere
    End If
  End If
  
  OrderCalulation  'to calculate the Jobwork Order Value
  
  HelpDirectItem = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub sprItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim tmp3 As Variant
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

'  If Mode = 1 Then: GoTo ExitHere
'
'  sprItem.GetText eItem.mSentQty, Row, tmp  'OrderQty
'  sprItem.GetText eItem.mRate, Row, tmp1  'Rate
'  sprItem.GetText eItem.mRatePer, Row, tmp3  'Rateper
'  sprItem.SetText eItem.mGoodsValue, Row, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
'  sprItem.GetText eItem.mAssRate, Row, tmp1 'ARate
'  sprItem.SetText eItem.mAssGoodsValue, Row, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
'
'  OrderCalulation  'to calculate the Jobwork Order Value

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mSentQty, eItem.mRate, eItem.mAssRate
          sprItem.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
'            Call StartCalcFormula(IIf(Trim$(txtGrandTotal.Text) = "", True, False))
          End If
          
          Select Case Col
            Case eItem.mSentQty
              If Val(retVal) > 0 Then
                sprItem.SetActiveCell eItem.mRate, Row
              End If
            Case eItem.mRate
              If Val(retVal) > 0 Then
                sprItem.SetActiveCell eItem.mAssRate, Row
              End If
            Case eItem.mAssRate
              If Val(retVal) > 0 Then
                sprItem.SetActiveCell eItem.msentItem, Row + 1
              End If
          End Select
      End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_KeyDown(sender as object, e as KeyDownEventArgs)
  
  Try
  Me.MousePointer = vbHourglass
  
    Select Case KeyCode
      Case vbKeyF2
        sprItem_DblClick sprItem.ActiveCol, sprItem.ActiveRow
      Case 46
        If sprItem.ActiveCol = 2 And Trim$(SpVal(sprItem, sprItem.ActiveRow, eItem.mStatus)) = "" Then
          Spread_Clear_Row sprItem, 1, sprItem.MaxCols, sprItem.ActiveRow, sprItem.ActiveRow
          sprItem.DeleteRows sprItem.ActiveRow, 1
          sprItem.MaxRows = sprItem.DataRowCnt + 1
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retVal As Variant
  Dim tmp As Variant
  
  Try
  
  sprItem.GetText eItem.mFirstRow, NewRow, retVal
  If txtordertyping.Text <> "Open" Then
    If Row <> NewRow And NewRow >= 1 And Val(retVal) = NewRow Then
      SetSheet (NewRow)
      sprItem.GetText eItem.msentItem, NewRow, retVal
      txtItemCode.Text = Trim$(retVal)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spSchedule_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retVal As Variant
  
  Try
  
    If Col = eSch.mSchDate Then
      Spread_Clear_Duplicate spSchedule, Row, eSch.mSchDate
      spSchedule.MaxRows = spSchedule.DataRowCnt + 1
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtDelIns_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Delivery Instruction' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Delivery Instruction", "Description") = True Then
    txtDelIns.Text = getvalue(1)
    txtDelIns.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtForm_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select description,Code From TERMSMASTER Where typeflg = 'Form' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Form's", "Description") = True Then
    txtForm.Text = getvalue(1)
    txtForm.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_Change()
   
  Try
  Me.MousePointer = vbHourglass
  
  If sprItem.DataRowCnt <= 0 Then
    GoTo ExitHere                 'No need to calculate formula
  End If
  
  If Len(Trim$(txtFormula.Tag)) > 0 Then
    If StartCalcFormula = False Then
      GoTo ExitHere
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select FrmLName, FrmID from COM_FORMULAMAST_HDR"
  Sql = Sql & " where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula's", "Formula Name") = True Then
    txtFormula.Tag = getvalue(2)
    txtFormula.Text = getvalue(1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFormula_DblClick
End If
End Sub

Private Sub txtFreight_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description,Code From TERMSMASTER Where typeflg = 'FreightCode'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of FreightCode", "Description") = True Then
    txtFreight.Text = getvalue(1)
    txtFreight.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  GoTo ExitHere
End Sub

Private Sub txtInsurance_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Insurance' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Insurance", "Description") = True Then
    txtInsurance.Text = getvalue(1)
    txtInsurance.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtMode_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description,Code From TERMSMASTER Where typeflg = 'Mode' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtMode.Text = getvalue(1)
    txtMode.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  GoTo ExitHere
End Sub

Private Sub txtPack_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Packing & Forwarding' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtPack.Text = getvalue(1)
    txtPack.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtTerms_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Payment Terms' "
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Payment Terms", "Description") = True Then
    txtTerms.Text = getvalue(1)
    txtTerms.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  GoTo ExitHere
End Sub

Private Sub txtPack_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtPack_DblClick
End If
End Sub

Private Sub txtInsurance_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtInsurance_DblClick
End If
End Sub

Private Sub txtDelIns_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtDelIns_DblClick
End If
End Sub

Private Sub txtForm_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtForm_DblClick
End If
End Sub

Private Sub txtFreight_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFreight_DblClick
End If
End Sub

Private Sub txtMode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtMode_DblClick
End If
End Sub

Private Sub txtTerms_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtTerms_DblClick
End If
End Sub

Private Sub txtordnum_DblClick()
  Dim orddetrec As AceADODB.Recordset
  Dim ret As String
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  Set orddetrec = New AceADODB.Recordset

  Sql = " Select distinct a.OrderType,a.OrderNo,a.OrderDate,a.CustomerCode,(Select b.PartyName from"
  Sql = Sql & " FAS_PARTY_MASTER b Where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
  Sql = Sql & " and compid='" & gCompID & "') PartyName,a.ValidFrom,a.ValidTo, a.ReferenceNo,a.OrderingType, a.Finyear"
  Sql = Sql & " From dbo.fn_Sal_jobworkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a where "
  Sql = Sql & " a.Status='A' order by a.OrderDate,a.OrderType,a.OrderNo"
  If FetchValue_New(False, db.ConnectionString, Sql, "1110101", "List of Order", "Order Type", "Order No", "Order Date", "Customer", "PO No") = True Then
    txtOrdType = Trim$(getvalue(1))
    txtordnum = getvalue(2)
    txtorddt = getvalue(3)
    txtcustcd.Tag = getvalue(4)
    txtcustcd = getvalue(5)
    dpvalidfrom.Value = getvalue(6)
    dpvalidto.Value = getvalue(7)
    txtordertyping.Text = getvalue(9)
    cmdOk.Tag = getvalue(9)
    txtFinyear.Text = getvalue(10)
    
    If Not Trim$(getvalue(7)) = "" Then
      If dpamddate > CDate(getvalue(7)) Then
        ret = MsgBox("Amendment Date is greater than Valid till Date of Order" & vbCrLf & " Do you want to continue Amending the Order ?", vbYesNo + vbQuestion, App.Title)

        If Val(ret) <> 6 Then txtordnum = ""
txtordnum.SetFocus
GoTo ExitHere
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set orddetrec = Nothing
End Try
End Sub

Private Sub txtordnum_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyF2 Then
 txtordnum_DblClick
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retVal As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(sprItem.DataRowCnt - 1)
      ReDim SCValList(sprItem.DataRowCnt * 2 - 1)
      
      For rCnt = 1 To sprItem.DataRowCnt
        'Item List
        sprItem.GetText eItem.msentItem, rCnt, retVal
        Item = Trim$(retVal)
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        sprItem.GetText eItem.mSentQty, rCnt, retVal
        Qty = Val(retVal)
        
        sprItem.GetText eItem.mStatus, rCnt, retVal
        If Trim$(retVal) <> "Active" Then
          Qty = 0                         'To Avoid the value of Deactive Items
        End If
        
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        
        sprItem.GetText eItem.mRate, rCnt, retVal
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        
        sprItem.GetText eItem.mAssRate, rCnt, retVal
        SCValList((rCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
      Next rCnt
  Else
    ReDim ItemList(0)
    ItemList(0) = "0" & Chr(145) & "0"
    
    ReDim SCValList(1)
    SCValList(0) = eParamID.mGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Text)
    SCValList(1) = eParamID.mAGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Tag)
  End If

  If CalculateFormula(Val(txtcustcd.Tag), Val(txtFormula.Tag), txtUserKey, SCValList, ItemList, txtGrandTotal, sprCalc, eCalc.mDesc, eCalc.mValue, , "Sales Order", False, ShowPromptScreen) = False Then

    
    GoTo ExitHere
  End If
  
  Call CalcButtonStatus(Me, False)

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CalcButtonStatus(ByVal FrmObj As Form,ByVal  Enable As Boolean) As Boolean
  
  Try
  
  CalcButtonStatus = False
  
    If Enable Then
      FrmObj.cmdCalc.Enabled = True
      FrmObj.sprCalc.Visible = False
      FrmObj.txtGrandTotal.Visible = False
      FrmObj.lblGrandTotal.Visible = False
    Else
      FrmObj.cmdCalc.Enabled = False
      FrmObj.sprCalc.Visible = True
      FrmObj.txtGrandTotal.Visible = True
      FrmObj.lblGrandTotal.Visible = True
    End If
  
  CalcButtonStatus = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdCalc_Click()

  Try
  Me.MousePointer = vbHourglass

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  Call CalcButtonStatus(Me, False)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetSprGoodsValue(ByVal fRow As Long)
  Dim retVal(1 To 3) As Variant
  Dim tmpTot As Double
    
  Try
  
    sprItem.GetText eItem.mSentQty, fRow, retVal(1)
    retVal(2) = 1
    
    sprItem.GetText eItem.mRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    sprItem.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    sprItem.GetText eItem.mAssRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    sprItem.SetText eItem.mAssGoodsValue, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(sprItem, eItem.mGoodsValue), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(sprItem, eItem.mAssGoodsValue), "0.00")
End Sub

Private Function SprTotCalc(ByRef fSpread As AceSpread,ByVal  fCol As Long) As Double
  Dim rCnt As Long
  Dim retVal As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, rCnt, retVal
      Total = Total + Val(retVal)
    Next rCnt
  
  SprTotCalc = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_spSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spSchedule


spSchedule.LeaveCell +=  new EventHandler(spSchedule_LeaveCell)

End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.DblClick +=  new EventHandler(sprItem_DblClick)
sprItem.EditMode +=  new EventHandler(sprItem_EditMode)
sprItem.KeyDown +=  new EventHandler(sprItem_KeyDown)
sprItem.LeaveCell +=  new EventHandler(sprItem_LeaveCell)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_spSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrderAmendment.frm", spSchedule, "")

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrderAmendment.frm", sprItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\JobworkOrderAmendment.frm", sprCalc, "")

End Sub
