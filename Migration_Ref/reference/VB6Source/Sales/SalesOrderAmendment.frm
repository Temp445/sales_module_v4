VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesOrderAmendment 
   Caption         =   "Sales Order Amendment"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
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
   ScaleHeight     =   8595
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFinyear 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   435
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   59
      Top             =   8025
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Frame fraSch 
      Height          =   1875
      Left            =   270
      TabIndex        =   46
      Top             =   5790
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
      Top             =   5265
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5160
      TabIndex        =   18
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5265
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3870
      TabIndex        =   17
      ToolTipText     =   "Click to Save the Data"
      Top             =   5265
      Width           =   1300
   End
   Begin VB.Frame fraAmd 
      Height          =   720
      Left            =   315
      TabIndex        =   24
      Top             =   105
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
         Format          =   61341697
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
      Top             =   105
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
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Sales Order Type"
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
         MultiLine       =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Select the Sales Order. Press F2 or Double Click for Help"
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
         ToolTipText     =   "Sales Order Date"
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
         Left            =   1770
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
Begin AceSpread spAmd
      Height          =   2085
      Left            =   300
      TabIndex        =   16
      ToolTipText     =   "Enter the Amendment Details"
      Top             =   3120
      Width           =   11220
      _Version        =   458752
      _ExtentX        =   19791
      _ExtentY        =   3678
      _StockProps     =   64
      Enabled         =   0   'False
      ColsFrozen      =   2
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
      MaxCols         =   12
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
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
      Height          =   2205
      Left            =   300
      TabIndex        =   29
      Top             =   795
      Width           =   11220
      Begin VB.TextBox txtordertyping 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   52
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   1230
      End
      Begin VB.TextBox txtReLevel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8700
         MaxLength       =   30
         TabIndex        =   50
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
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3675
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
         ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
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
         Left            =   90
         TabIndex        =   55
         ToolTipText     =   "Enter the Valid From Date"
         Top             =   1815
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
         Format          =   61341699
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   1470
         TabIndex        =   56
         ToolTipText     =   "Enter the Valid Till Date"
         Top             =   1815
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
         Format          =   61341699
         CurrentDate     =   37250
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
         Height          =   210
         Left            =   90
         TabIndex        =   58
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valid Till"
         Height          =   210
         Left            =   1470
         TabIndex        =   57
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OrderTyping"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   7455
         TabIndex        =   53
         Top             =   1110
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "RevisionLevel"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   8700
         TabIndex        =   51
         Top             =   1110
         Width           =   1005
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
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7965
      TabIndex        =   38
      Top             =   5145
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   42
         Top             =   105
         Width           =   900
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   120
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   40
         Top             =   120
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
         TabIndex        =   39
         Top             =   2550
         Visible         =   0   'False
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   2085
         Left            =   45
         TabIndex        =   43
         Top             =   450
         Visible         =   0   'False
         Width           =   3555
         _Version        =   458752
         _ExtentX        =   6271
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
         Top             =   135
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   975
         TabIndex        =   44
         Top             =   2595
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.TextBox txtOrdSour 
      Height          =   315
      Left            =   3555
      TabIndex        =   54
      Top             =   330
      Width           =   330
   End
End
Attribute VB_Name = "frmSalesOrderAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SKey As String

'Column Nos for Order Spread
Private Enum eItem
  mItem = 1
  mDesc = 2
  mPartNo = 3
  mRevNo = 4
  mUOM = 5
  mQty = 6
  mRate = 7
  mARate = 8
  mGV = 9
  mAGV = 10
  mStatus = 11
  mFlag = 12
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

Private Function DataChanged(ByVal fRow As Long) As Boolean
  Dim OldValue As Variant
  Dim NewValue As Variant
  Dim cCnt As Long
  
  Try
  DataChanged = False
  
    spAmd.Row = fRow
    For cCnt = 1 To eItem.mStatus
      If cCnt <= eItem.mARate Or cCnt = eItem.mStatus Then
        OldValue = Spread_Get_CellTag(spAmd, cCnt, fRow)
        spAmd.GetText cCnt, fRow, NewValue
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

Private Function ValidateData() As Boolean
  Dim retval As Variant
  Dim flg As Boolean
  Dim i As Long
  Dim sCnt As Long
  Dim sql As String
  Dim CurDate As Date
 
  Try
 
  CurDate = ServerDate
  
  flg = False
  ValidateData = False
  
  For i = 1 To spAmd.DataRowCnt
    If CmdOk.Tag <> "Open" Then
    spAmd.GetText eItem.mQty, i, retval
      If Val(retval) < Val(Spread_Get_CellTag(spAmd, eItem.mQty, i)) Then
        sShowMessage "Amend Qty should be greater than Invoice Qty (" & Val(spAmd.CellTag) & ")"
        spAmd.SetFocus
        Spread_Set_Focus spAmd, i, eItem.mQty
        GoTo ExitHere
      End If
    End If
  Next i
  
  For i = 1 To spAmd.DataRowCnt
    spSchedule.sheet = i
    If fraSch.Visible = True Then
      If spSchedule.DataRowCnt = 0 Then
        sShowMessage "Enter the Schedule Details"
        GoTo ExitHere
      End If
      For sCnt = 1 To spSchedule.DataRowCnt
        spSchedule.GetText eSch.mSchDate, sCnt, retval
        If Len(retval) = 0 Then
          sShowMessage "Enter the ScheduleDate"
          spSchedule.ActiveSheet = i
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
          GoTo ExitHere
        End If
        spSchedule.GetText eSch.mSchDate, sCnt, retval
        If IsDate(retval) = 0 Then
          sShowMessage "invalid ScheduleDate"
          spSchedule.ActiveSheet = i
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
          GoTo ExitHere
        End If
        spSchedule.GetText eSch.mSchDate, sCnt, retval
        If sCnt > spSchedule.GetColItemData(eSch.mSchDate) Then
          If Format(retval, "yyyyMMdd") < Format(CurDate, "yyyyMMdd") Then
            sShowMessage "invalid ScheduleDate"
            spSchedule.ActiveSheet = i
            Spread_Set_Focus spSchedule, sCnt, eSch.mSchDate
            GoTo ExitHere
          End If
        End If
        spSchedule.GetText eSch.mSchQty, sCnt, retval
        If Val(retval) = 0 Then
          sShowMessage "Enter the ScheduleQty"
          spSchedule.ActiveSheet = i
          Spread_Set_Focus spSchedule, sCnt, eSch.mSchQty
          GoTo ExitHere
        End If
      Next sCnt
      spAmd.GetText eItem.mQty, i, retval
      If Val(retval) <> Spread_Cell_Total(spSchedule, eSch.mSchQty, 1, eSch.mSchQty, spSchedule.DataRowCnt) Then
        sShowMessage "Schedule TotalQty Should be OrderQty"
        spSchedule.ActiveSheet = i
        Spread_Set_Focus spSchedule, 1, eSch.mSchQty
        GoTo ExitHere
      End If
    End If
  Next i
  flg = False
  For i = 1 To spAmd.DataRowCnt
    If Trim$(SpVal(spAmd, i, eItem.mStatus)) = "Active" Then
      flg = True
    End If
  Next i

  If flg = False Then
    sShowMessage "Atlease one Item should be Active"
    spAmd.SetFocus
    Spread_Set_Focus spAmd, 1, eItem.mStatus
    GoTo ExitHere
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub SetSheet(ByVal fRow As Long)
' On Error Resume Next
  
    If spSchedule.SheetCount < fRow Then
      spSchedule.SheetCount = fRow
      Call CopySpreadStructure(spSchedule, 1, , fRow)
      spSchedule.sheet = fRow
      spSchedule.MaxRows = 1
    End If
    spSchedule.sheet = fRow
    spSchedule.ActiveSheet = fRow
End Sub

Private Sub OrderCalulation()
  Dim tot As Double
  Dim atot As Double
  Dim i As Long
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If spAmd.DataRowCnt = 0 Then
 GoTo ExitHere
End If
  
  tot = 0
atot = 0
  For i = 1 To spAmd.DataRowCnt
    If Trim$(SpVal(spAmd, i, eItem.mStatus)) = "Active" Then
      tot = tot + Val(SpVal(spAmd, i, eItem.mGV))
      atot = atot + Val(SpVal(spAmd, i, eItem.mAGV))
    End If
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub TermsDetails(ByVal CurDate As Date)
  Dim snap As New AceADODB.Recordset
  Dim sql As String

  Try
  Me.MousePointer = vbHourglass

  sql = " select Distinct a.FormCode,(select b.Description from TERMSMASTER b where "
  sql = sql + " b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid "
  sql = sql + " and b.Compid=a.Compid) FormDesc, a.PackingCode,(select b1.Description "
  sql = sql + " from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding' "
  sql = sql + " and b1.Code=a.PackingCode and b1.Branchid=a.Branchid "
  sql = sql + " and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,(select b2.Description "
  sql = sql + " from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction' "
  sql = sql + " and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid "
  sql = sql + " and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,(select b3.Description "
  sql = sql + " from TERMSMASTER b3 where b3.TypeFlg='Payment Terms'"
  sql = sql + " and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid "
  sql = sql + " and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,(select "
  sql = sql + " b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance'"
  sql = sql + " and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid "
  sql = sql + " and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,(select "
  sql = sql + " b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  sql = sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid "
  sql = sql + " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,(select b6.Description "
  sql = sql + " from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode "
  sql = sql + " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc,a.FormulaCode,"
  sql = sql + " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where "
  sql = sql + " b7.ModuleCode = '" & gModuleCode & "' and b7.FrmId = a.FormulaCode and b7.Branchid=a.Branchid and b7.Compid=a.Compid) FrmLName"
  sql = sql + " from dbo.fn_SAL_ORDER_HDR('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a where a.OrderType ='" + Trim$(txtOrdType) + "' "
  sql = sql + " and a.OrderNo = " & Val(txtordnum) & " and a.OrderDate = '" + Format(txtorddt, "yyyy-MM-dd") + "'"
  sql = sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.CustomerCode = " & Val(txtcustcd.Tag) & ""
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then GoTo ExitHere
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
  
  txtamdno.Text = GetDocNo("SAL_Order_Amendment", "AmendmentType", "AmendmentNo", Trim$(cmbamdtype), "A")
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
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='A'"
'  Sql = Sql + " and ModuleName='SAL' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
'  If RSOpen(amdtyperec, Sql) = False Then GoTo ExitHere
'  If Not amdtyperec.EOF Then
'    cmbamdtype.Clear
'    Do Until amdtyperec.EOF
'      cmbamdtype.AddItem amdtyperec("DocumentType")
'      amdtyperec.MoveNext
'    Loop
'     cmbamdtype.ListIndex = 0
'     txtamdno = GetNo("SAL_AMENDMENT_HDR", "AmendmentType", "AmendmentNo", Trim$(cmbamdtype), "A", "SAL")
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
  Dim sql As String
 
  Try
  Me.MousePointer = vbHourglass
  
  spSchedule.SheetCount = 1
  Form_Clear Me
  Call LoadDocType(mSalOrderAmendt)
  Call SetDocNo
  fraAmd.Enabled = True
  fraOrder.Enabled = True
  fraTerms.Enabled = False
  spAmd.Enabled = False
  cmdSave.Enabled = False
  cmbamdtype.SetFocus
  spAmd.MaxRows = 1
  fraSch.Visible = False
  CurrDate Me
'  LoadAmdtype
  spSchedule.MaxRows = 1
'  Sql = "delete from COM_FORMULA_TMP where SKey = '" + SKey + "'"
'  Sql = Sql + " and Userid ='" & gUserID & "' and Branchid='" & gBranchID & "' "
'  Sql = Sql + " and Compid='" & gCompID & "'"
'  db.Execute Sql
  spSchedule.Enabled = False
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

Private Sub CmdOk_Click()
  Dim orddetrec As AceADODB.Recordset
  Dim ordSch As AceADODB.Recordset
  Dim sql As String
  Dim i As Long
  Dim sCnt As Long
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
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

  SKey = Format(Now, "yyyyMMddHHmmss") & gUserID

'  Set orddetrec = New ADODB.Recordset
'
  TermsDetails (CurDate) 'To display th terms details
  
  sql = "select distinct a.ItemCode,c.ItemDescription,a.RevisionNo,(select b.PartNo from SAL_Customer_ItemMaster b "
  sql = sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.Compid "
  sql = sql + " and b.CustomerCode=a.CustomerCode) PartNo, c.UOMDesc, c.UOM,"
  sql = sql + " c.UOMDecimal,a.OrderQty,a.Rate,a.AssessableRate,a.ItemStatus Status,a.OrderingType, "
  sql = sql + " isnull((select sum(d.InvoiceQty) from SAL_Invoice_Dtl d where d.OrderType = a.OrderType "
  sql = sql + " and d.OrderNo = a.OrderNo and d.OrderDate = a.OrderDate and d.ItemCode = a.ItemCode"
  sql = sql + " and d.Branchid = a.Branchid and d.Compid = a.Compid),0) InvoiceQty from dbo.fn_SAL_SalesOrder ('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')   a left outer join"
  sql = sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid"
  sql = sql + " and a.Compid=C.Compid where a.OrderType='" & Trim$(txtOrdType) & "' "
  sql = sql + " and a.OrderNo=" & Val(txtordnum) & " and a.OrderDate="
  sql = sql + " convert(datetime,'" + Format(txtorddt, "yyyy-MM-dd") + "',111)  "
  sql = sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  
  Set orddetrec = New AceADODB.Recordset
  If RSOpen(orddetrec, sql) = False Then GoTo ExitHere
  If Not orddetrec.EOF Then
    orddetrec.MoveLast
orddetrec.MoveFirst
    spAmd.MaxRows = orddetrec.RecordCount
    CmdOk.Tag = orddetrec("OrderingType")
    For i = 1 To orddetrec.RecordCount
      spAmd.SetText eItem.mItem, i, Trim$(orddetrec("ItemCode"))
      Spread_Set_CellTag spAmd, Trim$(orddetrec("ItemCode")), eItem.mItem, i
      
      spAmd.SetText eItem.mDesc, i, orddetrec("ItemDescription")
      Spread_Set_CellTag spAmd, orddetrec("ItemDescription"), eItem.mDesc, i
      
      spAmd.SetText eItem.mPartNo, i, CheckNull(orddetrec("PartNo"))
      Spread_Set_CellTag spAmd, CheckNull(orddetrec("PartNo")), eItem.mPartNo, i
      
      spAmd.SetText eItem.mRevNo, i, orddetrec("RevisionNo")
      Spread_Set_CellTag spAmd, orddetrec("RevisionNo"), eItem.mRevNo, i
      
      spAmd.SetText eItem.mUOM, i, orddetrec("UOMDesc")
      Spread_Set_CellTag spAmd, orddetrec("UOMDesc"), eItem.mUOM, i
      
      Call Spread_Set_Decimal(spAmd, Val(CheckNull(orddetrec("UOMDecimal"))), eItem.mQty, i, eItem.mQty)
      
      spAmd.SetText eItem.mQty, i, orddetrec("OrderQty")   'OrderQty
      Spread_Set_CellTag spAmd, orddetrec("OrderQty"), eItem.mQty, i
      
      Call Spread_Set_Min(spAmd, orddetrec("InvoiceQty"), eItem.mQty, i)  'InvoiceQty
      Spread_Set_CellTag spAmd, orddetrec("InvoiceQty"), eItem.mQty, i
      
      
      spAmd.SetText eItem.mRate, i, orddetrec("Rate")  'Rate
      Spread_Set_CellTag spAmd, orddetrec("Rate"), eItem.mRate, i
      Call spAmd_EditMode(eItem.mRate, i, 0, True)
      
      spAmd.SetText eItem.mARate, i, orddetrec("AssessableRate")  'Assessable Rate
      Spread_Set_CellTag spAmd, orddetrec("AssessableRate"), eItem.mARate, i
      
      spAmd.SetText eItem.mGV, i, (orddetrec("OrderQty") * orddetrec("Rate"))
      spAmd.SetText eItem.mAGV, i, (orddetrec("OrderQty") * orddetrec("AssessableRate"))
      spAmd.SetText eItem.mStatus, i, IIf(orddetrec("Status") <> "C", "Active", "Deactive")  'Status
      Spread_Set_CellTag spAmd, IIf(orddetrec("Status") <> "C", "Active", "Deactive"), eItem.mStatus, i 'Status
      txtItemcode.Text = Trim$(orddetrec("ItemCode"))
      
      If orddetrec("OrderingType") <> "Open" Then
        SetSheet (i)
        spSchedule.sheet = i
        
        sql = "select a.ScheduleDate,ScheduleQty,a.Location"
        sql = sql & " from SAL_Order_Schedule a "
        sql = sql & " where a.OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo=" & Trim$(txtordnum.Text) & " and orderdate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "'"
        sql = sql & " and a.ItemCode='" & Trim$(orddetrec("ItemCode")) & "' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
        
        Set ordSch = New AceADODB.Recordset
        If RSOpen(ordSch, sql) = False Then GoTo ExitHere
          If Not ordSch.EOF Then
            spSchedule.MaxRows = ordSch.RecordCount
            For sCnt = 1 To ordSch.RecordCount
              spSchedule.SetText eSch.mSchDate, sCnt, Format(ordSch("ScheduleDate"), "dd/MM/yyyy")
              spSchedule.SetText eSch.mSchQty, sCnt, ordSch("ScheduleQty")
              spSchedule.SetText eSch.mLocation, sCnt, Trim$(ordSch("Location"))
              ordSch.MoveNext
            Next sCnt
          Else
            spSchedule.MaxRows = 1
          End If
        spSchedule.SetColItemData eSch.mSchDate, spSchedule.DataRowCnt
        fraSch.Visible = True
        Spread_Lock spAmd, False, eItem.mQty, 0, eItem.mQty, 500
      Else
        Spread_Lock spAmd, True, eItem.mQty, 0, eItem.mQty, 500
      End If
      orddetrec.MoveNext
    Next i
    spAmd.SetColItemData eItem.mFlag, spAmd.DataRowCnt
    
    Call SetTotalGoodsValue
    
    fraAmd.Enabled = False
    fraOrder.Enabled = False
    fraTerms.Enabled = True
    spAmd.Enabled = True
    cmdSave.Enabled = True
    spSchedule.Enabled = True
    txtPack.SetFocus
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set orddetrec = Nothing
  Set ordSch = Nothing
End Try
End Sub

Private Sub cmdSave_Click()
  Dim retval As Variant
  Dim tmp As Variant
  Dim amdnorec As AceADODB.Recordset
  Dim cmd As AceADODB.Command
  Dim flg As Boolean
  Dim i As Long
  Dim sCnt As Long
  Dim sql As String
  Dim CurDate As Date
  
  Try
 
  Set amdnorec = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  CurDate = ServerDate
  
  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetDocNo
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  '****SAL_Ins_Order_Amendment
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_Ins_Order_Amendment"
  cmd.Parameters.Append cmd.CreateParameter("AmendmentType", adVarChar, adParamInput, 3, Trim$(cmbamdtype))
  cmd.Parameters.Append cmd.CreateParameter("AmendmentNo", adInteger, adParamInput, , Val(txtamdno))
  cmd.Parameters.Append cmd.CreateParameter("AmendmentDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Text))
  cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType))
  cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(Trim$(txtordnum)))
  cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt, "yyyy-MM-dd"))
  Call AddCommonCmdParameters(cmd, , True)
  cmd.Execute
  '*****
  
  '@@@@@@@SAL_INS_ORDER_HDR
  
  sql = "insert into SAL_Order_Hdr"
  sql = sql & " Select Distinct OrderType,OrderNo,OrderDate,OrderSource,CustomerCode,ReferenceNo,ReferenceDate,'" & Format(dpvalidfrom.Value, "yyyy-MM-dd") & "','" & Format(dpvalidto.Value, "yyyy-MM-dd") & "',"
  sql = sql & " OrderingType,Remarks,'" & txtForm.Tag & "','" & txtPack.Tag & "','" & txtDelIns.Tag & "','" & txtTerms.Tag & "','" & txtInsurance.Tag & "','" & txtMode.Tag & "',"
  sql = sql & " '" & txtFreight.Tag & "','" & txtFormula.Tag & "','" & txtGrandTotal.Text & "',PostingAc,'" & txtReLevel.Text & "',GeneralDetails,'M',"
  sql = sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "' , Status, CreatedDate ,'" & gUserID & "','" & gUserID & "',"
  sql = sql & " Branchid, Compid, Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  sql = sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') "
  sql = sql & " Where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo='" & Val(txtordnum.Text) & "' and OrderDate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
  sql = sql & " and Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  'Update Selected Order Finyear for Generate DocNo to this Financial year
  
  sql = " Update SAL_Order_Hdr Set Finyear = '" & Trim$(txtFinyear.Text) & "'"
  sql = sql & " Where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo='" & Val(txtordnum.Text) & "' and OrderDate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
  sql = sql & " and Branchid='" & gBranchID & "' and compid='" & gCompID & "' and Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  '@@@@@@@@
  
  '!!!!!!!!!!!SAL_INS_ORDER_DTL
    For i = 1 To spAmd.DataRowCnt
      If i <= spAmd.GetColItemData(eItem.mFlag) Then
        If DataChanged(i) = True Then
          commClear cmd
          cmd.CommandText = "SAL_INS_ORDER_DTL"
          cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
          cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
          cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
          spAmd.GetText eItem.mItem, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
          spAmd.GetText eItem.mDesc, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("ItemDescription", adVarChar, adParamInput, 50, Trim$(tmp))
          spAmd.GetText eItem.mPartNo, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
          spAmd.GetText eItem.mRevNo, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adVarChar, adParamInput, 5, Trim$(tmp))
          spAmd.GetText eItem.mUOM, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("UOM", adInteger, adParamInput, , Val(tmp))
          spAmd.GetText eItem.mQty, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
          spAmd.GetText eItem.mRate, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
          spAmd.GetText eItem.mARate, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
          spAmd.GetText eItem.mStatus, i, tmp
          cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, IIf(Trim$(tmp) = "Active", "M", "C"))
          cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
          Call AddCommonCmdParameters(cmd, , True)
          cmd.Execute
        End If
      Else
        commClear cmd
        cmd.CommandText = "SAL_INS_ORDER_DTL"
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
        spAmd.GetText eItem.mItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        spAmd.GetText eItem.mDesc, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemDescription", adVarChar, adParamInput, 50, Trim$(tmp))
        spAmd.GetText eItem.mPartNo, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(tmp))
        spAmd.GetText eItem.mRevNo, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adVarChar, adParamInput, 5, Trim$(tmp))
        spAmd.GetText eItem.mUOM, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("UOM", adInteger, adParamInput, , Val(tmp))
        spAmd.GetText eItem.mQty, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(tmp))
        spAmd.GetText eItem.mRate, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
        spAmd.GetText eItem.mARate, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
        spAmd.GetText eItem.mStatus, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, IIf(Trim$(tmp) = "Active", "M", "C"))
        cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd HH:mm:ss"))
        'Call AddCommonCmdParameters(cmd, , Trim$(txtFinyear.Text))
        Call AddCommonCmdParameters(cmd, , True)
        cmd.Execute
      End If
     '!!!!!!!!
  If CmdOk.Tag <> "Open" Then
    With spSchedule
      spAmd.sheet = i
      sql = "delete from SAL_Order_Schedule"
      sql = sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo=" & Val(txtordnum.Text) & " and orderdate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
      spAmd.GetText eItem.mItem, i, retval
      sql = sql & " and ItemCode='" & Trim$(retval) & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      
      For sCnt = 1 To spAmd.DataRowCnt
        cmd.CommandText = "Sal_Ins_Order_Schedule"
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
        spAmd.GetText eItem.mItem, i, retval
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(retval))
        spAmd.GetText eItem.mPartNo, i, retval
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retval))
        spAmd.GetText eSch.mSchDate, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(retval, "yyyy-MM-dd"))
        spAmd.GetText eSch.mSchQty, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(retval))
        spAmd.GetText eSch.mLocation, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(retval))
        Call AddCommonCmdParameters(cmd, , True)
        cmd.Execute
      Next sCnt
      spSchedule.sheet = i
      sql = "delete from SAL_Order_Schedule"
      sql = sql & " where OrderType='" & Trim$(txtOrdType.Text) & "' and OrderNo=" & Val(txtordnum.Text) & " and orderdate='" & Format(txtorddt.Text, "yyyy-MM-dd") & "' "
      spAmd.GetText eItem.mItem, i, retval
      sql = sql & " and ItemCode='" & Trim$(retval) & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      
      For sCnt = 1 To spSchedule.DataRowCnt
        cmd.CommandText = "Sal_Ins_Order_Schedule"
        commClear cmd
        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(txtOrdType.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(txtordnum.Text))
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(txtorddt.Text, "yyyy-MM-dd"))
        spAmd.GetText eItem.mItem, i, retval
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(retval))
        spAmd.GetText eItem.mPartNo, i, retval
        cmd.Parameters.Append cmd.CreateParameter("PartNo", adVarChar, adParamInput, 20, Trim$(retval))
        spSchedule.GetText eSch.mSchDate, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleDate", adDBTimeStamp, adParamInput, 5, Format(retval, "yyyy-MM-dd"))
        spSchedule.GetText eSch.mSchQty, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("ScheduleQty", adDouble, adParamInput, , Val(retval))
        spSchedule.GetText eSch.mLocation, sCnt, retval
        cmd.Parameters.Append cmd.CreateParameter("Location", adVarChar, adParamInput, 50, Trim$(retval))
        Call AddCommonCmdParameters(cmd, , True)
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
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If

  sShowMessage "Amendment Number :" & Trim$(cmbamdtype) & " " & Trim$(txtamdno) & "  Is Generated"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set amdnorec = Nothing
  Set cmd = Nothing
End Try
End Sub

Private Function SaveCalcDetails() As Boolean
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    sql = "delete from COM_FormulaCalc_Hdr where DocType = '" & Trim$(txtOrdType) & "' "
    sql = sql + " and DocNo = " & Val(txtordnum) & " and DocDate = '" & Format(txtorddt, "yyyy-MM-dd") & "'"
    sql = sql + " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
    sql = "delete from COM_FormulaCalc_Dtl where DocType = '" & Trim$(txtOrdType) & "' "
    sql = sql + " and DocNo = " & Val(txtordnum) & " and DocDate = '" & Format(txtorddt, "yyyy-MM-dd") & "'"
    sql = sql + " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    CurDate = ServerDate
    
    sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(txtOrdType) & "' DocType," & Val(txtordnum) & " DocNo,"
    sql = sql & "'" & Format(txtorddt, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    sql = sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,'',"
    sql = sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear , '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    sql = sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    sql = sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    sql = sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    sql = sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    sql = "Insert into COM_FormulaCalc_Dtl"
    sql = sql & " Select '" & Trim$(txtOrdType) & "', " & Val(txtordnum) & ", '" & Format(txtorddt, "yyyy-MM-dd") & "',"
    sql = sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, "
    sql = sql & " a.IsTypePercentage, a.IsValueEditable, a.ValueUsed, a.IsValueEditable,'', a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    sql = sql & " from COM_tmpFormulaCalc_Dtl a"
    sql = sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    sql = sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  SaveCalcDetails = True
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
  
  Call SetDocNo
'  LoadAmdtype
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

Private Sub spAmd_ComboSelChange(sender as object, e as ComboSelChangeEvent)

  Try
  Me.MousePointer = vbHourglass

  If Col = eItem.mStatus Then
    Call SetTotalGoodsValue
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spAmd_DblClick(sender as object, e as DblClickEventArgs)
  Try
  Me.MousePointer = vbHourglass

  Sprhlp
  Avoid_Duplicate spAmd, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Sprhlp()
  Dim sql As String
  Dim i As Long
  Dim sRow As Long
  
  Try

  If spAmd.ActiveCol = 1 Then
    sql = "select distinct a.ItemCode,c.ItemDescription,a.PartNo,c.UOMDesc,c.UOMDecimal,a.Rate,a.AssessableRate,a.FormulaCode,"
    sql = sql + "(select top 1 b.FrmLName from COM_FormulaMast_Hdr b where b.ModuleCode = '" & gModuleCode & "' and b.FrmId = a.FormulaCode"
    sql = sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid) FromulaName, c.UOM from SAL_Customer_ItemMaster a,MATERIAL_MASTER c"
    sql = sql + " where CustomerCode = '" & Trim$(txtcustcd.Tag) & "' and a.rate is not null and a.Branchid = '" & gBranchID & "' "
    sql = sql + " and a.compid='" & gCompID & "' and a.ItemCode = c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
    If FetchValue_New(True, db.ConnectionString, sql, "1111010000", "List of Items for the Customer", "Item Code", "Item Description", "Part No", "UOM", "Rate") = True Then
      For i = 1 To UBound(getvalue, 1)
        If getvalue(i, 1) <> "" And (getvalue(1, 8) = getvalue(i, 8)) Then
          sRow = spAmd.DataRowCnt + 1
          spAmd.SetText eItem.mItem, sRow, getvalue(i, 1)  'ItemCode
          spAmd.SetText eItem.mDesc, sRow, getvalue(i, 2)  'ItemDesc
          spAmd.SetText eItem.mPartNo, sRow, getvalue(i, 3)  'Part No
          spAmd.SetText eItem.mUOM, sRow, getvalue(i, 4)  'UOMDesc
          spAmd.Col = eItem.mUOM
spAmd.Row = sRow
          spAmd.CellTag = getvalue(i, 10)  'UOMCode
          If Trim$(CmdOk.Tag) = "Open" Then
            Call Spread_Lock(spAmd, True, eItem.mQty, sRow)
          Else
            Call Spread_Lock(spAmd, False, eItem.mQty, sRow)
          End If
          Call Spread_Set_Decimal(spAmd, Val(getvalue(i, 5)), eItem.mQty, sRow, eItem.mQty)  'UOM Decimal
          spAmd.SetText eItem.mQty, sRow, 0 ' 0 Qty
          spAmd.SetText eItem.mQty, sRow, 0 ' 0 Qty
          spAmd.SetText eItem.mRate, sRow, Val(getvalue(i, 6))  'Rate
          spAmd.SetText eItem.mARate, sRow, Val(getvalue(i, 7))  'Asseassable Rate
          spAmd.SetText eItem.mStatus, sRow, "Active"     'Status
          spAmd.MaxRows = spAmd.DataRowCnt
          OrderCalulation  'to calculate the Order Value
        End If
      Next i
      If Trim(SpVal(spAmd, 1, eItem.mItem)) <> "" Then
 Spread_Set_Focus spAmd, 1, eItem.mQty
End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spAmd_EditMode(sender as object, e as EditModeEventArgs)
  Dim retval As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mQty, eItem.mARate, eItem.mRate
          spAmd.GetText Col, Row, retval
          
          If Val(retval) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
          End If
          
          Select Case Col
            Case eItem.mQty
              If Val(retval) > 0 Then
                spAmd.SetActiveCell eItem.mRate, Row
              End If
            Case eItem.mRate
              If Val(retval) > 0 Then
                spAmd.SetActiveCell eItem.mARate, Row
              End If
            Case eItem.mARate
              If Val(retval) > 0 Then
                spAmd.SetActiveCell eItem.mItem, Row + 1
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

Private Sub spAmd_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If KeyCode = 13 And spAmd.ActiveCol = eItem.mStatus Then
    spAmd.MaxRows = spAmd.DataRowCnt + 1
    If CmdOk.Tag = "Open" Then
      Spread_Lock spAmd, True, eItem.mQty, 0, eItem.mQty, spAmd.MaxRows
    Else
      Spread_Lock spAmd, False, eItem.mQty, 0, eItem.mQty, spAmd.MaxRows
    End If
  End If
End Sub

Private Sub spAmd_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retval As Variant
  Dim sRow As Long
  
  Try
  
  If NewRow = -1 Then
    sRow = spAmd.ActiveRow
  Else
    sRow = NewRow
  End If
  
  If NewRow <> Row Then
    Call SetSheet(sRow)
    spAmd.GetText eItem.mItem, sRow, retval
    txtItemcode.Text = Trim$(retval)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spSchedule_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim retval As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If spSchedule.ActiveCol = eSch.mLocation Then
      If KeyCode = vbKeyReturn Then
          spSchedule.sheet = spSchedule.ActiveSheet
          spSchedule.GetText eSch.mSchDate, spSchedule.ActiveRow, retval
          If IsDate(Format(retval, "dd/MM/yyyy")) = True Then
            spSchedule.GetText eSch.mSchQty, spSchedule.ActiveRow, retval
            If Val(retval) <= 0 Then
              GoTo ExitHere
            End If
            spSchedule.MaxRows = spSchedule.DataRowCnt + 1
          Else
            GoTo ExitHere
          End If
       End If
  End If
  
  If KeyCode = vbKeyDelete Then
    If spSchedule.ActiveCol = eSch.mSchDate Then
      spSchedule.DeleteRows spSchedule.ActiveRow, 1
      spSchedule.MaxRows = spSchedule.DataRowCnt
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spSchedule_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim retval As Variant
  Dim sRow As Variant
  
  Try
  
    If spSchedule.ActiveCol = eSch.mSchDate Then
      spSchedule.GetText eSch.mSchDate, spSchedule.ActiveRow, retval
      retval = spSchedule.SearchCol(eSch.mSchDate, -1, spSchedule.DataRowCnt, retval, SearchFlagsPartialMatch)
      sRow = sRow + retval
      If spSchedule.ActiveRow <> retval Then
        If Val(retval) <> -1 Then
          spSchedule.ClearRange eSch.mSchDate, spSchedule.ActiveRow, eSch.mLocation, spSchedule.ActiveRow, True
        End If
      Else
        spSchedule.GetText eSch.mSchDate, spSchedule.ActiveRow, retval
        retval = spSchedule.SearchCol(eSch.mSchDate, sRow, spSchedule.DataRowCnt, retval, SearchFlagsPartialMatch)
        If spSchedule.ActiveRow <> retval Then
          If Val(retval) <> -1 Then
            spSchedule.ClearRange eSch.mSchDate, spSchedule.ActiveRow, eSch.mLocation, spSchedule.ActiveRow, True
          End If
        End If
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtDelIns_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description,Code From TERMSMASTER Where typeflg = 'Delivery Instruction' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Delivery Instruction", "Description") = True Then
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
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select description,Code From TERMSMASTER Where typeflg = 'Form' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Form's", "Description") = True Then
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
  
  If spAmd.DataRowCnt <= 0 Then
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
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  sql = "Select Distinct FrmLName, FrmID from COM_FORMULAMAST_HDR"
  sql = sql & " where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Formula's", "Formula Name") = True Then
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
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description,Code From TERMSMASTER Where typeflg = 'FreightCode'"
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of FreightCode", "Description") = True Then
    txtFreight.Text = getvalue(1)
    txtFreight.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtInsurance_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description,Code From TERMSMASTER Where typeflg = 'Insurance' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Insurance", "Description") = True Then
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
  Dim sql As String
 
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Mode' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtMode.Text = getvalue(1)
    txtMode.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtPack_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description,Code From TERMSMASTER Where typeflg = 'Packing & Forwarding' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Packing & Forwarding", "Description") = True Then
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
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = " Select Description, Code From TERMSMASTER Where typeflg = 'Payment Terms' "
  sql = sql & " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Payment Terms", "Description") = True Then
    txtTerms.Text = getvalue(1)
    txtTerms.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
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
  Dim sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  Set orddetrec = New AceADODB.Recordset
  
  CurDate = ServerDate

  sql = " Select distinct a.OrderType,a.OrderNo,a.OrderDate,a.CustomerCode,("
  sql = sql & " Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and compid='" & gCompID & "') PartyName,"
  sql = sql & " a.ValidTo, a.ReferenceNo,a.Revisionlevel,a.OrderingType,a.OrderSource,a.ValidFrom,a.ValidTo, a.Finyear "
  sql = sql & " From dbo.fn_SAL_ORDER_HDR('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
  sql = sql & " Where Status='A'"
  If FetchValue_New(False, db.ConnectionString, sql, "1110101", "List of Order", "Order Type", "Order No", "Order Date", "Customer", "PO No") = True Then
    txtOrdType = Trim$(getvalue(1))
    txtordnum = getvalue(2)
    txtorddt = getvalue(3)
    txtcustcd.Tag = getvalue(4)
    txtcustcd = getvalue(5)
    txtReLevel.Text = getvalue(8)
    txtordertyping.Text = getvalue(9)
    txtOrdSour.Text = getvalue(10)
    dpvalidfrom.Value = getvalue(11)
    dpvalidto.Value = getvalue(12)
    CmdOk.Tag = Trim$(getvalue(9))
    txtFinyear.Text = getvalue(13) 'Finyear
    
    If Not Trim$(getvalue(6)) = "" Then
      If dpamddate > CDate(getvalue(6)) Then
        ret = MsgBox("Amendment Date is greater than Valid till Date of Order" & vbCrLf & " Do you want to continue Amending the Order ?", vbYesNo + vbQuestion, App.Title)

        If Val(ret) <> 6 Then
 Call CmdClear_Click
txtordnum.SetFocus
GoTo ExitHere
End If
      End If
    End If
  If Trim$(getvalue(9)) = "Open" Then
    fraSch.Visible = False
    Spread_Lock spAmd, True, eItem.mQty, -1, eItem.mQty
  Else
    fraSch.Visible = True
    Spread_Lock spAmd, False, eItem.mQty, -1, eItem.mQty
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
  Dim retval As Variant
  Dim ItemList() As String
  Dim SCValList() As String
'  Dim ActiveItemCount As Integer
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(spAmd.DataRowCnt - 1)
      ReDim SCValList(spAmd.DataRowCnt * 2 - 1)

'      ActiveItemCount = 0
      
      For rCnt = 1 To spAmd.DataRowCnt
        spAmd.GetText eItem.mQty, rCnt, retval
        Qty = retval
        spAmd.GetText eItem.mStatus, rCnt, retval
        
        'Item List
        spAmd.GetText eItem.mItem, rCnt, retval
        Item = Trim$(retval)
       
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(retval)
        
        'SCValue List
        spAmd.GetText eItem.mQty, rCnt, retval
        Qty = Val(retval)
        
        If Trim$(retval) <> "Active" Then
          Qty = 0   'To avoid the value of Deactive Items
        End If
        
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        spAmd.GetText eItem.mRate, rCnt, retval
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retval)
        
        spAmd.GetText eItem.mARate, rCnt, retval
        SCValList((rCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retval)
      Next rCnt
      
'      If ActiveItemCount <= 0 Then
'        GoTo ExitHere
'      End If
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

  StartCalcFormula = True
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
  Dim retval(1 To 3) As Variant
  Dim tmpTot As Double
    
  Try
  
    spAmd.GetText eItem.mQty, fRow, retval(1)
    retval(2) = 1
    
    spAmd.GetText eItem.mRate, fRow, retval(3)
    tmpTot = (Val(retval(1)) / Val(retval(2)) * Val(retval(3)))
    spAmd.SetText eItem.mGV, fRow, Format(tmpTot, "0.00")
    
    spAmd.GetText eItem.mARate, fRow, retval(3)
    tmpTot = (Val(retval(1)) / Val(retval(2)) * Val(retval(3)))
    spAmd.SetText eItem.mAGV, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(spAmd, eItem.mGV), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(spAmd, eItem.mAGV), "0.00")
End Sub

Private Function SprTotCalc(ByRef fSpread As AceSpread,ByVal  fCol As Long) As Double
  Dim rCnt As Long
  Dim retval As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText eItem.mStatus, rCnt, retval
      If Trim$(retval) = "Active" Then
        fSpread.GetText fCol, rCnt, retval
        Total = Total + Val(retval)
      End If
    Next rCnt
  
  SprTotCalc = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_spSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spSchedule


spSchedule.KeyDown +=  new EventHandler(spSchedule_KeyDown)
spSchedule.LeaveCell +=  new EventHandler(spSchedule_LeaveCell)

End Sub

Private Sub Design_spAmd(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spAmd


spAmd.ComboSelChange +=  new EventHandler(spAmd_ComboSelChange)
spAmd.DblClick +=  new EventHandler(spAmd_DblClick)
spAmd.EditMode +=  new EventHandler(spAmd_EditMode)
spAmd.KeyDown +=  new EventHandler(spAmd_KeyDown)
spAmd.LeaveCell +=  new EventHandler(spAmd_LeaveCell)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_spSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAmendment.frm", spSchedule, "")

	Call Design_spAmd("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAmendment.frm", spAmd, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAmendment.frm", sprCalc, "")

End Sub
