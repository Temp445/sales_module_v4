VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSalesPlan 
   Caption         =   "Sales Plan"
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
   Begin VB.CommandButton cmdSetZero 
      Caption         =   "Set &Zero"
      Enabled         =   0   'False
      Height          =   400
      Left            =   10530
      TabIndex        =   26
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6990
      Width           =   1300
   End
   Begin VB.Frame fraDetails 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   1185
      TabIndex        =   5
      Top             =   180
      Width           =   9075
      Begin VB.Frame fraItemNature 
         Caption         =   "Select If Filter Need"
         Height          =   2445
         Left            =   4710
         TabIndex        =   23
         Top             =   120
         Width           =   4320
Begin AceSpread sprItem
            Height          =   2130
            Left            =   105
            TabIndex        =   24
            Top             =   210
            Width           =   4110
            _Version        =   458752
            _ExtentX        =   7250
            _ExtentY        =   3757
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
            MaxCols         =   2
            MaxRows         =   1
            OperationMode   =   4
         End
      End
      Begin VB.Frame fraCommon 
         Caption         =   "Select"
         Height          =   2445
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   4650
         Begin VB.Frame fraPeriod 
            Caption         =   "Period"
            Height          =   585
            Left            =   495
            TabIndex        =   17
            Top             =   180
            Width           =   3600
            Begin VB.OptionButton optWeekly 
               Caption         =   " Weekly"
               Height          =   360
               Left            =   1845
               TabIndex        =   21
               Top             =   195
               Width           =   885
            End
            Begin VB.OptionButton optDaily 
               Caption         =   "Daily"
               Height          =   225
               Left            =   2820
               TabIndex        =   20
               Top             =   255
               Width           =   645
            End
            Begin VB.OptionButton optyearly 
               Caption         =   "Yearly"
               Height          =   360
               Left            =   90
               TabIndex        =   19
               Top             =   180
               Value           =   -1  'True
               Width           =   810
            End
            Begin VB.OptionButton optMonthly 
               Caption         =   "Monthly"
               Height          =   225
               Left            =   915
               TabIndex        =   18
               Top             =   255
               Width           =   855
            End
         End
         Begin VB.Frame fraMonth 
            Height          =   750
            Left            =   90
            TabIndex        =   8
            Top             =   705
            Width           =   4470
            Begin VB.TextBox TxtPlanType 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   75
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   315
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtPlanNo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   315
               Width           =   825
            End
            Begin VB.ComboBox CmbPlanType 
               Height          =   330
               ItemData        =   "frmSalesPlan.frx":02E0
               Left            =   75
               List            =   "frmSalesPlan.frx":02E2
               TabIndex        =   9
               Text            =   "CmbPlanType"
               ToolTipText     =   "Select the Production Plan Type"
               Top             =   315
               Width           =   765
            End
            Begin MSComCtl2.DTPicker dtpPlanMonth 
               Height          =   315
               Left            =   2835
               TabIndex        =   11
               ToolTipText     =   "Select the Production Plan Month"
               Top             =   315
               Width           =   1545
               _ExtentX        =   2725
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
               OLEDropMode     =   1
               CustomFormat    =   "MMMM-yyyy"
               Format          =   140509187
               CurrentDate     =   37174
            End
            Begin MSComCtl2.DTPicker dtpPlanDate 
               Height          =   315
               Left            =   1695
               TabIndex        =   12
               Top             =   315
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               OLEDropMode     =   1
               Format          =   112459779
               CurrentDate     =   37174
            End
            Begin VB.Label lblyearly 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Plan Month"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   2835
               TabIndex        =   16
               Top             =   135
               Width           =   780
            End
            Begin VB.Label lblPlanNo 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Plan No"
               ForeColor       =   &H000000C0&
               Height          =   210
               Left            =   1125
               TabIndex        =   15
               Top             =   135
               Width           =   540
            End
            Begin VB.Label lblPlanDate 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Plan Date"
               ForeColor       =   &H000000C0&
               Height          =   210
               Left            =   1695
               TabIndex        =   14
               Top             =   135
               Width           =   675
            End
            Begin VB.Label lblPlanType 
               AutoSize        =   -1  'True
               Caption         =   "Plan Type"
               Height          =   210
               Left            =   75
               TabIndex        =   13
               Top             =   135
               Width           =   705
            End
         End
         Begin VB.CheckBox chkCustItem 
            Caption         =   "From Customer Item Master"
            Height          =   225
            Left            =   1170
            TabIndex        =   7
            Top             =   2130
            Value           =   1  'Checked
            Width           =   2265
         End
Begin AceSpread sprType
            Height          =   525
            Left            =   1065
            TabIndex        =   22
            Top             =   1500
            Width           =   2445
            _Version        =   458752
            _ExtentX        =   4313
            _ExtentY        =   926
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
            MaxRows         =   1
            ScrollBars      =   0
         End
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   480
      Left            =   10260
      TabIndex        =   4
      Top             =   2265
      Width           =   705
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6555
      TabIndex        =   0
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6990
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5265
      TabIndex        =   1
      ToolTipText     =   "Click to Clear the Data"
      Top             =   6990
      Width           =   1300
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3975
      TabIndex        =   2
      ToolTipText     =   "Click to Save the Data"
      Top             =   6990
      Width           =   1300
   End
Begin AceSpread sprPlan
      Height          =   4095
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Enter the Production Plan Details"
      Top             =   2805
      Width           =   11745
      _Version        =   458752
      _ExtentX        =   20717
      _ExtentY        =   7223
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   4
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   9
      MaxRows         =   100
      SelectBlockOptions=   0
      UserResize      =   1
      VisibleCols     =   6
   End
End
Attribute VB_Name = "frmSalesPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public SalesPlan As String

Const DefaultCol = 9

Private Enum ePlan
  mCustCode = 1
  mCustName = 2
  mPartNo = 3
  mItem = 4
  mDesc = 5
  mUOM = 6
  mRate = 7
  mPlanQty = 8
  mPlanValue = 9
End Enum

Private Enum ePlanFor
  mYear = 1
  mMonth = 2
  mWeek = 3
  mDay = 4
End Enum

Private Sub ExistTypeLoad()
  Dim Sql As String
  Dim rs As AceADODB.Recordset

  Dim rCnt As Long
  Dim PlanFor As Integer
  
  Try
  
'  Sql = " Select Distinct DocumentType, Description, ModuleName"
'  Sql = Sql & " From Com_DocumentType_Master "
'  Sql = Sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and ModuleName = 'SAL'"
'  Sql = Sql & " and PertainingTo in ('DP', 'MP', 'WP', 'YP') "
  
  If optyearly.Value = True Then
    PlanFor = ePlanFor.mYear ' "Year"
  ElseIf optMonthly.Value = True Then
    PlanFor = ePlanFor.mMonth '"Month"
  ElseIf optWeekly.Value = True Then
    PlanFor = ePlanFor.mWeek '"Week"
  ElseIf optDaily.Value = True Then
    PlanFor = ePlanFor.mDay '"Day"
  End If
  
  Sql = " Select * From"
    'This Query Existing Plan Not Applied
  Sql = Sql & " (Select 'N'DocType, 'NONE' as Description, 0 as seq"
  Sql = Sql & " Union All"
  ' This Query Retrieve If the year Plan is Exists for this year
  Sql = Sql & " Select Distinct a.DocType, b.Description, 1 as seq"
  Sql = Sql & " From Sal_Salesplan a"
  Sql = Sql & " Left Outer Join Com_DocumentType_Master b"
  Sql = Sql & " On a.DocType = b.DocumentType and a.BranchID = b.BranchID and a.CompID = b.CompID"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and ModuleName = 'SAL'"
  Sql = Sql & " and PlanFor = 1 and year(PlanPeriod) = " & Year(dtpPlanMonth) & "" 'PlanFor = 1 For year Plan
  Sql = Sql & " Union All"
  'This Query For except year Plan Loading
  Sql = Sql & " Select Distinct a.DocType, b.Description, 2 as seq"
  Sql = Sql & " From Sal_Salesplan a"
  Sql = Sql & " Left Outer Join Com_DocumentType_Master b"
  Sql = Sql & " On a.DocType = b.DocumentType and a.BranchID = b.BranchID and a.CompID = b.CompID"
  Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and ModuleName = 'SAL'"
  Sql = Sql & " and PlanFor < " & Val(PlanFor) & " and Month(PlanPeriod) = " & Month(dtpPlanMonth) & " and year(PlanPeriod) = " & Year(dtpPlanMonth) & ""
  Sql = Sql & " )x"
  Sql = Sql & " Order By Seq desc"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
    
  If Not rs.EOF Then
      sprType.TypeComboBoxClear 1, 1 'To Clear ComboBox in Spread
      sprType.TypeComboBoxClear 2, 1
      
      sprType.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprType.Row = 1
sprType.Col = 1
sprType.TypeComboBoxString = rs("Description")

        sprType.Row = 1
sprType.Col = 2
sprType.TypeComboBoxString = rs("DocType")
        rs.MoveNext
      Next rCnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub
Private Function GetMax(ByVal TableName As String,ByVal  MaxFieldName As String,ByVal  ParamArray FldNameAndDataList()) As Long
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim pCnt As Integer
  
  Try

  Sql = "Select Max(Convert(int, " & Trim$(MaxFieldName) & ")) from " & Trim$(TableName) & " where "
  
  If Not IsMissing(FldNameAndDataList) Then
    For pCnt = 0 To UBound(FldNameAndDataList) Step 2
      Sql = Sql & FldNameAndDataList(pCnt) & "='" & FldNameAndDataList(pCnt + 1) & "' and "
    Next pCnt
  End If
  
  Sql = Sql & " BranchID='" & gBranchID & "' and CompID = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
      GetMax = Val(rs(0))
    Else
      GetMax = 0
    End If
  Else
    GetMax = 0
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub TypeLoad()
  Dim Sql As String
  Dim adoAddItem As AceADODB.Recordset
  Dim Pertaining As String

  Try

  Sql = "Select DocumentType from COM_DOCUMENTTYPE_MASTER Where "
  If optyearly.Value = True Then
    Pertaining = "YP"
  ElseIf optMonthly.Value = True Then
    Pertaining = "MP"
  ElseIf optWeekly.Value = True Then
    Pertaining = "WP"
  Else
    Pertaining = "DP"
  End If
  Sql = Sql & " Pertainingto = '" & Trim$(Pertaining) & "' and ModuleName = 'SAL' "
  Sql = Sql & " and Branchid = '" & gBranchID & "' and compid = '" & gCompID & "'"
  Set adoAddItem = New AceADODB.Recordset
  If RSOpen(adoAddItem, Sql) = False Then
 GoTo ExitHere
End If

  If Not adoAddItem.EOF Then
    CmbPlanType.Clear
    Do While Not adoAddItem.EOF = True
      CmbPlanType.AddItem Trim$(adoAddItem("DocumentType"))
      adoAddItem.MoveNext
    Loop
    CmbPlanType.ListIndex = 0
    'txtPlanNo = GetDocNo("Sal_SalesPlan", "DocType", "DocNo", Trim$(CmbPlanType), Trim$(Pertaining))
  End If
  adoAddItem.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoAddItem = Nothing
End Try
End Sub

Private Sub ItemNatureLoad()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Integer
  Dim ItemNature As String
  
  Try
  
  Sql = " Select distinct ItemNature From Inv_Material_Master"
  Sql = Sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and Status=1"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  If Not rs.EOF Then
      sprItem.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprItem.SetText 1, rCnt, Trim$(rs("ItemNature"))
        ItemNature = SetItemNature(Trim$(rs("ItemNature")))
        sprItem.SetText 2, rCnt, Trim$(ItemNature)
        rs.MoveNext
      Next rCnt
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub PlanQtyFill(ByVal fRow As Long,ByVal  DataFor As ePlanFor,ByVal  PlanFor As ePlanFor,ByVal  PlanPeriod As Date,ByVal  PlanQty As Double,ByVal  UOMDecimal As Integer)
  Dim ColStart As Long
  Dim ColEnd As Long
  Dim fCol As Long
  Dim fCnt As Integer
  
  Dim retVal As Variant
  
  Try
  
    If optyearly.Value = True Then
      ColStart = DefaultCol - 1
      ColEnd = ColStart
    Else
      ColStart = DefaultCol + 1
      ColEnd = sprPlan.MaxCols
    End If
    
'    Spread_Decimal_Set sprPlan, ColStart, ColEnd, fRow, fRow, Val(UOMDecimal) 'UOM Decimal
    Call Spread_Set_Decimal(sprPlan, Val(UOMDecimal), ColStart, fRow, ColEnd)
    Select Case PlanFor
      Case mYear
          Select Case DataFor
            Case mYear
              sprPlan.SetText ePlan.mPlanQty, fRow, PlanQty
            Case Else
              'No need to fill
          End Select
      Case mMonth
          Select Case DataFor
            Case mYear
              For fCol = ColStart To ColEnd
                sprPlan.SetText fCol, fRow, (PlanQty \ 12)
              Next fCol
              
              sprPlan.SetText ColStart, fRow, (PlanQty \ 12) + (PlanQty - Spread_Cell_Total(sprPlan, DefaultCol + 1, fRow, ColEnd, fRow))
            Case mMonth
              fCol = Month(PlanPeriod)    'Month No
              fCol = IIf(fCol <= 3, fCol + 12, fCol) - 3  'Add 12 to Jan, Feb, Mar and Subtract 3 to get seq
              fCol = DefaultCol + fCol
              
              sprPlan.SetText fCol, fRow, PlanQty
            Case Else
              'No need to fill
          End Select
      Case mWeek
          Select Case DataFor
            Case mYear
              For fCol = ColStart To ColEnd
                sprPlan.SetText fCol, fRow, (PlanQty \ (12 * 4))
              Next fCol
              
              sprPlan.SetText ColStart, fRow, (PlanQty \ (12 * 4)) + ((PlanQty \ 12) - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
            Case mMonth
              For fCol = ColStart To ColEnd
                sprPlan.SetText fCol, fRow, (PlanQty \ 4)
              Next fCol
              
              sprPlan.SetText ColStart, fRow, (PlanQty \ 4) + (PlanQty - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
            Case mWeek
              fCol = (Day(PlanPeriod) \ 7) + 1    'Week No
              fCol = IIf(fCol > 4, fCol - 1, fCol)  'If 5th week present make it as 4th week
              fCol = DefaultCol + fCol
              
              sprPlan.SetText fCol, fRow, PlanQty
            Case Else
              'No need to fill
          End Select
      Case mDay
          Select Case DataFor
            Case mYear
            
              fCnt = 0
              For fCol = ColStart To ColEnd
                sprPlan.GetText fCol + fCnt, 1, retVal
                Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                  fCnt = fCnt + 1
                  sprPlan.GetText fCol + fCnt, 1, retVal
                Loop
                
                sprPlan.SetText fCol + fCnt, fRow, (PlanQty \ (12 * 28))
              Next fCol
            
              fCnt = ColStart
              sprPlan.GetText fCnt, 1, retVal
              Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                fCnt = fCnt + 1
                sprPlan.GetText fCnt, 1, retVal
              Loop
              
              sprPlan.GetText fCnt, fRow, retVal
              sprPlan.SetText fCnt, fRow, (retVal \ (12 * 28)) + ((PlanQty \ 12) - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
              
'              fCnt = 0
'              For fCol = ColStart To ColEnd
'                .GetText fCol, fRow, retVal
'                If Trim$(retVal) = "H" Then
'                  fCnt = fCnt + 1
'                End If
'                .SetText fCol + fCnt, fRow, (PlanQty \ (12 * 28))
'              Next fCol
'              .SetText ColStart, fRow, (PlanQty \ (12 * 28)) + ((PlanQty \ 12) - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
            Case mMonth
              
              fCnt = 0
              For fCol = ColStart To ColEnd
                sprPlan.GetText fCol + fCnt, 1, retVal
                Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                  fCnt = fCnt + 1
                  sprPlan.GetText fCol + fCnt, 1, retVal
                Loop
                
                sprPlan.SetText fCol + fCnt, fRow, (PlanQty \ (28))
              Next fCol
              
               fCnt = ColStart
              sprPlan.GetText fCnt, 1, retVal
              Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                fCnt = fCnt + 1
                sprPlan.GetText fCnt, 1, retVal
              Loop
              
              sprPlan.GetText fCnt, fRow, retVal
              sprPlan.SetText fCnt, fRow, Val(retVal \ 28) + (PlanQty - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
              
'              fCnt = 0
'              For fCol = ColStart To ColEnd
'                .GetText fCol, fRow, retVal
'                If Trim$(retVal) = "H" Then
'                  fCnt = fCnt + 1
'                End If
'                .SetText fCol + fCnt, fRow, (PlanQty \ 28)
'              Next fCol
'
'              .SetText ColStart, fRow, (PlanQty \ 28) + (PlanQty - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
            Case mWeek
              fCol = ((Day(PlanPeriod) - 1) \ 7) + 1  'Week No
              fCol = IIf(fCol > 4, fCol - 1, fCol)  'If 5th week present make it as 4th week
            
              ColStart = DefaultCol + ((fCol * 7) - 6)
              ColEnd = ColStart + 5

              fCnt = 0
              For fCol = ColStart To ColEnd
                sprPlan.GetText fCol + fCnt, 1, retVal
                Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                  fCnt = fCnt + 1
                  sprPlan.GetText fCol + fCnt, 1, retVal
                Loop
                
                sprPlan.SetText fCol + fCnt, fRow, (PlanQty \ (6))
              Next fCol
              
              fCnt = ColStart
              sprPlan.GetText fCnt, 1, retVal
              Do While Trim$(retVal) = "H" Or Trim$(retVal) = "S"
                fCnt = fCnt + 1
                sprPlan.GetText fCnt, 1, retVal
              Loop
              
              sprPlan.GetText fCnt, fRow, retVal
              sprPlan.SetText fCnt, fRow, Val(retVal) + (PlanQty - Spread_Cell_Total(sprPlan, ColStart, fRow, ColEnd, fRow))
            Case mDay
              fCol = Day(PlanPeriod)    'Date
              fCol = DefaultCol + fCol
              
              sprPlan.SetText fCol, fRow, PlanQty
            Case Else
              'No need to fill
          End Select
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub RunTimeSpreadDesign(ByVal PlanFor As Integer)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim cCnt  As Integer
  Dim rCnt As Integer
  Dim NoofDays As Integer
  
  Dim Mth As String
  Dim Holiday As Date
  
  Try
  
    sprPlan.Row = 1
    sprPlan.RowHidden = True
    Select Case PlanFor
      Case 1 'Year
        sprPlan.BlockMode = True
        sprPlan.BackColor = vbWhite
        sprPlan.BlockMode = False
        sprPlan.MaxCols = DefaultCol
      Case 2 'Month
        sprPlan.BlockMode = True
        sprPlan.BackColor = vbWhite
        sprPlan.BlockMode = False
        sprPlan.MaxCols = DefaultCol + 12
        rCnt = DefaultCol
        For cCnt = 1 To 12
          'Mth = MonthName(((cCnt + 2) Mod 12) + 1)
          Mth = Format(DateAdd("m", cCnt - 1, "01-" & Left(gFinyear, 7)), "MM-yyyy")
           rCnt = rCnt + 1
          sprPlan.SetText rCnt, 0, Trim$(Mth)
        Next cCnt
        
       'Spread set celltype as Number, Align as Right
        For cCnt = (DefaultCol + 1) To sprPlan.MaxCols
          sprPlan.Col = cCnt
sprPlan.Col2 = cCnt

          sprPlan.Row = 2
sprPlan.Row2 = -1
          sprPlan.BlockMode = True
          sprPlan.CellType = CellTypeNumber
          sprPlan.TypeHAlign = TypeHAlignRight
          sprPlan.BlockMode = False
         Next cCnt
      Case 3 'Week
        sprPlan.MaxCols = DefaultCol + 4
        For cCnt = (DefaultCol + 1) To sprPlan.MaxCols
          sprPlan.SetText cCnt, 0, "Week - " & (cCnt - DefaultCol)
           'Spread set celltype as Number, Align as Right
          sprPlan.Col = cCnt
sprPlan.Col2 = cCnt

          sprPlan.Row = 2
sprPlan.Row2 = -1
          sprPlan.BlockMode = True
          sprPlan.BackColor = vbWhite
          sprPlan.CellType = CellTypeNumber
          sprPlan.TypeHAlign = TypeHAlignRight
          sprPlan.BlockMode = False
        Next cCnt
       
      Case 4 'Day
        sprPlan.BlockMode = True
        sprPlan.BackColor = vbWhite
        sprPlan.BlockMode = False
        NoofDays = Day(DateAdd("d", -1, DateAdd("m", 1, Format(dtpPlanMonth.Value, "yyyy-MM-" & "01"))))
        sprPlan.MaxCols = DefaultCol + NoofDays
       
        For cCnt = (DefaultCol + 1) To sprPlan.MaxCols
          Holiday = (cCnt - DefaultCol) & Format(dtpPlanMonth, "-MM-yyyy")
          If Weekday(Holiday) = vbSunday Then
            sprPlan.BlockMode = True
            sprPlan.Col = cCnt
sprPlan.Col2 = cCnt

            sprPlan.Row = 2
sprPlan.Row2 = -1
            sprPlan.BackColor = vbYellow
            sprPlan.BlockMode = False
            sprPlan.SetText cCnt, 1, "H"
          End If
          sprPlan.SetText cCnt, 0, (cCnt - DefaultCol)
        Next cCnt
        'Spread Set celltype as Number, Align as Right
        For cCnt = (DefaultCol + 1) To sprPlan.MaxCols
          sprPlan.Col = cCnt
sprPlan.Col2 = cCnt

          sprPlan.Row = 2
sprPlan.Row2 = sprPlan.MaxRows
          sprPlan.BlockMode = True
          sprPlan.CellType = CellTypeNumber
          sprPlan.TypeHAlign = TypeHAlignRight
          sprPlan.BlockMode = False
        Next cCnt
          
        Sql = " Select HolidayDate From PAY_Holiday_Dtl"
        Sql = Sql & " Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and Month(HolidayDate) = '" & Month(dtpPlanMonth) & "'"
        Sql = Sql & " and Year(HolidayDate) = '" & Year(dtpPlanMonth) & "'"
        Set rs = New AceADODB.Recordset
        If RSOpen(rs, Sql) = False Then
          GoTo ExitHere
        End If
        If Not rs.EOF Then
          For cCnt = 1 To rs.RecordCount
            For rCnt = 1 To NoofDays
              If Day(rs("HolidayDate")) = rCnt Then
                sprPlan.Col = (DefaultCol + rCnt)
sprPlan.Col2 = sprPlan.Col

                sprPlan.Row = 1
sprPlan.Row2 = -1
                sprPlan.BlockMode = True
                sprPlan.BackColor = vbRed
                sprPlan.Lock = True
                sprPlan.BlockMode = False
                sprPlan.SetText (DefaultCol + rCnt), 1, "S"
              End If
            Next rCnt
            rs.MoveNext
          Next cCnt
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function SaveSalesPlanDetail(ByVal fRevNo As Integer,ByVal  CurDate As Date) As Boolean
  Dim Sql As String
  
  Dim retVal As Variant
  Dim rCnt As Long
  Dim cCnt As Long
  Dim PlanFor As Integer
  Dim Pertaining As String
 
  Dim Period As Date
  
  Dim ColStart As Long
  Dim ColEnd As Long
  
  Try
  SaveSalesPlanDetail = False
   
  If SalesPlan = "Create" Then
    If optyearly.Value = True Then
      Pertaining = "YP"
    ElseIf optMonthly.Value = True Then
      Pertaining = "MP"
    ElseIf optWeekly.Value = True Then
      Pertaining = "WP"
    Else
      Pertaining = "DP"
    End If
    txtPlanNo = GetDocNo("SAL_SalesPlan", "DocType", "DocNo", Trim$(CmbPlanType), Trim$(Pertaining))
  End If
  
    If optyearly.Value = True Then
      ColStart = DefaultCol - 1
      ColEnd = ColStart
    Else
      ColStart = DefaultCol + 1
      ColEnd = sprPlan.MaxCols
    End If
    
    For rCnt = 2 To sprPlan.DataRowCnt
      For cCnt = ColStart To ColEnd
'        .GetText ccnt, RCnt, retVal
'        If Val(retVal) > 0 Then
          Sql = "Execute SAL_Ins_SalesPlan"
          
          If SalesPlan = "Create" Then
            Sql = Sql & " @DocType = '" & Trim$(CmbPlanType.Text) & "',"
          Else
            Sql = Sql & " @DocType = '" & Trim$(TxtPlanType.Text) & "',"
          End If
          
          Sql = Sql & " @DocNo = " & Val(txtPlanNo.Text) & ","
          Sql = Sql & " @DocDate = '" & Format(dtpPlanDate, "yyyy-MM-dd") & "',"
          
          If optyearly.Value = True Then
            PlanFor = ePlanFor.mYear 'Year
            Period = Format(dtpPlanMonth, "yyyy") & "-04-01"
          ElseIf optMonthly.Value = True Then
            PlanFor = ePlanFor.mMonth 'Month
            ''Period = Format(dtpPlanMonth, "yyyy-") & Format((((cCnt - DefaultCol) + 2) Mod 12) + 1, "00") & "-01"
'            If Period = "2008-07-01" Then MsgBox "asdfasdf"
            sprPlan.GetText cCnt, 0, retVal
            Period = Format("01-" & retVal, "yyyy-MM-dd")
          ElseIf optWeekly.Value = True Then
            PlanFor = ePlanFor.mWeek 'Week
            Period = Format(dtpPlanMonth, "yyyy-MM") & "-" & Format(((((cCnt - DefaultCol) - 1) * 7) + 1), "00")
          ElseIf optDaily.Value = True Then
            PlanFor = ePlanFor.mDay  'Day
            Period = Format(dtpPlanMonth, "yyyy-MM") & "-" & Format((cCnt - DefaultCol), "00")
          End If
          
          Sql = Sql & " @PlanFor = '" & Val(PlanFor) & "',"
          Sql = Sql & " @PlanPeriod = '" & Format(Period, "yyyy-MM-dd") & "',"
          
          sprPlan.GetText ePlan.mCustCode, rCnt, retVal
          Sql = Sql & " @CustomerCode  = " & Trim$(retVal) & ","
        
          sprPlan.GetText ePlan.mItem, rCnt, retVal
          Sql = Sql & " @ItemCode  = '" & Trim$(retVal) & "',"
         
          sprPlan.GetText cCnt, rCnt, retVal  ' Plan Qty
          Sql = Sql & " @PlanQty = " & Val(retVal) & ","
        
          Sql = Sql & " @RevisionNo = " & fRevNo & ","
        
'          .GetText cCnt, rCnt, retVal ' Plan Qty
'          If Val(retVal) = 0 Then
'            Sql = Sql & " @EffStatus = '0'," ' if Plan Qty = 0
'          Else
          Sql = Sql & " @EffStatus = '1'," ' if Plan Qty > 0
'          End If
          Sql = Sql & " @EffDate = '" & Format(CurDate, "yyyy-MM-dd") & "',"
          Sql = Sql & " @CreatedDate = '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
          Sql = Sql & " @UserID = '" & gUserID & "',"
          Sql = Sql & " @BranchID = '" & gBranchID & "',"
          Sql = Sql & " @CompID = '" & gCompID & "',"
          Sql = Sql & " @Finyear = '" & gFinyear & "'"
          If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
'        End If
      Next cCnt
    Next rCnt

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateData() As Boolean
  Dim rsValidate As AceADODB.Recordset
  Dim Sql As String
  Dim ItemCode As String
  Dim PlanFor As Integer
  Dim rCnt As Long
  Dim cCnt As Long
  Dim PQty As Double
  Dim retVal As Variant
  Dim CurrDate As Date
  
  Try
  ValidateData = False
  
  CurrDate = ServerDate
  
  If optyearly.Value = True Then
    PlanFor = ePlanFor.mYear ' "Year"
  ElseIf optMonthly.Value = True Then
    PlanFor = ePlanFor.mMonth '"Month"
  ElseIf optWeekly.Value = True Then
    PlanFor = ePlanFor.mWeek '"Week"
  ElseIf optDaily.Value = True Then
    PlanFor = ePlanFor.mDay '"Day"
  End If
  
  'Validate the Plan is Exists/Not
  If SalesPlan = "Create" Then
    Sql = " Select DocType, (DocType + Convert(varchar, DocNo)) Document, PlanFor, PlanPeriod"
    Sql = Sql & " From dbo.fn_Sal_SalesPlan ('" & Format(Now, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
    Set rsValidate = New AceADODB.Recordset
    If RSOpen(rsValidate, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rsValidate.EOF Then
      For rCnt = 1 To rsValidate.RecordCount
        If rsValidate("PlanFor") = PlanFor And rsValidate("DocType") = Trim$(CmbPlanType.Text) And Format(rsValidate("Planperiod"), dtpPlanMonth.CustomFormat) = Format(dtpPlanMonth, dtpPlanMonth.CustomFormat) Then
          sShowMessage "Plan Already Exists"
          GoTo ExitHere
        End If
        rsValidate.MoveNext
      Next rCnt
    End If
  End If
  
    If sprPlan.DataRowCnt <= 1 Then ' Row 1 is hotCoded For Holiday
      sShowMessage "No Data Found."
      GoTo ExitHere
    End If
    
    If sprPlan.DataRowCnt > 1 Then
      For rCnt = 2 To sprPlan.DataRowCnt ' Row 1 is hotCoded For Holiday
        sprPlan.GetText ePlan.mPlanQty, rCnt, retVal
        If retVal > 0 Then
          sprPlan.GetText ePlan.mCustName, rCnt, retVal
          If retVal = "" Then
            sShowMessage "Enter CustomerName"
            GoTo ExitHere
          End If
          sprPlan.GetText ePlan.mItem, rCnt, retVal
          If retVal = "" Then
            sShowMessage "Enter ItemCode"
            GoTo ExitHere
          End If
        End If
        sprPlan.GetText ePlan.mItem, rCnt, retVal
        ItemCode = retVal
        sprPlan.GetText ePlan.mPlanQty, rCnt, retVal
        If ItemCode <> "" And retVal < 0 Then
           sShowMessage "Enter Plan qty"
           GoTo ExitHere
        End If
      Next rCnt
    End If
    
   
'    If optyearly.Value = True Then
'      For rCnt = 2 To .DataRowCnt
'        .GetText ePlan.mPlanQty, rCnt, retVal
'        PQty = retVal
'        If PQty = 0 Then
'          sShowMessage "Please enter the Plan Qty"
'         .SetFocus
'          Spread_Set_Focus sprPlan, rCnt, ePlan.mPlanQty
'          GoTo ExitHere
'        End If
'      Next rCnt
'    Else
'      For rCnt = 2 To .DataRowCnt
'        For cCnt = DefaultCol + 1 To .MaxCols
'          .GetText cCnt, rCnt, retVal
'          PQty = retVal
'          If PQty = 0 Then
'            sShowMessage "Please enter the Plan Qty"
'            .SetFocus
'            Spread_Set_Focus sprPlan, rCnt, cCnt
'            GoTo ExitHere
'          End If
'        Next cCnt
'      Next rCnt
'    End If
   
 
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsValidate = Nothing
End Try
End Function

Private Sub CmbPlanType_Click()

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
  Dim Pertaining As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If optyearly.Value = True Then
    Pertaining = "YP"
  ElseIf optMonthly.Value = True Then
    Pertaining = "MP"
  ElseIf optWeekly.Value = True Then
    Pertaining = "WP"
  Else
    Pertaining = "DP"
  End If
    
  txtPlanNo = GetDocNo("SAL_SalesPlan", "DocType", "DocNo", Trim$(CmbPlanType), Trim$(Pertaining))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  Call Form_Clear(Me)
  Call Spread_Clear(sprPlan)
  
  sprPlan.MaxRows = 2
  sprPlan.MaxCols = DefaultCol
  
  Call Spread_Clear(sprItem)
  sprItem.MaxRows = 1
  
  ' Spread ExistType Setting
    sprType.Enabled = True
  
    sprType.Col = 1
    sprType.Row = 1
    sprType.TypeComboBoxCurSel = -1
    
    sprType.Col = 2
    sprType.Row = 1
    sprType.TypeComboBoxCurSel = -1
    
    sprType.TypeComboBoxClear 1, 1
    sprType.TypeComboBoxClear 1, 2
    
    sprType.TypeComboBoxClear 2, 1
    sprType.TypeComboBoxClear 2, 2
  
  CurDate = ServerDate
  
  dtpPlanMonth.Value = CurDate
  dtpPlanDate.Value = CurDate
  
  dtpPlanMonth.Value = Format(CurDate, "dd-MM-yyyy")
  dtpPlanMonth.CustomFormat = "yyyy"
  optyearly.Value = True
  chkCustItem.Value = 1
 
  fraDetails.Enabled = True
  cmdOK.Enabled = True
  sprPlan.Enabled = False
   
  If SalesPlan = "Create" Then
    Call TypeLoad ' DocType Load
'    Call LoadDocType(mSalPlanCreate)
'    Call SetDocNo
    Call ExistTypeLoad
  ElseIf SalesPlan = "Change" Then
    CmbPlanType.Text = ""
    optyearly.Value = False
    optMonthly.Value = False
    optWeekly.Value = False
    optDaily.Value = False
    lblPlanNo.ForeColor = &HFF0000
  End If
  cmdSetZero.Enabled = False
  Call ItemNatureLoad
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
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rsValidate As AceADODB.Recordset
  
  Dim retVal As Variant
  Dim TotQty As Variant
  Dim TotValue As Variant
  
  Dim isCreate As Integer
  Dim rCnt As Long
  Dim Rate As Double
  
  Dim CurItem As String
  Dim PlanFor As ePlanFor  ' For Current Plan
  Dim FinStart As Date
  Dim CurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  FinStart = "01-" & Left$(gFinyear, 7)
  CurrDate = Format(ServerDate, "yyyy-MM-dd")
    
  If optyearly.Value = True Then
   PlanFor = ePlanFor.mYear ' "Year"
  ElseIf optMonthly.Value = True Then
    PlanFor = ePlanFor.mMonth '"Month"
  ElseIf optWeekly.Value = True Then
    PlanFor = ePlanFor.mWeek '"Week"
  ElseIf optDaily.Value = True Then
    PlanFor = ePlanFor.mDay '"Day"
  End If
  
  If PlanFor <> ePlanFor.mYear Then 'Not Equal to year For Lock TotPlanQty
      sprPlan.BlockMode = True
      sprPlan.Col = ePlan.mPlanQty
sprPlan.Col2 = ePlan.mPlanQty

      sprPlan.Row = 2
sprPlan.Row2 = -1
      sprPlan.Lock = True
      sprPlan.BlockMode = False  End If
  
  'Validate the Plan is Exists/Not
  If SalesPlan = "Create" Then
    Sql = " Select DocType, (DocType + Convert(varchar, DocNo)) Document, PlanFor, PlanPeriod"
    Sql = Sql & " From dbo.fn_Sal_SalesPlan ('" & Format(Now, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
    Set rsValidate = New AceADODB.Recordset
    If RSOpen(rsValidate, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rsValidate.EOF Then
      For rCnt = 1 To rsValidate.RecordCount
        If rsValidate("PlanFor") = PlanFor And rsValidate("DocType") = Trim$(CmbPlanType.Text) And (Format(rsValidate("Planperiod"), dtpPlanMonth.CustomFormat) & "-" & rsValidate("Planfor")) = (Format(dtpPlanMonth, dtpPlanMonth.CustomFormat) & "-" & rsValidate("PlanFor")) Then
          sShowMessage "Plan Already Exists"
          GoTo ExitHere
        End If
        rsValidate.MoveNext
      Next rCnt
    End If
  End If
  
  sprType.GetText 2, 1, retVal 'retVal is DocType
  If SalesPlan = "Create" And PlanFor <> ePlanFor.mYear And retVal <> "N" Then
    isCreate = 1
    'Retrive Type of Plan need to display
    sprType.GetText 2, 1, retVal 'retVal is DocType
    Sql = " Execute SAL_Get_SalesPlan '" & (isCreate) & "', '" & Trim$(retVal) & "', " & PlanFor & ",  '" & Format(dtpPlanMonth, "yyyy-MM-dd") & "', '" & Format(FinStart, "yyyy-MM-dd") & "',  '" & Format(Now, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    If rs.RecordCount = 0 Then
      MsgBox "No Records Exist"
    '  GoTo ExitHere
    End If
  
    If Not rs.EOF Then
        sprPlan.Row = 1
        sprPlan.RowHidden = True
        sprPlan.MaxRows = rs.RecordCount + 1
        CurItem = ""
        For rCnt = 2 To rs.RecordCount + 1
          If rs.EOF Then Exit For
          If CurItem <> rs("ItemCode") Then
            If rCnt > 2 Then
              If PlanFor = 1 Then
                sprPlan.GetText ePlan.mPlanQty, rCnt - 1, retVal
                TotQty = Val(retVal)
                TotValue = Val(TotQty) * Val(Rate)
                sprPlan.SetText ePlan.mPlanValue, rCnt - 1, Val(TotValue)
              Else
                TotQty = Spread_Row_Total(sprPlan, rCnt - 1, DefaultCol + 1, sprPlan.MaxCols)
                sprPlan.SetText ePlan.mPlanQty, rCnt - 1, Val(TotQty)
                TotValue = Val(TotQty) * Val(Rate)
                sprPlan.SetText ePlan.mPlanValue, rCnt - 1, Val(TotValue)
              End If
            End If
            sprPlan.SetText ePlan.mCustCode, rCnt, rs("CustomerCode")
            sprPlan.SetText ePlan.mCustName, rCnt, rs("CustomerName")
            sprPlan.SetText ePlan.mItem, rCnt, rs("ItemCode")
            sprPlan.SetText ePlan.mPartNo, rCnt, rs("Partno")
            sprPlan.SetText ePlan.mDesc, rCnt, rs("ItemDescription")
            sprPlan.SetText ePlan.mUOM, rCnt, rs("UomDesc")
            sprPlan.SetText ePlan.mRate, rCnt, rs("Rate")
             Rate = rs("Rate")
            CurItem = rs("ItemCode")
          End If
          
          TotQty = 0
TotValue = 0
          Do While CurItem = rs("ItemCode")
            Call PlanQtyFill(rCnt, Val(rs("PlanFor")), PlanFor, rs("PlanPeriod"), Val(rs("PlanQty")), Val(rs("UOMDecimal"))) 'To Fill PlanQty and PlanValue in Spread
            rs.MoveNext
            If rs.EOF Then
              If PlanFor = 1 Then
                sprPlan.GetText ePlan.mPlanQty, rCnt, retVal
                TotQty = Val(retVal)
                 TotValue = Val(TotQty) * Val(Rate)
                sprPlan.SetText ePlan.mPlanValue, rCnt, Val(TotValue)
              Else
                TotQty = Spread_Row_Total(sprPlan, rCnt, DefaultCol + 1, sprPlan.MaxCols)
                sprPlan.SetText ePlan.mPlanQty, rCnt, Val(TotQty)
                TotValue = Val(TotQty) * Val(Rate)
                sprPlan.SetText ePlan.mPlanValue, rCnt, Val(TotValue)
              End If
              Exit Do
            End If
          Loop
        Next rCnt
        sprPlan.MaxRows = sprPlan.DataRowCnt + 1
      fraDetails.Enabled = False
      cmdOK.Enabled = False
      sprPlan.Enabled = True
'    Else
'     fraDetails.Enabled = False
'     cmdOk.Enabled = False
'     sprPlan.Enabled = True
    End If
  Else
    fraDetails.Enabled = False
    cmdOK.Enabled = False
    sprPlan.Enabled = True
  End If
  cmdSetZero.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set rs = Nothing
End Try
End Sub

Private Sub sprPlan_Click(sender as object, e as ClickEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

    If sprPlan.ActiveRow >= 1 And sprPlan.ActiveCol >= 10 Then
      cmdSetZero.Enabled = True
    Else
      cmdSetZero.Enabled = False
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSetZero_Click()
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
    If sprPlan.ActiveRow >= 1 And sprPlan.ActiveCol >= 10 Then
    
      If MsgBox("Do you want to set Zero Value for all Records ?", vbInformation + vbYesNo, App.Title) = vbYes Then
        For i = 1 To sprPlan.DataRowCnt
          sprPlan.SetText sprPlan.ActiveCol, i, "0"
        Next i
      End If

    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim CurDate As Date
  Dim revNo As Integer
  Dim PlanFor As String
  
  Try
  Me.MousePointer = vbHourglass
   
  If ValidateData = False Then
    GoTo ExitHere
  End If
 
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
  'To generate RevNo
  If SalesPlan = "Create" Then 'Revno = 0
    revNo = GetMax("SAL_SalesPlan", "RevisionNo", "DocType", CmbPlanType.Text, "DocNo", txtPlanNo.Text, "DocDate", Format(dtpPlanDate.Value, "yyyy-MM-dd"))
  Else 'When Change RevNo is increased
    revNo = GetMax("SAL_SalesPlan", "RevisionNo", "DocType", TxtPlanType.Text, "DocNo", txtPlanNo.Text, "DocDate", Format(dtpPlanDate.Value, "yyyy-MM-dd")) + 1
  End If
  
  If SaveSalesPlanDetail(revNo, CurDate) = False Then
    GoTo ExitHere
  End If
 
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
    
  Select Case (InStr(1, (optyearly & optMonthly & optWeekly & optDaily), "True", vbTextCompare) \ 5) + 1
    Case mYear
      PlanFor = "Yearly"
    Case mMonth
      PlanFor = "Monthly"
    Case mWeek
      PlanFor = "Weekly"
    Case mDay
      PlanFor = "Daily"
    Case Else
      'No data to Fill
  End Select
  
  sShowMessage "Sales Plan :" & Trim$(txtPlanNo.Text) & " Is Generated For " & Trim$(PlanFor)
  Call cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpPlanMonth_Change()
  Call ExistTypeLoad
  Call SetDocNo
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try

  CurDate = ServerDate
  
  dtpPlanMonth.Value = CurDate
  dtpPlanDate.Value = CurDate
   
  dtpPlanMonth.Value = Format(CurDate, "dd-MM-yyyy")
  dtpPlanMonth.CustomFormat = "yyyy"
  
  sprPlan.MaxCols = DefaultCol
  Call SpreadHeadColor(sprPlan, "Customer Name", vbBlue, ePlan.mCustName, 0)
  Call SpreadHeadColor(sprPlan, "Item Code", vbBlue, ePlan.mItem, 0)
  
  If SalesPlan = "Create" Then
'    Call TypeLoad ' DocType Load
'    Call SetDocNo 'DocNo Generate
    Call ExistTypeLoad 'Existing DocType Loading
    lblPlanNo.ForeColor = &HC0&
  ElseIf SalesPlan = "Change" Then
    lblPlanNo.ForeColor = &HFF0000
    txtPlanNo.Enabled = True
    lblPlanType.ForeColor = &HC0&
    TxtPlanType.Visible = True
    fraPeriod.Enabled = False
    optyearly.Value = False
    sprType.Enabled = False
    dtpPlanMonth.Enabled = False
    cmdOK.Visible = False
  End If
 
  sprPlan.Enabled = False
  
  Call ItemNatureLoad ' LoadItemNature
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

Private Sub optDaily_Click()
  Dim PlanFor As Integer
  
' On Error Resume Next
  
  If optDaily.Value = True Then
    lblyearly.Caption = " Daily"
    dtpPlanMonth.CustomFormat = "MMM-yyyy"
    If SalesPlan = "Create" Then
      Call TypeLoad ' DocType Load
'      Call SetDocNo
      Call ExistTypeLoad ' Existing PlanType Load
    End If
    sprPlan.MaxCols = DefaultCol
    PlanFor = ePlanFor.mDay 'Day
    Call RunTimeSpreadDesign(PlanFor) 'Daily
  End If
End Sub

Private Sub optMonthly_Click()
  Dim PlanFor As Integer
  
' On Error Resume Next
  
  If optMonthly.Value = True Then
    lblyearly.Caption = " Monthly"
    dtpPlanMonth.CustomFormat = "MMM-yyyy"
    If SalesPlan = "Create" Then
      Call TypeLoad ' DocType Load
      'Call SetDocNo
      Call ExistTypeLoad ' Existing PlanType Load
    End If
    sprPlan.MaxCols = DefaultCol
    PlanFor = ePlanFor.mMonth ' Month
    Call RunTimeSpreadDesign(PlanFor) ' Month
  End If
End Sub

Private Sub optWeekly_Click()
  Dim PlanFor As Integer
  
' On Error Resume Next
  
  If optWeekly.Value = True Then
    lblyearly.Caption = " Weekly"
    dtpPlanMonth.CustomFormat = "MMM-yyyy"
    If SalesPlan = "Create" Then
      Call TypeLoad ' DocType Load
'      Call SetDocNo
      Call ExistTypeLoad ' Existing PlanType Load
    End If
    sprPlan.MaxCols = DefaultCol
    PlanFor = ePlanFor.mWeek 'Week
    Call RunTimeSpreadDesign(PlanFor) 'Weekly
  End If
End Sub

Private Sub optyearly_Click()
  Dim PlanFor As Integer
  
' On Error Resume Next
  
  If optyearly.Value = True Then
    lblyearly.Caption = "Yearly"
    dtpPlanMonth.CustomFormat = "yyyy"
    If SalesPlan = "Create" Then
      Call TypeLoad ' DocType Load
'      Call SetDocNo
      Call ExistTypeLoad ' Existing PlanType Load
    End If
    sprPlan.MaxCols = DefaultCol
    PlanFor = ePlanFor.mYear  'Year
    Call RunTimeSpreadDesign(PlanFor) 'Yearly
  End If
End Sub

Private Sub sprPlan_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  
  Dim rCnt As Long
  Dim PrevRowCnt As Long
  Dim sRow As Long
  Dim fRow As Long
  
  Dim retVal As Variant
  Dim Custcode As String
  Dim CustName As String
  
  Dim FinStart As Date
  Dim CurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If Col = ePlan.mCustName Then ' Customer Name Help
    Sql = " Select PartyName, PartyCode From FAS_Party_Master"
    Sql = Sql & " Where AcType in ('D', 'Y') and BranchID = '" & gBranchID & "'"
    Sql = Sql & " and CompID = '" & gCompID & "' and EffStatus = '1'"
    If FetchValue_New(False, db.ConnectionString, Sql, "1", "List of Customers.", "Customer Name") = True Then
        sprPlan.SetText 1, sprPlan.ActiveRow, Val(getvalue(2))
        sprPlan.SetText 2, sprPlan.ActiveRow, Trim$(getvalue(1))    End If
  End If
  
  If Col = ePlan.mItem Then 'ItemCode Help
    FinStart = "01-" & Left$(gFinyear, 7)
    CurrDate = Format(ServerDate, "yyyy-MM-dd")
    PrevRowCnt = sprPlan.DataRowCnt
    sprPlan.GetText 1, Row, retVal
    If Trim$(retVal) = "" Then
      sShowMessage "Select Customer Name"
      Call Spread_Set_Focus(sprPlan, Row, 2)
      GoTo ExitHere
    Else
      If chkCustItem.Value = 1 Then
        Sql = " Select ItemCode, PartNo, ItemDescription, UomDesc, Rate, NoofDecimal "
        Sql = Sql & " From (Select ItemCode, PartNo, ItemDescription, UomDesc, "
        Sql = Sql & " dbo.fn_Sal_Item_Rate (ItemCode, CustomerCode, '', 0, Null, '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')Rate, NoofDecimal"
        Sql = Sql & " From(Select a.ItemCode, b.ItemDescription, b.UomDesc, "
        Sql = Sql & " (Select Top 1 PartNo From PartMast Where StageCode = a.ItemCode and Branchid = a.Branchid and Compid = a.Compid)PartNo,"
        Sql = Sql & " a.CustomerCode, b.UomDecimal NoofDecimal  "
        Sql = Sql & " From SAL_Customer_ItemMaster a"
        Sql = Sql & " Left Outer  Join Material_Master b"
        Sql = Sql & " On b.ItemCode = a.ItemCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
        Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and b.Status = 1"
        sprPlan.GetText ePlan.mCustCode, Row, retVal
        Sql = Sql & " and a.CustomerCode = '" & retVal & "')x)y"
        Sql = Sql & " Where Rate > 0"
      Else
        Sql = " Select a.ItemCode, "
        Sql = Sql & " (Select top 1 PartNo From PartMast Where StageCode = a.ItemCode and Branchid = a.Branchid and Compid = a.Compid)PartNo,"
        Sql = Sql & " a.ItemDescription, a.UomDesc, "
        Sql = Sql & " dbo.fn_WtAvgRate (a.ItemCode, '" & Format(FinStart, "yyyy-MM-dd") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')Rate, a.UomDecimal NoofDecimal"
        Sql = Sql & " From MATERIAL_MASTER a"
        Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and a.Status = 1"
        Sql = Sql & " and Rate > 0"
        If sprItem.DataRowCnt > 0 Then
          If sprItem.SelectionCount > 0 Then 'sShowMessage "Please select atlease one item": Exit Sub
            Sql = Sql & " and a.ItemNature in("
              sRow = 0
              For rCnt = 1 To sprItem.SelectionCount
                sRow = sprItem.GetMultiSelItem(sRow)
                sprItem.GetText 1, sRow, retVal
                Sql = Sql & " '" & retVal & "',"
              Next rCnt
              Sql = Left$(Sql, Len(Sql) - 1)
              Sql = Sql & ")"          End If
        End If
      End If
      
      rCnt = 0
      If FetchValue_New(True, db.ConnectionString, Sql, "11101", "List of Items", "ItemCode", "Part No", "Description", "Rate") = True Then
           sprPlan.MaxRows = sprPlan.DataRowCnt + UBound(getvalue, 1)
          For rCnt = 1 To UBound(getvalue, 1)
            fRow = Row + (rCnt - 1)
            If rCnt = 1 Then
              sprPlan.GetText ePlan.mCustCode, fRow, retVal
              Custcode = retVal
              sprPlan.GetText ePlan.mCustName, fRow, retVal
              CustName = retVal
            Else
              sprPlan.SetText ePlan.mCustCode, fRow, Custcode
              sprPlan.SetText ePlan.mCustName, fRow, Trim$(CustName)
            End If
            
            sprPlan.SetText ePlan.mItem, fRow, Trim$(getvalue(rCnt, 1))
            sprPlan.SetText ePlan.mPartNo, fRow, Trim$(getvalue(rCnt, 2))
            If Spread_Clear_Duplicate(sprPlan, sprPlan.DataRowCnt, ePlan.mCustCode, ePlan.mItem) = False Then
              sprPlan.SetText ePlan.mDesc, fRow, Trim$(getvalue(rCnt, 3))
              sprPlan.SetText ePlan.mUOM, fRow, getvalue(rCnt, 4)
              sprPlan.SetText ePlan.mRate, fRow, Val(getvalue(rCnt, 5))
              sprPlan.GetText ePlan.mPlanQty, fRow, retVal ' Qty
              If retVal > 0 Then
                sprPlan.SetText ePlan.mPlanValue, fRow, (retVal * Val(getvalue(rCnt, 5))) ' Qty * Rate
              End If
            End If
            
          Next rCnt
'        Call Spread_Lock(sprPlan, False, ePlan.mTotQty 3, 6, PrevRowCnt + 1, sprPlan.DataRowCnt)
          sprPlan.MaxRows = sprPlan.DataRowCnt + 1
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprPlan_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim TotPlanValue As Double
  Dim TotPlanQty As Double
  Dim cCnt As Long
  Dim ColStart As Long
  Dim ColEnd As Long
  
  Try

  If Mode = 1 Then
 Exit Sub
End If
  
    If optyearly.Value <> True Then
      If sprPlan.DataRowCnt > 1 Then
        ColStart = DefaultCol + 1
        ColEnd = sprPlan.MaxCols
        TotPlanQty = 0
        For cCnt = ColStart To ColEnd
          sprPlan.GetText cCnt, Row, tmp
          TotPlanQty = TotPlanQty + Val(tmp)
          sprPlan.SetText ePlan.mPlanQty, Row, TotPlanQty
        Next cCnt
      End If
    End If
    
    sprPlan.GetText ePlan.mRate, Row, tmp    'Rate
    sprPlan.GetText ePlan.mPlanQty, Row, tmp1   'Plan Qty
    If tmp1 <> "" Then
      sprPlan.SetText ePlan.mPlanValue, Row, Format(Val(tmp) * Val(tmp1), "0.00")
    End If
    TotPlanValue = TotPlanValue + Val(tmp) * Val(tmp1)
    
    'For New Entry
    sprPlan.GetText 1, Row, tmp
    If tmp <> "" Then
      sprPlan.GetText ePlan.mRate, Row, tmp 'Rate
      sprPlan.GetText ePlan.mPlanQty, Row, tmp1  'Tot Plan Qty
      sprPlan.SetText ePlan.mPlanValue, Row, Format(Val(tmp) * Val(tmp1), "0.00")
      TotPlanValue = TotPlanValue + Val(tmp) * Val(tmp1)
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprPlan_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim retVal As Variant

  Try
  Me.MousePointer = vbHourglass

    If sprPlan.ActiveRow > 2 And sprPlan.ActiveRow = sprPlan.DataRowCnt + 1 Then
      If KeyCode = vbKeyUp And Shift = vbCtrlMask Then
        sprPlan.GetText 1, sprPlan.ActiveRow - 1, retVal
        sprPlan.SetText 1, sprPlan.ActiveRow, retVal
        sprPlan.GetText 2, sprPlan.ActiveRow - 1, retVal
        sprPlan.SetText 2, sprPlan.ActiveRow, retVal
      End If
    End If
    If sprPlan.EditMode = False Then
      If KeyCode = vbKeyDelete Then
        If Shift = vbShiftMask Then
          sprPlan.DeleteRows sprPlan.ActiveRow, 1
          sprPlan.MaxRows = sprPlan.DataRowCnt + 1
        Else
          If optyearly.Value = True Then
            If sprPlan.ActiveCol = ePlan.mPlanQty Then
              sprPlan.SetText ePlan.mPlanQty, sprPlan.ActiveRow, ""
            End If
          Else
            If sprPlan.ActiveCol > DefaultCol Then
              sprPlan.SetText sprPlan.ActiveCol, sprPlan.ActiveRow, ""
              sprPlan_EditMode sprPlan.ActiveCol, sprPlan.ActiveRow, 0, False
            End If
          End If
        End If
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprPlan_KeyUp(sender as object, e as KeyUpEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  Call sprPlan_Click(sprPlan.ActiveCol, sprPlan.ActiveRow)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprType_ComboCloseUp(sender as object, e as ComboCloseUpEvent)
  Dim retVal As Variant
 
  Try
  Me.MousePointer = vbHourglass

     sprType.Col = 1
sprType.Row = Row
    retVal = sprType.TypeComboBoxCurSel
 
    sprType.Col = 2
sprType.Row = Row
     sprType.TypeComboBoxCurSel = Val(retVal)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub
Private Sub txtPlanNo_DblClick()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Long
  Dim TotQty As Double
  Dim TotValue As Double
  Dim Rate As Double
  
  Dim retVal As Variant
  
  Dim PlanFor As ePlanFor
  Dim CurItem As String
  
  Dim FinStart As Date
  Dim CurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If SalesPlan = "Change" Then
    FinStart = "01-" & Left$(gFinyear, 7)
    CurrDate = ServerDate
    Sql = "Select Distinct a.DocType, a.DocNo, a.DocDate,Revisionno"
    Sql = Sql & " From dbo.fn_Sal_SalesPlan('" & Format(CurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and a.Finyear = '" & gFinyear & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "11110001", "List of Documents.", "Doc Type", "Doc No", "Doc Date", "Revisionno", "Customer Name") = True Then
      TxtPlanType.Text = Trim$(getvalue(1))
      txtPlanNo.Text = getvalue(2)
      dtpPlanDate.Value = Format(getvalue(3), "dd-MM-yyyy")
       
      Sql = " Select Distinct a.PlanFor, a.PlanPeriod, a.CustomerCode, (Select PartyName From FAS_Party_Master "
      Sql = Sql & " Where PartyCode = a.CustomerCode and BranchID = a.BranchID and CompID = a.CompID)CustomerName, "
      Sql = Sql & " a.ItemCode, b.ItemDescription, b.UomDesc, b.UomDecimal,"
      Sql = Sql & " (Select top 1 PartNo From PartMast Where StageCode = a.ItemCode and Branchid = a.Branchid and Compid = a.Compid)PartNo,"
      'Sql = Sql & " (Select Rate From SAL_Customer_ItemMaster Where CustomerCode = a.CustomerCode and ItemCode = a.ItemCode"
      'Sql = Sql & " and BranchID = a.BranchID and CompID = a.CompID) Rate, PlanQty"
      Sql = Sql & " dbo.fn_Sal_Item_Rate (a.ItemCode, a.CustomerCode, '', 0, Null, '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') Rate, a.PlanQty"
      Sql = Sql & " From dbo.fn_Sal_SalesPlan('" & Format(CurrDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
      Sql = Sql & " Left Outer Join Material_Master b"
      Sql = Sql & " On b.ItemCode = a.ItemCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
      Sql = Sql & " Where a.BranchID = '" & gBranchID & "' and a.CompID = '" & gCompID & "' and a.Finyear = '" & gFinyear & "' and EffStatus = '1'"
      Sql = Sql & " and a.DocType = '" & Trim$(TxtPlanType.Text) & "' and a.DocNo = " & txtPlanNo & " and b.Status = 1 "
      Sql = Sql & " Order By a.CustomerCode, a.ItemCode"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
        GoTo ExitHere
      End If
      
      If Not rs.EOF Then
          sprPlan.Row = 1
          sprPlan.RowHidden = True
          sprPlan.MaxRows = rs.RecordCount + 1
          CurItem = ""
          For rCnt = 2 To rs.RecordCount + 1
            If rs.EOF Then Exit For
            If CurItem <> rs("ItemCode") Then
              If rCnt > 2 Then
                If PlanFor = 1 Then
                  sprPlan.GetText ePlan.mPlanQty, rCnt - 1, retVal
                  TotQty = Val(retVal)
                  TotValue = Val(TotQty) * Val(Rate)
                  sprPlan.SetText ePlan.mPlanValue, rCnt - 1, Val(TotValue)
                Else
                  TotQty = Spread_Row_Total(sprPlan, rCnt - 1, DefaultCol + 1, sprPlan.MaxCols)
                  sprPlan.SetText ePlan.mPlanQty, rCnt - 1, Val(TotQty)
                  TotValue = Val(TotQty) * Val(Rate)
                  sprPlan.SetText ePlan.mPlanValue, rCnt - 1, Val(TotValue)
                End If
              End If
              sprPlan.SetText ePlan.mCustCode, rCnt, rs("CustomerCode")
              sprPlan.SetText ePlan.mCustName, rCnt, rs("CustomerName")
              sprPlan.SetText ePlan.mItem, rCnt, rs("ItemCode")
              sprPlan.SetText ePlan.mPartNo, rCnt, rs("partno")
              sprPlan.SetText ePlan.mDesc, rCnt, rs("ItemDescription")
              sprPlan.SetText ePlan.mUOM, rCnt, rs("UomDesc")
              sprPlan.SetText ePlan.mRate, rCnt, rs("Rate")
              Select Case rs("PlanFor")
                Case 1
                  dtpPlanMonth.Value = rs("PlanPeriod")
                  dtpPlanMonth.CustomFormat = "yyyy"
                  PlanFor = ePlanFor.mYear
                  optyearly.Value = True
                Case 2
                  dtpPlanMonth.Value = rs("PlanPeriod")
                  dtpPlanMonth.CustomFormat = "MMM-yyyy"
                  PlanFor = ePlanFor.mMonth
                  optMonthly.Value = True
                Case 3
                  dtpPlanMonth.Value = rs("PlanPeriod")
                  dtpPlanMonth.CustomFormat = "MMM-yyyy"
                  PlanFor = ePlanFor.mWeek
                  optWeekly.Value = True
                Case 4
                  dtpPlanMonth.Value = rs("PlanPeriod")
                  dtpPlanMonth.CustomFormat = "MMM-yyyy"
                  PlanFor = ePlanFor.mDay
                  optDaily.Value = True
              End Select
              Rate = rs("Rate")
              CurItem = rs("ItemCode")
            End If
            
            If rs("PlanFor") <> ePlanFor.mYear Then 'Not Equal to year
              sprPlan.BlockMode = True
              sprPlan.Col = ePlan.mPlanQty
sprPlan.Col2 = ePlan.mPlanQty

              sprPlan.Row = 2
sprPlan.Row2 = -1
              sprPlan.Lock = True
              sprPlan.BlockMode = False
            End If
          
            Do While CurItem = rs("ItemCode")
              Call PlanQtyFill(rCnt, Val(rs("PlanFor")), PlanFor, rs("PlanPeriod"), Val(rs("PlanQty")), Val(rs("UOMDecimal"))) 'To Fill PlanQty and PlanValue in Spread
              rs.MoveNext
              If rs.EOF Then
                If PlanFor = 1 Then
                  sprPlan.GetText ePlan.mPlanQty, rCnt, retVal
                  TotQty = Val(retVal)
                  TotValue = Val(TotQty) * Val(Rate)
                  sprPlan.SetText ePlan.mPlanValue, rCnt, Val(TotValue)
                Else
                  TotQty = Spread_Row_Total(sprPlan, rCnt, DefaultCol + 1, sprPlan.MaxCols)
                  sprPlan.SetText ePlan.mPlanQty, rCnt, Val(TotQty)
                  TotValue = Val(TotQty) * Val(Rate)
                  sprPlan.SetText ePlan.mPlanValue, rCnt, Val(TotValue)
                End If
                Exit Do
              End If
            Loop
          Next rCnt
          sprPlan.MaxRows = sprPlan.DataRowCnt + 1
      End If
    End If
    sprPlan.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set rs = Nothing
End Try
End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem

End Sub

Private Sub Design_sprType(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprType


sprType.ComboCloseUp +=  new EventHandler(sprType_ComboCloseUp)

End Sub

Private Sub Design_sprPlan(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprPlan


sprPlan.Click +=  new EventHandler(sprPlan_Click)
sprPlan.DblClick +=  new EventHandler(sprPlan_DblClick)
sprPlan.EditMode +=  new EventHandler(sprPlan_EditMode)
sprPlan.KeyDown +=  new EventHandler(sprPlan_KeyDown)
sprPlan.KeyUp +=  new EventHandler(sprPlan_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesPlan.frm", sprItem, "")

	Call Design_sprType("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesPlan.frm", sprType, "")

	Call Design_sprPlan("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesPlan.frm", sprPlan, "")

End Sub
