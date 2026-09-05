Attribute VB_Name = "mdlSpreadReport"
Option Explicit

Public Enum eSprAlignConst
  'Value Mod 3 i.e Reminder  Gives the Vertical Alignment
  'Value \ 3 i.e Quotient Gives the Horizontal Alignment
  '0 - Top, Left; 1 - Bottom, Right; 2 - Center, Center
  
  mAlignleftTop = 0
  mAlignLeftBottom = 1
  mAlignLeftCenter = 2
  
  mAlignRightTop = 3
  mAlignRightBottom = 4
  mAlignRightCenter = 5
  
  mAlignCenterTop = 6
  mAlignCenterBottom = 7
  mAlignCenterCenter = 8
End Enum

Public Enum eSprBorderConst
  mBorderNone = 0
  mBorderLeft = 1
  mBorderRight = 2
  mBorderTop = 4
  mBorderBottom = 8
  mBorderAllSides = 15
  mBorderOutline = 16
End Enum

Public Enum eRem
  mPageNo = 0
  mPrintDate = 1
  mERPNo = 2
  mTsNo = 3
  mCompanyname = 4
  mReportName = 5
  mHeader = 6     '1st Line
End Enum

Public Enum eRemOpt
  mColStart = 0
  mRowStart = 1
  mNoOfCols = 2
  mNoOfRows = 3
End Enum

'Declare the array for the maximum no in members of eRem Enum
' First dimension to hold type index; second dimension for Block of Type(i.e. PageNo, PrintDate etc...)
Public  SprRptRemoveRows(eRem.mHeader, eRemOpt.mNoOfRows) As Long

'/********** End of Export To Excel Options ****************/

Public Sub CopySpreadStructure_New(ByRef SourceSpread As AceSpread,ByVal  SourceSheet As Long,Optional ByRef DestnSpread As AceSpread,Optional ByVal DestnSheet As Long,Optional ByVal SetMaxCol As Boolean = True,Optional ByVal SetMaxRow As Boolean = True,Optional ByVal SetCol As Long = 1,Optional ByVal SetRow As Long = 1)
  Dim cCnt As Long
  Dim rCnt As Long
  Dim sSheet As Long
  Dim FistSetCol As Long
  Dim dSheet As Long
  Dim sSpread As AceSpread
  Dim dSpread As AceSpread
  Dim retVal As Variant
  Dim fHAlign As FPSpreadADO.TypeHAlignConstants
  Dim fVAlign As FPSpreadADO.TypeVAlignConstants
  Dim fCellType As FPSpreadADO.CellTypeConstants
  Dim PvCol As Variant
  Dim PvRow As Variant
  Dim PvNoCol As Variant
  Dim PvNoRow As Variant
  Dim SBordColr As Variant
  Dim SBordStyl As Variant
  Dim SPvNoCol As Long
  Dim SPvNoRow As Long
  
  Try
  
  Set sSpread = SourceSpread
  sSheet = SourceSheet
  
  If DestnSheet = 0 Then
    dSheet = sSheet
  Else
    dSheet = DestnSheet
  End If
  
  FistSetCol = SetCol
  
  If TypeName(DestnSpread) = "Nothing" Then     'For Optional Object Variable isMissing not works
    Set dSpread = SourceSpread
  Else
    Set dSpread = DestnSpread
  End If
  
  'Setting Max Column
  If SetMaxCol = True Then
    sSpread.Sheet = sSheet
    retVal = sSpread.MaxCols
    dSpread.Sheet = dSheet
    dSpread.MaxCols = retVal + (SetCol - 1)
  End If
  
  'Setting Max Row
  If SetMaxRow = True Then
    sSpread.Sheet = sSheet
    retVal = sSpread.MaxRows
    dSpread.Sheet = dSheet
    dSpread.MaxRows = 0
    dSpread.MaxRows = retVal + (SetRow - 1)
  End If

  For rCnt = 1 To sSpread.MaxRows
    For cCnt = 1 To sSpread.MaxCols
      dSpread.Row = SetRow
dSpread.Col = SetCol
dSpread.BlockMode = False
      'Setting Column Width for Each Column
      sSpread.Sheet = sSheet
      retVal = sSpread.ColWidth(cCnt)
      dSpread.Sheet = dSheet
      dSpread.ColWidth(SetCol) = retVal
  
      sSpread.Sheet = sSheet
      sSpread.Col = cCnt
sSpread.Row = rCnt
      retVal = sSpread.Lock
        dSpread.Sheet = dSheet
          dSpread.Col = SetCol
dSpread.Row = SetRow
dSpread.Col2 = SetCol
dSpread.Row2 = SetRow
        dSpread.Lock = retVal      
      If rCnt = 1 Then
        'Setting Column Header
'        dSpread.BlockMode = True
        sSpread.Sheet = sSheet
        sSpread.GetText cCnt, 0, retVal
        dSpread.Sheet = dSheet
        dSpread.SetText SetCol, 0, retVal
'        dSpread.BlockMode = False
      End If
      
      'Setting Cell Type for Each Column
      sSpread.Sheet = sSheet
      sSpread.Col = cCnt
sSpread.Row = rCnt
      fCellType = sSpread.CellType
      fHAlign = sSpread.TypeHAlign
      fVAlign = sSpread.TypeVAlign
        dSpread.Sheet = dSheet
          dSpread.Col = SetCol
dSpread.Row = rCnt
dSpread.Col2 = SetCol
dSpread.Row2 = rCnt
'        .BlockMode = False
        dSpread.CellType = fCellType
        Select Case fCellType
          Case CellTypeButton
              Call CopySpreadButtonStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeDate
              Call CopySpreadDateStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeEdit
              Call CopySpreadEditStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypePic
            
          Case CellTypeStaticText
            
          Case CellTypeTime
              Call CopySpreadTimeStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeComboBox
              Call CopySpreadComboStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypePicture
                dSpread.TypePictPicture = sSpread.TypePictPicture
                dSpread.TypePicMask = sSpread.TypePicMask
                dSpread.TypePictCenter = sSpread.TypePictCenter
                dSpread.TypePictMaintainScale = sSpread.TypePictMaintainScale
                dSpread.TypePictStretch = sSpread.TypePictStretch
          Case CellTypeCheckBox
              Call CopySpreadCheckStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeOwnerDrawn
    
          Case CellTypeCurrency
              Call CopySpreadCurrencyStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeNumber
              Call CopySpreadNumberStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypePercent
              Call CopySpreadPercentStructure_New(sSpread, sSheet, dSpread, dSheet, cCnt, rCnt, SetCol, SetRow)
          Case CellTypeScientific
          
        End Select
        dSpread.TypeHAlign = fHAlign
        dSpread.TypeVAlign = fVAlign
'        .BlockMode = True
      
      sSpread.Sheet = sSheet ' modify rajapriyan 12/03/2008
      retVal = sSpread.MaxCols
      sSpread.Sheet = sSheet
      sSpread.Col = cCnt
sSpread.Row = rCnt
      dSpread.Sheet = dSheet
      dSpread.RowHeight(SetRow) = sSpread.RowHeight(rCnt)
      'Setting Column WITH
      
      sSpread.Sheet = sSheet
      sSpread.Row = rCnt
sSpread.Col = cCnt
      dSpread.Sheet = dSheet
      dSpread.ColWidth(SetCol) = sSpread.ColWidth(cCnt)
      'Setting cell Font
      
'      dSpread.BlockMode = False
      sSpread.Sheet = sSheet
      sSpread.Row = rCnt
sSpread.Col = cCnt
      dSpread.Sheet = dSheet
      dSpread.Row = SetRow
dSpread.Col = SetCol

      dSpread.Lock = False
sSpread.Lock = False
      dSpread.Font.Bold = sSpread.Font.Bold
      dSpread.Font.Charset = sSpread.Font.Charset
      dSpread.Font.Italic = sSpread.Font.Italic
      dSpread.Font.Name = sSpread.Font.Name
      dSpread.Font.Size = sSpread.Font.Size
      dSpread.Font.Strikethrough = sSpread.Font.Strikethrough
      dSpread.Font.Underline = sSpread.Font.Underline
      dSpread.Font.Weight = sSpread.Font.Weight
      dSpread.FontBold = sSpread.FontBold
      dSpread.FontItalic = sSpread.FontItalic
      dSpread.ForeColor = sSpread.ForeColor
      dSpread.BackColor = sSpread.BackColor
      dSpread.BackColorStyle = sSpread.BackColorStyle
      dSpread.BlockMode = sSpread.BlockMode
      dSpread.BorderStyle = sSpread.BorderStyle
      dSpread.TypeEditCharCase = sSpread.TypeEditCharCase
      dSpread.TypeEditCharSet = sSpread.TypeEditCharSet
      dSpread.TypeEditMultiLine = sSpread.TypeEditMultiLine
      dSpread.TypeEditPassword = sSpread.TypeEditPassword
      dSpread.TypeEllipses = sSpread.TypeEllipses
      dSpread.TypeMaxEditLen = sSpread.TypeMaxEditLen
        
      sSpread.GetCellBorder cCnt, rCnt, CellBorderIndexBottom, SBordColr, SBordStyl
      dSpread.SetCellBorder SetCol, SetRow, SetCol, SetRow, CellBorderIndexBottom, SBordColr, SBordStyl
      sSpread.GetCellBorder cCnt, rCnt, CellBorderIndexLeft, SBordColr, SBordStyl
      dSpread.SetCellBorder SetCol, SetRow, SetCol, SetRow, CellBorderIndexLeft, SBordColr, SBordStyl
      sSpread.GetCellBorder cCnt, rCnt, CellBorderIndexRight, SBordColr, SBordStyl
      dSpread.SetCellBorder SetCol, SetRow, SetCol, SetRow, CellBorderIndexRight, SBordColr, SBordStyl
      sSpread.GetCellBorder cCnt, rCnt, CellBorderIndexTop, SBordColr, SBordStyl
      dSpread.SetCellBorder SetCol, SetRow, SetCol, SetRow, CellBorderIndexTop, SBordColr, SBordStyl
      
      retVal = sSpread.GetCellSpan(cCnt, rCnt, PvCol, PvRow, PvNoCol, PvNoRow)
      If cCnt <= PvCol And rCnt <= PvRow Then
        dSpread.AddCellSpan SetCol, SetRow, PvNoCol, PvNoRow
      End If
      dSpread.Text = sSpread.Text
      dSpread.BlockMode = True
      SetCol = SetCol + 1
    Next cCnt
    SetCol = FistSetCol
    SetRow = SetRow + 1
  Next rCnt
  
  'Setting Column Header height
  sSpread.Sheet = sSheet
  retVal = sSpread.RowHeight(0)
  dSpread.Sheet = dSheet
  dSpread.RowHeight(0) = retVal
  dSpread.PrintBorder = True
  dSpread.BorderStyle = BorderStyleFixedSingle
  
  'Setting Row Header Width
  sSpread.Sheet = sSheet
  retVal = sSpread.ColWidth(0)
  dSpread.Sheet = dSheet
  dSpread.ColWidth(0) = retVal
  
  'Setting Column Header
  sSpread.Sheet = sSheet
  sSpread.GetText cCnt, 0, retVal
  dSpread.Sheet = dSheet
  dSpread.SetText SetCol, 0, retVal
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_ERPnoPrint(ByRef mSpread As AceSpread,ByVal  mCol As Long,ByVal  mRow As Long)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try

  Sql = "Select ERPNo,TSNo from Sec_ModuleFormList where MenuName = '" & Trim$(mnuFlg) & "'"
  Set rs = New Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  If rs.RecordCount > 0 Then
      mSpread.AddCellSpan mCol, mRow, 3, 1
      mSpread.Col = mCol
mSpread.Row = mRow

      mSpread.SetText 1, mRow, "ERP Ref. No :" & rs("ErpNo")
mSpread.TypeHAlign = TypeHAlignLeft
      
      mSpread.AddCellSpan 4, mRow, mSpread.MaxCols, 1
      mSpread.Col = 4
mSpread.Row = mRow

      mSpread.SetText 4, mRow, "TS Ref. No : " & rs("TsNo")
mSpread.TypeHAlign = TypeHAlignRight
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub
Private Sub CopySpreadButtonStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
sSpread.Row = fRow
  retVal = sSpread.TypeButtonAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonBorderColor
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonBorderColor = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonColor
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonDarkColor
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = fRow
  dSpread.TypeButtonDarkColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonLightColor
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonLightColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonPicture
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonPicture = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonPictureDown
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonPictureDown = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonShadowSize
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonShadowSize = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonText
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonText = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonTextColor
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonTextColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeButtonType
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeButtonType = retVal
End Sub

Private Sub CopySpreadCheckStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
  Dim Cnt As Integer
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCheckType
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCheckType = retVal
  
  If retVal = TypeCheckTypeThreeState Then
    For Cnt = 0 To 5
      sSpread.Sheet = sSheet
      sSpread.Col = fCol
sSpread.Row = fRow
      retVal = sSpread.TypeCheckPicture(Cnt)
      If retVal <> 0 Then
        dSpread.Sheet = dSheet
        dSpread.Col = SetCol
dSpread.Row = SetRow
        dSpread.TypeCheckPicture(Cnt) = retVal
      End If
    Next
  Else
    For Cnt = 0 To 3
      sSpread.Sheet = sSheet
      sSpread.Col = fCol
sSpread.Row = fRow
      retVal = sSpread.TypeCheckPicture(Cnt)
      If retVal <> 0 Then
        dSpread.Sheet = dSheet
        dSpread.Col = SetCol
dSpread.Row = SetRow
        dSpread.TypeCheckPicture(Cnt) = sSpread.TypeCheckPicture(Cnt)
      End If
    Next
  End If
  
  'For Type Check Center Both the HAlign & VAlign used
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeHAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeVAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeVAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCheckText
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCheckText = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCheckTextAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCheckTextAlign = retVal
End Sub
Private Sub CopySpreadComboStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxAutoSearch
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxAutoSearch = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxCount
  If Val(retVal) > 0 Then
    dSpread.Sheet = dSheet
    dSpread.Col = SetCol
dSpread.Row = SetRow
    dSpread.TypeComboBoxList = sSpread.TypeComboBoxList
  End If

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxCurSel
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxCurSel = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxEditable
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxEditable = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxIndex
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxIndex = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxList
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxList = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxMaxDrop
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxMaxDrop = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxString
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxString = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeComboBoxWidth
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeComboBoxWidth = retVal
End Sub

Private Sub CopySpreadCurrencyStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyDecPlaces = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyMax
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyMin
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyNegStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyPosStyle
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyPosStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencySeparator
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencySeparator = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyShowSep
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyShowSep = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencyShowSymbol
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencyShowSymbol = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeCurrencySymbol
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeCurrencySymbol = retVal
End Sub

Private Sub CopySpreadDateStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeDateCentury
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeDateCentury = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeDateFormat
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeDateFormat = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeDateMax
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeDateMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeDateMin
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeDateMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeDateSeparator
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeDateSeparator = retVal
End Sub

Private Sub CopySpreadEditStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEditCharCase
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEditCharCase = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEditCharSet
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEditCharSet = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEditMultiLine
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEditMultiLine = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEditPassword
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEditPassword = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEllipses
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEllipses = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeHAlign = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = SetCol
sSpread.Row = fRow
  retVal = sSpread.TypeMaxEditLen
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeMaxEditLen = retVal
End Sub

Private Sub CopySpreadNumberStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNegRed
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNegRed = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberDecPlaces = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberMax
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberMin
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberNegStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberSeparator
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberSeparator = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeNumberShowSep
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeNumberShowSep = retVal
End Sub

Private Sub CopySpreadPercentStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentDecPlaces = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentMax
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentMin
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypePercentNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypePercentNegStyle = retVal
End Sub

Private Sub CopySpreadTimeStructure_New(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  SetCol As Long,ByVal  SetRow As Long)
  Dim retVal As Variant
  Dim Cnt As Integer
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeTime24Hour
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeTime24Hour = retVal

  'For Type Time Center Both the HAlign & VAlign and Time Format used
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeHAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeVAlign
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeVAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeTimeSeconds
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeTimeSeconds = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeEllipses
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeEllipses = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
sSpread.Row = fRow
  retVal = sSpread.TypeSpin
  dSpread.Sheet = dSheet
  dSpread.Col = SetCol
dSpread.Row = SetRow
  dSpread.TypeSpin = retVal
End Sub
Public Function spSetBorder_TwoRow(ByVal fRow As Long,ByVal  sCol As Long,Optional ByVal Endborder As Boolean = False)
  If GridCount = 53 Then
    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow, mBorderAllSides)
    If Endborder Then
      Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow + 1, frmPrint.sprPrint.MaxCols, fRow + 1, mBorderAllSides)
    End If
'    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow, mBorderLeft)
'    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow, mBorderRight)
'    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow, mBorderBottom)
  Else
    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow - 1, frmPrint.sprPrint.MaxCols, fRow - 1, mBorderBottom)
    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow + 1, mBorderLeft)
    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow, frmPrint.sprPrint.MaxCols, fRow + 1, mBorderRight)
    Call spSetCellBorder(frmPrint.sprPrint, sCol, fRow + 1, frmPrint.sprPrint.MaxCols, fRow + 1, mBorderBottom)
  End If
End Function
Public Function spSetCellBorder(ByRef mSpread As AceSpread,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2,Optional ByVal mBorderType As eSprBorderConst = mBorderOutline,Optional ByVal mBorderColor As ColorConstants = 0,Optional ByVal mBorderStyle As CellBorderStyleConstants = CellBorderStyleSolid)
  Try
  
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    Call mSpread.SetCellBorder(mCol1, mRow1, mCol2, mRow2, mBorderType, mBorderColor, mBorderStyle)
    
'    If mBorderIndex = 16 Then
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexLeft, 0, CellBorderStyleSolid
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexRight, 0, CellBorderStyleSolid
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexTop, 0, CellBorderStyleSolid
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexBottom, 0, CellBorderStyleSolid
'    ElseIf mBorderIndex = 1 Then
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexLeft, 0, CellBorderStyleSolid
'    ElseIf mBorderIndex = 2 Then
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexRight, 0, CellBorderStyleSolid
'    ElseIf mBorderIndex = 4 Then
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexTop, 0, CellBorderStyleSolid
'    ElseIf mBorderIndex = 8 Then
'      .SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexBottom, 0, CellBorderStyleSolid
'    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function SpSetCellSpan(ByRef mSpread As AceSpread,ByVal  ParamArray ColXRowXNoColXNoRow() As Variant) As Boolean
  Dim Cnt As Long
  
' On Error Resume Next
  
  If IsMissing(ColXRowXNoColXNoRow) Then
    sShowMessage "Missing Arguemnts (Col,Row,NoofCol and NoofRow)."
  End If
  
  If UBound(ColXRowXNoColXNoRow) Mod 4 = 0 Then
    sShowMessage "Arguemnts (Col,Row,NoofCol and NoofRow) not Passed Properly."
  End If
  
    For Cnt = 0 To UBound(ColXRowXNoColXNoRow) - 1 Step 4
      mSpread.AddCellSpan ColXRowXNoColXNoRow(Cnt), ColXRowXNoColXNoRow(Cnt + 1), ColXRowXNoColXNoRow(Cnt + 2), ColXRowXNoColXNoRow(Cnt + 3)
    Next Cnt
  
End Function

Public Function spSetColWidth(ByRef mSpread As AceSpread,ByVal  ParamArray ColNoColWidthList() As Variant) As Boolean
  'ColNoAndColWidth should be passed like this 1, 25, 2, 30, 6, 25, 10, 45
  'ColNo_x, ColWdith_x, ColNo_y, ColWidth_y,...
  
  Dim Cnt As Long
  Dim tmpCol As Long
  
' On Error Resume Next
  
  If IsMissing(ColNoColWidthList) Then
    sShowMessage "Missing Arguemnts (ColNo and ColWidth)."
  End If
  
  If UBound(ColNoColWidthList) Mod 2 = 0 Then
    sShowMessage "Arguemnts (ColNo and ColWidth) not Passed Properly."
  End If
  
    tmpCol = (UBound(ColNoColWidthList) + 1) / 2
    If mSpread.MaxCols > tmpCol Then mSpread.MaxCols = tmpCol
    
    For Cnt = 0 To UBound(ColNoColWidthList) - 1 Step 2
      mSpread.ColWidth(ColNoColWidthList(Cnt)) = ColNoColWidthList(Cnt + 1)
    Next Cnt
  
End Function

Public Function Common_spSetColWidth(ByRef mSpread As AceSpread,ByVal  ParamArray ColNoColWidthList() As Variant) As Boolean
  'ColNoAndColWidth should be passed like this 1, 25, 2, 30, 6, 25, 10, 45
  'ColNo_x, ColWdith_x, ColNo_y, ColWidth_y,...
  
  Dim Cnt As Long
  Dim tmpCol As Long
  
' On Error Resume Next
  
  If IsMissing(ColNoColWidthList) Then
    sShowMessage "Missing Arguemnts (ColNo and ColWidth)."
  End If
  
  If UBound(ColNoColWidthList) Mod 2 = 0 Then
    sShowMessage "Arguemnts (ColNo and ColWidth) not Passed Properly."
  End If
  
    tmpCol = (UBound(ColNoColWidthList) + 1) / 2
    
    If mSpread.MaxCols < tmpCol Then
      mSpread.MaxCols = tmpCol
    End If
    
    For Cnt = 0 To UBound(ColNoColWidthList) - 1 Step 2
      mSpread.ColWidth(ColNoColWidthList(Cnt)) = ColNoColWidthList(Cnt + 1)
    Next Cnt
  
End Function

Public Function spSetFont(ByRef sSpread As AceSpread,ByVal  sCol As Long,ByVal  sCol1 As Long,ByVal  sRow As Long,ByVal  sRow1 As Long,Optional ByVal SName As String = "Arial",Optional ByVal sSize As Double = 8,Optional ByVal sBold As Boolean = False)
  Dim fCol As Long
  Dim fRow As Long
  
  For fRow = sRow To sRow1
    For fCol = sCol To sCol1
      sSpread.Row = fRow
sSpread.Col = fCol
sSpread.Font = SName
sSpread.FontSize = sSize
sSpread.FontBold = sBold
    Next fCol
  Next fRow
  

End Function

Public Function spSetMultiline(ByRef mSpread As AceSpread,ByVal  fRow As Long,ByVal  ParamArray ColVisStatus() As Variant) As Boolean
  Dim sCount As Integer
  Dim i As Integer
  
' On Error Resume Next
  
  
    If ColVisStatus(0) = "*" Then
      sCount = mSpread.MaxCols
    Else
      sCount = Len(ColVisStatus(0))
    End If
    mSpread.Row = fRow
    For i = 1 To sCount
      mSpread.Col = i
      mSpread.CellType = CellTypeEdit
      mSpread.TypeEditMultiLine = IIf(Trim$(ColVisStatus(0)) = "*", 1, Mid$(ColVisStatus(0), i, 1))
    Next i
    
End Function

Public Function spSetPicture(ByRef mSpread As AceSpread,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  sPath As String) As Boolean

  Try
  
  spSetPicture = False
  
  
    mSpread.Row = fRow
mSpread.Col = fCol
    mSpread.CellType = CellTypePicture
    mSpread.TypePictPicture = LoadPicture(sPath)
    mSpread.TypePicMask = 1
    mSpread.TypePictCenter = True
    mSpread.TypePictMaintainScale = True
    mSpread.TypePictStretch = True
    
  
  spSetPicture = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function spSetText(ByRef mSpread As AceSpread,ByVal  fCol As Long,ByVal  fRow As Long,ByVal  strText As String,Optional ByVal AlignConst As eSprAlignConst = mAlignleftTop,Optional ByVal sFontBold As Boolean = False,Optional ByVal ColToSpan As Long = 1,Optional ByVal RowToSpan As Long = 1) As Boolean
  
' On Error Resume Next
  
    If fRow > mSpread.MaxRows Then
      mSpread.MaxRows = fRow
    End If
    
    If fCol > mSpread.MaxCols Then
      mSpread.MaxCols = fCol
    End If
  
    mSpread.Row = fRow
mSpread.Col = fCol
    mSpread.SetText fCol, fRow, strText
    
    mSpread.TypeHAlign = AlignConst \ 3
    mSpread.TypeVAlign = AlignConst Mod 3
    
    mSpread.FontBold = sFontBold
    
    If ColToSpan > 1 Or RowToSpan > 1 Then
      mSpread.AddCellSpan fCol, fRow, ColToSpan, RowToSpan
    End If
  
  
  spSetText = True
  
End Function

Public Function spSetTextColArray(ByRef mSpread As AceSpread,ByVal  fStartCol As Long,ByVal  fRow As Long,ByVal  ColIncrement As Integer,ByVal  AlignConst As eSprAlignConst,ByVal  sFontBold As Boolean,ByVal  ParamArray strTextList() As Variant) As Boolean
  Dim cCnt As Long
  Dim sCol As Long
  
' On Error Resume Next
  
    If fRow > mSpread.MaxRows Then
      mSpread.MaxRows = fRow
    End If
    
    sCol = fStartCol + UBound(strTextList)
    If sCol > mSpread.MaxCols Then
      mSpread.MaxCols = sCol
    End If
  
    mSpread.Row = fRow
    
    For cCnt = fStartCol To sCol
      mSpread.Col = cCnt
      mSpread.SetText cCnt, fRow, strTextList(cCnt - fStartCol)
      
      mSpread.TypeHAlign = AlignConst \ 3
      mSpread.TypeVAlign = AlignConst Mod 3
      
      mSpread.FontBold = sFontBold
    Next cCnt
  
  spSetTextColArray = True
  
End Function

Public Sub Common_RowProcess(ByVal sForm As Form,ByVal  spRow As Long,Optional ByVal SheetNo As Integer)
' On Error Resume Next
  
  If GridCount >= spHeight Then
      If SheetNo > 0 Then
        sForm.fSpr.Sheet = SheetNo
      Else
        'Continue
      End If
      
      GridCount = 1
      sForm.fSpr.MaxRows = sForm.fSpr.MaxRows + spHeight + 1
      sForm.fSpr.Row = sForm.fSpr.Row + 1:
      spRow = spRow + 1     'One line space at the end of Page
      sForm.fSpr.RowPageBreak = 1
      iPageNo = iPageNo + 1
      sForm.Header
  Else
    'Continue
  End If
End Sub

Public Sub SetExcelExportRemoveBlock(ByVal fValueSetFor As eRem,Optional ByVal fStartCol As Long = -1,Optional ByVal fStartRow As Long = -1,Optional ByVal fNoOfCols As Integer = 0,Optional ByVal fNoOfRows As Integer = 0)
' On Error Resume Next
  
  If fStartCol > -1 Then
    SprRptRemoveRows(fValueSetFor, eRemOpt.mColStart) = fStartCol
  End If
  If fStartRow > -1 Then
    SprRptRemoveRows(fValueSetFor, eRemOpt.mRowStart) = fStartRow
  End If
  If fNoOfCols > 0 Then
    SprRptRemoveRows(fValueSetFor, eRemOpt.mNoOfCols) = fNoOfCols
  End If
  If fNoOfRows > 0 Then
    SprRptRemoveRows(fValueSetFor, eRemOpt.mNoOfRows) = fNoOfRows
  End If
End Sub

Public Function Common_RptStandard_Header(ByVal sRightMarginCol As Long,ByVal  StartingRow As Long,ByVal  ErpNo As String,ByVal  TsNo As String,Optional ByVal PrintDate As Boolean = True,Optional ByVal sERPnoAndTSno As Boolean = True,Optional ByVal CurDate As Variant,Optional ByRef sSpread As AceSpread,Optional ByVal SetHeadingsLeftAligned As Boolean = False,Optional ByVal isSetMaxAsRigthMarginCol As Boolean = True,Optional ByVal IsSmartPrintReport As Boolean = False) As Long




  
  Dim sColLeft As Long
  Dim sColRight As Long
  Dim sdate As Date
  Dim isMultiLineRptHeader As Boolean
  
  If sSpread Is Nothing Then
    If IsSmartPrintReport = True Then
      Set sSpread = frmSpreadReport_Smart.sprPrint
    Else
      Set sSpread = frmSpreadReport.sprPrint
    End If
  Else
    'No Need
  End If
  
    sSpread.BlockMode = False
    
    If sSpread.MaxRows < (iPageNo * spHeight) Then
      sSpread.MaxRows = iPageNo * spHeight
    End If
    
    If isSetMaxAsRigthMarginCol = True Then
      sSpread.MaxCols = sRightMarginCol
    Else
      If iPageNo <= 1 And sRightMarginCol > sSpread.MaxCols Then
        sSpread.MaxCols = sRightMarginCol
      End If
    End If
    
    'Set Page Number
    sColLeft = sRightMarginCol / 2
    sColRight = sRightMarginCol - sColLeft
    
    sSpread.AddCellSpan 1, StartingRow, sColLeft, 1
    sSpread.AddCellSpan sColLeft + 1, StartingRow, sColRight, 1

    'Set Page Number
    sSpread.SetText 1, StartingRow, "Page Number :" + str(iPageNo)
    
    If iPageNo <= 1 Then
      Call SetExcelExportRemoveBlock(eRem.mPageNo, 1, StartingRow, 1, 1)
    End If
    sSpread.Col = 1
sSpread.Row = StartingRow
    
    'After Export To Excel For Alignment Purpose-Added On 03/08/2010
    sSpread.CellType = CellTypeEdit
       
    sSpread.TypeHAlign = TypeHAlignLeft
sSpread.TypeVAlign = TypeVAlignCenter
    
    'Set Print Date
    If PrintDate = True Then
      sSpread.Col = sColLeft + 1
sSpread.Row = StartingRow
      If IsMissing(CurDate) Then
        sdate = SetServerDate
        If sdate < FinStart Then
 GoTo ExitHere
End If
      Else
        sdate = CurDate
      End If
        
      sSpread.SetText sColLeft + 1, StartingRow, "Printed Date : " & Format(sdate, "dd-MM-yyyy HH:MM")
      
      If iPageNo <= 1 Then
        Call SetExcelExportRemoveBlock(eRem.mPrintDate, sColLeft + 1, StartingRow, 1, 1)
      End If
        
      'After Export To Excel For Alignment Purpose-Added On 03/08/2010
      sSpread.CellType = CellTypeEdit
      
      sSpread.TypeHAlign = TypeHAlignRight
sSpread.TypeVAlign = TypeVAlignCenter
    End If
    
    'Set Erp & Ts No
    If sERPnoAndTSno = True Then
      StartingRow = StartingRow + 1
      
      sSpread.AddCellSpan 1, StartingRow, sColLeft, 1
      sSpread.AddCellSpan sColLeft + 1, StartingRow, sColRight, 1
      
      sSpread.Col = 1
sSpread.Row = StartingRow
sSpread.FontUnderline = True
      sSpread.SetText 1, StartingRow, "ERP Ref. No   :" & Space(1) & Trim$(ErpNo)
      If iPageNo <= 1 Then
        Call SetExcelExportRemoveBlock(eRem.mERPNo, 1, StartingRow, 1, 1)
      End If
      
      'After Export To Excel For Alignment Purpose-Added On 03/08/2010
      sSpread.CellType = CellTypeEdit
      
      sSpread.TypeHAlign = TypeHAlignLeft
sSpread.TypeVAlign = TypeVAlignCenter
        
      sSpread.Col = sColLeft + 1
sSpread.Row = StartingRow
sSpread.FontUnderline = True
      
      If gReportID = "KUN" Then
        sSpread.SetText sColLeft + 1, StartingRow, "Format No" & Chr(32) & Chr(32) & " : " & TsNo & String(Val(14 - Len(TsNo)), "_")
      Else
        sSpread.SetText sColLeft + 1, StartingRow, "TS Ref. No" & Chr(32) & Chr(32) & " : " & TsNo & String(Val(14 - Len(TsNo)), "_")
      End If
      
      If iPageNo <= 1 Then
        Call SetExcelExportRemoveBlock(eRem.mTsNo, sColLeft + 1, StartingRow, 1, 1)
      End If
      
      'After Export To Excel For Alignment Purpose-Added On 03/08/2010
      sSpread.CellType = CellTypeEdit
      
      sSpread.TypeHAlign = TypeHAlignRight
sSpread.TypeVAlign = TypeVAlignCenter
    End If
    'Set Company Name
    StartingRow = StartingRow + 1
    
    sSpread.Col = 1
sSpread.Row = StartingRow
    sSpread.AddCellSpan 1, StartingRow, sRightMarginCol, 1
    sSpread.FontBold = True
sSpread.FontSize = 12
    
    'After Export To Excel For Alignment Purpose-Added On 03/08/2010
    sSpread.CellType = CellTypeEdit
    
    If SetHeadingsLeftAligned Then sSpread.TypeHAlign = TypeHAlignLeft Else sSpread.TypeHAlign = TypeHAlignCenter
    sSpread.TypeVAlign = TypeVAlignCenter
    
    sSpread.SetText 1, StartingRow, gCompName
    If iPageNo <= 1 Then
      Call SetExcelExportRemoveBlock(eRem.mCompanyname, 1, StartingRow - 1, 1, 2)
    End If
    
    'Set ReportHeading
    StartingRow = StartingRow + 2
    
    sSpread.Col = 1
sSpread.Row = StartingRow
    sSpread.TypeEditMultiLine = True
    
    If InStr(1, gStrReportHeading, Chr(10), vbTextCompare) > 0 Then
      isMultiLineRptHeader = True
      sSpread.AddCellSpan 1, StartingRow, sRightMarginCol, 2
    Else
      isMultiLineRptHeader = False
      sSpread.AddCellSpan 1, StartingRow, sRightMarginCol, 1
    End If
    
    sSpread.FontBold = True
sSpread.FontSize = 10
sSpread.FontUnderline = True
    
    'After Export To Excel For Alignment Purpose-Added On 03/08/2010
    sSpread.CellType = CellTypeEdit
    sSpread.TypeMaxEditLen = 500
    
    If SetHeadingsLeftAligned Then sSpread.TypeHAlign = TypeHAlignLeft Else sSpread.TypeHAlign = TypeHAlignCenter
    sSpread.TypeVAlign = TypeVAlignCenter
    
    sSpread.SetText 1, StartingRow, gStrReportHeading
    If iPageNo <= 1 Then
      Call SetExcelExportRemoveBlock(eRem.mReportName, 1, StartingRow - 1, CInt(sRightMarginCol), IIf(isMultiLineRptHeader, 3, 2))
    End If
   
    If isMultiLineRptHeader Then
      StartingRow = StartingRow + 3
    Else
      StartingRow = StartingRow + 2
    End If
  
  Common_RptStandard_Header = StartingRow
  
'  DoEvents
Finally
  On Error GoTo 0
  Exit Function
End Function

