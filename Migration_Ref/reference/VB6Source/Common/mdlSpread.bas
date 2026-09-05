Attribute VB_Name = "mdlSpread"
Option Explicit

Public ssSpread As AceSpread

Public Enum anRowOrColumn
  'Anand.S on 17/01/2005
  'required for Spread Functions
  anCol = 0
  anRow = 1
End Enum

Public Enum eCondition
  mEquals = 0
  mNotEqual = 1
  mLesserThan = 2
  mGreaterThan = 3
  mLesserThanEqual = 4
  mGreaterThanEqual = 5
End Enum

Public Enum eSearchFlag
  mEqualNoCase = 0            'Searches for exact match to complete text in cell, but ignores cases
  mGreaterOrEqual = 1         'Search for text or a value greater than or equal to the specified text or value
  mPartialMatchAny = 2        'Search for text that matches or partially matches the specified text in any position
  mPartialMatchFromStart = 3  'Search for text that matches or partially matches the specified text only from Start
  mValue = 4                  'Searches for text as if the text provided is equivalent to using the Value property
  mEqualCase = 8              'Search for text that exactly matches the text and case specified
End Enum

Public Function ExportSpreadToExcel(ByRef fSpr As AceSpread,ByVal  fCommonDialog As CommonDialog,Optional ByVal FileName As String) As Boolean
  Dim temp As String
  Dim Temp1 As String
  Dim retVal As Variant
  Dim X As Long
  Dim dlgSave As CommonDialog
  
  Try
  ExportSpreadToExcel = False
  
  temp = GetFileName(Trim$(FileName))

  Set dlgSave = fCommonDialog
  dlgSave.CancelError = True
  dlgSave.FileName = temp
  dlgSave.Filter = "Excel Formated (*.xls)|*.xls"
  dlgSave.ShowSave
  Temp1 = dlgSave.FileName
  
  retVal = Dir$(Temp1)
  If retVal = dlgSave.FileTitle Then
    If MsgBox("File already Exists. Do you want to overwrite it?", vbYesNo) = vbNo Then

      GoTo ExitHere
    Else
      Kill dlgSave.FileName
    End If
  End If
  
  'ExcelPorting
  X = fSpr.ExportExcelBookEx(dlgSave.FileName, "", ExcelSaveFlagNoFormulas)
  
  ' Display result to user based on true/false value of x
  If X = True Then
    MsgBox "Export completed Succcessfully.", , "Result"
  Else
    MsgBox "Export did not succeed.", , "Result"
  End If
  
Catch ex As Exception
  If Err.Number = 70 Then
    sShowMessage "File has been opened by some other user. Pls close it before save."

  ElseIf Err.Number = 53 Then
    sShowMessage "Source File not found."
  Else
    If Err.Number <> 32755 Then
      Call ErrMsgShow(ex)
    End If
  End If

Finally

End Try
End Function

Public Sub Spread_Sort(ByRef Spr As AceSpread,ByVal  mCol As Integer,Optional ByVal isAscending As Boolean = True)
' On Error Resume Next
  
    Spr.Sort 1, 1, Spr.MaxCols, Spr.MaxRows, SortByRow, , IIf(isAscending, 0, 1)
End Sub

Public Sub Spread_Set_Decimal(ByRef mSpread As AceSpread,ByVal  NoOfDecimalPoints As Integer,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2,Optional ByVal mSheet1 As Integer = -1,Optional ByVal mSheet2 As Integer = -2)
' This Procedure is used to set the Decimals in Spread
  Dim rCnt As Integer
  
  Try

    If mSheet1 = -1 Then
      mSheet1 = mSpread.Sheet
    End If
    
    If mSheet2 = -2 Then
      mSheet2 = mSheet1
    ElseIf mSheet2 < 1 Then
      mSheet2 = mSpread.Sheet
    End If
    
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    For rCnt = mSheet1 To mSheet2
      mSpread.Sheet = rCnt
      mSpread.Col = mCol1
mSpread.Col2 = mCol2

      mSpread.Row = mRow1
mSpread.Row2 = mRow2
      mSpread.BlockMode = True
        mSpread.TypeNumberDecPlaces = NoOfDecimalPoints
      mSpread.BlockMode = False
    Next rCnt
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Label_Set_Color(ByRef Ctrl As Label, ByVal ColorConst As eOprModeColorConstants)
  'S.Anand on 17-01-2005
  'To Set the Label Fore Color
' On Error Resume Next
  Ctrl.ForeColor = ColorConst
End Sub
Public Sub SP_rowprocess(ByVal sForm As Form,ByVal  SpRow As Long)
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
SpRow = SpRow + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      sForm.Header
  End If
End Sub

Public Sub Spread_BkColour(ByRef mSpread As AceSpread,ByVal  mBackColor As String,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
  Try
  
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.BackColor = mBackColor
    mSpread.BlockMode = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_CellHorAlign_Set(ByRef mSpread As AceSpread,ByVal  mCellAlign As FPSpreadADO.TypeHAlignConstants,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
' This Procedure is used to set the CellType & Alignment
  Try

    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.TypeHAlign = mCellAlign
    mSpread.BlockMode = False
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_CellType_Set(ByRef mSpread As AceSpread,ByVal  mCellType As FPSpreadADO.CellTypeConstants,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal ByVal mCol2 As Long = -2,Optional ByVal ByVal mRow2 As Long = -2)
' This Procedure is used to set the CellType & Alignment
  Try

    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.CellType = mCellType
    mSpread.BlockMode = False
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_Clear(ByRef mSpread As AceSpread)
  mSpread.Row = 1
mSpread.Row2 = mSpread.MaxRows

  mSpread.Col = 1
mSpread.Col2 = mSpread.MaxCols
  mSpread.BlockMode = True
  mSpread.Action = 12
  mSpread.BlockMode = False
'  Spread_Set_Focus mSpread, 1, 1
End Sub

Public Sub Spread_Clear_Row(mSpread, mCol1, mCol2, mRow1, mRow2)
' This Procedure is used to clear the given row in Spread
' Here mSpread is the name of the Spread Control and
' R is the Row in Spread to be deleted

  mSpread.ClearRange mCol1, mRow1, mCol2, mRow2, True
  
'    mSpread.Row = mRow1
'    mSpread.Row2 = mRow2
'    mSpread.Col = mCol1
'    mSpread.Col2 = mCol2
'    mSpread.BlockMode = True
'    mSpread.Action = 3
'    mSpread.BlockMode = False
End Sub

Public Sub Spread_HideCol(ByRef mSpread As AceSpread,ByVal  HideIt As Boolean,ByVal  mCol1 As Long,Optional ByVal mCol2 As Long = -2)
  Dim rCnt As Long
  
  Try
  
    
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    For rCnt = mCol1 To mCol2
      mSpread.Col = rCnt
      mSpread.ColHidden = HideIt
    Next rCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_HideRow(ByRef mSpread As AceSpread,ByVal  HideIt As Boolean,ByVal  mRow1 As Long,Optional ByVal mRow2 As Long = -2)
  Dim rCnt As Long
  
  Try
  
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    For rCnt = mRow1 To mRow2
      mSpread.Row = rCnt
      mSpread.RowHidden = HideIt
    Next rCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_Lock(ByRef mSpread As AceSpread,ByVal  Lock_Status As Boolean,ByVal  mStart_Col As Long,ByVal  mStart_Row As Long,Optional ByVal mEnd_Col As Long = -2,Optional ByVal mEnd_Row As Long = -2,Optional ByVal mSheetNo As Long = 1)
' On Error Resume Next
  
  mSpread.Sheet = mSheetNo
  
  mSpread.Col = mStart_Col
  mSpread.Row = mStart_Row
   
  If mEnd_Col = -2 Then
    mSpread.Col2 = mStart_Col
  Else
    mSpread.Col2 = mEnd_Col
  End If
  
  If mEnd_Row = -2 Then
    mSpread.Row2 = mStart_Row
  Else
    mSpread.Row2 = mEnd_Row
  End If
  
  mSpread.BlockMode = True
  mSpread.Lock = Lock_Status
  mSpread.BlockMode = False
    
End Sub
Public Sub Spread_PagePrint(ByRef Spr As AceSpread,ByVal  fStartRow As Integer,ByVal  fEndRow As Integer,ByVal  fMaxCol As Integer,ByVal  PrintOrient As PrinterOrientationConstants)
  Dim sRow As Integer
  Dim i As Integer
  Dim IRow(10) As String
  
  Try
  
    sRow = fEndRow
Spr.Tag = 0
Spr.MaxCols = fMaxCol
    For i = fStartRow To Spr.MaxRows
      If i = sRow Or i = Spr.DataRowCnt Then
        If i = Spr.DataRowCnt Then
          sRow = Val(sRow) - i
          Spr.MaxRows = Val(Spr.DataRowCnt + sRow - 3)
          i = i + 1
          Spr.SetCellBorder 1, i, Spr.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
          Spr.SetCellBorder 1, i, Spr.MaxCols, Spr.MaxRows, CellBorderIndexLeft, 0, CellBorderStyleBlank
          Spr.SetCellBorder 1, i, Spr.MaxCols, Spr.MaxRows, CellBorderIndexRight, 0, CellBorderStyleBlank
          Spr.SetCellBorder 1, i + 1, Spr.MaxCols, Spr.MaxRows, CellBorderIndexTop, 0, CellBorderStyleBlank
          Spr.SetCellBorder 1, i, Spr.MaxCols, Spr.MaxRows, CellBorderIndexBottom, 0, CellBorderStyleBlank
          i = Spr.MaxRows
        Else
          Spr.InsertRows i, 1
          IRow(Val(Spr.Tag)) = i
        End If
        Spr.AddCellSpan 1, i, (Spr.MaxCols - 1), 1
        Spr.SetCellBorder 1, i, Spr.MaxCols, i, CellBorderIndexTop, 0, CellBorderStyleSolid
        Spr.Tag = Val(Spr.Tag) + 1
'        .Col = 1: .Row = i: .Font = "Arial": .FontBold = True: .FontSize = 12: .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignCenter
'        .SetText 1, i, gCompName
        Spr.SetCellBorder Spr.MaxCols, i, Spr.MaxCols, i, CellBorderIndexLeft, 0, CellBorderStyleSolid
        Spr.Col = Spr.MaxCols
Spr.Row = i
Spr.Font = "Arial"
Spr.FontBold = True
Spr.TypeVAlign = TypeVAlignCenter
Spr.TypeHAlign = TypeHAlignCenter

        Spr.SetText fMaxCol, i, "Page : " & Val(Spr.Tag)
        If i <> Spr.MaxRows Then
          Call Spread_HideRow(Spr, True, Val(sRow) + 1, Spr.DataRowCnt)
        End If
        If Val(PrintOrient) = 1 Then
          Spr.PrintOrientation = PrintOrientationPortrait
        Else
          Spr.PrintOrientation = PrintOrientationLandscape
        End If
        Spr.PrintPageOrder = PageOrderDownThenOver
        Spr.PrintFirstPageNumber = 1
        Spr.PrintPageStart = 1
        Spr.PrintScalingMethod = PrintScalingMethodSmartPrint
        Spr.BorderStyle = BorderStyleFixedSingle
        Spr.PrintBorder = True
        Spr.PrintCenterOnPageH = True
        Spr.PrintCenterOnPageV = True
        Spr.PrintMarginTop = 500
        Spr.PrintMarginLeft = 50
        Spr.PrintMarginRight = 1.25
        Spr.Action = ActionPrint
        Call Spread_HideRow(Spr, False, Val(fStartRow), Spr.DataRowCnt)
        If i <> Spr.DataRowCnt Then
          Call Spread_HideRow(Spr, True, Val(fStartRow), Val(sRow))
        End If
        sRow = sRow + fEndRow
      End If
    Next i
    For i = 0 To Val(Spr.Tag)
      If Val(IRow(i)) > 0 Then
        Spr.DeleteRows IRow(i), 1
      End If
    Next i
  
  sShowMessage "Print Job Over."
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Public Sub Spread_Set_CellBold(ByRef sSpread As AceSpread,ByVal  iCol As Long,ByVal  IRow As Long,ByVal  iCol2 As Long,ByVal  iRow2 As Long,ByVal  sBold As Boolean)
  Dim i As Integer
    sSpread.Row = IRow
sSpread.Row2 = iRow2
    For i = iCol To iCol2
      sSpread.Col = i
sSpread.FontBold = sBold
    Next i
End Sub

Public Sub Spread_Set_CellBorder(ByRef mSpread As AceSpread,ByVal  mBorderIndex As CellBorderIndexConstants,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
  Try
  
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    If mBorderIndex = 16 Then
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexLeft, 0, CellBorderStyleSolid
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexRight, 0, CellBorderStyleSolid
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexTop, 0, CellBorderStyleSolid
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexBottom, 0, CellBorderStyleSolid
    ElseIf mBorderIndex = 1 Then
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexLeft, 0, CellBorderStyleSolid
    ElseIf mBorderIndex = 2 Then
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexRight, 0, CellBorderStyleSolid
    ElseIf mBorderIndex = 4 Then
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexTop, 0, CellBorderStyleSolid
    ElseIf mBorderIndex = 8 Then
      mSpread.SetCellBorder mCol1, mRow1, mCol2, mRow2, CellBorderIndexBottom, 0, CellBorderStyleSolid
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_Set_CellFontSize(ByRef sSpread As AceSpread,ByVal  iCol As Long,ByVal  IRow As Long,ByVal  iCol2 As Long,ByVal  iRow2 As Long,ByVal  sFont As String,ByVal  sFontSize As Single)
  Dim i As Integer
    sSpread.Row = IRow
sSpread.Row2 = iRow2
    For i = iCol To iCol2
      sSpread.Col = i
sSpread.Font = sFont
sSpread.FontSize = sFontSize
    Next i
End Sub
Public Sub Spread_Set_CellVAlign(ByRef sSpread As AceSpread,ByVal  iCol As Long,ByVal  IRow As Long,ByVal  iCol2 As Long,ByVal  iRow2 As Long,ByVal  svAlin As TypeVAlignConstants)
  Dim i As Integer
    sSpread.Row = IRow
sSpread.Row2 = iRow2
    For i = iCol To iCol2
      sSpread.Col = i
sSpread.TypeVAlign = svAlin
    Next i
End Sub

Public Sub Spread_Set_ColHAlign(ByRef sSpread As AceSpread,ByVal  iCol As Long,ByVal  IRow As Long,ByVal  iCol2 As Long,ByVal  iRow2 As Long,ByVal  sHAlin As TypeHAlignConstants)
  Dim i As Integer
    sSpread.Row = IRow
sSpread.Row2 = iRow2
    For i = iCol To iCol2
      sSpread.Col = i
sSpread.TypeHAlign = sHAlin
    Next i
End Sub

Public Sub Spread_Set_Focus(ByRef mSpread As AceSpread,ByVal  Row As Long,ByVal  Col As Long)
  mSpread.SetFocus
  mSpread.SetActiveCell Col, Row
'  mSpread.Row = R
'  mSpread.Row2 = R
'  mSpread.Col = C
'  mSpread.Col2 = C
'  mSpread.BlockMode = True
'  mSpread.Action = 0
'  mSpread.BlockMode = False
End Sub

'Public Sub Spread_Lock(mSpread, Lock_Status As Integer, mStart_Col, mEnd_Col, mStart_Row, mEnd_Row)
'  mSpread.Row = mStart_Row
'  mSpread.Row2 = mEnd_Row
'  mSpread.Col = mStart_Col
'  mSpread.Col2 = mEnd_Col
'  mSpread.BlockMode = True
'  If Lock_Status = 1 Then
'    mSpread.Lock = False
'  ElseIf Lock_Status = 2 Then
'    mSpread.Lock = True
'  End If
'  mSpread.BlockMode = False
'End Sub

Public Sub Spread_Set_HeaderColor(ByRef Spr As AceSpread, ByVal ColorConst As eOprModeColorConstants, ByVal RowOrCol_Hdr As anRowOrColumn, ByVal RowOrCol As Long,Optional ByVal NthHeader As Long = 0,Optional ByVal EndRowOrCol As Long = -2)
  'S.Anand on 17-01-2005
  'To Set the Spread Header Color
  Dim var As String
  Dim EndRowCol As Long
  Dim cnt As Long
  
' On Error Resume Next
  
  If EndRowOrCol = -2 Then
    EndRowCol = RowOrCol
  Else
    EndRowCol = EndRowOrCol
  End If
  
    For cnt = RowOrCol To EndRowCol
      If RowOrCol_Hdr = anRow Then
        Spr.Col = NthHeader
Spr.Row = cnt
        var = IIf(Spr.CellType = CellTypeButton, Spr.TypeButtonText, Spr.Text)
        Spr.CellType = CellTypeButton
        Spr.TypeButtonTextColor = ColorConst
        Spr.TypeButtonText = Trim$(var)
      Else
        Spr.Col = cnt
Spr.Row = NthHeader
        var = IIf(Spr.CellType = CellTypeButton, Spr.TypeButtonText, Spr.Text)
        Spr.CellType = CellTypeButton
        Spr.TypeButtonTextColor = ColorConst
        Spr.TypeButtonText = Trim$(var)
      End If
    Next cnt
End Sub

Public Sub Spread_Set_Max(ByRef mSpread As AceSpread,ByVal  MaxValue As Double,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
' This Procedure is used to set the Decimals in Spread
  Try
  
    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.TypeNumberMax = MaxValue
    mSpread.BlockMode = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_Set_Min(ByRef mSpread As AceSpread,ByVal  MinValue As Double,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
' This Procedure is used to set the Decimals in Spread
  Try

    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.TypeNumberMin = MinValue
    mSpread.BlockMode = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub Spread_SortOld(ByRef Spr As AceSpread,ByVal  mCol As Integer)
    Spr.SortBy = SortByRow
    Spr.SortKey(1) = mCol
    Spr.SortKeyOrder(1) = SortKeyOrderAscending
    Spr.Col = 1
Spr.Col2 = Spr.MaxCols

    Spr.Row = 1
Spr.Row2 = Spr.MaxRows
    Spr.Action = ActionSort
End Sub

Public Function AllCellContainData(ByRef Spr As AceSpread,ByVal  TestIn As anRowOrColumn,ByVal  TestNo As Long,ByVal  StartNo As Long,ByVal  EndNo As Long,ByVal  ParamArray SkipNo() As Variant) As Boolean
  'TestNo is Row or Col No to which is tested based on the TestIn parameter.
  'StartNo & EndNo is Row or Col opposite to the TestIn Parameter
  
  Dim retVal As Variant
  Dim cnt As Long
  Dim isSkipFound As Boolean
  Dim SkipNoCnt As Long
  Dim SkipIt As Boolean

  Try
  AllCellContainData = False
  
  If IsMissing(SkipNo) Then
    isSkipFound = False
  Else
    isSkipFound = True
  End If
  
  For cnt = StartNo To EndNo
    If isSkipFound Then
      SkipIt = False
      For SkipNoCnt = 0 To UBound(SkipNo)
        If SkipNo(SkipNoCnt) = cnt Then
          SkipIt = True
          Exit For
        End If
      Next SkipNoCnt
    Else
      SkipIt = False
    End If
    
    If SkipIt = False Then
      If TestIn = anCol Then
        Spr.GetText TestNo, cnt, retVal
      Else
        Spr.GetText cnt, TestNo, retVal
      End If
      'If it is Empty returns false
      If Trim$(retVal) = "" Then
        GoTo ExitHere
      End If
    End If
  Next cnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Avoid_Duplicate(ByVal mSpread As Control,ByVal  mCol As Integer
Dim      i As Integer
Dim  j As Integer

Dim      mString As String
Dim  mString1 As String
Dim  mString2 As String
Dim  mString3 As String


    If IsMissing(nCol) Then
      For i = 1 To mSpread.DataRowCnt - 1
        mSpread.Row = i
mSpread.Col = mCol
        mString = Trim$(CStr(mSpread.Text))
        If Not (mString) = "" Then
          For j = i + 1 To mSpread.DataRowCnt
            mSpread.Row = j
mSpread.Col = mCol
            mString1 = Trim$(CStr(mSpread.Text))
            If StrComp(mString, mString1) = 0 Then
              mSpread.Col = mCol
              mSpread.Action = 5
              Avoid_Duplicate = True
              Screen.MousePointer = 0
            End If
          Next j
        End If
      Next i
    Else
      For i = 1 To mSpread.DataRowCnt - 1
        mSpread.Row = i
mSpread.Col = mCol
        mString = Trim$(CStr(mSpread.Text))
        mSpread.Col = nCol
mString2 = Trim$(CStr(mSpread.Text))
        For j = i + 1 To mSpread.DataRowCnt
          mSpread.Row = j
mSpread.Col = mCol
          mString1 = Trim$(CStr(mSpread.Text))
          mSpread.Col = nCol
mString3 = Trim$(CStr(mSpread.Text))
          If StrComp(mString, mString1) = 0 And StrComp(mString2, mString3) = 0 Then
            mSpread.Col = mCol
mSpread.Action = 5
            Avoid_Duplicate = True
            Screen.MousePointer = 0
            Exit Function
          End If
        Next j
      Next i
    End If
End Function

Public Function CopySpreadData(ByRef SourceSpread As AceSpread,ByVal  SourceSheet As Long,Optional ByRef DestnSpread As AceSpread,Optional ByVal DestnSheet As Long,Optional ByVal SourceColStart As Long = 1,Optional ByVal SourceRowStart As Long = 1,Optional ByVal SourceColEnd As Long = -2,Optional ByVal SourceRowEnd As Long = -2,Optional ByVal DestnColStart As Long = -2,Optional ByVal DestnRowStart = -2) As Boolean



  'In locked cells data cannot be copied
  
  Dim sSheet As Long
  Dim dSheet As Long
  Dim sSpread As AceSpread
  Dim dSpread As AceSpread
  Dim sCol As Long
  Dim sCol2 As Long
  Dim sRow As Long
  Dim sRow2 As Long
  Dim dCol As Long
  Dim dCol2 As Long
  Dim dRow As Long
  Dim dRow2 As Long

  Try
  
  Set sSpread = SourceSpread
  sSheet = SourceSheet
  
  If DestnSheet = 0 Then
    dSheet = sSheet
  Else
    dSheet = DestnSheet
  End If
  
  sSpread.Sheet = sSheet
  
  If SourceColStart <= 0 Then
    sCol = 1
  Else
    sCol = SourceColStart
  End If
  
  If SourceRowStart <= 0 Then
    sRow = 1
  Else
    sRow = SourceRowStart
  End If
  
  If SourceColEnd <= -2 Then
    sCol2 = SourceColStart
  ElseIf SourceColEnd <= 0 Then
    sCol2 = sSpread.MaxCols
  Else
    sCol2 = SourceColEnd
  End If
  
  If SourceRowEnd <= -2 Then
    sRow2 = SourceRowStart
  ElseIf SourceRowEnd <= 0 Then
    sRow2 = sSpread.MaxRows
  Else
    sRow2 = SourceRowEnd
  End If
  
  If DestnColStart <= -2 Then
    dCol = SourceColStart
  ElseIf DestnColStart <= 0 Then
    dCol = 1
  Else
    dCol = DestnColStart
  End If
  
  If DestnRowStart <= -2 Then
    dRow = SourceRowStart
  ElseIf DestnRowStart <= 0 Then
    dRow = 1
  Else
    dRow = DestnRowStart
  End If
  
  
  If TypeName(DestnSpread) = "Nothing" Then     'For Optional Object Variable isMissing not works
    Set dSpread = SourceSpread
  Else
    Set dSpread = DestnSpread
  End If
  
  dCol2 = sCol2 - sCol + dCol
  dRow2 = sRow2 - sRow + dRow
  
    sSpread.Sheet = sSheet
    sSpread.SetSelection sCol, sRow, sCol2, sRow2
    sSpread.ClipboardCopy
    sSpread.ClearSelection  
    dSpread.Sheet = dSheet
    dSpread.SetSelection dCol, dRow, dCol2, dRow2
    dSpread.ClipboardPaste
    dSpread.ClearSelection
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub CopySpreadStructure(ByRef SourceSpread As AceSpread,ByVal  SourceSheet As Long,Optional ByRef DestnSpread As AceSpread,Optional ByVal DestnSheet As Long)
  Dim cCnt As Long
  Dim sSheet As Long
  Dim dSheet As Long
  Dim sSpread As AceSpread
  Dim dSpread As AceSpread
  Dim retVal As Variant
  Dim fHAlign As FPSpreadADO.TypeHAlignConstants
  Dim fVAlign As FPSpreadADO.TypeVAlignConstants
  Dim fCellType As FPSpreadADO.CellTypeConstants
  
  Try
  
  Set sSpread = SourceSpread
  sSheet = SourceSheet
  
  If DestnSheet = 0 Then
    dSheet = sSheet
  Else
    dSheet = DestnSheet
  End If
  
  If TypeName(DestnSpread) = "Nothing" Then     'For Optional Object Variable isMissing not works
    Set dSpread = SourceSpread
  Else
    Set dSpread = DestnSpread
  End If
  
  'Setting Max Column
  sSpread.Sheet = sSheet
  retVal = sSpread.MaxCols
  dSpread.Sheet = dSheet
  dSpread.MaxCols = retVal

  'Setting Max Row
  dSpread.Sheet = dSheet
  dSpread.MaxRows = 1000
  
  'commented for future use
  sSpread.Sheet = sSheet
  retVal = sSpread.RowHeight(1)
    dSpread.Sheet = dSheet
    dSpread.Col = cCnt
dSpread.Row = 1
dSpread.Col2 = cCnt
dSpread.Row2 = -1
    dSpread.BlockMode = True
    dSpread.RowHeight(-1) = Val(retVal)
    dSpread.BlockMode = False
  For cCnt = 1 To retVal
    'Setting Column Width for Each Column
    sSpread.Sheet = sSheet
    retVal = sSpread.ColWidth(cCnt)
    dSpread.Sheet = dSheet
    dSpread.ColWidth(cCnt) = retVal
    
    sSpread.Sheet = sSheet
    sSpread.Col = cCnt
    retVal = sSpread.Lock
      dSpread.Sheet = dSheet
      dSpread.Col = cCnt
dSpread.Row = 1
dSpread.Col2 = cCnt
dSpread.Row2 = -1
      dSpread.BlockMode = True
      dSpread.Lock = retVal
      dSpread.BlockMode = False
    'Setting Cell Type for Each Column
    sSpread.Sheet = sSheet
    sSpread.Col = cCnt
    fCellType = sSpread.CellType
    fHAlign = sSpread.TypeHAlign
    fVAlign = sSpread.TypeVAlign
      dSpread.Sheet = dSheet
      dSpread.Col = cCnt
dSpread.Row = 1
dSpread.Col2 = cCnt
dSpread.Row2 = -1
      dSpread.BlockMode = True
      dSpread.CellType = fCellType
      Select Case fCellType
        Case CellTypeButton
          Call CopySpreadButtonStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeDate
          Call CopySpreadDateStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeEdit
          Call CopySpreadEditStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypePic
          
        Case CellTypeStaticText
          
        Case CellTypeTime
          Call CopySpreadTimeStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeComboBox
          Call CopySpreadComboStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypePicture
  
        Case CellTypeCheckBox
          Call CopySpreadCheckStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeOwnerDrawn
  
        Case CellTypeCurrency
          Call CopySpreadCurrencyStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeNumber
          Call CopySpreadNumberStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypePercent
          Call CopySpreadPercentStructure(sSpread, sSheet, dSpread, dSheet, cCnt)
        Case CellTypeScientific
        
      End Select
      dSpread.TypeHAlign = fHAlign
      dSpread.TypeVAlign = fVAlign
      dSpread.BlockMode = False

    'Setting Column Header
    sSpread.Sheet = sSheet
    sSpread.GetText cCnt, 0, retVal
    dSpread.Sheet = dSheet
    dSpread.SetText cCnt, 0, retVal
  Next cCnt
  
  'Setting Column Header height
  sSpread.Sheet = sSheet
  retVal = sSpread.RowHeight(0)
  dSpread.Sheet = dSheet
  dSpread.RowHeight(0) = retVal
  
  'Setting Row Header Width
  sSpread.Sheet = sSheet
  retVal = sSpread.ColWidth(0)
  dSpread.Sheet = dSheet
  dSpread.ColWidth(0) = retVal
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
'Function to Distinct Item List in the spread for the given column
'Result is zero based array but ignore the zero-eth element
Public Function GetDistinctItemList(ByRef fSpread As AceSpread,ByVal  fCol As Long,Optional ByVal ConditionCol As Long = -2,Optional ByVal ConditionValue As String = 0,Optional ByVal Condition As eCondition = mEquals) As String()
  Dim ResList() As String
  Dim ValStr As String
  Dim retVal As Variant
  Dim rCnt As Long
  Dim uBnd As Integer
  Dim CondPass As Boolean

  Try

  ReDim ResList(0)
  uBnd = 0
  ResList(0) = ""
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, rCnt, retVal
      ValStr = Trim$(retVal)
      If ValStr <> ResList(uBnd) Then
        CondPass = False
        If ConditionCol >= 0 Then
          fSpread.GetText ConditionCol, rCnt, retVal
          Select Case Condition
            Case mEquals
              If Trim$(retVal) = ConditionValue Then CondPass = True
            Case mNotEqual
              If Trim$(retVal) <> ConditionValue Then CondPass = True
            Case mLesserThan
              If Trim$(retVal) < ConditionValue Then CondPass = True
            Case mLesserThanEqual
              If Trim$(retVal) <= ConditionValue Then CondPass = True
            Case mGreaterThan
              If Trim$(retVal) > ConditionValue Then CondPass = True
            Case mGreaterThanEqual
              If Trim$(retVal) >= ConditionValue Then CondPass = True
          End Select
        Else
          CondPass = True
        End If
        
        If CondPass = True Then
          ReDim Preserve ResList(uBnd + 1)
          uBnd = uBnd + 1
          ResList(uBnd) = ValStr
        End If
      End If
    Next rCnt

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function GetSprText(ByRef Spr As AceSpread,ByVal  Col As Long,ByVal  Row As Long) As String
  Dim TypeVar As Variant
  
  Try

  Spr.GetText Col, Row, TypeVar
  GetSprText = Trim$(TypeVar)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function NoOfOccurence(ByRef mSpread As AceSpread,ByVal  fSearchCol As Long,ByVal  fSearchString As String,Optional ByVal ByVal fStartRow As Long = 1,Optional ByVal ByVal fEndRow As Long = -1,Optional ByVal ByVal isContigious As Boolean = False) As Integer
  Dim retVal As Long
  Dim tmpVar As Variant
  Dim TotFind As Integer
  
  Try
  NoOfOccurence = 0
  
    If fStartRow < 1 Then
      fStartRow = 1
    End If
    
    If fEndRow = -1 Then
      fEndRow = mSpread.MaxRows
    End If
    
    TotFind = 0
    
    If isContigious Then
      retVal = fStartRow
      Do While retVal <= fEndRow
        mSpread.GetText fSearchCol, retVal, tmpVar
        If Trim$(tmpVar) = fSearchString Then
          TotFind = TotFind + 1
        Else
          Exit Do
        End If
        retVal = retVal + 1
      Loop
    Else
      retVal = fStartRow - 1
      Do While retVal <> -1
        TotFind = TotFind + 1
        retVal = mSpread.SearchCol(fSearchCol, retVal, fEndRow, fSearchString, SearchFlagsNone)
      Loop
      TotFind = TotFind - 1   'Since the Loop returns One more than the actual found
    End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function
Public Sub rowprocess(ByVal sForm As Form,ByVal  SpRow As Long,Optional ByVal SheetNo As Integer)
' On Error Resume Next
  If GridCount >= pheight Then
      If SheetNo > 0 Then
        frmPrint.sprPrint.Sheet = SheetNo
      End If
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
SpRow = SpRow + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      sForm.Header
  End If
End Sub

Public Sub SetSprTxt(ByRef Spr As AceSpread,ByVal  Col As Long,ByVal  Row As Long,ByVal  Val As Variant)
' On Error Resume Next
  Spr.SetText Col, Row, Val
End Sub

Public Sub SortSpread(ByRef Spr As AceSpread,ByVal  SortCol As Integer,ByVal  isAsc As Boolean)
    Spr.SortBy = SortByRow
    Spr.SortKey(1) = SortCol
    If isAsc Then
      Spr.SortKeyOrder(1) = SortKeyOrderAscending
    Else
      Spr.SortKeyOrder(1) = SortKeyOrderDescending
    End If
    Spr.Col = 1
Spr.Col2 = Spr.MaxCols

    Spr.Row = 1
Spr.Row2 = Spr.MaxRows
    Spr.Action = ActionSort
End Sub

Public Function Spread_Cell_Total(ByRef fSpread As AceSpread,Optional ByVal fColStart As Long = 1,Optional ByVal fRowStart As Long = 1,Optional ByVal fColEnd As Long = -1,Optional ByVal fRowEnd As Long = -1,Optional ByVal fStatusCol As Long = -1,Optional ByVal fStatusTrueVal As String = "1") As Double
  'fStatusCol is a Column which contains T/F.  Based on that True value Rows only added
  Dim Total As Double
  Dim cCnt As Long
  Dim rCnt As Long
  Dim retVal As Variant
  Dim retVal2 As Variant
  
  Try
  
  Spread_Cell_Total = 0
  
  If fColEnd = -2 Then
    fColEnd = fColStart
  ElseIf fColEnd = -1 Then
    fColEnd = fSpread.DataColCnt
  End If
  
  If fRowEnd = -2 Then
    fRowEnd = fRowStart
  ElseIf fRowEnd = -1 Then
    fRowEnd = fSpread.DataRowCnt
  End If
  
    For rCnt = fRowStart To fRowEnd
      For cCnt = fColStart To fColEnd
        If fStatusCol = -1 Then       'No Condition to be verified
          fSpread.GetText cCnt, rCnt, retVal
          Total = Total + Val(retVal)
        Else                          'Based on this T/F Cell Value to be added
          fSpread.GetText fStatusCol, rCnt, retVal2
          If Val(retVal2) = fStatusTrueVal Then
            fSpread.GetText cCnt, rCnt, retVal
            Total = Total + Val(retVal)
          End If
        End If
      Next cCnt
    Next rCnt
  
  Spread_Cell_Total = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Clear_CellTag(ByRef Spr As AceSpread,Optional ByVal fStartCol As Long = 0,Optional ByVal fStartRow As Long = 0,Optional ByVal fEndRow As Long = -1) As Boolean

  Try
  Spread_Clear_CellTag = False
  
    Spr.Col = fStartCol
Spr.Row = fStartRow

    Spr.Col2 = fEndCol
Spr.Row2 = fEndRow
    Spr.BlockMode = True
      Spr.CellTag = ""
    Spr.BlockMode = False
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Clear_Duplicate(ByRef mSpread As AceSpread,ByVal  SearchRow As Long,ByVal  ParamArray KeyColList()) As Boolean
  Dim DuplicateRow As Long
  Dim lstArray() As Variant
  
  Try
  Spread_Clear_Duplicate = False
  
  If IsMissing(KeyColList) Then
    GoTo ExitHere
  End If
  
  If (mSpread.DataRowCnt = 1) Or (SearchRow > mSpread.DataRowCnt) Then
    GoTo ExitHere
  End If
  
  ReDim lstArray(UBound(KeyColList)) As Variant
  lstArray = KeyColList
  DuplicateRow = Spread_Find_Duplicate(mSpread, SearchRow, 1, -1, lstArray)
  If DuplicateRow <> -1 Then
      mSpread.ClearRange 1, SearchRow, mSpread.MaxCols, SearchRow, True
      Spread_Clear_Duplicate = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Clear_Duplicate_Nto1(ByRef mSpread As AceSpread,ByVal  SearchRow As Long,ByVal  ValCol As Long,ByVal  ParamArray KeyColList()) As Boolean
  Dim DuplicateRow As Long
  Dim lstArray() As Variant
  
  Try
  Spread_Clear_Duplicate_Nto1 = False
  
  If (mSpread.DataRowCnt = 1) Or (SearchRow > mSpread.DataRowCnt) Then
    GoTo ExitHere
  End If
  
  ReDim lstArray(UBound(KeyColList)) As Variant
  lstArray = KeyColList
  DuplicateRow = Spread_Find_Duplicate_Nto1(mSpread, SearchRow, ValCol, 1, -1, lstArray)
  If DuplicateRow <> -1 Then
      mSpread.ClearRange 1, SearchRow, mSpread.MaxCols, SearchRow, True
      Spread_Clear_Duplicate_Nto1 = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Clear_Sheet(ByRef mSpr As AceSpread,ByVal  mCol As Long,ByVal  mRow As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2,Optional ByVal mSheet As Integer = 1,Optional ByVal mSheet2 As Integer = -2)
  Dim sCnt As Integer

  Try

    If mCol2 = -2 Then
      mCol2 = -mCol
    End If
    
    If mRow2 = -2 Then
      mRow2 = -mRow
    End If
    
    If mSheet2 = -2 Then
      mSheet2 = mSheet
    ElseIf mSheet2 = -1 Then
      mSheet2 = mSpr.SheetCount
    End If
    
    For sCnt = mSheet To mSheet2
      mSpr.Sheet = sCnt
      mSpr.ClearRange mCol, mRow, mCol2, mRow2, True
    Next sCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Col_Total(ByRef fSpread As AceSpread,ByVal  fCol As Long,Optional ByVal fRowStart As Long = 1,Optional ByVal fRowEnd As Long = -1,Optional ByVal fStatusCol As Long = -1,Optional ByVal fStatusTrueVal As String = "1",Optional ByVal fStatusCondition As eCondition = mEquals) As Double
  'fStatusCol is a Column which contains T/F.  Based on that True value Rows only added
  
  Dim Total As Double
  Dim rCnt As Long
  Dim retVal As Variant
  Dim retVal2 As Variant
  
  Try
  
  Spread_Col_Total = 0
  
  If fRowEnd = -2 Then
    fRowEnd = fRowStart
  ElseIf fRowEnd = -1 Then
    fRowEnd = fSpread.DataRowCnt
  End If
  
    For rCnt = fRowStart To fRowEnd
      If fStatusCol = -1 Then       'No Condition to be verified
        fSpread.GetText fCol, rCnt, retVal
        Total = Total + Val(retVal)
      Else                          'Based on this T/F Cell Value to be added
        fSpread.GetText fStatusCol, rCnt, retVal2
        Select Case fStatusCondition
        
          Case 0 ' Equal
            If Trim$(retVal2) = fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case 1 'Not Equal
            If Trim$(retVal2) <> fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case 2 'LesserThan
            If Trim$(retVal2) < fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case 3 'Greater then
            If Trim$(retVal2) > fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case 4 'Lesserthan or Equal
            If Trim$(retVal2) <= fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case 5 'Greaterthan or Equal
            If Trim$(retVal2) >= fStatusTrueVal Then
              fSpread.GetText fCol, rCnt, retVal
              Total = Total + Val(retVal)
            End If
          Case Else
            'Nothing
        End Select
      End If
    Next rCnt
  
  Spread_Col_Total = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Col_Total_FAS(ByRef fSpread As AceSpread,ByVal  fCol As Long,Optional ByVal fRowStart As Long = 1,Optional ByVal fRowEnd As Long = -1,Optional ByVal isConsiderCrDr As Boolean = False,Optional ByVal fStatusCol As Long = -1,Optional ByVal fStatusTrueVal As String = "1") As Double
  'fStatusCol is a Column which contains T/F.  Based on that True value Rows only added
  Dim Total As Double
  Dim rCnt As Long
  Dim retVal As Variant
  Dim retVal2 As Variant
  
  Try
  
  Spread_Col_Total_FAS = 0
  
  If fRowEnd = -2 Then
    fRowEnd = fRowStart
  ElseIf fRowEnd = -1 Then
    fRowEnd = fSpread.DataRowCnt
  End If
  
    For rCnt = fRowStart To fRowEnd
      retVal = 0
      If fStatusCol = -1 Then       'No Condition to be verified
        fSpread.GetText fCol, rCnt, retVal
      Else                          'Based on this T/F Cell Value to be added
        fSpread.GetText fStatusCol, rCnt, retVal2
        If Val(retVal2) = fStatusTrueVal Then
          fSpread.GetText fCol, rCnt, retVal
        End If
      End If
      If isConsiderCrDr Then
        If Right(Trim$(retVal), 2) = "Cr" Then
          Total = Total - Val(retVal)
        Else
          Total = Total + Val(retVal)
        End If
      Else
        Total = Total + Val(retVal)
      End If
    Next rCnt
  
  Spread_Col_Total_FAS = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Export(ByRef fpSpr As AceSpread,ByVal  dlg As CommonDialog,ByVal  Caption As String) As Boolean
  Dim dlgSave As CommonDialog
  Dim retVal As String
  Dim ResultStatus As Boolean
  
  Try
  Spread_Export = False

  Set dlgSave = dlg
  dlgSave.CancelError = True
  dlgSave.FileName = "Report-" & Format(ServerDate, "yyyyMMdd") & ".xls"
  dlgSave.Filter = "Excel (*.xls)|*.xls"
  dlgSave.ShowSave
  retVal = Dir$(dlgSave.FileName)
  If retVal = dlgSave.FileTitle Then
    If MsgBox("File already Exists. Do you want to overwrite it?", vbYesNo) = vbNo Then
      GoTo ExitHere
    Else
' On Error Resume Next
      Kill dlgSave.FileName
    End If
  End If
  
  ' Export Excel file and set result to x
  ResultStatus = fpSpr.ExportToExcel(dlgSave.FileName, Caption, "")
  ' Display result to user based on true/false value of ResultStatus
  If ResultStatus = True Then
    MsgBox "Export complete. File Location: " & dlgSave.FileName, , "Result"
  Else
    MsgBox "Export did not succeed.", , "Result"
  End If
    
'  FileCopy Me.Tag, dlgSave.FileName

Catch ex As Exception
  If Err.Number = 70 Then
    sShowMessage "File has been opened by some other user. Pls close it before save."
  ElseIf Err.Number = 53 Then
    sShowMessage "Source File not found."
  Else
    Call ErrMsgShow(ex)
  End If

Finally

End Try
End Function

Public Function Spread_Find_Duplicate(ByRef mSpread As AceSpread,ByVal  SearchRow As Long,ByVal  mRow As Long,ByVal  mRow2 As Long,ByVal  KeyColList()) As Long
  'To Avoid the duplicate row of SearchRow in the spread for the given keycolumn
  
  Dim TotKeyCol As Integer
  Dim RefArray() As Variant       'To store key text of the given searchrow
  Dim MatchFoundRow As Long       'To Store the Match Found Row
  Dim retRow As Variant           'Text Found Row
  Dim EndRow As Long              'To Restrict Search to Single Row
  Dim cCnt As Long
  
  Try
  Spread_Find_Duplicate = -1
  
    If (mSpread.DataRowCnt = 1) Or (mSpread.DataRowCnt < SearchRow) Then
      GoTo ExitHere
    End If
  
    TotKeyCol = UBound(KeyColList) + 1
    
    If TotKeyCol = 0 Then          'then take all column as key column
      TotKeyCol = mSpread.MaxCols
      ReDim RefArray(1 To TotKeyCol, 1 To 2)        'RefArray(x, 1) is ColumnNo, RefArray(x, 2) is Text
      For cCnt = 1 To TotKeyCol
        RefArray(cCnt, 1) = cCnt
        mSpread.GetText RefArray(cCnt, 1), SearchRow, retRow
        RefArray(cCnt, 2) = Trim$(retRow)
      Next cCnt
    Else
      ReDim RefArray(1 To TotKeyCol, 1 To 2)        'RefArray(x, 1) is ColumnNo, RefArray(x, 2) is Text,
      For cCnt = 1 To TotKeyCol
        RefArray(cCnt, 1) = KeyColList(cCnt - 1)
        mSpread.GetText RefArray(cCnt, 1), SearchRow, retRow
        RefArray(cCnt, 2) = Trim$(retRow)
      Next cCnt
    End If
    
    cCnt = 0
    retRow = mRow - 1
EndRow = mRow2
    
    Do While True
    
      If cCnt >= TotKeyCol Then
        Exit Do
      End If
      
      If retRow = mSpread.DataRowCnt Then
        GoTo ExitHere
      End If
      
      cCnt = cCnt + 1
      
      retRow = mSpread.SearchCol(RefArray(cCnt, 1), retRow, EndRow, RefArray(cCnt, 2), SearchFlagsNone)
      
      'To exclude the SearchRow in Search
      If retRow = SearchRow Then
        retRow = mSpread.SearchCol(RefArray(cCnt, 1), retRow, EndRow, RefArray(cCnt, 2), SearchFlagsNone)
      End If
      
      'No Match Found
      If (retRow = -1) Then
        If cCnt = 1 Then
          GoTo ExitHere
        Else
          cCnt = 0       'To start search from 1st column in next row
          retRow = MatchFoundRow
          EndRow = mRow2
        End If
      Else
        'First Column Match Found
        If cCnt = 1 Then
          MatchFoundRow = retRow
        End If
        'Match Found. So next time search in that row
        EndRow = retRow
        retRow = retRow - 1
      End If
    Loop
  
  Spread_Find_Duplicate = MatchFoundRow
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Find_Duplicate_Nto1(ByRef mSpread As AceSpread,ByVal  SearchRow As Long,ByVal  mValCol As Long,ByVal  mRow As Long,ByVal  mRow2 As Long,ByVal  KeyColList()) As Long
  'To Avoid the duplicate row of SearchRow in the spread for the given keycolumn
  
  Dim TotKeyCol As Integer
  Dim RefArray() As Variant       'To store key text of the given searchrow
  Dim MatchFoundRow As Long       'To Store the Match Found Row
  Dim retRow As Variant           'Text Found Row
  Dim EndRow As Long              'To Restrict Search to Single Row
  Dim cCnt As Long
  Dim ValRow As Long
  
  Try
  Spread_Find_Duplicate_Nto1 = -1
  
    If (mSpread.DataRowCnt = 1) Or (mSpread.DataRowCnt < SearchRow) Then
      GoTo ExitHere
    End If
  
    TotKeyCol = UBound(KeyColList) + 1
    
    mSpread.GetText mValCol, SearchRow, retRow
    ValRow = Val(retRow)
    
    If TotKeyCol = 0 Then          'then take all column as key column
      TotKeyCol = mSpread.MaxCols
      ReDim RefArray(1 To TotKeyCol, 1 To 2)        'RefArray(x, 1) is ColumnNo, RefArray(x, 2) is Text
      For cCnt = 1 To TotKeyCol
        RefArray(cCnt, 1) = cCnt
        mSpread.GetText RefArray(cCnt, 1), ValRow, retRow
        RefArray(cCnt, 2) = Trim$(retRow)
      Next cCnt
    Else
      ReDim RefArray(1 To TotKeyCol, 1 To 2)        'RefArray(x, 1) is ColumnNo, RefArray(x, 2) is Text,
      For cCnt = 1 To TotKeyCol
        RefArray(cCnt, 1) = KeyColList(cCnt - 1)
        mSpread.GetText RefArray(cCnt, 1), ValRow, retRow
        RefArray(cCnt, 2) = Trim$(retRow)
      Next cCnt
    End If
    
    cCnt = 0
    retRow = mRow - 1
EndRow = mRow2
    
    Do While True
    
      If cCnt >= TotKeyCol Then
        Exit Do
      End If
      
      If retRow = mSpread.DataRowCnt Then
        GoTo ExitHere
      End If
      
      cCnt = cCnt + 1
      
      retRow = mSpread.SearchCol(RefArray(cCnt, 1), retRow, EndRow, RefArray(cCnt, 2), SearchFlagsNone)
      
      'To exclude the SearchRow in Search
      If retRow = SearchRow Then
        retRow = mSpread.SearchCol(RefArray(cCnt, 1), retRow, EndRow, RefArray(cCnt, 2), SearchFlagsNone)
      End If
      
      'No Match Found
      If (retRow = -1) Then
        If cCnt = 1 Then
          GoTo ExitHere
        Else
          cCnt = 0       'To start search from 1st column in next row
          retRow = MatchFoundRow
          EndRow = mRow2
        End If
      Else
        'First Column Match Found
        If cCnt = 1 Then
          MatchFoundRow = retRow
        End If
        'Match Found. So next time search in that row
        EndRow = retRow
        retRow = retRow - 1
      End If
    Loop
  
  Spread_Find_Duplicate_Nto1 = MatchFoundRow
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Find_Values(ByRef mSpread As AceSpread,ByVal  mRow As Long,ByVal  mRow2 As Long,ByVal  ParamArray ColValueList()) As Long
  Dim cCnt As Long
  Dim rRow As Long
  Dim fRow As Long

  Try
  Spread_Find_Values = -1
  
  If UBound(ColValueList) = -1 Then
    sShowMessage "Column SearchValue List not given"
    GoTo ExitHere
  End If
  
  If (UBound(ColValueList) + 1) Mod 2 <> 0 Then
    sShowMessage "Column SearchValue Pair not matched given"
    GoTo ExitHere
  End If
  
  rRow = -1
  
    fRow = mRow
    Do While True
      fRow = mSpread.SearchCol(ColValueList(0), fRow - 1, mRow2, ColValueList(1), SearchFlagsNone)
      
      If fRow <> -1 Then
        rRow = fRow
        For cCnt = 3 To UBound(ColValueList) Step 2
          rRow = mSpread.SearchCol(ColValueList(cCnt), rRow - 1, rRow, ColValueList(cCnt + 1), SearchFlagsNone)
          If rRow = -1 Then
            Exit For
          End If
        Next cCnt
      
        If rRow <> -1 Then
          Exit Do
        Else
          fRow = fRow + 1
        End If
      Else
        Exit Do
      End If
    Loop
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Get_CellTag(ByRef Spr As AceSpread,Optional ByVal fCol As Long = -2,Optional ByVal fRow As Long = -2) As Variant

  Try

  
    If fCol = -2 Then
      Spr.Col = Spr.ActiveCol
    Else
      Spr.Col = fCol
    End If
    
    If fRow = -2 Then
      Spr.Row = Spr.ActiveRow
    Else
      Spr.Row = fRow
    End If
    
    Spread_Get_CellTag = Spr.CellTag
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Get_ListAll(ByRef Spr As AceSpread,ByVal  ListCol As Long,Optional ByVal StartRow As Long = 1) As String
  Dim sRow As Long
  Dim retVal As Variant
  Dim Sql As String
  Dim i As Long
  
  Try

    Sql = "("
    For sRow = StartRow To Spr.DataRowCnt
      Spr.GetText ListCol, sRow, retVal
      Sql = Sql & " '" & retVal & "',"
    Next sRow
    Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql & ")"
  
  Spread_Get_ListAll = Sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Get_SelectionList(ByRef Spr As AceSpread,ByVal  ListCol As Long,Optional ByVal StartRow As Long = 1) As String
  Dim sRow As Long
  Dim retVal As Variant
  Dim Sql As String
  Dim i As Long
  
  Try

    sRow = StartRow - 1
    Sql = "("
    For i = 1 To Spr.SelectionCount
      sRow = Spr.GetMultiSelItem(sRow)
      Spr.GetText ListCol, sRow, retVal
      Sql = Sql & " '" & retVal & "',"
    Next i
    Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql & ")"
  
  Spread_Get_SelectionList = Sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Get_SheetNoByName(ByRef fSpread As AceSpread,ByVal  fSheetName As String) As Long
  Dim sCnt As Long
  
  Try
  Spread_Get_SheetNoByName = -1
  
    For sCnt = 1 To fSpread.SheetCount
      fSpread.Sheet = sCnt
      If fSheetName = Trim$(fSpread.SheetName) Then
        Spread_Get_SheetNoByName = sCnt
        GoTo ExitHere
      End If
    Next sCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_PreView(ByRef FsSpread As AceSpread,ByVal  sReportName As String,Optional ByVal spOrientation As PrinterOrientationConstants = cdlPortrait) As Boolean

  Try
  Spread_PreView = False
  
'  spreadpreview.fpSpreadPreview1.hWndSpread = FsSpread.hWnd
'  Set ssSpread = FsSpread
'  spreadpreview.fpSpreadPreview1.BorderStyle = BorderStyleFixedSingle
'  ssSpread.PrintOrientation = spOrientation
'  spreadpreview.Caption = sReportName & "- Print Preview"
'  spreadpreview.Show vbModal
  
  Spread_PreView = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Row_Total(ByRef fSpread As AceSpread,ByVal  fRow As Long,Optional ByVal fColStart As Long = 1,Optional ByVal fColEnd As Long = -1,Optional ByVal fStatusRow As Long = -1,Optional ByVal fStatusTrueVal As String = "1") As Double
  'fStatusCol is a Column which contains T/F.  Based on that True value Rows only added
  Dim Total As Double
  Dim cCnt As Long
  Dim retVal As Variant
  Dim retVal2 As Variant
  
  Try
  
  Spread_Row_Total = 0
  
  If fColEnd = -2 Then
    fColEnd = fColStart
  ElseIf fColEnd = -1 Then
    fColEnd = fSpread.MaxCols
  End If
  
    For cCnt = fColStart To fColEnd
      If fStatusRow = -1 Then       'No Condition to be verified
        fSpread.GetText cCnt, fRow, retVal
        Total = Total + Val(retVal)
      Else                          'Based on this T/F Cell Value to be added
        fSpread.GetText fStatusRow, cCnt, retVal2
        If Val(retVal2) = fStatusTrueVal Then
          fSpread.GetText cCnt, fRow, retVal
          Total = Total + Val(retVal)
        End If
      End If
    Next cCnt
  
  Spread_Row_Total = Total
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Search_Col(ByRef fSpr As AceSpread,ByVal  fSearchCol As Long,ByVal  fSearchString As String,Optional ByVal fSearchFlag As eSearchFlag = mPartialMatchFromStart,Optional ByVal fRowStart As Long = 1,Optional ByVal SetTopRow As Boolean = False) As Long
  Dim Bkmk As Long
  Dim rCnt As Long
  Dim CellText As Variant
  
  Try
  Spread_Search_Col = -1
  
    Bkmk = fSpr.ActiveRow
    If fSearchFlag = mPartialMatchFromStart Then
      For rCnt = fRowStart To fSpr.MaxRows
        fSpr.GetText fSearchCol, rCnt, CellText
        If UCase(Left$(CellText, Len(fSearchString))) = UCase(fSearchString) Then
          Bkmk = rCnt
          Exit For
        End If
      Next rCnt
    Else
      rCnt = fSpr.SearchCol(fSearchCol, fRowStart, -1, fSearchString, fSearchFlag)
      If rCnt > fRowStart Then
        Bkmk = rCnt
      End If
    End If
    
    If SetTopRow Then
      fSpr.SetActiveCell 1, Bkmk
      fSpr.TopRow = Bkmk
    End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_SelectDeSelect(ByRef fSpr As AceSpread,ByVal  isSelect As Boolean) As Boolean
  Dim rCnt As Long

  Try
  Spread_SelectDeSelect = False
  
    If Not (fSpr.OperationMode = OperationModeExtended Or fSpr.OperationMode = OperationModeMulti) Then
      sShowMessage "Spread Operation Mode should be Multi or Extended"
      GoTo ExitHere
    End If
  
    For rCnt = 1 To fSpr.DataRowCnt
      fSpr.Row = rCnt
      fSpr.SelModeSelected = isSelect
    Next rCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Spread_Set_CellTag(ByRef Spr As AceSpread,ByVal  TagText As String,Optional ByVal fCol As Long = -2,Optional ByVal fRow As Long = -2) As Boolean
  
  Try

  Spread_Set_CellTag = False
  
  
    If fCol = -2 Then
      Spr.Col = Spr.ActiveCol
    Else
      Spr.Col = fCol
    End If
    
    If fRow = -2 Then
      Spr.Row = Spr.ActiveRow
    Else
      Spr.Row = fRow
    End If
    
    Spr.CellTag = TagText
    
  
  Spread_Set_CellTag = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function SpreadCellDataDiffrent(ByRef sSpread As AceSpread,ByVal  fRow As Long,ByVal  ParamArray sCol() As Variant) As Boolean
  Dim retVal(1 To 2) As Variant
  Dim i As Long
  
  'create function by rajapriyan on 09/07/2008
  
  Try
  SpreadCellDataDiffrent = True
  
    For i = LBound(sCol()) To UBound(sCol())
      sSpread.Col = sCol(i)
sSpread.Row = fRow
retVal(1) = sSpread.Text
'      .GetText SCol(i), fRow, retVal(1)
      retVal(2) = Spread_Get_CellTag(sSpread, CLng(sCol(i)), fRow)
      
      If IsNumeric(retVal(1)) Then
        If Val(retVal(1)) <> Val(retVal(2)) Then
          GoTo ExitHere
        End If
      Else
        If Trim$(retVal(1)) <> Trim$(retVal(2)) Then
          GoTo ExitHere
        End If
      End If
    Next i
  
  SpreadCellDataDiffrent = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Public Function SpreadCellDataDiffrent(sSpread As fpSpread, fRow As Long, ParamArray sCol() As Variant) As Boolean
'  Dim retVal(1 To 2) As Variant
'  Dim i As Long
'
'  'create function by rajapriyan on 09/07/2008
'
'  On Error GoTo ErrHandler
'  SpreadCellDataDiffrent = True
'
'  With sSpread
'    For i = 1 To UBound(sCol)
'      .GetText sCol(i), fRow, retVal(1)
'      retVal(2) = Spread_Get_CellTag(sSpread, CLng(sCol(i)), fRow)
'      If Trim$(retVal(1)) <> Trim$(retVal(2)) Then
'        GoTo ExitHere
'      End If
'    Next i
'  End With
'
'  SpreadCellDataDiffrent = False
'
'ExitHere:
'  On Error GoTo 0
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Function

Public Sub SpreadHeadColor(ByRef SprName As AceSpread,ByVal  SprTitle As String,ByVal  TitleColor As ColorConstants,ByVal  SprCol As Long,ByVal  SprRow As Long)
  Try

    SprName.Row = SprRow
    SprName.Col = SprCol
    SprName.CellType = CellTypeButton
    SprName.TypeTextWordWrap = True
    SprName.TypeButtonText = SprTitle
    SprName.TypeButtonTextColor = 0
    SprName.TypeButtonTextColor = TitleColor
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Function SpVal(ByVal  lrow As Long,ByVal  lcol As Long) As String
  mSpread.Row = lrow
  mSpread.Col = lcol
  SpVal = mSpread.Text
End Function

Private Sub CopySpreadButtonStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonBorderColor
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonBorderColor = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonColor
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonDarkColor
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonDarkColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonLightColor
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonLightColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonPicture
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonPicture = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonPictureDown
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonPictureDown = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonShadowSize
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonShadowSize = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonText
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonText = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonTextColor
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonTextColor = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeButtonType
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeButtonType = retVal
End Sub
Private Sub CopySpreadCheckStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
  Dim cnt As Integer
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCheckType
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCheckType = retVal
  
  If retVal = TypeCheckTypeThreeState Then
    For cnt = 0 To 5
      sSpread.Sheet = sSheet
      sSpread.Col = fCol
      retVal = sSpread.TypeCheckPicture(cnt)
      If retVal <> 0 Then
        dSpread.Sheet = dSheet
        dSpread.Col = fCol
        dSpread.TypeCheckPicture(cnt) = retVal
      End If
    Next
  Else
    For cnt = 0 To 3
      sSpread.Sheet = sSheet
      sSpread.Col = fCol
      retVal = sSpread.TypeCheckPicture(cnt)
      If retVal <> 0 Then
        dSpread.Sheet = dSheet
        dSpread.Col = fCol
        dSpread.TypeCheckPicture(cnt) = retVal
      End If
    Next
  End If
  
  'For Type Check Center Both the HAlign & VAlign used
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeHAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeVAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeVAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCheckText
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCheckText = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCheckTextAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCheckTextAlign = retVal
End Sub

Private Sub CopySpreadComboStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxAutoSearch
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxAutoSearch = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxCount
  If Val(retVal) > 0 Then
    dSpread.Sheet = dSheet
    dSpread.Col = fCol
    dSpread.TypeComboBoxList = sSpread.TypeComboBoxList
  End If

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxCurSel
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxCurSel = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxEditable
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxEditable = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxIndex
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxIndex = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxList
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxList = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxMaxDrop
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxMaxDrop = retVal

'  sSpread.Sheet = sSheet
'  sSpread.Col = fCol
'  retVal = sSpread.TypeComboBoxString
'  dSpread.Sheet = dSheet
'  dSpread.Col = fCol
'  dSpread.TypeComboBoxString = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeComboBoxWidth
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeComboBoxWidth = retVal
  dSpread.ClearRange 1, 1, -1, -1, True
End Sub

Private Sub CopySpreadCurrencyStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyDecPlaces = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyMax
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyMin
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyNegStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyPosStyle
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyPosStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencySeparator
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencySeparator = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyShowSep
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyShowSep = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencyShowSymbol
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencyShowSymbol = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeCurrencySymbol
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeCurrencySymbol = retVal
End Sub

Private Sub CopySpreadDateStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeDateCentury
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeDateCentury = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeDateFormat
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeDateFormat = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeDateMax
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeDateMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeDateMin
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeDateMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeDateSeparator
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeDateSeparator = retVal
End Sub

Private Sub CopySpreadEditStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEditCharCase
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEditCharCase = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEditCharSet
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEditCharSet = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEditMultiLine
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEditMultiLine = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEditPassword
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEditPassword = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEllipses
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEllipses = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeHAlign = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeMaxEditLen
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeMaxEditLen = retVal
End Sub

Private Sub CopySpreadNumberStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNegRed
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNegRed = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberDecPlaces = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberMax
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberMin
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberNegStyle = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberSeparator
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberSeparator = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeNumberShowSep
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeNumberShowSep = retVal
End Sub

Private Sub CopySpreadPercentStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentDecimal
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentDecimal = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentDecPlaces
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentDecPlaces = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentLeadingZero
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentLeadingZero = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentMax
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentMax = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentMin
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentMin = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypePercentNegStyle
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypePercentNegStyle = retVal
End Sub

Private Sub CopySpreadTimeStructure(ByRef sSpread As AceSpread,ByVal  sSheet As Long,ByRef  dSpread As AceSpread,ByVal  dSheet As Long,ByVal  fCol As Long)
  Dim retVal As Variant
  Dim cnt As Integer
  
' On Error Resume Next
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeTime24Hour
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeTime24Hour = retVal

  'For Type Time Center Both the HAlign & VAlign and Time Format used
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeHAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeHAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeVAlign
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeVAlign = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeTimeSeconds
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeTimeSeconds = retVal

  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeEllipses
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeEllipses = retVal
  
  sSpread.Sheet = sSheet
  sSpread.Col = fCol
  retVal = sSpread.TypeSpin
  dSpread.Sheet = dSheet
  dSpread.Col = fCol
  dSpread.TypeSpin = retVal
End Sub

Public Function ADORS_Spread_DataFill(ByVal fRs As AceADODB.Recordset,ByRef  fSpr As AceSpread,Optional ByVal isHdrFill As Boolean = False,Optional ByVal fStCol As Long = 1,Optional ByVal fStRow As Long = 1,Optional ByVal fSheet As Integer = 1,Optional ByVal isPreserveDesignLayout As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim cCnt As Long
  Dim fMaxCol As Long
  Dim fMaxRow As Long
  Dim DtType As Long
  
  Try
  ADORS_Spread_DataFill = False
  
  If fRs.RecordCount = 0 Then
    sShowMessage "No Records Found"
    GoTo ExitHere
  End If
  
  If isPreserveDesignLayout = False Then
    Spread_Clear fSpr
    fSpr.MaxCols = 1
  End If
  
  If fStCol < 0 Or fStRow < 0 Then
    sShowMessage "Invalid Coloum/Row"
    GoTo ExitHere
  End If
  
  If fSheet < 0 Then
    sShowMessage "Invalid Sheet"
    GoTo ExitHere
  End If
  
    ' Set Sheet No
    fSpr.Sheet = fSheet
    ' Set MaxCol and MaxRow
    fMaxCol = fRs.Fields.count + fStCol - 1
    fMaxRow = fRs.RecordCount + fStRow - 1
    
    If fSpr.MaxCols < fMaxCol Then
      fSpr.MaxCols = fMaxCol
    End If
    
    If fSpr.MaxRows < fMaxRow Then
      fSpr.MaxRows = fMaxRow
    End If
    
    fSpr.ClearRange fStCol, fStRow, fMaxCol, fMaxRow, True
    
    ' Fill Data
    For rCnt = fStRow To fMaxRow
      If rCnt = fStRow Then
        fSpr.Row = -1 ': fSpr.Row2 = -1
        
        For cCnt = fStCol To fMaxCol
          DtType = fRs.Fields(cCnt - fStCol).Type
          fSpr.Col = (cCnt - fStCol) + 1 ': fSpr.BlockMode = False
          Select Case DtType
            Case 2, 3, 4, 5, 14, 20, 131 'Numeric Format
              fSpr.CellType = CellTypeNumber
              fSpr.TypeNumberDecPlaces = NoOfDecimal(CheckNull(fRs(cCnt - fStCol), 0))
            Case 6  'Currency Format
              fSpr.CellType = CellTypeNumber
              fSpr.TypeNumberDecPlaces = 2
            Case 129, 200 ', 135 'String Format
              fSpr.CellType = CellTypeEdit
              fSpr.TypeMaxEditLen = 3500
            Case 133, 134, 135  'Date Format
              fSpr.CellType = CellTypeDate
              fSpr.TypeDateFormat = TypeDateFormatDDMMYY
            Case 11       ' Check Box for Bit Field
              fSpr.CellType = CellTypeCheckBox
              fSpr.TypeCheckType = TypeCheckTypeNormal
              fSpr.TypeCheckCenter = True
            Case Else
              fSpr.CellType = CellTypeEdit
          End Select
          '.BlockMode = False
        Next cCnt
      End If
      For cCnt = fStCol To fMaxCol
        DtType = fRs.Fields(cCnt - fStCol).Type
        Select Case DtType
          Case 2, 3, 4, 5, 14, 20, 131 'Numeric Format
              fSpr.Row = rCnt
fSpr.Col = cCnt
              fSpr.TypeNumberDecPlaces = NoOfDecimal(CheckNull(fRs(cCnt - fStCol), 0))
              fSpr.SetText cCnt, rCnt, fRs(cCnt - fStCol)
          Case 135
            If Val(Format(fRs(cCnt - fStCol), "yyyyMMdd")) = Val(Format(fRs(cCnt - fStCol), "yyyyMMdd.HHmmss")) Then
              fSpr.SetText cCnt, rCnt, Format(fRs(cCnt - fStCol), "dd/MM/yyyy") 'Date Format
            Else  'Date with Time Format and Time Format
              If Val(Format(fRs(cCnt - fStCol), "yyyyMM")) <= Val(190001) Then  'Time Format
                fSpr.SetText cCnt, rCnt, Format(fRs(cCnt - fStCol), "HH:mm:ss")
              Else    'Date with Time Format
                fSpr.SetText cCnt, rCnt, Format(fRs(cCnt - fStCol), "dd/MM/yyyy HH:mm:ss")
              End If
            End If
          Case 11 'Bit Value
            fSpr.SetText cCnt, rCnt, IIf(fRs(cCnt - fStCol), 1, 0)
          Case 129, 200 ', 135 'String Format
            fSpr.SetText cCnt, rCnt, fRs(cCnt - fStCol)
          Case Else
            fSpr.SetText cCnt, rCnt, fRs(cCnt - fStCol)
        End Select
      Next cCnt
      fRs.MoveNext
    Next
    
    ' Header Fill
    If isHdrFill = True Then
      For cCnt = fStCol To fMaxCol
        fSpr.SetText cCnt, 0, fRs.Fields(cCnt - fStCol).Name
        If isPreserveDesignLayout = False Then
          ' Set the width of column to the returned value
          fSpr.ColWidth(cCnt) = fSpr.MaxTextColWidth(cCnt)
        End If
      Next cCnt
    Else
      If isPreserveDesignLayout = False Then
        For cCnt = fStCol To fMaxCol
          ' Set the width of column to the returned value
          fSpr.ColWidth(cCnt) = fSpr.MaxTextColWidth(cCnt)
        Next cCnt
      End If
    End If

  
  ADORS_Spread_DataFill = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

