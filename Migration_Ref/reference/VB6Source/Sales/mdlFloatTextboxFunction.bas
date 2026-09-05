Attribute VB_Name = "mdlFloatIntegerTextboxFunction"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

 
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Public Const GWL_WNDPROC = (-4)

Public Const WM_CONTEXTMENU = &H7B

Global lpPrevWndProc As Long
'Global gHW As Long

Public Sub NumValMouseDown(ByVal txtBoxHandle As Long)
  'Hook the Textbox control
  Dim gHW As Long
  
' On Error Resume Next
'  gHW = txtBox.hWnd
  gHW = txtBoxHandle
  lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf gWindowProc)
  Debug.Print "Hooked"
End Sub

Public Sub NumValMouseUp(ByVal txtBoxHandle As Long)
  'Unhook the Textbox Control
  Dim gHW As Long
  Dim temp As Long
  
' On Error Resume Next
  gHW = txtBoxHandle
  temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
  Debug.Print "Unhooked"
End Sub

Public Function gWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If Msg = WM_CONTEXTMENU Then
      gWindowProc = True
   Else
      ' Send all other messages to the default message handler
      gWindowProc = CallWindowProc(lpPrevWndProc, hWnd, Msg, wParam, lParam)
   End If
End Function

Public Function NumValKeyDownPress(ByVal fTxtBox As TextBox,ByVal  fKeyCode As Integer,Optional ByVal isKeyPress As Boolean = True,Optional ByVal isCtlKeyPressed As Integer = 0,Optional ByVal MaxVal As Variant,Optional ByVal MinVal As Variant,Optional ByVal DecPlaces As Variant) As Integer
  Dim TextAfter As String
  Dim DecPos As Integer
  Dim isDecNoLimit As Boolean
  Dim NoOfDec As Integer
  
  Try
  NumValKeyDownPress = 0
  
    Select Case fKeyCode
      Case vbKey0 To vbKey9, vbKeyDecimal, vbKeySubtract, vbKeyDelete, vbKeyBack, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
        'Proceed
      Case vbKeyNumpad0 To vbKeyNumpad9, 189, 190
        If isKeyPress = True Then
          GoTo ExitHere
        Else
          If isKeyPress = False Then
            If isCtlKeyPressed > 0 Then
              GoTo ExitHere
            End If
          End If
        End If
      Case Else
        GoTo ExitHere
    End Select
  
    Select Case fKeyCode
      Case vbKey0 To vbKey9
        TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Chr(fKeyCode) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
      Case vbKeyNumpad0 To vbKeyNumpad9
        TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Chr(fKeyCode - 48) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
      Case vbKeySubtract, 189   'Subtract sign non numpad key
        TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Chr(fKeyCode - 48) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
      Case vbKeyDecimal, 190    'Decimal point not in Numerical Pad
        TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & "." & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
      Case vbKeyDelete
        If Len(fTxtBox.Text) > 0 Then
          If fTxtBox.SelLength > 0 Then
            TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
          Else
            TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + 2))
          End If
        End If
      Case vbKeyBack
        If Len(fTxtBox.Text) > 0 Then
          If fTxtBox.SelLength > 0 Then
            TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + fTxtBox.SelLength + 1))
          Else
            TextAfter = Left$(fTxtBox.Text, fTxtBox.SelStart - 1) & Mid$(fTxtBox.Text, (fTxtBox.SelStart + 1))
          End If
        End If
      Case Else
        TextAfter = fTxtBox.Text
    End Select
    
    If IsMissing(DecPlaces) Then
      isDecNoLimit = True
    Else
      isDecNoLimit = False
      
      If Val(DecPlaces) < 0 Then
        NoOfDec = 0
      Else
        NoOfDec = Val(DecPlaces)
      End If
    End If
    
    DecPos = InStr(1, Trim$(TextAfter), ".", vbTextCompare)
    
    If isDecNoLimit = False Then
      'For integer decimal point not allowed
      If NoOfDec = 0 Then
        If DecPos > 0 Then
          GoTo ExitHere
        End If
      End If
      
      'Length of decimal places should not exceed if it is a constraint
      If DecPos > 0 Then
        If Len(Mid$(TextAfter, DecPos + 1)) > DecPlaces Then
          GoTo ExitHere
        End If
      End If
    End If
    
    'two decimal points not allowed
    If InStr(DecPos + 1, Trim$(TextAfter), ".", vbTextCompare) > 0 And NoOfDec > 0 Then
      GoTo ExitHere
    End If
    
    'Subtract sign should be the first character
    If InStr(1, TextAfter, "-", vbTextCompare) > 1 Then
      GoTo ExitHere
    End If
    
    'First digit Zero should be followed by decimal point
    If Left$(TextAfter, 1) = "0" And Mid$(TextAfter, 2, 1) <> "." And Len(TextAfter) > 1 Then
      GoTo ExitHere
    End If
    
    'Text after value should be less than the Maximum Value
    If Not IsMissing(MaxVal) Then
      If IsNumeric(MaxVal) Then
        If CheckMaxMinLimits(TextAfter, , Val(MaxVal)) = False Then
          GoTo ExitHere
        End If
      End If
    End If
  
  NumValKeyDownPress = fKeyCode
Catch ex As Exception

Finally

End Try
End Function

Public Function NumValValidate(ByVal TestVal As String,Optional ByVal isValueMandatory As Boolean = False,Optional ByVal MaxVal As Variant,Optional ByVal MinVal As Variant) As Boolean

  Try

  NumValValidate = CheckMaxMinLimits(TestVal, isValueMandatory, MaxVal, MinVal)
Catch ex As Exception
  MsgBox Err.NUMBER & ": " & Err.Description

Finally

End Try
End Function

Private Function CheckMaxMinLimits(ByVal TestVal As String,Optional ByVal isValueMandatory As Boolean = False,Optional ByVal MaxVal As Variant,Optional ByVal MinVal As Variant) As Boolean
  Try
  CheckMaxMinLimits = False
  
  If Trim$(TestVal) = "" Then
    If isValueMandatory Then
'      MsgBox "Value is Mandatory"
      GoTo ExitHere
    End If
  Else
    If Not IsMissing(MinVal) Then
      If IsNumeric(MinVal) Then
        If Val(TestVal) < Val(MinVal) Then
'          MsgBox "Value less than the Minimum Value " & Val(MinVal)
          GoTo ExitHere
        End If
      Else
'        MsgBox "Minimum Value should be a numeric value"
        GoTo ExitHere
      End If
    End If
    
    If Not IsMissing(MaxVal) Then
      If IsNumeric(MaxVal) Then
        If Val(TestVal) > Val(MaxVal) Then
'          MsgBox "Value greater than the Maximum Value " & Val(MaxVal)
          GoTo ExitHere
        End If
      Else
'        MsgBox "Maximum Value should be a numeric value"
        GoTo ExitHere
      End If
    End If
  End If
  
Catch ex As Exception
  MsgBox Err.NUMBER & ": " & Err.Description

Finally

End Try
End Function
