VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_PREPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Option Compare Text

'Pre-Print Property Variable For Font Name, size & Style
Private l_PPRNFName As String
Private l_PPRNFSize As Single
Private l_PPRNFStyle As PPRNFontStyleConst
Private l_PPRNPSize As PPRNPaperSizeConst
Private l_PPRNPOrient As PPRNOrientConst
Private l_PPRNPQuality As PPRNPrintQualityConst
Private l_PPRNPSMode As PPRNScaleModeConst
Private l_PPRNTxtCtrl As PPRNTextControlConst
Private l_PPRNPWidth As Double
Private l_PPRNPHeight As Double
Private l_PPRNFlagSetPaper As Boolean
Private l_PPRNPaperBin As PPRNPaperBinConst

Private l_PPRNPagePerPaper As Long
Private l_PPRNPagePerHeight As Single

Private l_PPRNMarTopAdj As Single   'Margin Top Adjacent
Private l_PPRNMarLeftAdj As Single  'Margin Left Adjacent

'Pre-Print font Styles
Public Enum PPRNFontStyleConst
  PPRNFSNone = -7
  PPRNFSBold = 1
  PPRNFSItalic = 2
  PPRNFSUnderline = 4
End Enum

'Pre-Print Orientations
Public Enum PPRNOrientConst
  PPRNORPortrait = 1  ' Documents print with the top at the narrow side of the paper
  PPRNORLandscape = 2 ' Documents print with the top at the wide side of the paper
End Enum

'Pre-Print Print Quality
Public Enum PPRNPrintQualityConst
  PPRNPQDraft = -1  ' Draft print quality
  PPRNPQLow = -2    ' Low print quality
  PPRNPQMedium = -3 ' Medium print quality
  PPRNPQHigh = -4   ' High print quality
End Enum

'Pre-Print Print ScaleMode
Public Enum PPRNScaleModeConst
  PPRNSMUser = 0        'Indicates that one or more of the ScaleHeight, ScaleWidth, ScaleLeft, and ScaleTop properties are set to custom values.
  PPRNSMTwips = 1       'Twip (1440 twips per inch; 567 twips per centimeter). (This is the default.)
  PPRNSMPoints = 2      'Point (72 points per logical inch).
  PPRNSMPixels = 3      'Pixel (smallest unit of monitor or printer resolution).
  PPRNSMCharacters = 4  'Character (horizontal = 120 twips per unit; vertical = 240 twips per unit).
  PPRNSMInches = 5      'Inch.
  PPRNSMMillimeters = 6 'Millimeter.
  PPRNSMCentimeters = 7 'Centimeter.
End Enum

'Pre-Print Paper Size
Public Enum PPRNPaperSizeConst
  PPRNPSLetter = 1                ' Letter, 8 1/2 x 11 in
  PPRNPSLetterSmall = 2           ' +A611Letter Small, 8 1/2 x 11 in
  PPRNPSTabloid = 3               ' Tabloid, 11 x 17 in
  PPRNPSLedger = 4                ' Ledger, 17 x 11 in
  PPRNPSLegal = 5                 ' Legal, 8 1/2 x 14 in
  PPRNPSStatement = 6             ' Statement, 5 1/2 x 8 1/2 in
  PPRNPSExecutive = 7             ' Executive, 7 1/2 x 10 1/2 in
  PPRNPSA3 = 8                    ' A3, 297 x 420 mm
  PPRNPSA4 = 9                    ' A4, 210 x 297 mm
  PPRNPSA4Small = 10              ' A4 Small, 210 x 297 mm
  PPRNPSA5 = 11                   ' A5, 148 x 210 mm
  PPRNPSB4 = 12                   ' B4, 250 x 354 mm
  PPRNPSB5 = 13                   ' B5, 182 x 257 mm
  PPRNPSFolio = 14                ' Folio, 8 1/2 x 13 in
  PPRNPSQuarto = 15               ' Quarto, 215 x 275 mm
  PPRNPS1H14 = 16                 ' 10 x 14 in
  PPRNPS11x17 = 17                ' 11 x 17 in
  PPRNPSNote = 18                 ' Note, 8 1/2 x 11 in
  PPRNPSEnv9 = 19                 ' Envelope #9, 3 7/8 x 8 7/8 in
  PPRNPSEnv10 = 20                ' Envelope #10, 4 1/8 x 9 1/2 in
  PPRNPSEnv11 = 21                ' Envelope #11, 4 1/2 x 10 3/8 in
  PPRNPSEnv12 = 22                ' Envelope #12, 4 1/2 x 11 in
  PPRNPSEnv14 = 23                ' Envelope #14, 5 x 11 1/2 in
  PPRNPSCSheet = 24               ' C size sheet
  PPRNPSDSheet = 25               ' D size sheet
  PPRNPSESheet = 26               ' E size sheet
  PPRNPSEnvDL = 27                ' Envelope DL, 110 x 220 mm
  PPRNPSEnvC3 = 29                ' Envelope C3, 324 x 458 mm
  PPRNPSEnvC4 = 30                ' Envelope C4, 229 x 324 mm
  PPRNPSEnvC5 = 28                ' Envelope C5, 162 x 229 mm
  PPRNPSEnvC6 = 31                ' Envelope C6, 114 x 162 mm
  PPRNPSEnvC65 = 32               ' Envelope C65, 114 x 229 mm
  PPRNPSEnvB4 = 33                ' Envelope B4, 250 x 353 mm
  PPRNPSEnvB5 = 34                ' Envelope B5, 176 x 250 mm
  PPRNPSEnvB6 = 35                ' Envelope B6, 176 x 125 mm
  PPRNPSEnvItaly = 36             ' Envelope, 110 x 230 mm
  PPRNPSEnvMonarch = 37           ' Envelope Monarch, 3 7/8 x 7 1/2 in
  PPRNPSEnvPersonal = 38          ' Envelope, 3 5/8 x 6 1/2 in
  PPRNPSFanfoldUS = 39            ' U.S. Standard Fanfold, 14 7/8 x 11 in
  PPRNPSFanfoldStdGerman = 40     ' German Standard Fanfold, 8 1/2 x 12 in
  PPRNPSFanfoldLglGerman = 41     ' German Legal Fanfold, 8 1/2 x 13 in
'  PPRNPS8d5x11 = 54               ' 12240 X 15840 twips , 8.5 x 11 in            d=.(Dot)
'  PPRNPS8d27x11d69 = 55           ' 11904 X 16832 twips ,  8.27  x   11.69 in
'  PPRNPS16d53x23d27 = 120         ' 23816 X 33664 twips , 16.53   x 23.37 in
'  PPRNPS16d53x11d69 = 121         ' 23816 X 16832 twips , 16.53 x 11.69 in
'  PPRNPS8x6 = 122                 ' 11520 X 8640 twips ,  8 x 6 in
'  PPRNPS14d87x11d6 = 123          ' 21424 X 16832 twips , 14.87 x 11.6 in
'  PPRNPS8d26x12 = 124             ' 11904 X 17280 twips , 8.26  x 12 in
'  PPRNPS14d09x12 = 125            ' 20296 X 17280 twips , 14.09 x 12 in
'  PPRNPS15x12 = 126               ' 21600 X 17280 twips , 15  x 12 in
'  PPRNPS16d53x23d37 = 127         ' 23816 X 33664 twips , 16.53 x 23.37 in
'  PPRNPS4d13x5d83 = 128           ' 5952 X 8392 twips , 4.13  x 5.83 in
'  PPRNPSFanfoldStd8x6 = 122      ' German Standard Fanfold, 8 x 6 in
  'PPRNPSFanfoldStdHalf = 134      ' German Standard Fanfold, 8 x 6 in
  PPRNPSUser = 256                ' User-defined
End Enum

'Pre-Print Textin Box Align
Public Enum PPRNAlignConst
  'Value Mod 3 i.e Reminder Gives the Vertical Alignment
  'Value \ 3 i.e Quotient Gives the Horizontal Alignment
  '0 - Top, Left; 1 - Bottom, Right; 2 - Center, Center
  
  PPRNALleftTop = 0
  PPRNALLeftBottom = 1
  PPRNALLeftCenter = 2
  
  PPRNALRightTop = 3
  PPRNALRightBottom = 4
  PPRNALRightCenter = 5
  
  PPRNALCenterTop = 6
  PPRNALCenterBottom = 7
  PPRNALCenterCenter = 8
End Enum

'Pre-Print Text Control's
Public Enum PPRNTextControlConst
  PPRNTCNone = 0
  PPRNTCWrap = 1
  PPRNTCFit = 2
End Enum

Public Enum PPRNPaperBinConst
  PPRNBNUpper = 1           'paper from the upper bin.
  PPRNBNLower = 2           'Use paper from the lower bin.
  PPRNBNMiddle = 3          'Use paper from the middle bin.
  PPRNBNManual = 4          'Wait for manual insertion of each sheet of paper.
  PPRNBNEnvelope = 5        'Use envelopes from the envelope feeder.
  PPRNBNEnvManual = 6       'Use envelopes from the envelope feeder, but wait for manual insertion.
  PPRNBNAuto = 7            'Use paper from the current default bin. (This is the default.)
  PPRNBNTractor = 8         'Use paper fed from the tractor feeder.
  PPRNBNSmallFmt = 9        'Use paper from the small paper feeder.
  PPRNBNLargeFmt = 10       'Use paper from the large paper bin.
  PPRNBNLargeCapacity = 11  'Use paper from the large capacity feeder.
  PPRNBNCassette = 14       'Use paper from the attached cassette cartridge.
End Enum

Public Function PPRN_NewPage()

' On Error Resume Next
  
  
'  If l_PPRNPaperBin = PPRNBNManual Then
'
''    If l_PPRNPHeight > Printer.CurrentY Then
'      l_PPRNPagePerHeight = l_PPRNPagePerHeight + l_PPRNPHeight
'      If (l_PPRNPagePerHeight / l_PPRNPHeight) Mod l_PPRNPagePerPaper = 0 Then
'        Printer.NewPage
'        l_PPRNPagePerHeight = 0
'      Else
'        Printer.CurrentY = l_PPRNPHeight
'      End If
''    End If
'  Else
    Printer.NewPage
'  End If

End Function

Public Function PPRN_EndDocument()
' On Error Resume Next
  
  Printer.EndDoc
  
End Function
Public Function PPRN_KillDocument()
' On Error Resume Next
  
  Printer.KillDoc
  
End Function
Public Property Get PPRNFontName() As String
  PPRNFontName = l_PPRNFName
End Property

Public Property Let PPRNFontName(ByVal vNewValue As String)
  l_PPRNFName = vNewValue
End Property

Public Property Get PPRNFontSize() As Double
  PPRNFontName = l_PPRNFSize
End Property

Public Property Let PPRNFontSize(ByVal vNewValue As Double)
  l_PPRNFSize = vNewValue
End Property

Public Property Get PPRNFontStyle() As PPRNFontStyleConst
  PPRNFontStyle = l_PPRNFStyle
End Property

Public Property Let PPRNFontStyle(ByVal vNewValue As PPRNFontStyleConst)
  l_PPRNFStyle = vNewValue
End Property

Public Property Get PPRNPaperSize() As PPRNPaperSizeConst
  PPRNPaperSize = l_PPRNPSize
End Property

Public Property Let PPRNPaperSize(ByVal vNewValue As PPRNPaperSizeConst)
  l_PPRNPSize = vNewValue
End Property

Public Property Get PPRNOrientation() As PPRNOrientConst
  PPRNOrientation = l_PPRNPOrient
End Property

Public Property Let PPRNOrientation(ByVal vNewValue As PPRNOrientConst)
  l_PPRNPOrient = vNewValue
End Property

Public Property Get PPRNPrintQuality() As PPRNPrintQualityConst
  PPRNPrintQuality = l_PPRNPQuality
End Property

Public Property Let PPRNPrintQuality(ByVal vNewValue As PPRNPrintQualityConst)
  l_PPRNPQuality = vNewValue
End Property

Public Property Get PPRNScaleMode() As PPRNScaleModeConst
  PPRNScaleMode = l_PPRNPSMode
End Property

Public Property Let PPRNScaleMode(ByVal vNewValue As PPRNScaleModeConst)
  l_PPRNPSMode = vNewValue
End Property

Public Property Get PPRNPaperWidth() As Single
  PPRNPaperWidth = l_PPRNPWidth
End Property

Public Property Let PPRNPaperWidth(ByVal vNewValue As Single)
  l_PPRNPWidth = vNewValue
End Property

Public Property Get PPRNPaperHeight() As Single
  PPRNPaperHeight = l_PPRNPHeight
End Property

Public Property Let PPRNPaperHeight(ByVal vNewValue As Single)
  l_PPRNPHeight = vNewValue
End Property

Public Property Get PPRNTextControl() As PPRNTextControlConst
  PPRNTextControl = l_PPRNTxtCtrl
End Property

Public Property Let PPRNTextControl(ByVal vNewValue As PPRNTextControlConst)
  l_PPRNTxtCtrl = vNewValue
End Property

Public Property Get PPRNMarginTopAdj() As Single
  PPRNMarginTopAdj = l_PPRNMarTopAdj
End Property

Public Property Let PPRNMarginTopAdj(ByVal vNewValue As Single)
  l_PPRNMarTopAdj = vNewValue
End Property

Public Property Get PPRNMarginLeftAdj() As Single
  PPRNMarginLeftAdj = l_PPRNMarLeftAdj
End Property

Public Property Let PPRNMarginLeftAdj(ByVal vNewValue As Single)
  l_PPRNMarLeftAdj = vNewValue
End Property

Public Function PPRN_PrintText(ByVal strText As String,ByVal  X_Pos As Single,Optional ByVal Y_Pos As Single,Optional ByVal PPRNFName As String,Optional ByVal PPRNFSize As Single,Optional ByVal PPRNFStyle As PPRNFontStyleConst) As Boolean

  
  Try

  PPRN_PrintText = False
  
  If l_PPRN_SetFontToPrint(PPRNFName, PPRNFSize, PPRNFStyle) = False Then
 GoTo ExitHere
End If
  
  Printer.CurrentX = (X_Pos + l_PPRNMarLeftAdj) '* 566.929133858
  Printer.CurrentY = (Y_Pos + l_PPRNMarTopAdj) + l_PPRNPagePerHeight
  Printer.Print strText
  
  PPRN_PrintText = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call PPRN_KillDocument

Finally

End Try
End Function

Private Function l_PPRN_SetFontToPrint(Optional ByVal PPRNFName As String,Optional ByVal PPRNFSize As Single,Optional ByVal PPRNFStyle As PPRNFontStyleConst) As Boolean
  Dim vStyle As Long

  Try
  
  l_PPRN_SetFontToPrint = False

  If Trim$(PPRNFName) = "" Then
    Printer.FontName = l_PPRNFName
  Else
    Printer.FontName = PPRNFName
  End If
  
  If Val(PPRNFSize) = 0 Then
    Printer.FontSize = l_PPRNFSize
  Else
    Printer.FontSize = PPRNFSize
  End If
  
  If IsMissing(PPRNFStyle) Then
    vStyle = l_PPRNFStyle
  Else
    vStyle = PPRNFStyle
  End If
  
  Select Case vStyle
  
    Case 0
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontUnderline = False
      
    Case 1
      Printer.FontBold = True
      Printer.FontItalic = False
      Printer.FontUnderline = False
      
    Case 2
      Printer.FontBold = False
      Printer.FontItalic = True
      Printer.FontUnderline = False
    
    Case 3
      Printer.FontBold = True
      Printer.FontItalic = True
      Printer.FontUnderline = False
    
    Case 4
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontUnderline = True
    
    Case 5
      Printer.FontBold = True
      Printer.FontItalic = False
      Printer.FontUnderline = True
    
    Case 6
      Printer.FontBold = False
      Printer.FontItalic = True
      Printer.FontUnderline = True
    
    Case 7
      Printer.FontBold = True
      Printer.FontItalic = True
      Printer.FontUnderline = True
    
  End Select
  
  l_PPRN_SetFontToPrint = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function l_PPRN_SetPaperToPrint() As Boolean
  Dim vStyle As Long
  Dim FormNo As Integer
  Dim PaperSize As SIZEL

  Try
  
  l_PPRN_SetPaperToPrint = False

'  If l_PPRNFlagSetPaper = False Then
    Printer.ScaleMode = l_PPRNPSMode
    'Printer.PaperBin = l_PPRNPaperBin
    
    If l_PPRNPSize <> PPRNPSUser Then
      Printer.PaperSize = l_PPRNPSize
    Else
      'Get Paper Size Index From Print Server Properties
      FormNo = l_PPRN_GetPaperSizeNo
      If Val(FormNo) <> 0 Then
        Printer.PaperSize = FormNo
      Else
        sShowMessage "Please run the PrintServerUtilities.exe, located in ACEProfit_Setup folder, in the System where " & vbLf & Printer.DeviceName & " Printer attached with Administrative Privilege. In that application " & vbLf & "Add Paper Size as " & Round(Printer.ScaleX(l_PPRNPWidth, l_PPRNPSMode, PPRNSMInches), 2) & " in X " & Round(Printer.ScaleY(l_PPRNPHeight, l_PPRNPSMode, PPRNSMInches), 2) & " in. [Width x Height ]"



        GoTo ExitHere
      End If
'      Printer.Width = Printer.ScaleX(l_PPRNPWidth, l_PPRNPSMode, PPRNSMTwips)
'      Printer.Height = Printer.ScaleY(l_PPRNPHeight, l_PPRNPSMode, PPRNSMTwips)
    End If
    
    l_PPRNPagePerPaper = Floor(Printer.ScaleY(Printer.Height, PPRNSMTwips, l_PPRNPSMode) / l_PPRNPHeight)

    l_PPRNFlagSetPaper = True
'  End If
  
  l_PPRN_SetPaperToPrint = True
Catch ex As Exception
  If Err.Number = 380 Then
    sShowMessage "The " & Printer.DeviceName & " Does't Support For this Print"
  Else
    Call ErrMsgShow(ex)
  End If

Finally

End Try
End Function

Private Function l_PPRN_GetPaperSizeNo() As Integer
  Dim PrinterHandle As Long
  Dim PaperSize As SIZEL
  Dim Dm As DEVMODE

  Try
  
  l_PPRN_GetPaperSizeNo = 0

  If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
  
    PaperSize.cx = Round(Printer.ScaleX(l_PPRNPWidth, l_PPRNPSMode, PPRNSMMillimeters) * 1000)
    PaperSize.cy = Round(Printer.ScaleY(l_PPRNPHeight, l_PPRNPSMode, PPRNSMMillimeters) * 1000)
    
    l_PPRN_GetPaperSizeNo = GetFormName(PrinterHandle, PaperSize, "")
    
    ClosePrinter (PrinterHandle)
    
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function PPRN_SetFont(ByVal PPRNFName As String,ByVal  PPRNFSize As Single,Optional ByVal PPRNFStyle As PPRNFontStyleConst = -999,Optional ByVal PPRNTxtCtrl As PPRNTextControlConst = -999) As Boolean
  
  Try
  
  PPRN_SetFont = False

  l_PPRNFName = PPRNFName
  l_PPRNFSize = PPRNFSize
  
  If PPRNFStyle <> -999 Then
    l_PPRNFStyle = PPRNFStyle
  End If
  
  If PPRNTxtCtrl <> -999 Then
    l_PPRNTxtCtrl = PPRNTxtCtrl
  End If
  
  If l_PPRN_SetFontToPrint(PPRNFName, PPRNFSize, PPRNFStyle) = False Then
 GoTo ExitHere
End If
  
  PPRN_SetFont = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function PPRN_SetPaper(ByVal PPRNPSize As PPRNPaperSizeConst,Optional ByVal PPRNSMode As PPRNScaleModeConst = -999,Optional ByVal PPRNPWidth As Single = -999,Optional ByVal PPRNPHeight As Single = -999,Optional ByVal PPRNOrient As PPRNOrientConst = -999,Optional ByVal PPRNQuality As PPRNPrintQualityConst = -999,Optional ByVal PPRNPBin As PPRNPaperBinConst = -999) As Boolean
  
  '-999 is Nothing

  Try
  
  PPRN_SetPaper = False

  l_PPRNPSize = PPRNPSize
  
  If PPRNSMode <> -999 Then
    l_PPRNPSMode = PPRNSMode
  End If
  
  If PPRNPWidth <> -999 Then
    l_PPRNPWidth = PPRNPWidth
    If PPRNPBin = -999 Then
 l_PPRNPSize = PPRNPSUser
End If
  End If
  
  If PPRNPHeight <> -999 Then
    l_PPRNPHeight = PPRNPHeight
    If PPRNPBin = -999 Then
 l_PPRNPSize = PPRNPSUser
End If
  End If
  
  If PPRNOrient <> -999 Then
    l_PPRNPOrient = PPRNOrient
  End If
  
  If PPRNQuality <> -999 Then
    l_PPRNPQuality = PPRNQuality
  End If
  
  If PPRNPBin <> -999 Then
    l_PPRNPaperBin = PPRNPBin
  End If
  
  'Set to Printer
  If l_PPRNFlagSetPaper = False Then
    If l_PPRN_SetPaperToPrint = False Then
 GoTo ExitHere
End If
  End If
  
  PPRN_SetPaper = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function PPRN_PrintTextInBox(ByVal X_Pos As Single,ByVal  Y_Pos As Single,ByVal  strText As String,ByVal  Box_Height As Double,ByVal  Box_Width As Double,ByVal  BoxTxtAlign As PPRNAlignConst,Optional ByVal TextCtrl As PPRNTextControlConst,Optional ByVal PPRNFName As String,Optional ByVal PPRNFSize As Single,Optional ByVal PPRNFStyle As PPRNFontStyleConst) As Boolean




  Try
  
  PPRN_PrintTextInBox = False
  
  If l_PPRN_SetFontToPrint(PPRNFName, PPRNFSize, PPRNFStyle) = False Then
 GoTo ExitHere
End If
  
  Printer.CurrentX = X_Pos
  Printer.CurrentY = y_Pos
  Printer.Print strText
  
  PPRN_PrintTextInBox = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call PPRN_KillDocument

Finally

End Try
End Function

Private Sub Class_Initialize()
' On Error Resume Next
  
  l_PPRNFName = "Arial"
  l_PPRNFSize = 11
  l_PPRNFStyle = PPRNFSNone
  l_PPRNTxtCtrl = PPRNTCNone
  
  l_PPRNPSize = PPRNPSA4
  l_PPRNPSMode = PPRNSMMillimeters
  l_PPRNPWidth = 210  ' mm
  l_PPRNPHeight = 297 ' mm
  
  l_PPRNPOrient = PPRNORPortrait
  l_PPRNPQuality = PPRNPQDraft
  
  l_PPRNMarTopAdj = 0
  l_PPRNMarLeftAdj = 0
  
  l_PPRNPaperBin = PPRNBNAuto
  l_PPRNPagePerHeight = 0
  
  Call l_PPRN_SetPaperToPrint
  l_PPRNFlagSetPaper = False
  
End Sub

Private Sub Class_Terminate()
' On Error Resume Next
  l_PPRNFlagSetPaper = False
End Sub

