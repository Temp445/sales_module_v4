VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrint 
   Caption         =   "Printing Screen"
   ClientHeight    =   8190
   ClientLeft      =   1710
   ClientTop       =   45
   ClientWidth     =   11880
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
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   675
      Left            =   1890
      Picture         =   "Print.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   960
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   675
      Left            =   945
      Picture         =   "Print.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
   Begin VB.CommandButton cmdPPrint 
      Caption         =   "Print"
      Height          =   675
      Left            =   0
      Picture         =   "Print.frx":05F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   960
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   525
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
Begin AceSpread sprPrint
      Height          =   7125
      Left            =   15
      TabIndex        =   0
      Top             =   660
      Width           =   11955
      _Version        =   458752
      _ExtentX        =   21087
      _ExtentY        =   12568
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      SelectBlockOptions=   0
      VisibleCols     =   500
      VisibleRows     =   500
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Menu mnMain 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Save As"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Clear"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Print"
         Index           =   4
      End
      Begin VB.Menu mnFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Copies As Integer

Private Enum ePaper 'Page Height
  mPaperA4 = 63  'For A4 Sheet
  mPaper80Col = 75 'For 80 Col Paper
End Enum

Private Sub cmdDoc_Click()
  WordPorting
End Sub

Private Sub cmdDOS_Click()
  'dos_print1
  Dim cols As Integer
  Dim L As Integer
  Dim i As Long
  Dim k As Long
  cols = 0
  L = 0
  Printer.Font.Name = "Draft 10cpi"
  Printer.Orientation = vbPRORPortrait
  Printer.PaperSize = vbPRPSFanfoldUS
  Printer.PrintQuality = vbPRPQDraft
    For i = 1 To sprPrint.DataRowCnt
      sprPrint.Row = i
      For k = 1 To sprPrint.MaxCols
        sprPrint.Col = k
        If sprPrint.Text <> "" Then
          If k = 1 Then
            Printer.Print ; Tab(1); sprPrint.Text;
          ElseIf k = sprPrint.MaxCols Then
            Printer.Print ; Tab(cols); sprPrint.Text
          Else
            If IsNumeric(sprPrint.Text) = False Then
              If sprPrint.TypeHAlign = TypeHAlignLeft Then
                Printer.Print ; Tab(cols); sprPrint.Text;
              ElseIf sprPrint.TypeHAlign = TypeHAlignRight Then
                Printer.Print ; Tab(cols + sprPrint.ColWidth(k) - Len(sprPrint.Text)); sprPrint.Text;
              End If
            Else
              If sprPrint.TypeHAlign = TypeHAlignLeft Then
                Printer.Print ; Tab(cols); sprPrint.Text;
              Else
                Printer.Print ; Tab(cols + sprPrint.ColWidth(k) - Len(Format(sprPrint.Text, "0.00"))); sprPrint.Text;
              End If
            End If
          End If
          cols = cols + sprPrint.ColWidth(k) + 3
        Else
          cols = cols + sprPrint.ColWidth(k) + 3
          L = L + 1
        End If
      Next k
      cols = 0
      If L = sprPrint.MaxCols Then
        Printer.Print ; " "
      End If
      L = 0
      If i > 6 Then
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
      End If
    Next i
  Printer.EndDoc

End Sub

Private Sub cmdExcel_Click()
  Dim X As Boolean
  Dim Temp As Variant
  Dim Temp1 As Variant
  Dim CurrDate As Date

  Dim dlgSave As CommonDialog
  Dim retVal As String
  
  Try
  Me.MousePointer = vbHourglass

  CurrDate = ServerDate
  Temp = Year(CurrDate) & Month(CurrDate) & Day(CurrDate)

  Set dlgSave = Me.dlgMain
  dlgSave.CancelError = True
  dlgSave.FileName = Temp
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
  
  ' ExcelPorting
  X = sprPrint.ExportExcelBook((dlgSave.FileName), "")  'Without Log File
  
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

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOpen_Click()
  OpenSpreadFile
End Sub

'Private Sub cmdPPrint_Click()
'  SprPrint.BorderStyle = 0
'  On Error GoTo ErrHand

''  '  dlgMain.FLAGS = &H2&
''  '
''  '  dlgMain.Min = 1
''  '  dlgMain.Max = totalpages
''  '  dlgMain.CancelError = True
''  '  dlgMain.FromPage = 1
''  '  dlgMain.ToPage = totalpages
''  '  If Orient = 1 Then
''  '    dlgMain.Orientation = cdlPortrait
''  '  ElseIf Orient = 2 Then
''  '    dlgMain.Orientation = cdlLandscape
''  '  End If
''  '  dlgMain.ShowSave   ' .ShowPrinter
'''    If dlgMain.FLAGS = &H0& Then
'''      sprPrint.PrintPageStart = dlgMain.FromPage
'''      sprPrint.PrintPageEnd = dlgMain.ToPage
'''      dlgMain.FLAGS = &H4&
'''    Else
'''      sprPrint.PrintPageStart = dlgMain.FromPage
'''      sprPrint.PrintPageEnd = dlgMain.ToPage
'''    End If
'''
'''    If Orient = 1 Then
'''      sprPrint.PrintMarginLeft = 1240
'''      sprPrint.PrintMarginTop = 1300
'''    ElseIf Orient = 2 Then
'''      sprPrint.PrintMarginLeft = 1150
'''      sprPrint.PrintMarginTop = 1300
'''    End If
'''    sprPrint.PrintBorder = True
'''    sprPrint.PrintType = 3
'''    sprPrint.Col = 1: sprPrint.Col2 = sprPrint.MaxCols

'''    sprPrint.Row = 1: sprPrint.Row2 = sprPrint.MaxRows
'''    sprPrint.BlockMode = True
'''    sprPrint.PrintColor = True
'''
'''    For Copies = 1 To dlgMain.Copies
'''      If Orient = 1 Then
'''        sprPrint.Action = 13
'''      ElseIf Orient = 2 Then
'''        sprPrint.Action = ActionSmartPrint
'''      End If
'''    Next
'
'''    sprPrint.BlockMode = False
'''    sShowMessage "Print Job Over "
'''    Exit Sub
'
'ErrHand:
'  If Err.Number = 32755 Then Exit Sub
'  MsgBox Err.Number & "  " & Err.Description
'End Sub

Private Sub CmdSave_Click()
  SaveSpreadFile
End Sub

Private Sub cmdSaveAs_Click()
  SaveSpreadFile
End Sub

Private Sub cmdPPrint_Click()
  Dim i As Long
  
  On Error GoTo ErrHand
  sprPrint.BorderStyle = 0
  
    dlgMain.Flags = &H2&
    
    dlgMain.Min = 1
    dlgMain.Max = totalpages
    dlgMain.CancelError = True
    dlgMain.FromPage = 1
    dlgMain.ToPage = totalpages
    dlgMain.Orientation = cdlLandscape
    
    dlgMain.ShowPrinter      'Show Printer Dialogue Box
    
    If dlgMain.Flags = &H0& Then
      sprPrint.PrintPageStart = dlgMain.FromPage
      sprPrint.PrintPageEnd = dlgMain.ToPage
      dlgMain.Flags = &H4&
    Else
      sprPrint.PrintPageStart = dlgMain.FromPage
      sprPrint.PrintPageEnd = dlgMain.ToPage
    End If  
    sprPrint.PrintMarginLeft = 1240
    sprPrint.PrintMarginTop = 1300
    sprPrint.PrintBorder = False
    sprPrint.PrintType = 3
    
    sprPrint.Col = 1
sprPrint.Col2 = sprPrint.MaxCols

    sprPrint.Row = 1
sprPrint.Row2 = sprPrint.MaxRows
    
    sprPrint.BlockMode = True
      sprPrint.PrintColor = True
    
      For Copies = 1 To dlgMain.Copies    'Copies
        sprPrint.Action = ActionSmartPrint
      Next Copies
    
    sprPrint.BlockMode = False
    sShowMessage "Print Job Over "
  
Finally
  On Error GoTo 0
  Exit Sub
ErrHand:
  If Err.Number = 32755 Then Exit Sub
  MsgBox Err.Number & "  " & Err.Description
End Sub

Private Sub cmdSprint_Click()
  sprPrint.BorderStyle = 0
  On Error GoTo ErrHand
  dlgMain.Flags = &H2&

  dlgMain.Min = 1
  dlgMain.Max = totalpages
  dlgMain.CancelError = True
  dlgMain.FromPage = 1
  dlgMain.ToPage = totalpages
  dlgMain.Orientation = cdlLandscape
  dlgMain.ShowPrinter
  If dlgMain.Flags = &H0& Then
    sprPrint.PrintPageStart = dlgMain.FromPage
    sprPrint.PrintPageEnd = dlgMain.ToPage
    dlgMain.Flags = &H4&
  Else
    sprPrint.PrintPageStart = dlgMain.FromPage
    sprPrint.PrintPageEnd = dlgMain.ToPage
  End If

  sprPrint.PrintMarginLeft = 1240
  sprPrint.PrintMarginTop = 1300
  sprPrint.PrintBorder = False
  sprPrint.PrintType = 3
  sprPrint.Col = 1
sprPrint.Col2 = sprPrint.MaxCols

  sprPrint.Row = 1
sprPrint.Row2 = sprPrint.MaxRows
  sprPrint.BlockMode = True
  sprPrint.PrintColor = True

  For Copies = 1 To dlgMain.Copies
      sprPrint.Action = ActionSmartPrint
  Next
  sprPrint.BlockMode = False
  sShowMessage "Print Job Over "
  Exit Sub

ErrHand:
  If Err.Number = 32755 Then Exit Sub
  MsgBox Err.Number & "  " & Err.Description

End Sub

Private Sub cmdText_Click()
  TextPorting
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

' On Error Resume Next
  
  iPageNo = 1
  If False Then
      sprPrint.Row = 1
      sprPrint.Row2 = sprPrint.MaxRows
      sprPrint.Col = 1
      sprPrint.Col2 = sprPrint.MaxCols
      sprPrint.BlockMode = True
      sprPrint.Lock = True
      sprPrint.BlockMode = False
  End If
  
  
  
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub mnFile_Click(ByVal Index As Integer)
  Select Case Index
    Case 0
        OpenSpreadFile
    Case 1
        SaveSpreadFile
    Case 2
        SaveSpreadFile
    Case 3
        ClearSpread
    Case 4
        sprPrint.BorderStyle = 0
        PrintSpreadInfo
    Case 5
        ClearSpread
        Unload Me
  End Select
End Sub

' To Clear the Spread
Private Function ClearSpread()
' On Error Resume Next
  sprPrint.Col = 1
sprPrint.Col2 = sprPrint.MaxCols

  sprPrint.Row = 1
sprPrint.Row2 = sprPrint.MaxRows
  
  sprPrint.Action = 3

End Function

' To Open a Particular Spread File
Private Sub OpenSpreadFile()

  On Error GoTo ErrHand
  
  dlgMain.DialogTitle = " Choose the Spread Sheet File"
  dlgMain.Filter = "*.SS2"
  dlgMain.CancelError = True
  dlgMain.ShowOpen
  
  If Not sprPrint.LoadFromFile(dlgMain.FileName) Then
      MsgBox " Could not Open", vbInformation, App.Title
      Exit Sub
  End If
  
  frmPrint.mnFile(1).Enabled = True
  frmPrint.mnFile(2).Enabled = True
  frmPrint.mnFile(3).Enabled = True
  
  Exit Sub
  
ErrHand:
  If Err.Number = 32755 Then
    Exit Sub
  Else
    MsgBox Err.Number & "  " & Err.Description, vbCritical, App.Title
  End If
End Sub

' To print the Spread Information
Private Sub PrintSpreadInfo()

End Sub

' To Save the Spread Sheet information into a File
Private Sub SaveSpreadFile()

  On Error GoTo ErrHand
  
  dlgMain.DialogTitle = " File Name to Save"
  dlgMain.ShowSave
  dlgMain.CancelError = True
  sprPrint.SaveToFile dlgMain.FileName & ".ss2", False
  frmPrint.mnFile(1).Enabled = False
  
  Exit Sub
  
ErrHand:
  If Err.Number = 32755 Then Exit Sub
  MsgBox Err.Number & "  " & Err.Description, vbCritical, App.Title
End Sub

Private Sub TextPorting()

  On Error GoTo ErrHand
  
  dlgMain.Filter = "Text File (*.txt)|*.txt"
  dlgMain.ShowSave
  dlgMain.CancelError = True
  
  If Dir$(dlgMain.FileName) <> "" Then
      If MsgBox("Warning! File Already Exists", vbYesNo + vbCritical, "Purchase") = vbNo Then Exit Sub
  End If
  
  sprPrint.SaveTabFile dlgMain.FileName & ".txt"
  frmPrint.mnFile(1).Enabled = False
  Exit Sub
  
ErrHand:
  If Err.Number <> 32755 Then
    MsgBox Err.Number & "   " & Err.Description, vbCritical, App.Title
  End If
End Sub

Private Sub WordPorting()

On Error GoTo ErrHand

Dim Wrd As Object
Dim  Col As 
Dim  Row As 
Dim  Val


Set Wrd = CreateObject("Word.Basic")

    Wrd.FileNewDefault
    Wrd.Insert ""
    Wrd.Insert Chr(13)
    Wrd.Insert ""
    For Row = 1 To sprPrint.DataRowCnt
        For Col = 1 To sprPrint.MaxCols
            sprPrint.GetText Col, Row, Val
            Wrd.Insert CStr(Val)
        Next
        Wrd.Insert Chr(13)
        sprPrint.Row = Row
        If sprPrint.RowPageBreak = True Then Wrd.NewPage
    Next
    Wrd.FileSaveAs App.Path & "\" & Month(Now) & Year(Now) & ".DOC"
    MsgBox "Data Ported in " & App.Path & "\" & Month(Now) & Year(Now) & ".DOC", vbInformation, App.Title

Exit Sub

ErrHand:
    If Err.Number = 429 Then
      MsgBox "MS-Word Not Found in the Machine", vbCritical, App.Title
    End If
End Sub

Private Sub Design_sprPrint(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprPrint

End Sub

Private Sub InitializeSpreads()

	Call Design_sprPrint("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\Print.frm", sprPrint, "")

End Sub
