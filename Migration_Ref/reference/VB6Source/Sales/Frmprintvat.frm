VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frmprintvat 
   Caption         =   "Print Screen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   1050
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   8655
         Picture         =   "Frmprintvat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   165
         Width           =   885
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   480
         Left            =   7485
         Picture         =   "Frmprintvat.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Text File"
         Height          =   495
         Left            =   6270
         Picture         =   "Frmprintvat.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdSprint 
         Caption         =   "Smart Print"
         Height          =   495
         Left            =   3645
         Picture         =   "Frmprintvat.frx":06F6
         TabIndex        =   5
         Top             =   180
         Width           =   1065
      End
      Begin VB.CommandButton cmdPPrint 
         Caption         =   "Print"
         Height          =   495
         Left            =   2460
         Picture         =   "Frmprintvat.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   1350
         Picture         =   "Frmprintvat.frx":0E62
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   495
         Left            =   225
         Picture         =   "Frmprintvat.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdHtml 
         Caption         =   "HTML"
         Height          =   495
         Left            =   5055
         TabIndex        =   1
         Top             =   180
         Width           =   885
      End
   End
   Begin MSComDlg.CommonDialog printdlg 
      Left            =   5700
      Top             =   4065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmdPrint 
      Left            =   525
      Top             =   2205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
Begin AceSpread sprPrintvat
      Height          =   7425
      Left            =   0
      TabIndex        =   9
      Top             =   810
      Width           =   11805
      _Version        =   458752
      _ExtentX        =   20823
      _ExtentY        =   13097
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   499
      MaxRows         =   496
      SelectBlockOptions=   0
      VisibleCols     =   499
      VisibleRows     =   496
   End
End
Attribute VB_Name = "Frmprintvat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcel_Click()
  Dim X As Boolean
  Dim temp As Variant
  
  
  Try
  Me.MousePointer = vbHourglass

  temp = Year(ServerDate) & Month(ServerDate) & Day(ServerDate) & ".xls"
  X = Frmprintvat.sprPrintvat.ExportToExcel(App.Path & "\" & temp, "Test Sheet 1", "C:Program Files\Spread30\Samples\LOGFILE.TXT")
  ' Display result to user based on true/false value of x
  If X = True Then
    MsgBox "Export complete.", , "Result"
  Else
    MsgBox "Export did not succeed.", , "Result"
  End If
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

Private Sub cmdPPrint_Click()
Try
  Me.MousePointer = vbHourglass

  sprPrintvat.BorderStyle = 0
  cmdPrint.Flags = &H2&
  cmdPrint.Min = 1
  cmdPrint.Max = totalpages
  cmdPrint.CancelError = True
  cmdPrint.FromPage = 1
  Printer.PaperSize = vbPRPSA4
  cmdPrint.ToPage = totalpages
  cmdPrint.Orientation = cdlPortrait
  printdlg.ShowPrinter
  printdlg.Orientation = cdlPortrait
  If cmdPrint.Flags = &H0& Then
    sprPrintvat.PrintPageStart = cmdPrint.FromPage
    sprPrintvat.PrintPageEnd = cmdPrint.ToPage
    cmdPrint.Flags = &H4&
  Else
    sprPrintvat.PrintPageStart = cmdPrint.FromPage
    sprPrintvat.PrintPageEnd = cmdPrint.ToPage
  End If

  'sprPrintvat.PrintMarginLeft = 1240
  'sprPrintvat.PrintMarginTop = 1300
  sprPrintvat.PrintBorder = False
  sprPrintvat.PrintType = 3
  sprPrintvat.Col = 1
sprPrintvat.Col2 = sprPrintvat.MaxCols

  sprPrintvat.Row = 1
sprPrintvat.Row2 = sprPrintvat.MaxRows
  sprPrintvat.BlockMode = True
  sprPrintvat.PrintColor = True

  For Copies = 1 To cmdPrint.Copies
      sprPrintvat.Action = 13
  Next
  sprPrintvat.BlockMode = False
  sShowMessage "Print Job Over "
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprPrintvat(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprPrintvat

End Sub

Private Sub InitializeSpreads()

	Call Design_sprPrintvat("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\Frmprintvat.frm", sprPrintvat, "")

End Sub
