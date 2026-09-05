VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmReportViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Viewer"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19815
   ControlBox      =   0   'False
   Icon            =   "ReportViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleMode       =   0  'User
   ScaleTop        =   45
   ScaleWidth      =   19815
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraToolbar 
      BackColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   30
      TabIndex        =   1
      Top             =   -75
      Width           =   9855
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7545
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
      Begin VB.CommandButton cmdPPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6510
         TabIndex        =   9
         Top             =   165
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   645
         Left            =   2310
         TabIndex        =   8
         Top             =   0
         Width           =   30
      End
      Begin VB.TextBox txtTotalPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtToPageNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   7
         Top             =   240
         Width           =   570
      End
      Begin VB.TextBox txtFromPageNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3330
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5475
         TabIndex        =   12
         Top             =   165
         Width           =   1050
      End
      Begin VB.Label lbSheet 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "80 Col. Paper"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8730
         TabIndex        =   11
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblTotalPagesCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. Of Pages:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   2
         Top             =   255
         Width           =   1530
      End
      Begin VB.Label lblToPageCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Page:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3990
         TabIndex        =   6
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblFromPageCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Page:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2370
         TabIndex        =   4
         Top             =   270
         Width           =   930
      End
   End
   Begin SHDocVwCtl.WebBrowser wbrView 
      Height          =   4890
      Left            =   15
      TabIndex        =   0
      Top             =   585
      Width           =   9855
      ExtentX         =   17383
      ExtentY         =   8625
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdPPrint_Click()
  Dim iFileTag As Integer
  Dim strText As String
  Dim blnPageFound As Boolean
  Dim dlgPrinter As CommonDialog

  On Error GoTo Err_Handler
  
  blnPageFound = False
  iFileTag = FreeFile
  
  'TAG property stores the File Name
  Open Me.Tag For Input As #iFileTag
  Do While Not EOF(iFileTag)
    Input #iFileTag, strText
    If InStr(1, strText, "~~" & CStr(txtFromPageNo.Text) & "~~") <> 0 Then
      blnPageFound = True
      Set dlgPrinter = Me.dlgMain
      dlgPrinter.PrinterDefault = True
      dlgPrinter.CancelError = True
      dlgPrinter.Min = Val(txtFromPageNo)
      dlgPrinter.Max = Val(txtToPageNo)
      dlgPrinter.FromPage = Val(txtFromPageNo)
      dlgPrinter.ToPage = Val(txtToPageNo)
      dlgPrinter.ShowPrinter
      Printer.PrintQuality = -1
      Printer.FontName = "Draft 10cpi"
      
      If Printer.FontName <> "Draft 10cpi" Then
        Printer.FontName = "Roman 10cpi"
        If Printer.FontName <> "Roman 10cpi" Then
          Printer.FontName = "Courier New"
          Printer.FontSize = 10
        End If
      End If
      
      If pPaper = vbPRPSFanfoldStdGerman And Orient = 1 Then
        Printer.PaperSize = vbPRPSA4
'        Printer.PaperSize = vbPRPSFanfoldStdGerman
'        Printer.Width = 10 * 1440
'        Printer.Height = 12 * 1440
      ElseIf pPaper = vbPRPSFanfoldUS And Orient = 1 Then
        Printer.PaperSize = vbPRPSFanfoldUS
'        Printer.Width = 15 * 1440
'        Printer.Height = 12 * 1440
      Else
        Printer.PaperSize = pPaper
        Printer.Orientation = Orient
      End If
      
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      Printer.Print
      
      Do While Not EOF(iFileTag)
        Line Input #iFileTag, strText
        If InStr(1, strText, "~~" & CStr(Val(txtToPageNo.Text) + 1) & "~~") <> 0 Then

          Exit Do
        ElseIf InStr(1, strText, "") <> 0 Then
          Printer.NewPage
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.Print
        ElseIf InStr(1, strText, "~~") = 0 Then
          Printer.Print Tab(3); strText
        End If
      Loop
      
      Printer.EndDoc
      MsgBox "Print Job Over", vbInformation + vbOKOnly
      Exit Do
    End If
  Loop
  
  Close #iFileTag
  If Not blnPageFound Then
    MsgBox "The required page could not be located"
  End If
  wbrView.SetFocus
  Exit Sub

Err_Handler:
  Close #iFileTag
  If Err.Number = 55 Then 'File already open error - ignore it
    
  ElseIf Err.Number = cdlCancel Then
    MsgBox "Printing operation was cancelled by the user."
    wbrView.SetFocus
  End If

End Sub

Private Sub CmdSave_Click()
  Dim dlgSave As CommonDialog
  Dim retVal As String
  
  Try
  Me.MousePointer = vbHourglass

  Set dlgSave = Me.dlgMain
  dlgSave.CancelError = True
  dlgSave.FileName = "Report-" & Format(ServerDate, "yyyyMMdd") & ".txt"
  dlgSave.Filter = "Text (*.txt)|*.txt"
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
  FileCopy Me.Tag, dlgSave.FileName
Catch ex As Exception
  If Err.Number = 70 Then
    sShowMessage "File has been opened by some other user. Pls close it before save."

  ElseIf Err.Number = 53 Then
    sShowMessage "Source File not found."
  Else
    Call ErrMsgShow(ex)
  End If

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
  
  Try
  
  Select Case pPaper
    Case vbPRPSA4
      lbSheet.Caption = "A4 Paper"
    Case vbPRPSFanfoldStdGerman
      lbSheet.Caption = "80 Col. Paper"
    Case vbPRPSFanfoldUS
      lbSheet.Caption = "132 Col. Paper"
    Case Else
      lbSheet.Caption = "Custome Size"
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  
' On Error Resume Next
  
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    Kill strFilename
  End If
End Sub

Private Sub Form_Resize()
  
  Try
  
    wbrView.Left = 15
    wbrView.Top = 645 '630 + 15 add the height of the toolbar
    If Me.ScaleHeight - 645 > 0 Then
        wbrView.Height = Me.ScaleHeight - 1030 '(630 + 30) + 370 Toolbar height+taskbar height
    End If
    If Me.ScaleWidth - 30 > 0 Then
        wbrView.Width = Me.ScaleWidth - 30
    End If

    fraToolbar.Left = (Me.ScaleWidth / 2 - fraToolbar.Width / 2) + 15

    txtFromPageNo.Text = "1"
    txtToPageNo.Text = txtTotalPages.Text
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtFromPageNo_KeyPress(ByVal KeyAscii As Integer)
  
' On Error Resume Next
  
  If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8)) And ( KeyAscii <> 46) Then

      KeyAscii = 0
  End If
End Sub

Private Sub txtToPageNo_KeyPress(ByVal KeyAscii As Integer)
   
' On Error Resume Next
  
  If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8)) And ( KeyAscii <> 46) Then

      KeyAscii = 0
  End If
End Sub

