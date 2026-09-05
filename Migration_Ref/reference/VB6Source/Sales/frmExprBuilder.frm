VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExprBuilder 
   Caption         =   "Purchase Invoice Formula Master"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10875
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstExpression 
      Height          =   6450
      Left            =   300
      TabIndex        =   31
      Top             =   330
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   11377
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Expression"
      TabPicture(0)   =   "frmExprBuilder.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSave"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClear"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClose"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFormula"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Expression Builder"
      TabPicture(1)   =   "frmExprBuilder.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdExprCancel"
      Tab(1).Control(1)=   "fraExprBuilder"
      Tab(1).Control(2)=   "cmdExprApply"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Select Expression"
      TabPicture(2)   =   "frmExprBuilder.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraExprSelect"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdSeqSelect"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSeqCancel"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdNext"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdNext 
         Caption         =   "Select Expression"
         Height          =   390
         Left            =   -67755
         TabIndex        =   58
         Top             =   5845
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdExprCancel 
         Caption         =   "Cancel"
         Height          =   390
         Left            =   -69345
         TabIndex        =   26
         ToolTipText     =   "Cancels the Current Expression"
         Top             =   5845
         Width           =   1170
      End
      Begin VB.CommandButton cmdSeqCancel 
         Caption         =   "Cancel"
         Height          =   390
         Left            =   -69345
         TabIndex        =   29
         Top             =   5845
         Width           =   1170
      End
      Begin VB.CommandButton cmdSeqSelect 
         Caption         =   "Select"
         Height          =   390
         Left            =   -70500
         TabIndex        =   28
         Top             =   5845
         Width           =   1170
      End
      Begin VB.Frame fraExprSelect 
         Height          =   5125
         Left            =   -71970
         TabIndex        =   43
         Top             =   525
         Width           =   5220
Begin AceSpread sprExprSelect
            Height          =   4840
            Left            =   90
            TabIndex        =   27
            Top             =   195
            Width           =   5040
            _Version        =   458752
            _ExtentX        =   8890
            _ExtentY        =   8537
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
            MaxCols         =   3
            MaxRows         =   1
            OperationMode   =   5
         End
      End
      Begin VB.Frame fraExprBuilder 
         Height          =   5125
         Left            =   -74835
         TabIndex        =   36
         Top             =   525
         Width           =   10890
         Begin VB.Frame fraExpr 
            Height          =   1395
            Left            =   585
            TabIndex        =   38
            Top             =   135
            Width           =   9585
            Begin VB.TextBox txtExpr 
               Appearance      =   0  'Flat
               Height          =   990
               Left            =   75
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   7
               Top             =   330
               Width           =   9420
            End
            Begin VB.Label lblExpr 
               AutoSize        =   -1  'True
               Caption         =   "Expression"
               Height          =   195
               Left            =   60
               TabIndex        =   39
               Top             =   120
               Width           =   765
            End
         End
         Begin VB.Frame fraParameters 
            Height          =   2905
            Left            =   3150
            TabIndex        =   41
            Top             =   1485
            Width           =   7030
Begin AceSpread sprParam
               Height          =   2685
               Left            =   105
               TabIndex        =   63
               ToolTipText     =   "Double Click to Select Paramter"
               Top             =   150
               Width           =   6825
               _Version        =   458752
               _ExtentX        =   12039
               _ExtentY        =   4736
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
               MaxCols         =   9
               MaxRows         =   1
            End
         End
         Begin VB.Frame fraNumbers 
            Height          =   1200
            Left            =   600
            TabIndex        =   37
            Top             =   1485
            Width           =   2520
            Begin VB.CommandButton cmdNumber 
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   0
               Left            =   75
               TabIndex        =   8
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   1
               Left            =   555
               TabIndex        =   9
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   2
               Left            =   1035
               TabIndex        =   10
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   3
               Left            =   1515
               TabIndex        =   11
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   4
               Left            =   1995
               TabIndex        =   12
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   5
               Left            =   75
               TabIndex        =   13
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   6
               Left            =   555
               TabIndex        =   14
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   7
               Left            =   1035
               TabIndex        =   15
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   8
               Left            =   1515
               TabIndex        =   16
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   9
               Left            =   1995
               TabIndex        =   17
               Top             =   675
               Width           =   450
            End
         End
         Begin VB.Frame fraOperator 
            Height          =   1200
            Left            =   600
            TabIndex        =   40
            Top             =   2595
            Width           =   2520
            Begin VB.CommandButton cmdOprEdit 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   0
               Left            =   1995
               TabIndex        =   62
               ToolTipText     =   "All Clear"
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdOprEdit 
               Caption         =   "C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   1
               Left            =   1515
               TabIndex        =   61
               ToolTipText     =   "Cancel Last Operation"
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   3
               Left            =   1995
               TabIndex        =   60
               ToolTipText     =   "Divide"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "�"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   6
               Left            =   1035
               TabIndex        =   59
               ToolTipText     =   "% Expression Block Close"
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   0
               Left            =   555
               TabIndex        =   18
               ToolTipText     =   "Add"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   1
               Left            =   1035
               TabIndex        =   19
               ToolTipText     =   "subtract"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   2
               Left            =   1515
               TabIndex        =   20
               ToolTipText     =   "Multiply"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   "("
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   4
               Left            =   75
               TabIndex        =   21
               ToolTipText     =   "Open Paranthesis"
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdOpr 
               Caption         =   ")"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   5
               Left            =   555
               TabIndex        =   22
               ToolTipText     =   "Close Paranthesis"
               Top             =   675
               Width           =   450
            End
            Begin VB.CommandButton cmdNumber 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   10
               Left            =   75
               TabIndex        =   23
               Top             =   165
               Width           =   450
            End
         End
         Begin VB.Frame fraStep 
            Height          =   690
            Left            =   600
            TabIndex        =   42
            Top             =   3705
            Width           =   2520
            Begin VB.CommandButton cmdFormula 
               Caption         =   "Formula"
               Enabled         =   0   'False
               Height          =   450
               Left            =   1290
               TabIndex        =   57
               Top             =   165
               Width           =   1155
            End
            Begin VB.CommandButton cmdExprSeq 
               Caption         =   "Expr. Step"
               Enabled         =   0   'False
               Height          =   450
               Left            =   75
               TabIndex        =   24
               Top             =   165
               Width           =   1185
            End
         End
         Begin VB.Frame fraLogic 
            Height          =   690
            Left            =   600
            TabIndex        =   45
            Top             =   4320
            Width           =   9585
            Begin VB.CommandButton cmdLogic 
               Caption         =   "} ENDIF"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   3
               Left            =   3195
               TabIndex        =   56
               ToolTipText     =   "Logic If/Then/Else Block Close"
               Top             =   165
               Width           =   1005
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "NOT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   10
               Left            =   8775
               TabIndex        =   55
               ToolTipText     =   "Logical NOT"
               Top             =   165
               Width           =   675
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "OR"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   9
               Left            =   8070
               TabIndex        =   54
               ToolTipText     =   "Logical OR"
               Top             =   165
               Width           =   675
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "AND"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   8
               Left            =   7350
               TabIndex        =   53
               ToolTipText     =   "Logical AND"
               Top             =   165
               Width           =   675
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   7
               Left            =   6105
               TabIndex        =   52
               ToolTipText     =   "Lesser Than"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   6
               Left            =   5655
               TabIndex        =   51
               ToolTipText     =   "Greater Than"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "<>"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   5
               Left            =   5205
               TabIndex        =   50
               ToolTipText     =   "Not Equals"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   4
               Left            =   4755
               TabIndex        =   49
               ToolTipText     =   "Equals"
               Top             =   165
               Width           =   450
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "} ELSE {"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   2
               Left            =   2155
               TabIndex        =   48
               ToolTipText     =   "Logic Else"
               Top             =   165
               Width           =   1005
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "} THEN {"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   1
               Left            =   1115
               TabIndex        =   47
               ToolTipText     =   "Logic Then"
               Top             =   165
               Width           =   1005
            End
            Begin VB.CommandButton cmdLogic 
               Caption         =   "IF {"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   0
               Left            =   75
               TabIndex        =   46
               ToolTipText     =   "Logic IF"
               Top             =   165
               Width           =   1005
            End
         End
      End
      Begin VB.Frame fraFormula 
         Height          =   5205
         Left            =   165
         TabIndex        =   32
         Top             =   525
         Width           =   10890
         Begin VB.TextBox txtROffValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   6015
            MaxLength       =   30
            TabIndex        =   73
            Top             =   4815
            Width           =   1095
         End
         Begin VB.TextBox txtROffPosting 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            Left            =   7185
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   70
            Top             =   4815
            Width           =   3540
         End
         Begin VB.ComboBox cmbRoundOff 
            Height          =   315
            ItemData        =   "frmExprBuilder.frx":08A9
            Left            =   4110
            List            =   "frmExprBuilder.frx":08B9
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   4808
            Width           =   1830
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&Ok"
            Height          =   300
            Left            =   10275
            TabIndex        =   68
            Top             =   900
            Width           =   450
         End
         Begin VB.TextBox txtTemplate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3060
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   66
            Top             =   360
            Width           =   3150
         End
         Begin VB.ComboBox cmbCreateType 
            Height          =   315
            ItemData        =   "frmExprBuilder.frx":08DC
            Left            =   105
            List            =   "frmExprBuilder.frx":08E6
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   353
            Width           =   2955
         End
         Begin VB.TextBox txtFSName 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   0
            Top             =   360
            Width           =   960
         End
         Begin VB.TextBox txtFLName 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7215
            MaxLength       =   30
            TabIndex        =   1
            Top             =   360
            Width           =   3510
         End
         Begin VB.TextBox txtFDesc 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   105
            MaxLength       =   100
            TabIndex        =   2
            Top             =   900
            Width           =   10095
         End
Begin AceSpread sprExprBuilder
            Height          =   3285
            Left            =   105
            TabIndex        =   3
            Top             =   1290
            Width           =   10620
            _Version        =   458752
            _ExtentX        =   18732
            _ExtentY        =   5794
            _StockProps     =   64
            Enabled         =   0   'False
            BackColorStyle  =   1
            ButtonDrawMode  =   4
            ColsFrozen      =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   10
            MaxRows         =   1
         End
         Begin VB.Label lblFrmTotal 
            AutoSize        =   -1  'True
            Caption         =   "Rounding Off Option for Formula Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   75
            Top             =   4845
            Width           =   3915
         End
         Begin VB.Label lblROffValue 
            AutoSize        =   -1  'True
            Caption         =   "R. Off Value"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6015
            TabIndex        =   74
            Top             =   4635
            Width           =   870
         End
         Begin VB.Label lblROffOption 
            AutoSize        =   -1  'True
            Caption         =   "Rounding Off Options"
            Height          =   195
            Left            =   4110
            TabIndex        =   72
            Top             =   4635
            Width           =   1530
         End
         Begin VB.Label lblROffPosting 
            AutoSize        =   -1  'True
            Caption         =   "Rounding Off Posting Account"
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   7185
            TabIndex        =   71
            Top             =   4635
            Width           =   2160
         End
         Begin VB.Label lblTemplate 
            AutoSize        =   -1  'True
            Caption         =   "Template Formula"
            Height          =   195
            Left            =   3060
            TabIndex        =   67
            Top             =   150
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Creation Type"
            Height          =   195
            Left            =   105
            TabIndex        =   65
            Top             =   150
            Width           =   990
         End
         Begin VB.Label lblFSName 
            AutoSize        =   -1  'True
            Caption         =   "Short Name"
            Height          =   195
            Left            =   6240
            TabIndex        =   35
            Top             =   150
            Width           =   840
         End
         Begin VB.Label lblFLName 
            AutoSize        =   -1  'True
            Caption         =   "Long Name"
            Height          =   195
            Left            =   7215
            TabIndex        =   34
            Top             =   150
            Width           =   825
         End
         Begin VB.Label lblFDesc 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   105
            TabIndex        =   33
            Top             =   690
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cl&ose"
         Height          =   390
         Left            =   6180
         TabIndex        =   6
         Top             =   5845
         Width           =   1170
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   390
         Left            =   5025
         TabIndex        =   5
         Top             =   5845
         Width           =   1170
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   390
         Left            =   3870
         TabIndex        =   4
         Top             =   5845
         Width           =   1170
      End
      Begin VB.CommandButton cmdExprApply 
         Caption         =   "Apply "
         Height          =   390
         Left            =   -70500
         TabIndex        =   25
         ToolTipText     =   "Update the Current Expression"
         Top             =   5845
         Width           =   1170
      End
   End
Begin AceSpread sprExprBuildSeq
      Height          =   1530
      Left            =   300
      TabIndex        =   30
      Top             =   6840
      Visible         =   0   'False
      Width           =   5220
      _Version        =   458752
      _ExtentX        =   9208
      _ExtentY        =   2699
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
      MaxCols         =   10
      ScrollBars      =   2
      ClipboardOptions=   0
   End
Begin AceSpread sprCurExprBuildSeq
      Height          =   1530
      Left            =   5595
      TabIndex        =   44
      Top             =   6840
      Visible         =   0   'False
      Width           =   5955
      _Version        =   458752
      _ExtentX        =   10504
      _ExtentY        =   2699
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
      MaxCols         =   11
      ScrollBars      =   2
      ClipboardOptions=   0
   End
End
Attribute VB_Name = "frmExprBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const PercBlockOpen As String = "�"
Private Const PercBlockClose As String = "�"

Public isCallCreate As Boolean    'For identify whether it was called as Create / Modify

'For Operator Button Index
Private Enum eOperator
  mAdd = 0
  mSubtract = 1
  mMultiply = 2
  mDivide = 3
  
  mOpenBrace = 4
  mCloseBrace = 5
  
  mPercClose = 6
End Enum

'For Logic Button Index
Private Enum eLogic
  mIF = 0
  mTHEN = 1
  mELSE = 2
  mENDIF = 3
  
  mEqual = 4
  mNotEqual = 5
  mGreater = 6
  mLesser = 7
  
  mAND = 8
  mOR = 9
  mNOT = 10
  
  mBETWEEN = 10
End Enum

'For Tab No
Private Enum eTab
  mExpr = 0
  mExprBuild = 1
  mSelect = 2
End Enum

'For ExprBuilder Spread Col No
Private Enum eExBuild
  mExRoute = 1
  mExName = 2
  mExBuildBut = 3
  mExpr = 4
  mRndOff = 5
  mRndOffVal = 6
  mGrand = 7
  mAcCode = 8
  mAcName = 9
  mRemarks = 10
End Enum

'For Parameter Spread
Private Enum eParam
  mParamID = 1
  mSName = 2
  mLName = 3
  mDesc = 4
  mIsPerc = 5
  mIsEditable = 6
  mValinMethod = 7
  mDefined = 8
  mIsMulti = 9
End Enum

'For Full Expr Build Spread Col No
Private Enum eExBuildSeq
  mExSeq = 1
  mBlockSeq = 2
  mBlockType = 3
  mBlock = 4
  mValinMethod = 5
  mDefined = 6
  mValUsed = 7
  mIsPerc = 8
  mIsTypeEditable = 9
  mIsValueEditable = 10
  mPrevCharEnd = 11
End Enum

'For Expr Selection Spread Col No
Private Enum eExSel
  mExName = 1
  mExpr = 2
  mExCode = 3
End Enum

Private Sub cmbCreateType_Click()

  Try
  Me.MousePointer = vbHourglass

  If cmbCreateType.ListIndex = 0 Then   'New Formula
    lblTemplate.ForeColor = vbBlack
    lblTemplate.Enabled = False
    txtTemplate.Enabled = False
  Else                                  'From Template - Create from Old Formula
    lblTemplate.ForeColor = vbBlue
    lblTemplate.Enabled = True
    txtTemplate.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbRoundOff_Change()
' On Error Resume Next
  Call cmbRoundOff_Click
End Sub

Private Sub cmbRoundOff_Click()
' On Error Resume Next
  If cmbRoundOff.Text = "None" Then
    txtROffValue.Text = ""
    txtROffValue.Enabled = False
    lblROffValue.Enabled = False
    lblROffPosting.Enabled = False
    txtROffPosting.Enabled = False
  Else
    txtROffValue.Enabled = True
    lblROffValue.Enabled = True
  End If
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim rCnt As Long
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass
  
  If Len(Trim$(txtFSName.Text)) <= 0 Then
    sShowMessage "Please Enter the Formula Short Name"
    GoTo ExitHere
  End If
  
  If Len(Trim$(txtFLName.Text)) <= 0 Then
    sShowMessage "Please Enter the Formula Long Name"
    GoTo ExitHere
  End If
  
  Set rs = New AceADODB.Recordset
  If cmbCreateType.ListIndex = 1 Then 'Formula from Template
    Sql = "Select a.ExprSeq, a.ExprRoute, Case When a.ExprRoute < 0 then 'SUB TOTAL' When a.ExprRoute='0001' Then 'TOTAL GST' else c.ExprGroupName End as ExprName," 'isNull(c.ExprGroupName, 'SUB TOTAL') as ExprName,
    Sql = Sql & " dbo.fn_FrontEndExpression(a.ModuleCode, a.FrmID, a.ExprSeq, a.BranchID, a.CompID) as Expression,"
    Sql = Sql & " a.RoundOff, a.RoundOffValue, a.GrandTotalFlag, a.PostingAc,"
    Sql = Sql & " Case a.PostingAc When 0 Then '' When -1 then 'As Per Order' Else d.PartyName End as PartyName, a.Remarks"
    Sql = Sql & " from COM_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group c"
    Sql = Sql & " On c.Route+c.Width=a.ExprRoute and c.ModuleCode=a.ModuleCode and c.BranchID=a.BranchID and c.CompID=a.CompID"
    Sql = Sql & " Left Outer Join FAS_Party_Master d On d.PartyCode=a.PostingAc and d.BranchID=a.BranchID and d.CompId=a.CompID"
    Sql = Sql & " Where a.FrmID=" & Val(txtTemplate.Tag)
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    Sql = Sql & " Order by a.ExprSeq"
'    Set rs = New ADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
        sprExprBuilder.MaxRows = rs.RecordCount + 1
        For rCnt = 1 To rs.RecordCount
          sprExprBuilder.SetText eExBuild.mExRoute, rCnt, rs("ExprRoute")
          sprExprBuilder.SetText eExBuild.mExName, rCnt, rs("ExprName")
          sprExprBuilder.SetText eExBuild.mExpr, rCnt, rs("Expression")
          sprExprBuilder.SetText eExBuild.mRndOff, rCnt, GetRoundingOffText(rs("RoundOff"))
          
          If rs("RoundOff") <> "N" Then
            sprExprBuilder.SetText eExBuild.mRndOffVal, rCnt, rs("RoundOffValue")
          End If
          
          Select Case rs("GrandTotalFlag")
            Case 0
              sprExprBuilder.SetText eExBuild.mGrand, rCnt, "No"
              Call PostingAcActive(False, rCnt)
            Case 1
              sprExprBuilder.SetText eExBuild.mGrand, rCnt, "Yes (Add)"
              Call PostingAcActive(True, rCnt)
            Case -1
              sprExprBuilder.SetText eExBuild.mGrand, rCnt, "Yes (Subtract)"
              Call PostingAcActive(True, rCnt)
          End Select
          sprExprBuilder.SetText eExBuild.mAcCode, rCnt, rs("PostingAc")
          sprExprBuilder.SetText eExBuild.mAcName, rCnt, rs("PartyName")
          sprExprBuilder.SetText eExBuild.mRemarks, rCnt, rs("Remarks")
    
          Call Spread_Lock(sprExprBuilder, True, eExBuild.mExBuildBut, rCnt)
          rs.MoveNext
        Next rCnt
        
        Call Spread_Lock(sprExprBuilder, False, eExBuild.mExBuildBut, sprExprBuilder.DataRowCnt)    
    
      Sql = "Select a.ExprSeq, a.BlockSeq, a.BlockType, a.Block,"
      Sql = Sql & " b.ValInputMethod, b.DefinedValues, a.ValueUsed, a.isTypePercentage, a.isTypeEditable, a.isValueEditable"
      Sql = Sql & " from (Select ModuleCode, ExprSeq, BlockSeq, BlockType, Block, ValueUsed,"
      Sql = Sql & " isTypePercentage, isTypeEditable, isValueEditable, BranchID, CompID"
      Sql = Sql & " from COM_FormulaMast_Block_Dtl Where FrmID=" & Val(txtTemplate.Tag) & " and BlockType='Param' and ModuleCode=" & gModuleCode
      Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "') a"
      Sql = Sql & " Left Outer Join COM_FrmParamMast_Hdr b On Convert(int, a.Block)=b.ParamID and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
      Sql = Sql & " Union All"
      Sql = Sql & " Select ExprSeq, BlockSeq, BlockType, Block, '', '', ValueUsed, isTypePercentage, isTypeEditable, isValueEditable"
      Sql = Sql & " From COM_FormulaMast_Block_Dtl Where FrmID=" & Val(txtTemplate.Tag) & " and BlockType<>'Param' and ModuleCode=" & gModuleCode
      Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
      Sql = Sql & " Order By ExprSeq, BlockSeq"
      If RSOpen(rs, Sql) = False Then
        GoTo ExitHere
      End If
    
      If Not rs.EOF Then
          For rCnt = 1 To rs.RecordCount
            sprExprBuildSeq.SetText eExBuildSeq.mExSeq, rCnt, rs("ExprSeq")
            sprExprBuildSeq.SetText eExBuildSeq.mBlockSeq, rCnt, rs("BlockSeq")
            sprExprBuildSeq.SetText eExBuildSeq.mBlockType, rCnt, rs("BlockType")
            sprExprBuildSeq.SetText eExBuildSeq.mBlock, rCnt, rs("Block")
            sprExprBuildSeq.SetText eExBuildSeq.mValinMethod, rCnt, rs("ValInputMethod")
            sprExprBuildSeq.SetText eExBuildSeq.mDefined, rCnt, rs("DefinedValues")
            sprExprBuildSeq.SetText eExBuildSeq.mValUsed, rCnt, rs("ValueUsed")
            sprExprBuildSeq.SetText eExBuildSeq.mIsPerc, rCnt, rs("isTypePercentage")
            sprExprBuildSeq.SetText eExBuildSeq.mIsTypeEditable, rCnt, rs("isTypeEditable")
            sprExprBuildSeq.SetText eExBuildSeq.mIsValueEditable, rCnt, rs("isValueEditable")
           
            rs.MoveNext
          Next rCnt
      End If
    End If
  End If
  
  sprExprBuilder.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
' On Error Resume Next
  
  If isCallCreate = True Then
      sstExpression.TabEnabled(eTab.mExpr) = True
      sstExpression.TabEnabled(eTab.mExprBuild) = False
      sstExpression.TabEnabled(eTab.mSelect) = False
      sstExpression.Tab = eTab.mExpr
    
    cmbCreateType.Enabled = True
    cmbCreateType.ListIndex = 0
    
    cmbRoundOff.ListIndex = 0
  Else
  
    cmbCreateType.Enabled = False
  End If
  Call LoadParamSpread
'  Call SetDefaultGV
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(frmExprBuilder, UnloadMode) = False Then
    Cancel = 1
  End If
End Sub

Private Sub sprExprBuilder_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim retVal As Variant
  
' On Error Resume Next
  'Build Button Press
  If Col = eExBuild.mExBuildBut Then
    If CheckExprNameSelect(Row - 1) = False Then
      Exit Sub
    End If
    
    'If Prev Expr Found then the Expr Seq button to be enabled for selection
    If sprExprBuilder.DataRowCnt > 1 Then
      cmdExprSeq.Enabled = True
    Else
      cmdExprSeq.Enabled = False
    End If
    
    sprExprBuilder.GetText eExBuild.mExName, Row, retVal
    lblExpr.Caption = "Expression for " & Trim$(retVal)
    
    Call SetActiveTab(eTab.mExpr, eTab.mExprBuild)
    
    'Load Expression for Edit
    Call LoadExpressionForEdit(CInt(Row))
  End If
End Sub

Private Sub sprExprBuilder_ComboSelChange(sender as object, e as ComboSelChangeEvent)
  
  Try
  Me.MousePointer = vbHourglass
  
    If Col = eExBuild.mRndOff Then
      If Trim$(GetSprText(sprExprBuilder, eExBuild.mRndOff, Row)) = "None" Then
        sprExprBuilder.SetText eExBuild.mRndOffVal, Row, ""
        Call Spread_Lock(sprExprBuilder, True, eExBuild.mRndOffVal, Row)
      Else
        Call Spread_Lock(sprExprBuilder, False, eExBuild.mRndOffVal, Row)
      End If
    ElseIf Col = eExBuild.mGrand Then
      If Trim$(GetSprText(sprExprBuilder, eExBuild.mGrand, Row)) = "No" Then
        Call PostingAcActive(False, Row)
      Else
        Call Spread_BkColour(sprExprBuilder, "&HFFFFFF", eExBuild.mAcName, Row)
        Call PostingAcActive(True, Row)
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprExprBuilder_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If Col = eExBuild.mExName Then
      If Row = 1 Then
        retVal = "Proceed"
      Else
        sprExprBuilder.GetText eExBuild.mExpr, Row - 1, retVal
      End If
      
      If Trim$(retVal) <> "" Then
        Sql = "Select ExprGroupName, Route+Width as Path from COM_FormulaExpr_Group where isActive='1' and ExprGroupID > 28"
        Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
        If Row > 2 Then     'SUBTOTAL Required only after 2 rows
          Sql = Sql & " Union All"
          Sql = Sql & " Select 'SUBTOTAL', '-1'    "
        End If
	Sql = Sql & " Union All"
        Sql = Sql & " Select 'TOTAL GST', '0001'"
        Sql = Sql & " Order By ExprGroupName"
        
        If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Expression", "Expression Group") = True Then
          If Val(getvalue(2)) = -1 Then             'SUBTOTAL
            sprExprBuilder.SetText eExBuild.mExRoute, Row, -1 * Row
            sprExprBuilder.SetText eExBuild.mExName, Row, getvalue(1)
            sprExprBuilder.SetText eExBuild.mGrand, Row, "No"
            Call Spread_Lock(sprExprBuilder, True, eExBuild.mGrand, Row)
            Call sprExprBuilder_ComboSelChange(eExBuild.mGrand, Row)
          Else
            sprExprBuilder.SetText eExBuild.mExRoute, Row, getvalue(2)
            If Spread_Clear_Duplicate(sprExprBuilder, Row, eExBuild.mExRoute) = False Then
              sprExprBuilder.SetText eExBuild.mExName, Row, getvalue(1)
            End If
            sprExprBuilder.SetText eExBuild.mGrand, Row, "Yes (Add)"
            Call Spread_Lock(sprExprBuilder, False, eExBuild.mGrand, Row)
            Call sprExprBuilder_ComboSelChange(eExBuild.mGrand, Row)
          End If
          sprExprBuilder.MaxRows = sprExprBuilder.DataRowCnt + 1
          Call ExprBuildButton_Status(Row)
        End If
      End If
  ElseIf Col = eExBuild.mAcName Then
    If Trim$(GetSprText(sprExprBuilder, eExBuild.mGrand, Row)) <> "No" Then
      Sql = "Select LedgerName, LedgerCode from FAS_LedgerMaster where left(LrRoute,4) not in ('0206', '0303') "
      Sql = Sql & " and isActive = '1' and CompCode = '" & gCompID & gBranchID & "'"
      Sql = Sql & " Union All"
      Sql = Sql & " Select 'As per Order', -1"
      Sql = Sql & " Order by LedgerName"
      If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Accounts", "Account Name") = True Then
        sprExprBuilder.SetText eExBuild.mAcCode, Row, getvalue(2)    'AcCode
        sprExprBuilder.SetText eExBuild.mAcName, Row, getvalue(1)    'AcName
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprExprBuilder_EditChange(sender as object, e as EditChangeEventArgs)
' On Error Resume Next
  If Col = eExBuild.mExName Then
      sprExprBuilder.Col = Col
      sprExprBuilder.Row = Row
      sprExprBuilder.Text = Trim$(sprExprBuilder.Text)
  End If
End Sub

Private Sub sprExprBuilder_KeyUp(sender as object, e as KeyUpEventArgs)
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = vbKeyDelete And Shift = vbShiftMask Then
    If sprExprBuilder.EditMode = False Then
        If sprExprBuilder.ActiveRow > 1 And sprExprBuilder.ActiveRow < 3 Then
          sprExprBuildSeq.GetText eExBuildSeq.mBlock, 1, retVal
          If sprExprBuilder.ActiveRow = sprExprBuilder.DataRowCnt Then
            Call UnHideParamStatusWhileDelete(sprExprBuilder.ActiveRow)
          
            rCnt = sprExprBuildSeq.SearchCol(eExBuildSeq.mExSeq, 0, -1, sprExprBuilder.ActiveRow, SearchFlagsValue)
            If rCnt <> -1 Then
              sprExprBuildSeq.ClearRange 1, rCnt, sprExprBuildSeq.MaxCols, sprExprBuildSeq.DataRowCnt, True
            End If
            sprExprBuilder.DeleteRows sprExprBuilder.ActiveRow, 1
            sprExprBuilder.MaxRows = sprExprBuilder.DataRowCnt + 1
          End If
          
        Else
          If sprExprBuilder.ActiveRow = sprExprBuilder.DataRowCnt Then
            Call UnHideParamStatusWhileDelete(sprExprBuilder.ActiveRow)
            rCnt = sprExprBuildSeq.SearchCol(eExBuildSeq.mExSeq, 0, -1, sprExprBuilder.ActiveRow, SearchFlagsValue)
            If rCnt <> -1 Then
              sprExprBuildSeq.ClearRange 1, rCnt, sprExprBuildSeq.MaxCols, sprExprBuildSeq.DataRowCnt, True
            End If
            sprExprBuilder.DeleteRows sprExprBuilder.ActiveRow, 1
            sprExprBuilder.MaxRows = sprExprBuilder.DataRowCnt + 1
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

Private Sub CmdClear_Click()
' On Error Resume Next
  
  Call ClearForm(Me)
  Call ClearParamSpread
  Call LoadParamSpread
  sprExprBuilder.MaxRows = 1
  Call Spread_BkColour(sprExprBuilder, "&HFFFFFF", eExBuild.mAcName, 1)
  cmbRoundOff.ListIndex = 0
  cmbCreateType.ListIndex = 0
End Sub

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdExprApply_Click()
  Dim fRow As Long
  
  Try
  Me.MousePointer = vbHourglass

  If ValidateExpr(txtExpr.Text, sprCurExprBuildSeq.DataRowCnt) = False Then
    GoTo ExitHere
  Else
  
    'Update the expression in Expression Builder
    fRow = sprExprBuildSeq.SearchCol(eExBuildSeq.mExSeq, 0, -1, CStr(sprExprBuilder.ActiveRow), SearchFlagsNone)
    If fRow <= 0 Then
      fRow = sprExprBuildSeq.DataRowCnt + 1
    Else
      sprExprBuildSeq.ClearRange 1, fRow, -1, -1, True
    End If
  
    'Copy to Spread
    sprExprBuilder.SetText eExBuild.mExpr, sprExprBuilder.ActiveRow, Trim$(txtExpr.Tag)
    
      sprCurExprBuildSeq.SetSelection 1, 1, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt
      sprCurExprBuildSeq.ClipboardCopy
      sprCurExprBuildSeq.ClearSelection    
      sprExprBuildSeq.SetSelection 1, (fRow), sprExprBuildSeq.MaxCols, (sprExprBuildSeq.DataRowCnt + sprCurExprBuildSeq.DataRowCnt)
      sprExprBuildSeq.ClipboardPaste
      sprExprBuildSeq.ClearSelection  
    sprCurExprBuildSeq.ClearRange 1, 1, -1, -1, True
    
    'Tab change
    Call SetActiveTab(eTab.mExprBuild, eTab.mExpr)
    'Set Active Column to RoundingOff
      sprExprBuilder.SetFocus
'      'To lock the cell Expression
'      Call Spread_Lock(sprExprBuilder, True, eExBuild.mExpr, 1, eExBuild.mExpr, .MaxRows)
      'To unlock the cell  Remarks
      Call Spread_Lock(sprExprBuilder, False, eExBuild.mRemarks, sprExprBuilder.ActiveRow, eExBuild.mRemarks, sprExprBuilder.ActiveRow)
      sprExprBuilder.SetActiveCell eExBuild.mRndOff, sprExprBuilder.ActiveRow
      If GetSprText(sprExprBuilder, eExBuild.mRndOff, sprExprBuilder.ActiveRow) = "" Then
        sprExprBuilder.SetText eExBuild.mRndOff, sprExprBuilder.ActiveRow, "None"
        'To unlock the cell RoundingOff
        Call Spread_Lock(sprExprBuilder, False, eExBuild.mRndOff, sprExprBuilder.ActiveRow)
      End If
      If Val(GetSprText(sprExprBuilder, eExBuild.mExRoute, sprExprBuilder.ActiveRow)) > 0 Then
        sprExprBuilder.SetText eExBuild.mGrand, sprExprBuilder.ActiveRow, "Yes (Add)"   'Take for Grand Total
      End If
    
    'Clear Expression & Sequence Spread
    txtExpr.Text = ""
    txtExpr.Tag = ""
    
'    Call SetAllParamEnabled
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExprCancel_Click()
  Dim rCnt As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  txtExpr.Text = ""
txtExpr.Tag = ""
  Call UnHideParamStatusWhileAllClear
  sprCurExprBuildSeq.ClearRange 1, 1, -1, -1, True
  
'  With sprExprBuildSeq
'    rCnt = .SearchCol(eExBuildSeq.mExSeq, 0, -1, sprExprBuilder.ActiveRow, SearchFlagsValue)
'    If rCnt > -1 Then
'      .ClearRange 1, rCnt, .MaxCols, .DataRowCnt, True
'    End If
'  End With
  
    sprExprBuilder.SetText eExBuild.mExpr, sprExprBuilder.ActiveRow, ""
  
'  Call SetAllParamEnabled
  
  Call SetActiveTab(eTab.mExprBuild, eTab.mExpr)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExprSeq_Click()
  Dim rCnt As Long
  Dim rVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If ValidateSeq("ExprSeq") = True Then
      sprExprSelect.MaxRows = sprExprBuilder.ActiveRow - 1
      'Values copied to the Select Spread
      For rCnt = 1 To (sprExprBuilder.ActiveRow - 1)
        sprExprBuilder.GetText eExBuild.mExName, rCnt, rVal
        sprExprSelect.SetText eExSel.mExName, rCnt, rVal
        sprExprBuilder.GetText eExBuild.mExpr, rCnt, rVal
        sprExprSelect.SetText eExSel.mExpr, rCnt, rVal
      Next rCnt
      
      cmdNext.Visible = False
      sstExpression.TabCaption(eTab.mSelect) = "Select Expression"
      sprExprSelect.SetText eExSel.mExName, 0, "Expression Name"
      sprExprSelect.SetText eExSel.mExpr, 0, "Expression"
  
    Call SetActiveTab(eTab.mExprBuild, eTab.mSelect)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdFormula_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  Me.MousePointer = vbHourglass

  If ValidateSeq("ExprSeq") = True Then
    Sql = "Select FrmSName, FrmLName, FrmID from COM_FormulaMast_Hdr"
    Sql = Sql & " Where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If rs.EOF Then
      sShowMessage "No Formula availble to Load."
      GoTo ExitHere
    End If
    
      sprExprSelect.MaxRows = rs.RecordCount
      'Values copied to the Select Spread
      For rCnt = 1 To sprExprSelect.MaxRows
        sprExprSelect.SetText eExSel.mExName, rCnt, rs("FrmSName")
        sprExprSelect.SetText eExSel.mExpr, rCnt, rs("FrmLName")
        sprExprSelect.SetText eExSel.mExCode, rCnt, rs("FrmID")
        rs.MoveNext
      Next rCnt
      
      cmdNext.Visible = True
      cmdNext.Caption = "Select Expression"
      sstExpression.TabCaption(eTab.mSelect) = "Select Formula"
      sprExprSelect.SetText eExSel.mExName, 0, "Formula Short Name"
      sprExprSelect.SetText eExSel.mExpr, 0, "Formula Long Name"
    Call SetActiveTab(eTab.mExprBuild, eTab.mSelect)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdNext_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim SelectCode As Integer
  Dim SelectName As String
  
  Dim isExpression As Boolean
  
  Try
  Me.MousePointer = vbHourglass

  SelectCode = Val(GetSprText(sprExprSelect, eExSel.mExCode, sprExprSelect.ActiveRow))
  SelectName = Val(GetSprText(sprExprSelect, eExSel.mExName, sprExprSelect.ActiveRow))
  isExpression = IIf(cmdNext.Caption = "Select Expression", True, False)
  
  If isExpression = True Then
    Sql = "Select b.ExprGroupName as SName, dbo.fn_FrontEndExpression(" & SelectCode & ", a.ExprSeq, a.BranchID, a.CompID) as LName, a.ExprSeq as Code"
    Sql = Sql & " from COM_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group b"
    Sql = Sql & " On a.ExprRoute=b.Route+b.Width and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.FrmID=" & SelectCode
    Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Else
    Sql = "Select Distinct b.ParamShortName as SName, b.ParamLongName as LName, b.ParamID as Code"
    Sql = Sql & " from COM_FormulaMast_Block_Dtl a Left Outer Join COM_FrmParamMast_Hdr b"
    Sql = Sql & " On Convert(int, a.Block)=b.ParamID and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where a.ModuleCode=" & gModuleCode & " and a.FrmID=" & Val(sprExprSelect.Tag) & " and a.ExprSeq=" & SelectCode
    Sql = Sql & " and a.BlockType='Param' and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  End If

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    If isExpression = True Then
      sShowMessage "No Expression availble to Load."
    Else
      sShowMessage "No Parameters availble to Load."
    End If
    GoTo ExitHere
  End If
  
    If isExpression = True Then
      sprExprSelect.Tag = SelectCode           'Formula code saved for future reference
    End If
    
    sprExprSelect.MaxRows = rs.RecordCount
    'Values copied to the Select Spread
    For rCnt = 1 To sprExprSelect.MaxRows
      sprExprSelect.SetText eExSel.mExName, rCnt, rs("SName")
      sprExprSelect.SetText eExSel.mExpr, rCnt, rs("LName")
      sprExprSelect.SetText eExSel.mExCode, rCnt, rs("Code")
      rs.MoveNext
    Next rCnt
  
  If isExpression = True Then
    sstExpression.TabCaption(eTab.mSelect) = "Select Expression"
    sprExprSelect.SetText eExSel.mExName, 0, "Expression Name"
    sprExprSelect.SetText eExSel.mExpr, 0, "Expression"
    cmdNext.Caption = "Select Parameter"
  Else
    sstExpression.TabCaption(eTab.mSelect) = "Select Parameter"
    sprExprSelect.SetText eExSel.mExName, 0, "Parameter Short Name"
    sprExprSelect.SetText eExSel.mExpr, 0, "Parameter Long Name"
    cmdNext.Visible = False
  End If

  Call SetActiveTab(eTab.mExprBuild, eTab.mSelect)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdLogic_Click(ByVal Index As Integer)
  Dim cRow As Long
  Dim CurText As String
  Dim CurType As String
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = Trim$(cmdLogic(Index).Caption)
  
  Select Case Index
    Case eLogic.mIF
      CurType = "LogicIF"
    Case eLogic.mTHEN
      CurType = "LogicTHEN"
    Case eLogic.mELSE
      CurType = "LogicELSE"
    Case eLogic.mENDIF
      CurType = "LogicENDIF"
    Case eLogic.mAND, eLogic.mOR, eLogic.mNOT
      CurType = "LogicANDOR"
    Case eLogic.mNOT
      CurType = "LogicNOT"
    Case Else
      CurType = "LogicCompare"
  End Select
  
  If ValidateSeq(CurType) = True Then

      cRow = sprCurExprBuildSeq.DataRowCnt + 1
      sprCurExprBuildSeq.SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlockSeq, cRow, cRow
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, CurText
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlockType, cRow, CurType
      sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
      
      txtExpr.Text = txtExpr.Text & CurText
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdNumber_Click(ByVal Index As Integer)
  Dim CurType As String
  Dim CurText As String
  Dim retVal As Variant
  Dim cRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = Trim$(cmdNumber(Index).Caption)
  CurType = "NoBlock"
  
  If ValidateSeq(CurType, CurText) = True Then

      cRow = sprCurExprBuildSeq.DataRowCnt
    
      sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, cRow, retVal
      If Trim$(retVal) = "NoBlock" Then
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, cRow, retVal
        sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, Trim$(retVal) & CurText
      Else
        cRow = sprCurExprBuildSeq.DataRowCnt + 1
        sprCurExprBuildSeq.SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
        sprCurExprBuildSeq.SetText eExBuildSeq.mBlockSeq, cRow, cRow
        sprCurExprBuildSeq.SetText eExBuildSeq.mBlockType, cRow, "NoBlock"
        sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, CurText
        sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
      End If
      txtExpr.Text = txtExpr.Text & cmdNumber(Index).Caption
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub ClickParamSpr(ByVal fRow As Long)
  Dim fParamID As Long
  Dim CurType As String
  Dim CurText As String
  Dim SelVal As String
  Dim PercStr As String
  Dim retVal As Variant
  Dim cRow As Long
  Dim isPerc As Boolean
  Dim isEditable As Boolean
  Dim isPercBlockOpen As Boolean
  
  Try
  Me.MousePointer = vbHourglass
  
  sprParam.GetText eParam.mSName, fRow, retVal
  CurText = Trim$(retVal)
  CurType = "Param"
  
    If ValidateSeq(CurType) = True Then
      
      sprParam.GetText eParam.mParamID, fRow, retVal
      fParamID = retVal
      
      sprParam.GetText eParam.mIsEditable, fRow, retVal
      isEditable = retVal
      
      sprParam.GetText eParam.mIsPerc, fRow, retVal
      isPerc = retVal
      
      cRow = sprCurExprBuildSeq.DataRowCnt + 1
      sprParam.GetText eParam.mValinMethod, fRow, retVal
      Select Case Trim$(retVal)
        Case "DV"                         'Defined Values
          With frmPromptValues
            sprCurExprBuildSeq.lblValue.Enabled = False
            sprCurExprBuildSeq.txtValue.Enabled = False
            sprCurExprBuildSeq.txtValue.Tag = fParamID
            sprCurExprBuildSeq.optTypePerc.Value = isPerc
            sprCurExprBuildSeq.fraParamType.Enabled = isEditable
            sprCurExprBuildSeq.chkTypeEditable.Value = IIf(isEditable, 1, 0)
            sprCurExprBuildSeq.cmdOk.Caption = "&Proceed"
            sprCurExprBuildSeq.Show vbModal
            frmPromptValues.lblValue.Enabled = False
            frmPromptValues.txtValue.Enabled = False
            frmPromptValues.txtValue.Tag = fParamID
            frmPromptValues.optTypePerc.Value = isPerc
            frmPromptValues.fraParamType.Enabled = isEditable
            frmPromptValues.chkTypeEditable.Value = IIf(isEditable, 1, 0)
            frmPromptValues.cmdOk.Caption = "&Proceed"
            frmPromptValues.Show vbModal            
          If Val(getvalue(0)) = -1 Then           'User cancelled the prompt form
            txtExpr.Text = Trim$(txtExpr.Text)
            GoTo ExitHere
          Else
            .SetText eExBuildSeq.mValinMethod, cRow, "DV"
            .SetText eExBuildSeq.mValUsed, cRow, getvalue(3)                           'Value
            .SetText eExBuildSeq.mIsPerc, cRow, IIf(getvalue(1), "1", "0")             'Is Type Percentage
            .SetText eExBuildSeq.mIsTypeEditable, cRow, IIf(getvalue(2) = 1, "1", "0") 'Is Type Editable
            .SetText eExBuildSeq.mIsValueEditable, cRow, getvalue(4)                   'Is Value Editable
            .SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
            
            PercStr = IIf(getvalue(1), "%", "")             'If type is Percentage
            SelVal = "[" & getvalue(3) & PercStr & "]"      'Value
            
            txtExpr.Text = txtExpr.Text & CurText & SelVal
            
            If (getvalue(2) = 1) Or (getvalue(1) = True) Then 'Type is Editable or It is Percentage
              isPercBlockOpen = True
            Else
              isPercBlockOpen = False
            End If
          End If
          
        Case "PB"                         'Prompt While Building Expression
          sprParam.GetText eParam.mDefined, fRow, retVal
            If Trim$(retVal) = "None" Then
              frmPromptValues.lblValue.Enabled = True
              frmPromptValues.txtValue.Enabled = True
              frmPromptValues.cmdOk.Caption = "&Ok"
            Else
              frmPromptValues.lblValue.Enabled = False
              frmPromptValues.txtValue.Enabled = False
              frmPromptValues.cmdOk.Caption = "&Proceed"
            End If
            frmPromptValues.txtValue.Tag = fParamID
            frmPromptValues.optTypePerc.Value = isPerc
            frmPromptValues.fraParamType.Enabled = isEditable
            frmPromptValues.chkTypeEditable.Value = IIf(isEditable, 1, 0)
            frmPromptValues.Show vbModal
          
          If Val(getvalue(0)) = -1 Then         'User cancelled the prompt
            txtExpr.Text = Trim$(txtExpr.Text)
            GoTo ExitHere
          Else
            If Val(getvalue(3)) <> 0 Then     'While Building Expression, the given value 0 is equal to cancelled it
              .SetText eExBuildSeq.mValinMethod, cRow, "PB"
              .SetText eExBuildSeq.mValUsed, cRow, getvalue(3)                           'Value
              .SetText eExBuildSeq.mIsPerc, cRow, IIf(getvalue(1), "1", "0")             'Is Type Percentage
              .SetText eExBuildSeq.mIsTypeEditable, cRow, IIf(getvalue(2) = 1, "1", "0") 'Is Type Editable
              .SetText eExBuildSeq.mIsValueEditable, cRow, getvalue(4)                   'Is Value Editable
              .SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
              
              PercStr = IIf(getvalue(1), "%", "")             'If type is Percentage
              SelVal = "[" & getvalue(3) & PercStr & "]"      'Value
              
              txtExpr.Text = txtExpr.Text & CurText & SelVal
              
              If (getvalue(2) = 1) Or (getvalue(1) = True) Then 'Type is Editable or It is Percentage
                isPercBlockOpen = True
              Else
                isPercBlockOpen = False
              End If
            Else
              txtExpr.Text = Trim$(txtExpr.Text)
              GoTo ExitHere
            End If
          End If
          
        Case "PE"                         'Prompt While Executing the Formula
          .SetText eExBuildSeq.mValinMethod, cRow, "PE"
          .SetText eExBuildSeq.mValUsed, cRow, ""                                    'Value
          .SetText eExBuildSeq.mIsPerc, cRow, IIf(isPerc, "1", "0")                  'Is Type Percentage
          .SetText eExBuildSeq.mIsTypeEditable, cRow, IIf(isEditable, "1", "0")      'Is Type Editable
          .SetText eExBuildSeq.mIsValueEditable, cRow, "1"                           'Is Value Editable
          .SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
          
          If Len(PercStr) = 0 Then
            SelVal = ""
          Else
            SelVal = "[" & PercStr & "]"
          End If
          txtExpr.Text = txtExpr.Text & CurText & SelVal
          
          If (isEditable = True) Or (isPerc = True) Then 'Type is Editable or It is Percentage
            isPercBlockOpen = True
          Else
            isPercBlockOpen = False
          End If
        
        Case Else
          .SetText eExBuildSeq.mValinMethod, cRow, "SC"
          .SetText eExBuildSeq.mValUsed, cRow, ""                                    'Value
          .SetText eExBuildSeq.mIsPerc, cRow, IIf(isPerc, "1", "0")                  'Is Type Percentage
          .SetText eExBuildSeq.mIsTypeEditable, cRow, IIf(isEditable, "1", "0")      'Is Type Editable
          .SetText eExBuildSeq.mIsValueEditable, cRow, ""                           'Is Value Editable
          .SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
          
          If Len(PercStr) = 0 Then
            SelVal = ""
          Else
            SelVal = "[" & PercStr & "]"
          End If
          txtExpr.Text = txtExpr.Text & CurText & SelVal
          
          If (isEditable = True) Or (isPerc = True) Then 'Type is Editable or It is Percentage
            isPercBlockOpen = True
          Else
            isPercBlockOpen = False
          End If
      End Select
      
      .SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
      .SetText eExBuildSeq.mBlockSeq, cRow, cRow
      .SetText eExBuildSeq.mBlock, cRow, fParamID
      .SetText eExBuildSeq.mBlockType, cRow, "Param"
      
      If isPercBlockOpen = True Then      'For Percentage Block Open Symbol
        cRow = .DataRowCnt + 1
        .SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
        .SetText eExBuildSeq.mBlockSeq, cRow, cRow
        .SetText eExBuildSeq.mBlock, cRow, PercBlockOpen
        .SetText eExBuildSeq.mBlockType, cRow, "PercOpen"
        .SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
        
        txtExpr.Text = txtExpr.Text & PercBlockOpen
        
        sprParam.Row = fRow
        sprParam.RowHidden = True
      End If
      
      Call CheckMultiParamEnableStatus(False, fParamID)
    End If
  End With
  
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CheckMultiParamEnableStatus(Optional ByVal isClear As Boolean = False,Optional ByVal cParamID As Long = -1,Optional ByVal CurRow As Long = 1)

  Try

  If cParamID = -1 Then
    GoTo ExitHere
  Else
    Select Case cParamID
      Case eParamID.mQxR, eParamID.mQxAR
        If AllMultiParamCleared(CurRow - 1) = True Then
          Call ChangeParamStatus(eParamID.mGV, False)
          Call ChangeParamStatus(eParamID.mAGV, False)
        End If
      Case eParamID.mGV, eParamID.mAGV
        If AllMultiParamCleared(CurRow - 1) = True Then
          Call ChangeParamStatus(eParamID.mQxR, False)
          Call ChangeParamStatus(eParamID.mQxAR, False)
        End If
    End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function AllMultiParamCleared(ByVal fRow As Long) As Boolean
  Dim retVal As Variant
  Dim cRow As Long

  Try
  AllMultiParamCleared = False
  
  If fRow > 0 Then
      cRow = 0
      Do While cRow >= 0
        cRow = sprCurExprBuildSeq.SearchCol(eExBuildSeq.mBlockType, cRow, fRow - 1, "Param", SearchFlagsNone)
        If cRow > 0 Then
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, cRow, retVal
          Select Case Val(retVal)
            Case eParamID.mQxR, eParamID.mQxAR, eParamID.mGV, eParamID.mAGV
              GoTo ExitHere
          End Select
        End If
      Loop
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdSeqCancel_Click()
' On Error Resume Next
  
  Call SetActiveTab(eTab.mSelect, eTab.mExprBuild)
End Sub

Private Sub cmdSeqSelect_Click()
  Dim CurText As String
  Dim cRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = GetSprText(sprExprSelect, eExSel.mExName, sprExprSelect.ActiveRow)
  
    cRow = sprCurExprBuildSeq.DataRowCnt + 1
    sprCurExprBuildSeq.SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlockSeq, cRow, cRow
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, sprExprSelect.ActiveRow    'To retreive the ExprName we store the Expression Sequence No
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlockType, cRow, "ExprSeq"
    sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
  
  txtExpr.Text = txtExpr.Text & "#" & CurText & "#"
  
  Call SetActiveTab(eTab.mSelect, eTab.mExprBuild)
  Call UpdateExpression
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdOpr_Click(ByVal Index As Integer)
  Dim cRow As Long
  Dim CurType As String
  Dim CurText As String
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = Trim$(cmdOpr(Index).Caption)
  
  If Index = eOperator.mOpenBrace Then
    CurType = "OprOpenBrace"
  ElseIf Index = eOperator.mCloseBrace Then
    CurType = "OprCloseBrace"
  ElseIf Index = eOperator.mPercClose Then
    CurType = "PercClose"
  Else
    CurType = "OprOther"
  End If
  
  If ValidateSeq(CurType) = True Then

      cRow = sprCurExprBuildSeq.DataRowCnt + 1
      sprCurExprBuildSeq.SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlockSeq, cRow, cRow
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, CurText
      sprCurExprBuildSeq.SetText eExBuildSeq.mBlockType, cRow, CurType
      sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)

    txtExpr.Text = txtExpr.Text & CurText
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdOprEdit_Click(ByVal Index As Integer)
  Dim CurText As String
  Dim retVal As Variant
  Dim fRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = Trim$(cmdOprEdit(Index).Caption)
  
    If CurText = "C" Then
'      Call ClearLastEntry
      sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, sprCurExprBuildSeq.DataRowCnt, retVal
      Select Case Trim$(retVal)
        Case "PercOpen"       'If it PercOpen block then clear the parameter itself
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt - 1, retVal
          'To unhide the row
          fRow = sprParam.SearchCol(eParam.mParamID, 0, -1, Trim$(retVal), SearchFlagsNone)
          If fRow > 0 Then
            sprParam.Row = fRow
            sprParam.RowHidden = False
          End If
          
          Call CheckMultiParamEnableStatus(True, Val(retVal), sprCurExprBuildSeq.DataRowCnt - 1)
          
          sprCurExprBuildSeq.GetText eExBuildSeq.mPrevCharEnd, sprCurExprBuildSeq.DataRowCnt - 1, retVal
          sprCurExprBuildSeq.ClearRange 1, sprCurExprBuildSeq.DataRowCnt - 1, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
          
        Case "NoBlock"
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
          If Len(Trim$(retVal)) > 1 Then
            sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, Left$(Trim$(retVal), Len(Trim$(retVal)) - 1)
          Else
            sprCurExprBuildSeq.ClearRange 1, sprCurExprBuildSeq.DataRowCnt, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
          End If
          retVal = Len(txtExpr.Text) - 1
        Case Else
          If retVal = "Param" Then
            sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
            Call CheckMultiParamEnableStatus(True, Val(retVal))
          End If
          sprCurExprBuildSeq.GetText eExBuildSeq.mPrevCharEnd, sprCurExprBuildSeq.DataRowCnt, retVal
          sprCurExprBuildSeq.ClearRange 1, sprCurExprBuildSeq.DataRowCnt, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
      End Select
      txtExpr.Text = Left$(txtExpr.Text, Val(retVal))
    Else
      Call UnHideParamStatusWhileAllClear
    
      sprCurExprBuildSeq.ClearRange 1, 1, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
      txtExpr.Text = ""
txtExpr.Tag = ""
'      Call SetAllParamEnabled
    End If
  
  Call UpdateExpression
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub UnHideParamStatusWhileDelete(ByVal cRow As Long)
  Dim retVal As Variant
  Dim fRow As Long        'Search Row in Expression Builder Spread
  Dim srow As Long        'Search Row in Paramater Spread

  Try
  
    fRow = sprExprBuildSeq.SearchCol(eExBuildSeq.mExSeq, 0, -1, CStr(cRow), SearchFlagsNone)
    
    If fRow >= 0 Then
      fRow = sprExprBuildSeq.SearchCol(eExBuildSeq.mBlockType, fRow - 1, -1, "Param", SearchFlagsNone)
      
      Do While fRow >= 0
        sprExprBuildSeq.GetText eExBuildSeq.mBlock, fRow, retVal
        Select Case Val(retVal)
          Case eParamID.mAGV, eParamID.mGV, eParamID.mQxAR, eParamID.mQxR
          
          Case Else
            srow = sprParam.SearchCol(eParam.mParamID, 0, -1, Trim$(retVal), SearchFlagsNone)
            If srow >= 0 Then
              sprParam.Row = srow
              sprParam.RowHidden = False
            End If
        End Select
        fRow = sprExprBuildSeq.SearchCol(eExBuildSeq.mBlockType, fRow, -1, "Param", SearchFlagsNone)
      Loop
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub UnHideParamStatusWhileAllClear()
  Dim retVal As Variant
  Dim fRow As Long        'Search Row in Current Expression Builder Spread
  Dim srow As Long        'Search Row in Paramater Spread

  Try
  
    fRow = sprCurExprBuildSeq.SearchCol(eExBuildSeq.mBlockType, 0, -1, "Param", SearchFlagsNone)
    
    Do While fRow >= 0
      sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, fRow, retVal
      Select Case Val(retVal)
        Case eParamID.mAGV, eParamID.mGV, eParamID.mQxAR, eParamID.mQxR
        
        Case Else
          srow = sprParam.SearchCol(eParam.mParamID, 0, -1, Trim$(retVal), SearchFlagsNone)
          If srow >= 0 Then
            sprParam.Row = srow
            sprParam.RowHidden = False
          End If
      End Select
      fRow = sprCurExprBuildSeq.SearchCol(eExBuildSeq.mBlockType, fRow, -1, "Param", SearchFlagsNone)
    Loop
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetAllParamEnabled()
' On Error Resume Next
  
  Call ChangeParamStatus(eParamID.mQxR, True)
  Call ChangeParamStatus(eParamID.mQxAR, True)
  Call ChangeParamStatus(eParamID.mGV, True)
  Call ChangeParamStatus(eParamID.mAGV, True)
End Sub

Private Sub ChangeParamStatus(ByVal ParamID As Integer,ByVal  isEnable As Boolean)
  Try
  Dim cRow As Long
  
    cRow = sprParam.SearchCol(eParam.mParamID, cRow, -1, ParamID, SearchFlagsValue)
    If cRow > 0 Then
      sprParam.Row = cRow
      sprParam.RowHidden = Not isEnable
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub CmdSave_Click()
  Dim rCnt As Long
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim PrevBlockSeqEnd As Integer
  Dim CurExprSeq As Integer
  Dim tmpVar As Variant

  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If

  Set rs = New AceADODB.Recordset
  Sql = "Select FrmID from Com_FormulaMast_Hdr"
  Sql = Sql & " where FrmSName='" & Trim$(txtFSName.Text) & "' and ModuleCode=" & gModuleCode
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
'  Set rs = New ADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    MsgBox "Formula Short Name is already available. Please give unique name."
    GoTo ExitHere
  Else
    Sql = "Select Max(FrmID) from Com_FormulaMast_Hdr"
    Sql = Sql & " where ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
'    Set rs = New ADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    If Not rs.EOF Then
      txtFSName.Tag = IIf(IsNull(rs(0)), 0, rs(0)) + 1
    End If
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  'Formula Master Header
  Sql = "Execute Com_Ins_FormulaMast_Hdr"
  Sql = Sql & " @ModuleCode='" & gModuleCode & "',"
  Sql = Sql & " @FrmID=" & Val(txtFSName.Tag) & ","
  Sql = Sql & " @FrmSName='" & Trim$(txtFSName.Text) & "',"
  Sql = Sql & " @FrmLName='" & Trim$(txtFLName.Text) & "',"
  Sql = Sql & " @FrmDesc='" & Trim$(txtFDesc.Text) & "',"
  Sql = Sql & " @RoundOff='" & GetRoundingOffCode(Trim$(cmbRoundOff.Text)) & "',"
  Sql = Sql & " @RoundOffValue=" & Val(txtROffValue.Text) & ","
  Sql = Sql & " @PostingAc=" & Val(txtROffPosting.Tag) & ","
  Sql = Sql & " @isActive='1',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  'Formula Master Detail
  For rCnt = 1 To sprExprBuilder.DataRowCnt
    If Len(GetSprText(sprExprBuilder, eExBuild.mRndOff, rCnt)) > 0 Then
      Sql = "Execute Com_Ins_FormulaMast_Dtl"
      Sql = Sql & " @ModuleCode='" & gModuleCode & "',"
      Sql = Sql & " @FrmID=" & Val(txtFSName.Tag) & ","
      Sql = Sql & " @ExprSeq=" & rCnt & ","
      
      sprExprBuilder.GetText eExBuild.mExRoute, rCnt, tmpVar
      Sql = Sql & " @ExprRoute='" & Trim$(tmpVar) & "',"
      Sql = Sql & " @RoundOff='" & GetRoundingOffCode(GetSprText(sprExprBuilder, eExBuild.mRndOff, rCnt)) & "',"
      Sql = Sql & " @RoundOffValue=" & Val(GetSprText(sprExprBuilder, eExBuild.mRndOffVal, rCnt)) & ","
      
      Select Case Trim$(GetSprText(sprExprBuilder, eExBuild.mGrand, rCnt))
        Case "No"
          Sql = Sql & " @GrandTotalFlag=0,"
        Case "Yes (Add)"
          Sql = Sql & " @GrandTotalFlag=1,"
        Case "Yes (Subtract)"
          Sql = Sql & " @GrandTotalFlag=-1,"
      End Select
      
      Sql = Sql & " @PostingAc=" & Val(GetSprText(sprExprBuilder, eExBuild.mAcCode, rCnt)) & ","
      Sql = Sql & " @Remarks='" & GetSprText(sprExprBuilder, eExBuild.mRemarks, rCnt) & "',"
      Sql = Sql & " @UserID='" & gUserID & "',"
      Sql = Sql & " @BranchID='" & gBranchID & "',"
      Sql = Sql & " @CompID='" & gCompID & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    End If
  Next rCnt
  
  'Formula Master Block Detail
    PrevBlockSeqEnd = 0
CurExprSeq = 0
    For rCnt = 1 To sprExprBuildSeq.DataRowCnt
      Sql = "Execute Com_Ins_FormulaMast_Block_Dtl"
      Sql = Sql & " @ModuleCode='" & gModuleCode & "',"
      Sql = Sql & " @FrmID=" & Val(txtFSName.Tag) & ","
      
      'To get serial no group by the Expr Seq
      sprExprBuildSeq.GetText eExBuildSeq.mExSeq, rCnt, tmpVar
      If tmpVar <> CurExprSeq Then
        CurExprSeq = tmpVar
        PrevBlockSeqEnd = rCnt - 1
      End If
      Sql = Sql & " @ExprSeq=" & CurExprSeq & ","
      
      Sql = Sql & " @BlockSeq=" & (rCnt - PrevBlockSeqEnd) & ","
      sprExprBuildSeq.GetText eExBuildSeq.mBlockType, rCnt, tmpVar
      Sql = Sql & " @BlockType='" & Trim$(tmpVar) & "',"
      
      sprExprBuildSeq.GetText eExBuildSeq.mBlock, rCnt, tmpVar
      Sql = Sql & " @Block='" & Trim$(tmpVar) & "',"
      
      sprExprBuildSeq.GetText eExBuildSeq.mIsPerc, rCnt, tmpVar
      Sql = Sql & " @isTypePercentage='" & Val(tmpVar) & "',"
      
      sprExprBuildSeq.GetText eExBuildSeq.mIsValueEditable, rCnt, tmpVar
      Sql = Sql & " @isValueEditable='" & Val(tmpVar) & "',"
      
      If Val(tmpVar) <> 1 Then    'If Value is freezed Automatically Type is also locked Else get the Actual Status
        sprExprBuildSeq.GetText eExBuildSeq.mIsTypeEditable, rCnt, tmpVar
      End If
      Sql = Sql & " @isTypeEditable='" & Val(tmpVar) & "',"
      
      sprExprBuildSeq.GetText eExBuildSeq.mValUsed, rCnt, tmpVar
      Sql = Sql & " @ValueUsed=" & Val(tmpVar) & ","
      
      Sql = Sql & " @UserID='" & gUserID & "',"
      Sql = Sql & " @BranchID='" & gBranchID & "',"
      Sql = Sql & " @CompID='" & gCompID & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    Next rCnt
  
  If ValidateAfterSave(Val(txtFSName.Tag)) = False Then
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If

  sShowMessage "Record Saved Successfully."
  Call CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprExprSelect_DblClick(sender as object, e as DblClickEventArgs)
  Dim CurText As String
  Dim cRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  CurText = GetSprText(sprExprSelect, eExSel.mExName, Row)
  
    cRow = sprCurExprBuildSeq.DataRowCnt + 1
    sprCurExprBuildSeq.SetText eExBuildSeq.mExSeq, cRow, sprExprBuilder.ActiveRow
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlockSeq, cRow, cRow
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, cRow, Row    'To retreive the ExprName we store the Expression Sequence No
    sprCurExprBuildSeq.SetText eExBuildSeq.mBlockType, cRow, "ExprSeq"
    sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, cRow, Len(txtExpr.Text)
  
  txtExpr.Text = txtExpr.Text & "#" & CurText & "#"
  
  Call SetActiveTab(eTab.mSelect, eTab.mExprBuild)
  Call UpdateExpression
Catch ex As Exception
  Call ErrMsgShow(ex)
  Call RollBackExpression

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprParam_DblClick(sender as object, e as DblClickEventArgs)
  Dim retVal As Variant
  Dim ParamID As Long
  
  Try
  Me.MousePointer = vbHourglass
  
    sprParam.GetText eParam.mParamID, Row, retVal
    ParamID = Val(retVal)
    sprParam.GetText eParam.mSName, Row, retVal
    
    Call ClickParamSpr(Row)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFSName_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = ToUpper(KeyAscii)
End Sub

Public Sub LoadParamSpread()
  Dim Sql As String
  Dim rsParam As AceADODB.Recordset
  Dim cRow As Long
  
  Try
  
  Sql = "Select ParamID, ParamShortName + Case When isMultiple='1' Then 'n' Else '' End as ParamShortName, ParamLongName, ParamDesc,"
  Sql = Sql & " IsTypePercentage, IsTypeEditable, ValInputMethod, ValuesDefinedType, isMultiple"
  Sql = Sql & " from COM_FRMPARAMMAST_HDR"
  Sql = Sql & " where ParamID >= 4 and isActive='1' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " Order by ParamShortName"
  
  Set rsParam = New AceADODB.Recordset
  If RSOpen(rsParam, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rsParam.EOF Then
    sprParam.Enabled = False
    sShowMessage "No Parameters are defined or Active."
    GoTo ExitHere
  End If
  
    Do While Not rsParam.EOF
      sprParam.MaxRows = sprParam.DataRowCnt + 1
      cRow = sprParam.MaxRows
      
      sprParam.SetText eParam.mParamID, cRow, rsParam("ParamID")
      sprParam.SetText eParam.mSName, cRow, rsParam("ParamShortName")
      sprParam.SetText eParam.mLName, cRow, rsParam("ParamLongName")
      sprParam.SetText eParam.mDesc, cRow, rsParam("ParamDesc")
      sprParam.SetText eParam.mIsPerc, cRow, rsParam("isTypePercentage")
      sprParam.SetText eParam.mIsEditable, cRow, rsParam("isTypeEditable")
      sprParam.SetText eParam.mValinMethod, cRow, rsParam("ValInputMethod")
      sprParam.SetText eParam.mDefined, cRow, rsParam("ValuesDefinedType")
      sprParam.SetText eParam.mIsMulti, cRow, rsParam("isMultiple")
      rsParam.MoveNext
    Loop
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsParam = Nothing
End Try
End Sub

Public Sub ClearParamSpread()
' On Error Resume Next
    sprParam.ClearRange 1, 1, -1, -1, True
    sprParam.MaxRows = 1
    sprParam.Row = 1
sprParam.Row2 = -1
sprParam.BlockMode = True
sprParam.RowHidden = False
sprParam.BlockMode = False
End Sub

Private Sub ClearLastEntry()
  Dim retVal As Variant
  Dim Expr As String
  Dim Pos As Integer
  Dim cRow As Long
  
  Try
  
    
    sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, sprCurExprBuildSeq.DataRowCnt, retVal
    Select Case Trim$(retVal)
      Case "LogicIF"
        Pos = InStrRev(txtExpr.Text, "IF {", , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "LogicTHEN"
        Pos = InStrRev(txtExpr.Text, "} THEN {", , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "LogicELSE"
        Pos = InStrRev(txtExpr.Text, "} ELSE {", , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "LogicENDIF"
        Pos = InStrRev(txtExpr.Text, "} ENDIF", , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "PercOpen"
        sprCurExprBuildSeq.ClearRange 1, sprCurExprBuildSeq.DataRowCnt, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
        Expr = Trim$(retVal)
        Pos = InStrRev(txtExpr.Text, Expr, , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "ExprSeq"
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
        sprExprBuilder.GetText eExBuild.mExName, CLng(retVal), retVal
        Expr = "#" & Trim$(retVal) & "#"
        Pos = InStrRev(txtExpr.Text, Expr, , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
      Case "Param"
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
        cRow = sprParam.SearchCol(eParam.mParamID, 0, -1, Val(retVal), SearchFlagsValue)
        If cRow > 0 Then
          sprParam.GetText eParam.mSName, cRow, retVal
          Expr = Trim$(retVal)
          Pos = InStrRev(txtExpr.Text, Expr, , vbTextCompare)
          If Pos > 0 Then
            txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
          End If
        End If
      Case "NoBlock"
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
        txtExpr.Text = Left$(txtExpr.Text, Len(txtExpr.Text) - 1)
        If Len(Trim$(retVal)) > 1 Then
          sprCurExprBuildSeq.SetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, Left$(Trim$(retVal), Len(Trim$(retVal)) - 1)
          GoTo ExitHere
        End If
      Case Else   'Case "OprOpenBrace", "OprCloseBrace", "OprOther"
        sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt, retVal
        Expr = retVal
        Pos = InStrRev(txtExpr.Text, Expr, , vbTextCompare)
        If Pos > 0 Then
          txtExpr.Text = Left$(txtExpr.Text, Pos - 1)
        End If
    End Select
    sprCurExprBuildSeq.ClearRange 1, sprCurExprBuildSeq.DataRowCnt, sprCurExprBuildSeq.MaxCols, sprCurExprBuildSeq.DataRowCnt, True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function ValidateAfterSave(ByVal fFrmID As Integer) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  ValidateAfterSave = False
  
  Set rs = New AceADODB.Recordset
  
  'Header Table
  Sql = "Select * from COM_FormulaMast_Hdr"
  Sql = Sql & " Where FrmID=" & fFrmID & " and ModuleCode=" & gModuleCode
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
'  Set rs = New ADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "Formula Not Saved Properly in Header Table."
    GoTo ExitHere
  End If
  
  'Detail Table
  Sql = "Select Max(ExprSeq), Count(ExprSeq) from COM_FormulaMast_Dtl"
  Sql = Sql & " Where FrmID=" & fFrmID & " and ModuleCode=" & gModuleCode
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "Formula Not Saved in Detail Table."
    GoTo ExitHere
  Else
    If IsNull(rs(0)) Then
      sShowMessage "Formula Not Saved Properly in Detail Table."
      GoTo ExitHere
    ElseIf rs(0) <> rs(1) Then
      sShowMessage "Missing Sequence Found in Detail Table."
      GoTo ExitHere
    End If
  End If
  
  'Block Details Table - Overall
  Sql = "Select Max(ExprSeq), Count(Distinct ExprSeq) from COM_FormulaMast_Block_Dtl"
  Sql = Sql & " Where FrmID=" & fFrmID & " and ModuleCode=" & gModuleCode
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "Formula Not Saved in Expression Block Table."
    GoTo ExitHere
  Else
    If IsNull(rs(0)) Then
      sShowMessage "Formula Not Saved Properly in Expression Block Table."
      GoTo ExitHere
    ElseIf rs(0) <> rs(1) Then
      sShowMessage "Missing Sequence Found in Expression Block Table."
      GoTo ExitHere
    End If
  End If
  
  'Block Details Table - Individual Expression
  Sql = "Select Max(BlockSeq), Count(BlockSeq) from COM_FormulaMast_Block_Dtl"
  Sql = Sql & " Where FrmID=" & fFrmID & " and ModuleCode=" & gModuleCode
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " Group by ExprSeq"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "Formula Not Saved in Expression Block Table."
    GoTo ExitHere
  Else
    If IsNull(rs(0)) Then
      sShowMessage "Formula Not Saved Properly in Expression Block Table."
      GoTo ExitHere
    Else
      For rCnt = 1 To rs.RecordCount
        If rs(0) <> rs(1) Then
          sShowMessage "Missing Sequence Found in Expression Block Details Table."
          GoTo ExitHere
        Else
          rs.MoveNext
        End If
      Next rCnt
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function ValidateData() As Boolean
  Dim rCnt As Long
  Dim tmpExpr As String
  
  Try
  ValidateData = False
  
  If Len(Trim$(txtFSName.Text)) = 0 Then
    sShowMessage "Please Enter Formula Short Name."
    txtFSName.SetFocus
    GoTo ExitHere
  End If
  
  If Len(Trim$(txtFLName.Text)) = 0 Then
    sShowMessage "Please Enter Formula Long Name."
    txtFLName.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbRoundOff.Text) <> "None" Then
    If Trim$(txtROffValue.Text) = "" Or IsNumeric(Trim$(txtROffValue.Text)) = False Then
      sShowMessage "Rounding Off Value missing for Formula Total."
      txtROffValue.SetFocus
      GoTo ExitHere
    End If
    
    If Val(txtROffPosting.Tag) = 0 Then
      sShowMessage "Posting Ac for Formula Total Rounding Off is Missing."
      txtROffPosting.SetFocus
      GoTo ExitHere
    End If
  End If
  
    For rCnt = 1 To sprExprBuilder.DataRowCnt
      
      'Expression
      tmpExpr = Trim$(GetSprText(sprExprBuilder, eExBuild.mExpr, rCnt))
      
      If Trim$(tmpExpr) = "" Then
'        If rCnt <> .DataRowCnt Then
          'Missing Expression in between the sequence
          sShowMessage "Expression missing in between the whole sequence."
          GoTo ExitHere
'        End If
      Else
        'All Expression has the Expression Name
        If GetSprText(sprExprBuilder, eExBuild.mExName, rCnt) = "" Then
          sShowMessage "Expr Name missing for row " & rCnt & ""
          GoTo ExitHere
        End If
      End If
      
      'Check for Rounding Off Value
      If GetSprText(sprExprBuilder, eExBuild.mRndOff, rCnt) <> "None" Then
        If Val(GetSprText(sprExprBuilder, eExBuild.mRndOffVal, rCnt)) <= 0 Then
          sShowMessage "Invalid / Missing Rounding Off Value for the Expression in Row " & rCnt
          GoTo ExitHere
        End If
      End If
      
      'All the GrandTotal Expr should have PostingAc
      If Trim$(GetSprText(sprExprBuilder, eExBuild.mGrand, rCnt)) <> "No" Then
        If Val(GetSprText(sprExprBuilder, eExBuild.mAcCode, rCnt)) = 0 Then
          sShowMessage "Posting Account should given for the Expression in Row " & rCnt & " otherwise posting may mismatched."
          GoTo ExitHere
        End If
      Else
        If Val(GetSprText(sprExprBuilder, eExBuild.mAcCode, rCnt)) > 0 Then
          sShowMessage "Posting Account not required for Expr in Row " & rCnt & ". So it will be ignored."
          GoTo ExitHere
        End If
      End If
      
      'Check for Expression Validity
      If CheckBlockIntegrity(tmpExpr) = False Then
        sShowMessage "Expression Error in the Row " & rCnt
        GoTo ExitHere
      End If

    Next rCnt
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateSeq(ByVal CurInputType As String,Optional ByVal CurInputText As String) As Boolean
  Dim PrevType As Variant
  Dim Pos As Integer
  
  Try

  ValidateSeq = True
  
  If Trim$(txtExpr.Text) = "" Then
    PrevType = "NoPrevious"
  Else
    sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, sprCurExprBuildSeq.DataRowCnt, PrevType
    txtExpr.Text = Trim$(txtExpr.Text) & " "
  End If
  
  Select Case CurInputType
    Case "LogicIF"
      Select Case PrevType
        Case "LogicIF", "LogicTHEN", "LogicELSE", "LogicANDOR", "LogicNOT", "LogicCompare", "OprOther"
          GoTo ExitHere
        Case "NoPrevious", "OprOpenBrace", "PercOpen"
          GoTo TrimExitHere
      End Select
    Case "LogicTHEN"
      Select Case PrevType
        Case "LogicENDIF", "Param", "ExprSeq", "OprCloseBrace", "NoBlock", "PercClose"
          If GetStrCount("IF {", txtExpr.Text) > GetStrCount("THEN {", txtExpr.Text) Then
            GoTo TrimExitHere
          End If
      End Select
    Case "LogicELSE"
      Select Case PrevType
        Case "LogicENDIF", "Param", "ExprSeq", "OprCloseBrace", "NoBlock", "PercClose"
          If GetStrCount("THEN {", txtExpr.Text) > GetStrCount("ELSE {", txtExpr.Text) Then
            GoTo TrimExitHere
          End If
      End Select
    Case "LogicENDIF"
      Select Case PrevType
        Case "Param", "ExprSeq", "NoBlock", "LogicENDIF", "OprCloseBrace", "PercClose"
          Pos = InStrRev(txtExpr.Text, "{", , vbTextCompare)
          If Pos <> -1 Then   'LogicOpen should be must
            'All open paranthesis should be closed before the logic block closed
            If GetStrCount("(", Mid$(txtExpr.Text, Pos)) = GetStrCount(")", Mid$(txtExpr.Text, Pos)) Then
              'All Perc Block should be closed before the logic block closed
              If GetStrCount(PercBlockOpen, Mid$(txtExpr.Text, Pos)) = GetStrCount(PercBlockClose, Mid$(txtExpr.Text, Pos)) Then
                'Closing Logic should not exceed the Logic open at any point of time
                If GetStrCount("{", txtExpr.Text) > GetStrCount("}", txtExpr.Text) Then
                  GoTo TrimExitHere
                End If
              End If
            End If
          End If
      End Select
    Case "LogicANDOR"
      Select Case PrevType
        Case "Param", "ExprSeq", "OprCloseBrace", "NoBlock", "PercClose"
          GoTo ExitHere
      End Select
    Case "LogicNOT"
      Select Case PrevType
        Case "LogicIF", "LogicTHEN", "LogicELSE", "LogicANDOR", "LogicCompare", "PercOpen"
          GoTo ExitHere
        Case "OprOpenBrace"
          GoTo TrimExitHere
      End Select
    Case "LogicCompare"
      Select Case PrevType
        Case "Param", "ExprSeq", "NoBlock", "OprCloseBrace", "PercClose"
          GoTo ExitHere
      End Select
    Case "Param", "ExprSeq"
      Select Case PrevType
        Case "OprOther", "LogicANDOR", "LogicNOT"
          GoTo ExitHere
        Case "NoPrevious", "OprOpenBrace", "LogicIF", "LogicTHEN", "LogicELSE", "PercOpen"
          GoTo TrimExitHere
      End Select
    Case "NoBlock"
      Select Case PrevType
        Case "OprOther", "LogicCompare", "LogicANDOR", "LogicNOT"
          If CheckValidNumber("", CurInputText) = True Then
            GoTo ExitHere
          End If
        Case "NoBlock"
          If CheckValidNumber(GetSprText(sprCurExprBuildSeq, eExBuildSeq.mBlock, sprCurExprBuildSeq.DataRowCnt), CurInputText) = True Then
            GoTo TrimExitHere
          End If
        Case "NoPrevious", "OprOpenBrace", "LogicIF", "LogicTHEN", "LogicELSE", "PercOpen"
          If CheckValidNumber("", CurInputText) = True Then
            GoTo TrimExitHere
          End If
      End Select
    Case "OprOther"
      Select Case PrevType
        Case "Param", "ExprSeq", "NoBlock", "OprCloseBrace", "LogicENDIF", "PercClose"
          GoTo ExitHere
      End Select
    Case "OprOpenBrace"
      Select Case PrevType
        Case "OprOther", "LogicANDOR", "LogicNOT", "LogicCompare"
          GoTo ExitHere
        Case "NoPrevious", "LogicIF", "LogicTHEN", "LogicELSE", "OprOpenBrace", "PercOpen"
          GoTo TrimExitHere
      End Select
    Case "OprCloseBrace"
      Select Case PrevType
        Case "OprCloseBrace", "Param", "ExprSeq", "NoBlock", "PercClose"
          If GetStrCount("(", txtExpr.Text) > GetStrCount(")", txtExpr.Text) Then
            GoTo TrimExitHere
          End If
      End Select
    Case "PercClose"
      Select Case PrevType
        Case "OprCloseBrace", "Param", "ExprSeq", "NoBlock", "PercClose"
          Pos = InStrRev(txtExpr.Text, PercBlockOpen, , vbTextCompare)
          If Pos > 0 Then
            'All the Open Paranthesis should be closed befote the Perc Block Closed
            If GetStrCount("(", Mid$(txtExpr.Text, Pos)) = GetStrCount(")", Mid$(txtExpr.Text, Pos)) Then
              'Count of Perc Open should be greater than Perc Close
              If GetStrCount(PercBlockOpen, txtExpr.Text) > GetStrCount(PercBlockClose, txtExpr.Text) Then
                GoTo TrimExitHere
              End If
            End If
          End If
      End Select
      
  End Select
  
Catch ex As Exception
  Call ErrMsgShow(ex)
  Resume TrimExitHere
Finally

  Exit Function
TrimExitHere:
  txtExpr.Text = Trim$(txtExpr.Text)
  GoTo ExitHere
End Function

Private Function ValidateExpr(ByVal aExpression As String,Optional ByVal aRow As Long) As Boolean
  Dim retVal As Variant
  
  Try
  ValidateExpr = False
  
  If Trim$(aExpression) = "" Then
    sShowMessage "No Valid Expression Found."
    GoTo ExitHere
  End If
  
  If ValidateExpression(aExpression) = False Then
    GoTo ExitHere
  End If

    sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, aRow, retVal
    Select Case Trim$(retVal)
      Case "OprOther", "OprOpeningBrace", "LogicIF", "LogicTHEN", "LogicELSE", "LogicANDOR", "LogicNOT", "LogicCompare"
        sShowMessage "Invalid Expression."
        GoTo ExitHere
    End Select
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub SetActiveTab(ByVal PrevTab As Integer,ByVal  CurTab As Integer)
  
' On Error Resume Next
  
    sstExpression.TabEnabled(PrevTab) = False
    sstExpression.TabEnabled(CurTab) = True
    sstExpression.Tab = CurTab
End Sub

Private Sub UpdateExpression()
  
  Try
  
  txtExpr.Tag = txtExpr.Text
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub RollBackExpression()
  
  Try
  
  txtExpr.Text = txtExpr.Tag
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function CheckBraces(ByVal Expression As String) As Boolean
  Dim BraceCnt As Integer
  Dim Cnt As Integer

  Try
  CheckBraces = False

  For Cnt = 1 To Len(Expression)
    Select Case Mid$(Expression, Cnt, 1)
      Case "("
        BraceCnt = BraceCnt + 1
      Case ")"
        BraceCnt = BraceCnt - 1
    End Select
    If BraceCnt < 0 Then
      sShowMessage "Paranthesis not properly matched."
      GoTo ExitHere
    End If
  Next Cnt

  If BraceCnt > 0 Then
    sShowMessage "Closing Paranthesis Missing."
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function CheckPercBlock(ByVal Expression As String) As Boolean
  Dim BraceCnt As Integer
  Dim Cnt As Integer

  Try
  CheckPercBlock = False

  For Cnt = 1 To Len(Expression)
    Select Case Mid$(Expression, Cnt, 1)
      Case PercBlockOpen
        BraceCnt = BraceCnt + 1
      Case PercBlockClose
        BraceCnt = BraceCnt - 1
    End Select
    If BraceCnt < 0 Then
      sShowMessage "Percentage Block Open & Close not properly matched."
      GoTo ExitHere
    End If
  Next Cnt

  If BraceCnt > 0 Then
    sShowMessage "Closing Percentage Block Missing."
    GoTo ExitHere
  End If

  CheckPercBlock = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function CheckValidNumber(ByVal CurBlock As String,ByVal  CurText As String) As Boolean
  
  Try
  CheckValidNumber = False
  
  CurBlock = Trim$(CurBlock)
  
  Select Case Len(CurBlock)
    Case 0
      If CurText = "." Then
        GoTo ExitHere
      End If
    Case 1
      If CurBlock = "0" Then
        If CurText <> "." Then
          GoTo ExitHere
        End If
      End If
    Case Else
      If CurText = "." Then
        If InStr(1, CurBlock, ".", vbTextCompare) > 0 Then
          GoTo ExitHere
        End If
      End If
  End Select
  
  CheckValidNumber = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function CheckExprNameSelect(ByVal enRow As Long) As Boolean
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try
  CheckExprNameSelect = False
  
    For rCnt = 1 To enRow
      sprExprBuilder.GetText eExBuild.mExName, rCnt, retVal
      If Trim$(retVal) = "" Then
        sShowMessage "Please Select Expression Name for the Row " & rCnt
        GoTo ExitHere
      End If
    Next rCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub ExprBuildButton_Status(ByVal Row As Long)
  Dim tmpStr As String
  
  Try

  tmpStr = Trim$(GetSprText(sprExprBuilder, eExBuild.mExName, Row))
  If Len(tmpStr) = 0 Then
    Call Spread_Lock(sprExprBuilder, True, eExBuild.mExBuildBut, Row)            'Build Button
  Else
    If Row > 1 Then
      Call Spread_Lock(sprExprBuilder, True, eExBuild.mExBuildBut, Row - 1)
    End If
    Call Spread_Lock(sprExprBuilder, False, eExBuild.mExBuildBut, Row)
    sprExprBuilder.MaxRows = sprExprBuilder.DataRowCnt + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtROffPosting_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select LedgerName, LedgerCode from FAS_LedgerMaster where left(LrRoute,4) not in ('0206', '0303') "
  Sql = Sql & " and isActive = '1' and CompCode = '" & gCompID & gBranchID & "'"
  Sql = Sql & " Union All"
  Sql = Sql & " Select 'As per Order', -1"
  Sql = Sql & " Order by LedgerName"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Accounts", "Account Name") = True Then
    txtROffPosting.Tag = Val(getvalue(2))    'AcCode
    txtROffPosting.Text = Trim$(getvalue(1))    'AcName
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtROffPosting_KeyUp(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If Shift = 0 And KeyCode = vbKeyF2 Then
    Call txtROffPosting_DblClick
  End If
End Sub

Private Sub txtROffValue_Change()
' On Error Resume Next
  
  lblROffPosting.Enabled = False
  txtROffPosting.Enabled = False
  
  If Trim$(txtROffValue) <> "" Then
    If IsNumeric(Trim$(txtROffValue)) = True Then
      lblROffPosting.Enabled = True
      txtROffPosting.Enabled = True
    End If
  End If
End Sub

Private Sub txtTemplate_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select FrmLName, FrmID from COM_FormulaMast_Hdr"
  Sql = Sql & " Where ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " Order by FrmLName"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formulae", "Formula Name") = True Then
    txtTemplate.Text = Trim$(getvalue(1))
    txtTemplate.Tag = Val(getvalue(2))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub PostingAcActive(ByVal ActiveStatus As Boolean,ByVal  fRow As Long)

  Try

  If ActiveStatus = False Then
    Call Spread_BkColour(sprExprBuilder, "&H555555", eExBuild.mAcName, fRow)
  Else
    Call Spread_BkColour(sprExprBuilder, "&HFFFFFF", eExBuild.mAcName, fRow)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function LoadExpressionForEdit(ByVal fExprSeq As Integer) As Boolean
  Dim rPos As Long
  Dim rCnt As Long
  Dim retVal As Variant
  Dim Block As String
  Dim ValueUsed As String
  Dim CurExpr As String
  Dim tempStr As String

  Try
  LoadExpressionForEdit = False
  
  If fExprSeq = 1 Then
    rPos = 1
  Else
    rPos = sprExprBuildSeq.SearchCol(eExBuildSeq.mExSeq, 1, -1, CStr(fExprSeq), SearchFlagsNone)
    If rPos <= 0 Then
      GoTo ExitHere
    End If
  End If

    sprExprBuildSeq.SetSelection 1, rPos, sprExprBuildSeq.MaxCols, sprExprBuildSeq.DataRowCnt
    sprExprBuildSeq.ClipboardCopy
    sprExprBuildSeq.ClearSelection  
    sprCurExprBuildSeq.SetSelection 1, 1, sprExprBuildSeq.MaxCols, (sprExprBuildSeq.DataRowCnt - rPos + 1)
    sprCurExprBuildSeq.ClipboardPaste
    sprCurExprBuildSeq.ClearSelection  
    CurExpr = ""
    
    For rCnt = 1 To sprCurExprBuildSeq.DataRowCnt
      sprCurExprBuildSeq.GetText eExBuildSeq.mBlockType, rCnt, retVal
      
      Select Case Trim$(retVal)
        Case "ExprSeq"
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, rCnt, retVal
          
          sprExprBuilder.GetText eExBuild.mExName, Val(retVal), retVal
          tempStr = "#" & Trim$(retVal) & "#"
        Case "Param"
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, rCnt, retVal
          
          rPos = sprParam.SearchCol(eParam.mParamID, 1, -1, CStr(Val(retVal)), SearchFlagsNone)
          If rPos > 0 Then
            sprParam.GetText eParam.mSName, rPos, retVal
            Block = Trim$(retVal)
            
            sprCurExprBuildSeq.GetText eExBuildSeq.mValUsed, rCnt, retVal
            If Val(retVal) <> 0 Then
              ValueUsed = Trim$(retVal)
            Else
              ValueUsed = ""
            End If
            
            sprCurExprBuildSeq.GetText eExBuildSeq.mIsPerc, rCnt, retVal
            If Val(retVal) = 1 Then
              tempStr = Block & "[" & ValueUsed & "%]"
            Else
              If ValueUsed = "" Then
                tempStr = Block
              Else
                tempStr = Block & "[" & ValueUsed & "]"
              End If
            End If
          Else
            GoTo ExitHere
          End If
        Case Else
          sprCurExprBuildSeq.GetText eExBuildSeq.mBlock, rCnt, retVal
          tempStr = Trim$(retVal)
      End Select
        
      sprCurExprBuildSeq.SetText eExBuildSeq.mPrevCharEnd, rCnt, Len(CurExpr)
      
      If Trim$(tempStr) = ")" Then
        CurExpr = CurExpr & tempStr
      Else
        If Right$(CurExpr, 1) = "(" Then
          CurExpr = CurExpr & tempStr
        Else
         CurExpr = CurExpr & " " & tempStr
        End If
      End If
      
    Next rCnt
  
  txtExpr.Text = CurExpr
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprExprSelect(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprExprSelect


sprExprSelect.DblClick +=  new EventHandler(sprExprSelect_DblClick)

End Sub

Private Sub Design_sprParam(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprParam


sprParam.DblClick +=  new EventHandler(sprParam_DblClick)

End Sub

Private Sub Design_sprExprBuilder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprExprBuilder


sprExprBuilder.ButtonClicked +=  new EventHandler(sprExprBuilder_ButtonClicked)
sprExprBuilder.ComboSelChange +=  new EventHandler(sprExprBuilder_ComboSelChange)
sprExprBuilder.DblClick +=  new EventHandler(sprExprBuilder_DblClick)
sprExprBuilder.EditChange +=  new EventHandler(sprExprBuilder_EditChange)
sprExprBuilder.KeyUp +=  new EventHandler(sprExprBuilder_KeyUp)

End Sub

Private Sub Design_sprExprBuildSeq(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprExprBuildSeq

End Sub

Private Sub Design_sprCurExprBuildSeq(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCurExprBuildSeq

End Sub

Private Sub InitializeSpreads()

	Call Design_sprExprSelect("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprBuilder.frm", sprExprSelect, "")

	Call Design_sprParam("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprBuilder.frm", sprParam, "")

	Call Design_sprExprBuilder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprBuilder.frm", sprExprBuilder, "")

	Call Design_sprExprBuildSeq("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprBuilder.frm", sprExprBuildSeq, "")

	Call Design_sprCurExprBuildSeq("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprBuilder.frm", sprCurExprBuildSeq, "")

End Sub
