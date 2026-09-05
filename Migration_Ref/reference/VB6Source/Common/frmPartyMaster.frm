VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPartyMaster 
   Caption         =   "Vendor Master"
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
   Begin VB.TextBox tbPartyCode1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   105
      MaxLength       =   10
      TabIndex        =   67
      ToolTipText     =   "Party Code"
      Top             =   6270
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame fraAddress 
      Enabled         =   0   'False
      Height          =   4515
      Left            =   960
      TabIndex        =   40
      Top             =   1500
      Width           =   14490
      Begin VB.TextBox txtUdyamNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11340
         MaxLength       =   30
         TabIndex        =   104
         ToolTipText     =   "Enter the GST Registration No."
         Top             =   945
         Width           =   2985
      End
      Begin VB.TextBox txtUdyamType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11340
         MaxLength       =   20
         TabIndex        =   103
         ToolTipText     =   "Enter the GST Registration No."
         Top             =   585
         Width           =   2370
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   100
         ToolTipText     =   "Enter the City"
         Top             =   1635
         Width           =   2100
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   99
         ToolTipText     =   "Enter the City"
         Top             =   1275
         Width           =   2100
      End
      Begin VB.Frame fraDetails 
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   4080
         TabIndex        =   74
         Top             =   2025
         Width           =   5940
         Begin VB.TextBox txtexregno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   77
            ToolTipText     =   "Enter the Excise Reg. No."
            Top             =   1050
            Width           =   2415
         End
         Begin VB.TextBox txtPanNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   50
            TabIndex        =   82
            ToolTipText     =   "Enter the Pan No"
            Top             =   360
            Width           =   2370
         End
         Begin VB.TextBox txtTinNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3120
            MaxLength       =   50
            TabIndex        =   81
            ToolTipText     =   "Enter the Area Code"
            Top             =   15
            Width           =   2370
         End
         Begin VB.CheckBox ChkSub 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "SubContractor"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2655
            TabIndex        =   80
            ToolTipText     =   "Select as Subcontractor"
            Top             =   1830
            Width           =   1350
         End
         Begin VB.TextBox txtAName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   10
            TabIndex        =   79
            ToolTipText     =   "Enter the Alias Name"
            Top             =   1755
            Width           =   1005
         End
         Begin VB.TextBox txtContactPerson 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   50
            TabIndex        =   78
            ToolTipText     =   "Enter the Contact Person"
            Top             =   1395
            Width           =   3750
         End
         Begin VB.TextBox txteccno 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   76
            ToolTipText     =   "Enter the ECC No."
            Top             =   705
            Width           =   1950
         End
         Begin VB.TextBox txtareacd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   75
            ToolTipText     =   "Enter the Area Code"
            Top             =   15
            Width           =   1005
         End
         Begin MSComCtl2.DTPicker dpexregdt 
            Height          =   315
            Left            =   4470
            TabIndex        =   83
            ToolTipText     =   "Enter the Excise Reg. Date"
            Top             =   1065
            Width           =   1380
            _ExtentX        =   2434
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
            CheckBox        =   -1  'True
            Format          =   115539969
            CurrentDate     =   37246
         End
         Begin VB.Label lblPanNo 
            AutoSize        =   -1  'True
            Caption         =   "Pan No"
            Height          =   210
            Left            =   960
            TabIndex        =   91
            Top             =   405
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tin No"
            Height          =   210
            Left            =   2625
            TabIndex        =   90
            Top             =   75
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Alias Name"
            Height          =   210
            Left            =   660
            TabIndex        =   89
            Top             =   1800
            Width           =   810
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Contact Person"
            Height          =   210
            Left            =   360
            TabIndex        =   88
            Top             =   1440
            Width           =   1110
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Dated"
            Height          =   210
            Left            =   4020
            TabIndex        =   87
            Top             =   1110
            Width           =   420
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Excise Reg. No"
            Height          =   210
            Left            =   375
            TabIndex        =   86
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   210
            Left            =   690
            TabIndex        =   85
            Top             =   60
            Width           =   780
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "ECC No"
            Height          =   210
            Left            =   930
            TabIndex        =   84
            Top             =   750
            Width           =   540
         End
      End
      Begin VB.Frame fraGST 
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   4110
         TabIndex        =   92
         Top             =   1275
         Visible         =   0   'False
         Width           =   5925
         Begin VB.TextBox txtProvisionalID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   94
            ToolTipText     =   "Enter the Provisional ID"
            Top             =   375
            Width           =   2145
         End
         Begin VB.TextBox txtGSTRegNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   93
            ToolTipText     =   "Enter the GST Registration No."
            Top             =   30
            Width           =   2370
         End
         Begin MSComCtl2.DTPicker dtpGSTDate 
            Height          =   315
            Left            =   4455
            TabIndex        =   95
            ToolTipText     =   "Enter the GST Date"
            Top             =   30
            Width           =   1380
            _ExtentX        =   2434
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/mm/yyyy"
            Format          =   115539969
            CurrentDate     =   37246
         End
         Begin VB.Label lblProvisionalID 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Provisional ID"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   450
            TabIndex        =   98
            Top             =   420
            Width           =   960
         End
         Begin VB.Label lblGSTRegDate 
            AutoSize        =   -1  'True
            Caption         =   "Dated"
            Height          =   210
            Left            =   3990
            TabIndex        =   97
            Top             =   60
            Width           =   420
         End
         Begin VB.Label lblGSTRegNo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GST Reg. No"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   525
            TabIndex        =   96
            Top             =   105
            Width           =   930
         End
      End
      Begin VB.TextBox txtsupliercode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5610
         MaxLength       =   10
         TabIndex        =   70
         ToolTipText     =   "Enter the Address"
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox tbLst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5595
         MaxLength       =   25
         TabIndex        =   52
         ToolTipText     =   "Enter the LST No."
         Top             =   945
         Width           =   2370
      End
      Begin VB.TextBox tbCst 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5595
         MaxLength       =   25
         TabIndex        =   51
         ToolTipText     =   "Enter the CST No."
         Top             =   585
         Width           =   2370
      End
      Begin VB.TextBox tbEmail 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   50
         ToolTipText     =   "Enter the E-Mail Id"
         Top             =   3750
         Width           =   2775
      End
      Begin VB.TextBox tbCell 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   25
         TabIndex        =   49
         ToolTipText     =   "Enter the  Cell No."
         Top             =   3390
         Width           =   2340
      End
      Begin VB.TextBox tbFax 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   25
         TabIndex        =   48
         ToolTipText     =   "Enter the Fax No."
         Top             =   3045
         Width           =   2340
      End
      Begin VB.TextBox tbPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   47
         ToolTipText     =   "Enter the Phone No."
         Top             =   2700
         Width           =   2340
      End
      Begin VB.TextBox tbPin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   46
         ToolTipText     =   "Enter the Pincode"
         Top             =   2355
         Width           =   930
      End
      Begin VB.TextBox tbAdd2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   30
         TabIndex        =   45
         ToolTipText     =   "Enter the Address"
         Top             =   570
         Width           =   2790
      End
      Begin VB.TextBox tbAdd3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   30
         TabIndex        =   44
         ToolTipText     =   "Enter the Address"
         Top             =   915
         Width           =   2790
      End
      Begin VB.TextBox tbCity 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   43
         ToolTipText     =   "Enter the City"
         Top             =   1995
         Width           =   1875
      End
      Begin VB.TextBox tbAdd1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   30
         TabIndex        =   42
         ToolTipText     =   "Enter the Address"
         Top             =   225
         Width           =   2790
      End
      Begin VB.TextBox txtVendorCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1155
         MaxLength       =   10
         TabIndex        =   41
         ToolTipText     =   "Enter the Address"
         Top             =   4110
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker dtpCSTDate 
         Height          =   315
         Left            =   8550
         TabIndex        =   53
         ToolTipText     =   "Enter the CST Date"
         Top             =   570
         Width           =   1380
         _ExtentX        =   2434
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   115539969
         CurrentDate     =   37246
      End
      Begin MSComCtl2.DTPicker dtpLSTDate 
         Height          =   315
         Left            =   8550
         TabIndex        =   54
         ToolTipText     =   "Enter the LST Date"
         Top             =   930
         Width           =   1380
         _ExtentX        =   2434
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   115539969
         CurrentDate     =   37246
      End
      Begin MSComCtl2.DTPicker dtpUdyamDate 
         Height          =   315
         Left            =   11340
         TabIndex        =   105
         ToolTipText     =   "Enter the GST Date"
         Top             =   1305
         Width           =   1380
         _ExtentX        =   2434
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   115539969
         CurrentDate     =   37246
      End
      Begin VB.Label lblUdyamDate 
         AutoSize        =   -1  'True
         Caption         =   "Udyam Reg. Date"
         Height          =   210
         Left            =   10035
         TabIndex        =   108
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblUdyamNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Udyam No."
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10515
         TabIndex        =   107
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblUdyamType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Udyam Type"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10380
         TabIndex        =   106
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "State"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   720
         TabIndex        =   102
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Country"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   510
         TabIndex        =   101
         Top             =   1350
         Width           =   570
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Suplier Code"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4635
         TabIndex        =   71
         Top             =   285
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dated"
         Height          =   210
         Left            =   8085
         TabIndex        =   66
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Dated"
         Height          =   210
         Left            =   8085
         TabIndex        =   65
         Top             =   990
         Width           =   420
      End
      Begin VB.Label lblLst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LST No"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5025
         TabIndex        =   64
         Top             =   990
         Width           =   525
      End
      Begin VB.Label lblCst 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CST No"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5010
         TabIndex        =   63
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E- Mail ID"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   405
         TabIndex        =   62
         Top             =   3795
         Width           =   645
      End
      Begin VB.Label lblCell 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cell No"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   540
         TabIndex        =   61
         Top             =   3435
         Width           =   495
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fax No"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   525
         TabIndex        =   60
         Top             =   3090
         Width           =   510
      End
      Begin VB.Label lblPhone 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Phone No"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   390
         TabIndex        =   59
         Top             =   2745
         Width           =   690
      End
      Begin VB.Label lblPin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pincode"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   495
         TabIndex        =   58
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label lblCity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "City"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   780
         TabIndex        =   57
         Top             =   2040
         Width           =   270
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   450
         TabIndex        =   56
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblVendorCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendor Code"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   55
         Top             =   4095
         Width           =   960
      End
   End
   Begin VB.Frame fraAccount 
      Height          =   1455
      Left            =   1815
      TabIndex        =   24
      Top             =   45
      Width           =   7935
      Begin VB.TextBox txtRegisteredName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   105
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   72
         ToolTipText     =   "Enter the Registered  Name"
         Top             =   840
         Width           =   7665
      End
      Begin VB.ComboBox cmbAcType 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":0000
         Left            =   4710
         List            =   "frmPartyMaster.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select the Account Type"
         Top             =   315
         Width           =   1380
      End
      Begin VB.TextBox tbPartyName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   105
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Enter the Customer / Vendor Name. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   4575
      End
      Begin VB.ComboBox cmbStatus 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":0004
         Left            =   6090
         List            =   "frmPartyMaster.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select the Status"
         Top             =   315
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   7230
         TabIndex        =   3
         ToolTipText     =   "Click to OK"
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblRegisteredName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered  Name"
         Height          =   210
         Left            =   105
         TabIndex        =   73
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "* Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   27
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblAcType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Account Type"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   4710
         TabIndex        =   26
         ToolTipText     =   "Select Customer / Vendor / Party"
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   210
         Left            =   6090
         TabIndex        =   25
         Top             =   135
         Width           =   465
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6495
      TabIndex        =   22
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   9015
      Width           =   1300
   End
   Begin VB.Frame fraTerms 
      Enabled         =   0   'False
      Height          =   3030
      Left            =   1500
      TabIndex        =   28
      Top             =   5940
      Width           =   8745
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   13
         ToolTipText     =   "Enter the Excise Duty  Registration Number"
         Top             =   1410
         Width           =   2580
      End
      Begin VB.ComboBox cmbCreditDaysPeriod 
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":0024
         Left            =   1065
         List            =   "frmPartyMaster.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1402
         Width           =   1695
      End
      Begin VB.CommandButton cmdDefValue 
         Caption         =   "Restrict Defined Values"
         Height          =   315
         Index           =   1
         Left            =   4455
         TabIndex        =   19
         Top             =   2610
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.CommandButton cmdParamValue 
         Caption         =   "Parameter Values"
         Height          =   315
         Index           =   1
         Left            =   2895
         TabIndex        =   18
         Top             =   2610
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdDefValue 
         Caption         =   "Restrict Defined Values"
         Height          =   315
         Index           =   0
         Left            =   4455
         TabIndex        =   16
         Top             =   2085
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.CommandButton cmdParamValue 
         Caption         =   "Parameter Values"
         Height          =   315
         Index           =   0
         Left            =   2895
         TabIndex        =   15
         Top             =   2085
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtVendorFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   17
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help."
         Top             =   2610
         Width           =   2760
      End
      Begin VB.ComboBox CmbFrom 
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":006B
         Left            =   6600
         List            =   "frmPartyMaster.frx":006D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select the Form"
         Top             =   315
         Width           =   2070
      End
      Begin VB.ComboBox CmbMode 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":006F
         Left            =   1890
         List            =   "frmPartyMaster.frx":0071
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Select the Mode"
         Top             =   855
         Width           =   2010
      End
      Begin VB.ComboBox CmbInsurance 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":0073
         Left            =   1875
         List            =   "frmPartyMaster.frx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select the Insurance"
         Top             =   315
         Width           =   2010
      End
      Begin VB.ComboBox CmbTerms 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":0077
         Left            =   3885
         List            =   "frmPartyMaster.frx":0079
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Select the Payment Term"
         Top             =   855
         Width           =   4785
      End
      Begin VB.ComboBox CmbDelIns 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmPartyMaster.frx":007B
         Left            =   3885
         List            =   "frmPartyMaster.frx":007D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select the Delivery instruction"
         Top             =   315
         Width           =   2685
      End
      Begin VB.ComboBox CmbPackFoward 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select the  Packing & Forwarding"
         Top             =   315
         Width           =   1785
      End
      Begin VB.ComboBox CmbFreight 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select the Freight"
         Top             =   855
         Width           =   1800
      End
      Begin VB.TextBox txtCreditDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         MaxLength       =   30
         TabIndex        =   11
         ToolTipText     =   "Enter the Credit Days"
         Top             =   1410
         Width           =   915
      End
      Begin VB.TextBox txtCustomerFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help."
         Top             =   2085
         Width           =   2760
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Currency"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2790
         TabIndex        =   69
         Top             =   1215
         Width           =   675
      End
      Begin VB.Label lblCreditDaysPeriod 
         AutoSize        =   -1  'True
         Caption         =   "From the Date of"
         Height          =   195
         Left            =   1065
         TabIndex        =   39
         Top             =   1208
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   38
         Top             =   2430
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mode"
         Height          =   210
         Left            =   1890
         TabIndex        =   37
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Insurance"
         Height          =   210
         Left            =   1875
         TabIndex        =   36
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Form"
         Height          =   210
         Left            =   6600
         TabIndex        =   35
         Top             =   135
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Packing && Forwarding"
         Height          =   210
         Left            =   90
         TabIndex        =   34
         Top             =   135
         Width           =   1590
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Freight"
         Height          =   210
         Left            =   90
         TabIndex        =   33
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Terms of Payment"
         Height          =   210
         Left            =   3885
         TabIndex        =   32
         Top             =   675
         Width           =   1305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Instruction"
         Height          =   210
         Left            =   3885
         TabIndex        =   31
         Top             =   135
         Width           =   1380
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Credit Days"
         Height          =   210
         Left            =   90
         TabIndex        =   30
         Top             =   1215
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   90
         TabIndex        =   29
         Top             =   1905
         Width           =   1755
      End
   End
   Begin VB.CommandButton cbClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5205
      TabIndex        =   21
      ToolTipText     =   "Click to Clear the Data"
      Top             =   9015
      Width           =   1300
   End
   Begin VB.CommandButton cbSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3915
      TabIndex        =   20
      ToolTipText     =   "Click to Save the Data"
      Top             =   9015
      Width           =   1300
   End
   Begin MSComCtl2.DTPicker dtpPrevEffDt 
      Height          =   300
      Left            =   9015
      TabIndex        =   23
      Top             =   7965
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   115539971
      CurrentDate     =   38162
   End
   Begin VB.Label lblCode1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   30
      TabIndex        =   68
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmPartyMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim CurDate As Date

Private Enum eStatus
  mCheckUnSuccess = 0
  mNotExists = 1
  mFASOnly = 2
  mPurSalOnly = 3
  mBothAvailable = 4
End Enum

Private Enum eFrmCmd
  mCustomer = 0
  mVendor = 1
End Enum

Private Sub cbclear_Click()
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  Call ClearForm(Me)
  Call CmbLoad
  Call FillAccType
  
  'Get ServerDate
  CurDate = ServerDate
  
  dtpUdyamDate.Value = Format(CurDate, "dd/MM/yyyy")
  
  cmbAcType.Enabled = True
  fraAddress.Enabled = False
  fraTerms.Enabled = False
  cbSave.Enabled = False
  fraAccount.Enabled = True
  cmbStatus.Enabled = False
 
  cmbStatus.ListIndex = 1
  cmbStatus.Tag = ""
  
  tbPartyName.Tag = ""
  tbPartyName.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cbsave_Click()
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim CurDate As Date
  Dim LCode As Integer
  Dim PartyCode As Integer
  Dim GrName As String
  Dim ChkStatus As eStatus
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = Format(ServerDate, "dd/MM/yyyy HH:mm:ss")
  
  If ValidateData(CurDate) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
' Check whether the vendor already exists or not for New Vendor Creation
'  GrName = Trim$(ValidateVendor)
  
  ChkStatus = CheckPartyStatus
  
  If ChkStatus = mCheckUnSuccess Then
    sShowMessage "The availability of this party in FAS module Checking failed."
    GoTo ExitHere
  End If
  
  If ChkStatus = mNotExists Then
    tbPartyName.Tag = GenCode
  End If
  
  If ChkStatus = mFASOnly Or ChkStatus = mBothAvailable Then
    Sql = "Select LedgerCode, LrRoute from FAS_LedgerMaster Where LedgerName='" & Trim$(tbPartyName) & "' and CompCode='" & Left(gFasCompCode, 4) & "'"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not snap.EOF Then
      If snap(0) = 0 Then
        sShowMessage "Some Error Occurred while fetching FAS Code."
        GoTo ExitHere
      Else
        If ChkStatus = mBothAvailable Then
          If Val(tbPartyName.Tag) <> snap(0) Then
            sShowMessage "The Party Code and FAS Ledger Code does not Match. Contact System Administrator."
            GoTo ExitHere
          End If
        Else
          tbPartyName.Tag = snap(0)
        End If
      End If
    End If
  End If
  
  If ChkStatus = mPurSalOnly Then
    Sql = "Select LedgerName from FAS_LedgerMaster Where LedgerCode=" & Val(tbPartyName.Tag) & " and CompCode='" & Left(gFasCompCode, 4) & "'"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then
      GoTo ExitHere
    End If
    
    If Not snap.EOF Then
      sShowMessage "The Party Code already assigned for " & snap("LedgerName") & " in FAS Ledger Master. Contact System Administrator."
      GoTo ExitHere
    End If
  End If
  
  If ChkStatus = mFASOnly Or ChkStatus = mBothAvailable Then
    If SaveLedgerMaster(CurDate, Val(tbPartyName.Tag), True) = False Then
 GoTo ExitHere
End If
    
    If Trim$(tbAdd1) <> "" Or Trim$(txtTinNo.Text) <> "" Then
      If SaveLedgerAddress(CurDate, Val(tbPartyName.Tag), True) = False Then
        GoTo ExitHere
      End If
    End If
  Else
    If SaveLedgerMaster(CurDate, Val(tbPartyName.Tag), False) = False Then
 GoTo ExitHere
End If
    
    If Trim$(tbAdd1) <> "" Or Trim$(txtTinNo.Text) <> "" Then
      If SaveLedgerAddress(CurDate, Val(tbPartyName.Tag), False) = False Then
        GoTo ExitHere
      End If
    End If
  End If
  

  'to generate the AC_Code here
  If Trim$(txtVendorCode.Text) = "" Then
    If Trim$(cmbAcType.Text) = "Creditor" Then
      Sql = "Select Max(VendorCode) from FAS_PARTY_MASTER"
      Sql = Sql & " where AcType = 'P' and PartyName like '" + Trim$(Left(tbPartyName, 1)) & "%" + "'"
      Sql = Sql & " and Branchid ='" & gBranchID & "' and Compid= '" & gCompID & "'"
      If RSOpen(snap, Sql) = False Then
        GoTo ExitHere
      End If
      
      If snap.EOF Or (IsNull(snap(0)) Or snap(0) = "") Then
        txtVendorCode = "V" & UCase(Trim$(Left(tbPartyName, 1))) & "001"
      Else
        txtVendorCode = "V" & Trim$(Left(tbPartyName, 1)) & Format(Right(IIf(IsNull(Trim$(snap(0))), 0, snap(0)), 3) + 1, "000")
      End If
      snap.Close
    ElseIf cmbAcType = "Debtor" Then
      Sql = "Select Max(VendorCode) from FAS_PARTY_MASTER"
      Sql = Sql & " where AcType = 'D' and PartyName like '" & Trim$(Left(tbPartyName, 1)) & "%" & "'"
      Sql = Sql & " and Branchid ='" & gBranchID & "' and Compid= '" & gCompID & "'"
      If RSOpen(snap, Sql) = False Then
        GoTo ExitHere
      End If
      If snap.EOF Or IsNull(snap(0)) Then
        txtVendorCode = Trim$(Left(tbPartyName, 1)) & "01"
      Else
        txtVendorCode = Trim$(Left(tbPartyName, 1)) & Format(IIf(snap(0) = "", 0, Right(snap(0), 2)) + 1, "00")
      End If
      snap.Close
    End If
  End If
  
  If SavePartyMaster = False Then
    GoTo ExitHere
  End If
  
  If SavePartyAddress = False Then
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved Successfully"
  Call cbclear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbAcType_Click()

  Try
  Me.MousePointer = vbHourglass

  If cmbAcType.Text = "Debtor" Then
    ChkSub.Enabled = False
    txtCustomerFormula.Enabled = True
    txtVendorFormula.Enabled = False
    lblVendorCode.Caption = "Customer Code"
  ElseIf cmbAcType.Text = "Creditor" Then
    ChkSub.Enabled = True
    txtCustomerFormula.Enabled = False
    txtVendorFormula.Enabled = True
    lblVendorCode.Caption = "Vendor Code"
  ElseIf cmbAcType.Text = "Party" Then
    ChkSub.Enabled = True
    txtCustomerFormula.Enabled = True
    txtVendorFormula.Enabled = True
    lblVendorCode.Caption = "Vendor Code"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdOk_Click()

  Try
  Me.MousePointer = vbHourglass

  If Trim$(tbPartyName) = "" Then
    sShowMessage "Please enter the Account Name"
    tbPartyName.SetFocus
    GoTo ExitHere
  End If
  
  fraAddress.Enabled = True
  tbAdd1.SetFocus
  fraAccount.Enabled = False
  
  Call AddDetails
  
  fraTerms.Enabled = True
  cbSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub
Private Sub cmdParamValue_Click(ByVal Index As Integer)
  Dim fFormulaID As Integer
  Dim fUserKey As String
  Dim fPrevUserKey As String
  Dim ScreenFrom As String
  Dim MultiItemList(0) As String
  Dim IncludeSupplyThruCode As Boolean

  Try
  Me.MousePointer = vbHourglass

  If Index = eFrmCmd.mCustomer Then
    If Val(txtCustomerFormula.Tag) <= 0 Then
      sShowMessage "Please Select the Customer Formula."
      GoTo ExitHere
    End If
    
    fFormulaID = Val(txtCustomerFormula.Tag)
    ScreenFrom = "Vendor Master - Customer"
    
  ElseIf Index = eFrmCmd.mVendor Then
    If Val(txtVendorFormula.Tag) <= 0 Then
      sShowMessage "Please Select the Vendor Formula."
      GoTo ExitHere
    End If
    
    fFormulaID = Val(txtVendorFormula.Tag)
    ScreenFrom = "Vendor Master - Vendor"
    
  End If

  IncludeSupplyThruCode = False
  MultiItemList(0) = "0-0-0"

  If frmPromptValueEntry.GetPromptValuesHelp(Val(txtAName.Tag), fFormulaID, fUserKey, fPrevUserKey, ScreenFrom, MultiItemList, IncludeSupplyThruCode) = vbFalse Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dpexregdt_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyReturn And Shift = 0 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub dtpCSTDate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyReturn And Shift = 0 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub dtpGSTDate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyReturn And Shift = 0 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub dtpLSTDate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyReturn And Shift = 0 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next

  If TypeOf Me.ActiveControl Is FPSpreadADO.fpSpread Then
    Exit Sub
  Else
    If KeyAscii = vbKeyReturn Then
      SendKeys "{tab}"
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim CurDate As Date
  Try

  Call FillAccType
  cmbStatus.ListIndex = 1
  
  If CmbLoad = False Then GoTo ExitHere
  CurDate = ServerDate
  dtpCSTDate = Format(CurDate, "dd/MM/yyyy")
  dtpLSTDate = dtpCSTDate
  dtpGSTDate = dtpCSTDate
  dtpPrevEffDt = dtpCSTDate
  dpexregdt = dtpCSTDate
  dtpUdyamDate.Value = Format(CurDate, "dd/MM/yyyy")
  dtpCSTDate = ""
  dtpLSTDate = ""
  dtpGSTDate = ""
  dpexregdt = ""
  tbPartyName.Tag = ""
  
'  tbPartyName.Tag = GenCode         'Generate LedgerCode

  'Show GST Details based on ERP Option
  If IsGSTRequired = True Then
    fraGST.Visible = True
    fraDetails.Top = 1575
  End If
  
  dtpCSTDate = Null
  dtpLSTDate = Null
  dpexregdt = Null
  dtpGSTDate = Null

  If gModuleCode = eMdlCode.mSAL Then
    Me.Caption = "Customer Master"
    txtVendorCode.Enabled = True
    lblVendorCode.Enabled = True
    ChkSub.Visible = False
  ElseIf gModuleCode = eMdlCode.mPUR Then
    Me.Caption = "Vendor Master"
    If gReportID <> "PMI" Then
      txtVendorCode.Enabled = False
      lblVendorCode.Enabled = False
    End If
    ChkSub.Visible = True
  Else
    Unload Me
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  Call ClearTmpPromptTable
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub tbCell_KeyPress(ByVal KeyAscii As Integer)

  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumChkSpl(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbCity_DblClick()
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
    
  CurDate = ServerDate
  
  'Validate Country
  If contrl_valid(txtState, "Please Select the state.") = False Then
 GoTo ExitHere
End If
      
  Sql = " Execute FAS_PartyMaster_Hlp_City"
  Sql = Sql & " @StateCode=" & Val(txtState.Tag) & ","
  Sql = Sql & " @AsOnDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', "
  Sql = Sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', "
  Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "'"

  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of City's", "City Name") = True Then
    tbCity.Tag = getvalue(2)
    tbCity.Text = Trim$(getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbCity_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call tbCity_DblClick
  End If
  
End Sub

Private Sub tbCity_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = CheckAlphaNumeric(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbFax_KeyPress(ByVal KeyAscii As Integer)

  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumChkSpl(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbPartyName_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
''  If SetServerDate(dtpAsOnDate) < FinStart Then: GoTo ExitHere
  

  Sql = "Select PartyName, PartyCode, VendorCode, AcType, EffStatus, EffDate ,SuplierCode"
  Sql = Sql & " from FAS_PARTY_MASTER"
  Sql = Sql & " where Branchid ='" & gBranchID & "' and Compid ='" & gCompID & "'"
  If gModuleCode = eMdlCode.mSAL Then
    Sql = Sql & " and actype in ('D','Y') and EffStatus = 1"
  Else
    Sql = Sql & " and actype in ('P','Y') and EffStatus = 1"
  End If
  Sql = Sql & " Order by PartyName"
  
  If FetchValue_New(False, db.ConnectionString, Sql, "100000", "List of Accounts", "Account Name") = True Then
    tbPartyName = getvalue(1)
    tbPartyName.Tag = getvalue(2)
    txtVendorCode.Text = getvalue(3)
    cmbAcType.Text = AcType(CStr(UCase(getvalue(4))))
    cmbAcType.Enabled = False
    cmbStatus.Enabled = True
    cmbStatus.ListIndex = getvalue(5)
cmbStatus.Tag = getvalue(5)
    dtpPrevEffDt.Value = Format(getvalue(6), "yyyy-MM-dd")
    txtsupliercode = getvalue(7)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbPartyName_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 And Shift = 0 Then
    Call tbPartyName_DblClick
  End If
End Sub

Private Sub tbPhone_KeyPress(ByVal KeyAscii As Integer)

  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumChkSpl(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub tbPin_KeyPress(ByVal KeyAscii As Integer)

  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumValChk(KeyAscii)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCountry_DblClick()
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  Sql = " Execute FAS_PartyMaster_Hlp_Country"
  Sql = Sql & " @AsOnDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @BranchID = '" & gBranchID & "', "
  Sql = Sql & " @CompID = '" & gCompID & "'"
  
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Country's", "Country Name") = True Then
    txtCountry.Text = Trim$(getvalue(1))
    
    If Val(txtCountry.Tag) <> getvalue(2) Then
      txtCountry.Tag = getvalue(2)
      txtState.Text = ""
txtState.Tag = ""

      tbCity.Text = ""
tbCity.Tag = ""
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCountry_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txtCountry_DblClick
  End If
  
End Sub

Private Sub txtCreditDays_KeyPress(ByVal KeyAscii As Integer)

  Try
  Me.MousePointer = vbHourglass

  KeyAscii = NumValInt(txtCreditDays, KeyAscii, 999, 0)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCurrency_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select CurrencyCode, CurrencyName from COM_Currency_Master"
  If FetchValue_New(False, db.ConnectionString, Sql, "11", "List of Currency", "Currency Code", "Currency Name") = True Then
    txtCurrency.Tag = getvalue(1)
    txtCurrency.Text = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCurrency_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
 Call txtCurrency_DblClick
End If
End Sub
Private Sub txtCustomerFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
'  Sql = " select distinct FormulaName, FormulaCode"
'  Sql = Sql & " from COM_FORMULA_MASTER"
'  Sql = Sql & " where Branchid = '" & gBranchID & "' and Compid = '" + gCompID + "'"

  Sql = " select distinct FrmLName,FrmID from COM_FormulaMast_Hdr where ModuleCode = " & eMdlCode.mSAL & ""
  Sql = Sql + " and isActive = '1' and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula's", "Formula Name") = True Then
    txtCustomerFormula.Tag = getvalue(2)
    txtCustomerFormula.Text = getvalue(1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCustomerFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 And Shift = 0 Then
    Call txtCustomerFormula_DblClick
  End If
End Sub

Private Sub txtState_DblClick()
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
    
  CurDate = ServerDate
  
  'Validate Country
  If contrl_valid(txtCountry, "Please Select the Country.") = False Then
 GoTo ExitHere
End If
      
  Sql = " Execute FAS_PartyMaster_Hlp_State"
  Sql = Sql & " @CountryCode=" & Val(txtCountry.Tag) & ","
  Sql = Sql & " @AsOnDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', "
  Sql = Sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', "
  Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "'"

  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of State", "State Name") = True Then
    txtState.Text = Trim$(getvalue(1))
    
     If Val(txtState.Tag) <> getvalue(2) Then
      txtState.Tag = getvalue(2)
      tbCity.Text = ""
tbCity.Tag = ""
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtState_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txtState_DblClick
  End If
  
End Sub

Private Sub txtVendorFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select distinct FrmLName, FrmID from COM_FormulaMast_Hdr where isActive = '1' "
  Sql = Sql + " and ModuleCode = " & eMdlCode.mPUR & " and Branchid = '" & gBranchID & "' "
  Sql = Sql + " and Compid = '" + gCompID + "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Purchase Formula's", "Formula Name") = True Then
    txtVendorFormula.Tag = getvalue(2)
    txtVendorFormula.Text = getvalue(1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtVendorFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 And Shift = 0 Then
    Call txtVendorFormula_DblClick
  End If
End Sub

' To get the Description for Acc Types
Private Function AcType(ByVal AccountType As String) As String

  Try

  Select Case AccountType
    Case "D"
        AcType = "Debtor"
    Case "P"
        AcType = "Creditor"
    Case "Y"
        AcType = "Party"
'    Case "I"
'        AcType = "Sales"
'    Case "E"
'        AcType = "Purchase"
'    Case "M"
'        AcType = "Indirect Income"
'    Case "N"
'        AcType = "Indirect Expense"
'    Case "A"
'        AcType = "Asset"
'    Case "L"
'        AcType = "Current Liability"
'    Case "O"
'        AcType = "Loans"
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub AddDetails()
  Dim mRec As AceADODB.Recordset
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim CurDate As Date
  Dim uPos As Long
  
  Try
  
  CurDate = ServerDate
  
  Sql = " select a.FormCode,(select b.Description from TERMSMASTER b"
  Sql = Sql + " where b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid and b.Compid=a.Compid) FormDesc,"
  Sql = Sql + " a.PackingCode,(select b1.Description from TERMSMASTER b1"
  Sql = Sql + " where b1.TypeFlg='Packing & Forwarding' and b1.Code=a.PackingCode and b1.Branchid=a.Branchid and b1.Compid=a.Compid) PackingDesc,"
  Sql = Sql + " a.DeliveryCode,(select b2.Description "
  Sql = Sql + " from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction' "
  Sql = Sql + " and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid "
  Sql = Sql + " and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,(select b3.Description "
  Sql = Sql + " from TERMSMASTER b3 where b3.TypeFlg='Payment Terms'"
  Sql = Sql + " and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid "
  Sql = Sql + " and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,(select "
  Sql = Sql + " b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance'"
  Sql = Sql + " and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid "
  Sql = Sql + " and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,(select "
  Sql = Sql + " b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid "
  Sql = Sql + " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,(select b6.Description "
  Sql = Sql + " from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode "
  Sql = Sql + " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc,"
  Sql = Sql + " a.IsSubcontractor,a.AliasName,a.EffStatus,a.EffDate,a.CreditDays, a.CreditDaysPeriod, "
  Sql = Sql + " a.CurrencyCode, (select b8.CurrencyName from COM_Currency_Master b8 where b8.CurrencyCode = a.CurrencyCode) CurrencyName,"
  Sql = Sql + " a.CustomerFormulaCode, (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 "
  Sql = Sql + " where b7.isActive = '1' and b7.ModuleCode = " & eMdlCode.mSAL & " and b7.FrmID=a.CustomerFormulaCode "
  Sql = Sql + " and b7.Branchid=a.Branchid and b7.Compid=a.Compid) CustomerFormulaName, a.VendorFormulaCode, "
  Sql = Sql + " (select top 1 b7.FrmLName from COM_FormulaMast_Hdr b7 where isActive = '1' "
  Sql = Sql + " and b7.ModuleCode = " & eMdlCode.mPUR & " and b7.FrmID=a.VendorFormulaCode and b7.Branchid=a.Branchid "
  Sql = Sql + " and b7.Compid=a.Compid) VendorFormulaName, PanNo,TinNo from FAS_PARTY_MASTER a where "
  Sql = Sql + " a.PartyCode = " & Val(tbPartyName.Tag) & " and a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "'"
  
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    txtAName = IIf(IsNull(snap("aliasname")), "", snap("aliasname"))
    ChkSub = IIf(snap("IsSubcontractor") = True, 1, 0)
    If IsNull(snap("CreditDays")) = False Then
 txtCreditDays = snap("CreditDays")
End If
    If Not IsNull(snap("CurrencyCode")) = True Then txtCurrency.Tag = snap("CurrencyCode")
    If Not IsNull(snap("CurrencyName")) = True Then txtCurrency = snap("CurrencyName")
    If Not IsNull(snap("CustomerFormulaCode")) = True Then txtCustomerFormula.Tag = snap("CustomerFormulaCode")
    If Not IsNull(snap("CustomerFormulaName")) = True Then txtCustomerFormula.Text = snap("CustomerFormulaName")
    If Not IsNull(snap("VendorFormulaCode")) = True Then txtVendorFormula.Tag = snap("VendorFormulaCode")
    If Not IsNull(snap("VendorFormulaName")) = True Then txtVendorFormula.Text = snap("VendorFormulaName")
    If Not IsNull(snap("DeliveryDesc")) = True Then CmbDelIns.Text = snap("DeliveryDesc")
    If Not IsNull(snap("FormDesc")) = True Then CmbFrom.Text = snap("FormDesc")
    If Not IsNull(snap("PackingDesc")) = True Then CmbPackFoward.Text = snap("PackingDesc")
    If Not IsNull(snap("InsuranceDesc")) = True Then CmbInsurance.Text = snap("InsuranceDesc")
    If Not IsNull(snap("FreightDesc")) = True Then CmbFreight.Text = snap("FreightDesc")
    If Not IsNull(snap("PaymentDesc")) = True Then CmbTerms.Text = snap("PaymentDesc")
    If Not IsNull(snap("ModeDesc")) = True Then CmbMode.Text = snap("ModeDesc")
    If Not IsNull(snap("PanNo")) = True Then txtTinNo.Text = Trim$(snap("PanNo"))
    If Not IsNull(snap("TinNo")) = True Then txtTinNo.Text = Trim$(snap("TinNo"))
    If Trim$(snap("CreditDaysPeriod")) <> "" Then
      cmbCreditDaysPeriod.Text = snap("CreditDaysPeriod")
    End If
  End If
  
  'Get Party Address Data
  Sql = " Execute FAS_PartyMaster_Fil_Address"
  Sql = Sql & " @PartyCode =" & Val(tbPartyName.Tag) & ","
  Sql = Sql & " @AsOnDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "'"
  
'  Sql = "Select a.*, isnull(b.CountryCode,0) as CountryCode,  isnull(b.StateCode,0) as StateCode, isnull(b.CityCode, 0) as CityCode "
'  Sql = Sql & " from FAS_PARTY_ADDRESS a"
'  Sql = Sql & " Left Outer Join fn_COM_Get_CountryStateCity_Dtl('" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "','" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & gBranchID & "', '" & gCompID & "') b"
'  Sql = Sql & " On a.State= b.StateName and a.Country=b.CountryName and a.City=b.CityName"
'  Sql = Sql & " Where PartyCode =" & Val(tbPartyName.Tag) & " and Compid = '" & gCompID & "'"
'  Sql = Sql & " and Branchid='" & gBranchID & "'"
  
  Set mRec = New AceADODB.Recordset
  If RSOpen(mRec, Sql) = False Then GoTo ExitHere
  
  If mRec.RecordCount = 0 Then
 mRec.Close
GoTo ExitHere
End If
  
  tbAdd1 = IIf(IsNull(mRec("Add1")), "", mRec("Add1"))
  tbAdd2 = IIf(IsNull(mRec("Add2")), "", mRec("Add2"))
  tbAdd3 = IIf(IsNull(mRec("Add3")), "", mRec("Add3"))
  tbPin = IIf(IsNull(mRec("Pincode")), "", mRec("Pincode"))
  
  'Country
  txtCountry.Tag = mRec("CountryCode")
  txtCountry = IIf(IsNull(mRec("Country")), "", mRec("Country"))
    
  'State
  txtState.Tag = mRec("StateCode")
  txtState = IIf(IsNull(mRec("State")), "", mRec("State"))
  
  'City
  tbCity.Tag = mRec("CityCode")
  tbCity = IIf(IsNull(mRec("City")), "", mRec("City"))
  
  tbLst = IIf(IsNull(mRec("Lst")), "", mRec("Lst"))
  dtpLSTDate = Format(mRec("lstdate"), "dd/mm/yyyy")
  tbCst = IIf(IsNull(mRec("Cst")), "", mRec("Cst"))
  dtpCSTDate = Format(mRec("cstdate"), "dd/mm/yyyy")
  txteccno = IIf(IsNull(mRec("ECCNo")), "", mRec("ECCNo"))
  txtexregno = IIf(IsNull(mRec("EDRegno")), "", mRec("EDRegno"))
  dpexregdt = Format(mRec("EDRegDate"), "dd/mm/yyyy")
  txtContactPerson = IIf(IsNull(mRec("ContactPerson")), "", mRec("ContactPerson"))
  
  'GST RegNo, Date & Provisional ID
  txtGSTRegNo = IIf(IsNull(mRec("GSTRegNo")), "", mRec("GSTRegNo"))
  dtpGSTDate = Format(mRec("GSTRegDate"), "dd/mm/yyyy")
  txtProvisionalID = IIf(IsNull(mRec("ProvisionalID")), "", mRec("ProvisionalID"))
  
  'Get Udyam No. & Date
  If CheckNull(mRec("UdyamNo"), "") <> "" Then
    uPos = InStrRev(CheckNull(mRec("UdyamNo")), " ", Len(CheckNull(mRec("UdyamNo"))), vbTextCompare)
  Else
    uPos = 0
  End If
    
  If uPos > 0 Then
    txtUdyamType.Text = Mid(Trim$(CheckNull(mRec("UdyamNo"))), (uPos + 1), Len(Trim$(CheckNull(mRec("UdyamNo"), ""))))
    txtUdyamNo.Text = Mid(Trim$(CheckNull(mRec("UdyamNo"))), 1, (uPos - 1))
  Else
    txtUdyamNo.Text = Trim$(CheckNull(mRec("UdyamNo")))
  End If
    
  If CheckNull(mRec("UdyamDate"), "") <> "" Then
    dtpUdyamDate.Value = Format(CheckNull(mRec("UdyamDate")), "dd/mm/yyyy")
  Else
    'Continue
  End If
  
  txtareacd = IIf(IsNull(mRec("Areacode")), "", mRec("Areacode"))
  tbPhone = IIf(IsNull(mRec("Phone")), "", mRec("Phone"))
  tbFax = IIf(IsNull(mRec("Fax")), "", mRec("Fax"))
  tbEmail = IIf(IsNull(mRec("Email")), "", mRec("Email"))
  tbCell = IIf(IsNull(mRec("Cellular")), "", mRec("Cellular"))
  mRec.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set mRec = Nothing
  Set snap = Nothing
End Try
End Sub

Private Function CheckPartyStatus() As eStatus
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  CheckPartyStatus = mCheckUnSuccess
  
  Sql = " Select LedgerCode, isActive, (Select GroupName From FAS_GroupMaster "
  Sql = Sql & " Where (GrRoute+GrWidth) = a.LrRoute and CompCode = a.CompCode) as GroupName"
  Sql = Sql & " From FAS_LedgerMaster a"
  Sql = Sql & " Where LedgerName='" & Trim$(tbPartyName.Text) & "' and CompCode = '" & Left(gFasCompCode, 4) & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    If Val(tbPartyName.Tag) > 0 Then
      CheckPartyStatus = mBothAvailable
    Else
      CheckPartyStatus = mFASOnly
    End If
    
    If rs("isActive") <> IIf(cmbStatus.ListIndex = 1, True, False) Then
      If MsgBox("Status of this Party is differ in FAS and this module. Do you want update the Status in FAS Module?", vbYesNo) = vbYes Then
        Sql = "Update FAS_LedgerMaster Set isActive='" & cmbStatus.ListIndex & "'"
        Sql = Sql & " Where LedgerName='" & Trim$(tbPartyName.Text) & "' and CompCode = '" & Left(gFasCompCode, 4) & "'"
        If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
      End If
    End If
  Else
    If Val(tbPartyName.Tag) > 0 Then
      CheckPartyStatus = mPurSalOnly
    Else
      CheckPartyStatus = mNotExists
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function CmbLoad() As Boolean

  Try
  
  CmbLoad = False
  If LoadComboBox(CmbFrom, "Form") = False Then GoTo ExitHere
  If LoadComboBox(CmbPackFoward, "Packing & Forwarding") = False Then GoTo ExitHere
  If LoadComboBox(CmbDelIns, "Delivery Instruction") = False Then GoTo ExitHere
  If LoadComboBox(CmbInsurance, "Insurance") = False Then GoTo ExitHere
  If LoadComboBox(CmbFreight, "Freight") = False Then GoTo ExitHere
  If LoadComboBox(CmbMode, "Mode") = False Then GoTo ExitHere
  If LoadComboBox(CmbTerms, "Payment Terms") = False Then GoTo ExitHere
  CmbLoad = True
Catch ex As Exception
  Call ErrMsgShow(ex)
  GoTo ExitHere
Finally
End Try
End Function

' List of Account Types
Private Sub FillAccType()
  
  Try
  
  cmbAcType.Clear
  If gModuleCode = eMdlCode.mPUR Then
    cmbAcType.AddItem "Creditor"      '0303
  Else
    cmbAcType.AddItem "Debtor"        '0206
  End If
  cmbAcType.AddItem "Party"           '0206  Default
  
  cmbAcType.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function GenCode() As String
  Dim Sql As String
  
  Try
 
  Sql = "Select Max(LedgerCode) from FAS_LedgerMaster where CompCode = '" & Left(gFasCompCode, 4) & "'"
  GenCode = Val(SingleValRet(Sql)) + 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function GetCode(ByVal AccountType As String) As String
  Dim FType As String

  Try

  Select Case AccountType
    Case "Debtor"
      FType = "D"
    Case "Creditor"
      FType = "P"
    Case "Party"
      FType = "Y"
'    Case "I"
'        AcType = "Sales"
'    Case "E"
'        AcType = "Purchase"
'    Case "M"
'        AcType = "Indirect Income"
'    Case "N"
'        AcType = "Indirect Expense"
'    Case "A"
'        AcType = "Asset"
'    Case "L"
'        AcType = "Current Liability"
'    Case "O"
'        AcType = "Loans"
  End Select
  GetCode = FType
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function GetType(ByVal fAccountType As String) As String
  Dim FType As String

  Try
  
  Select Case fAccountType
    Case "Debtor", "Party"
      FType = "0206"
    Case "Creditor"
      FType = "0303"
'    Case "Party"
'      FType = "0206"
'    Case "Sales"
'      FType = "0E"
'    Case "Purchase"
'      FType = "0D"
'    Case "Indirect Income"
'      FType = "08"
'    Case "Indirect Expense"
'      FType = "07"
'    Case "Asset"
'      FType = "02"
'    Case "Current Liability"
'      FType = "03"
'    Case "Loans"
'      FType = "0204"
  End Select
  
  GetType = FType
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SaveLedgerAddress(ByVal CurDate As Date,ByVal  fPartyCode As Integer,Optional ByVal isUpdate As Boolean = False) As Boolean
  Dim Sql As String
  Dim retVal As Variant
  Try
  SaveLedgerAddress = False
  
  If isUpdate = False Then
    Sql = " Execute FAS_Ins_LrAddress"
  Else
    Sql = " Execute FAS_Up_LrAddress"
  End If
  Sql = Sql & " @LrCode=" & Val(tbPartyName.Tag) & ","
  Sql = Sql & " @add1='" & MaskSingleQuotes(Trim$(tbAdd1)) & "',"
  Sql = Sql & " @add2='" & MaskSingleQuotes(Trim$(tbAdd2)) & "',"
  Sql = Sql & " @add3='" & MaskSingleQuotes(Trim$(tbAdd3)) & "',"
  
  Sql = Sql & " @Country='" & Trim$(txtCountry.Text) & "',"
  Sql = Sql & " @State='" & Trim$(txtState.Text) & "',"
  Sql = Sql & " @City='" & Trim$(tbCity) & "',"
  
  Sql = Sql & " @Pincode='" & MaskSingleQuotes(Trim$(tbPin)) & "',"
  Sql = Sql & " @Phone='" & MaskSingleQuotes(Trim$(tbPhone)) & "',"
  Sql = Sql & " @EMail='" & MaskSingleQuotes(Trim$(tbEmail)) & "',"
  Sql = Sql & " @TinNo='" & MaskSingleQuotes(Trim$(txtTinNo)) & "',"
  Sql = Sql & " @PanNo='" & MaskSingleQuotes(Trim$(txtPanNo)) & "',"
  Sql = Sql & " @ECCNo='" & MaskSingleQuotes(Trim$(txteccno)) & "',"
  Sql = Sql & " @Range='" & MaskSingleQuotes(Trim$(txteccno)) & "',"
  
   'GST Details
  Sql = Sql & " @GSTRegNo='" & Trim$(txtGSTRegNo.Text) & "',"
  
  If IsNull(dtpGSTDate) Then
    Sql = Sql & " @GSTRegDate=Null,"
  Else
    Sql = Sql & " @GSTRegDate='" & Format(dtpGSTDate, "yyyy-MM-dd") & "',"
  End If
  
  Sql = Sql & " @ProvisionalID='" & Trim$(txtProvisionalID) & "',"
  
  'Udyam No. Type-No. & Date
  retVal = Trim$(txtUdyamType.Text)
  retVal = Trim$(txtUdyamNo.Text) & " " & Trim$(retVal)
  Sql = Sql & " @UdyamNo='" & Trim$(retVal) & "',"
    
  If CheckNull(dtpUdyamDate.Value, "") = "" Then
    Sql = Sql & " @UdyamDate=Null,"
  Else
    Sql = Sql & " @UdyamDate='" & Format(dtpUdyamDate.Value, "yyyy-MMM-dd") & "',"
  End If

  Sql = Sql & " @CompCode='" & Left$(gFasCompCode, 4) & "',"
  Sql = Sql & " @Userid='" & gUserID & "',"
  Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  If DBExecuteSql(db, Sql) = False Then GoTo ExitHere

  SaveLedgerAddress = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SaveLedgerMaster(ByVal CurDate As Date,ByVal  fPartyCode As Integer,Optional ByVal isUpdate As Boolean = False) As Boolean
  Dim Sql As String
  
  Try
  SaveLedgerMaster = False
  
  If isUpdate = False Then
    Sql = "Execute FAS_Ins_LedgerMaster"
    Sql = Sql & " @LedgerCode=" & fPartyCode & ","
    Sql = Sql & " @LedgerName='" & MaskSingleQuotes(Trim$(tbPartyName)) & "',"
    Sql = Sql & " @Alias='',"
    Sql = Sql & " @RegisteredName='" & Trim$(txtRegisteredName.Text) & "',"
    Sql = Sql & " @LrRoute='" & GetType(Trim$(cmbAcType)) & "',"
    Sql = Sql & " @MaxReached='0',"
    Sql = Sql & " @isBillwise='1',"
    Sql = Sql & " @CreditDays=" & Val(txtCreditDays) & ","
    Sql = Sql & " @OPbalance=0,"
    Sql = Sql & " @isCostCentreApplicable=0,"
    Sql = Sql & " @isActive='" & cmbStatus.ListIndex & "',"
    Sql = Sql & " @CompCode='" & Left$(gFasCompCode, 4) & "',"
    Sql = Sql & " @Userid='" & gUserID & "',"
    Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  Else
    Sql = "Update FAS_LedgerMaster Set isActive='" & cmbStatus.ListIndex & "',"
    Sql = Sql & " Userid='" & gUserID & "',"
    Sql = Sql & " RegisteredName='" & Trim$(txtRegisteredName.Text) & "',"
    Sql = Sql & " TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " Where "
    Sql = Sql & " LedgerCode=" & fPartyCode
    Sql = Sql & " and CompCode='" & Left$(gFasCompCode, 4) & "'"
    If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  End If
  

  SaveLedgerMaster = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SavePartyAddress() As Boolean
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim retVal As Variant
  
  Try
  SavePartyAddress = False

  Sql = " Select PartyCode from FAS_PARTY_ADDRESS "
  Sql = Sql & " where PartyCode = " & Val(tbPartyName.Tag) & " and Branchid ='" & gBranchID & "' and Compid='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not snap.EOF Then
    Sql = " Execute FAS_UP_PARTY_ADDRESS"
  Else
    Sql = " Execute FAS_INS_PARTY_ADDRESS"
  End If
  Sql = Sql & " @PartyCode=" & Val(tbPartyName.Tag) & ","
  Sql = Sql & " @add1='" & MaskSingleQuotes(Trim$(tbAdd1)) & "',"
  Sql = Sql & " @add2='" & MaskSingleQuotes(Trim$(tbAdd2)) & "',"
  Sql = Sql & " @add3='" & MaskSingleQuotes(Trim$(tbAdd3)) & "',"
  
  Sql = Sql & " @Country='" & Trim$(txtCountry.Text) & "',"
  Sql = Sql & " @State='" & Trim$(txtState.Text) & "',"
  Sql = Sql & " @City='" & Trim$(tbCity) & "',"
  
  Sql = Sql & " @pincode='" & Trim$(tbPin) & "',"
  Sql = Sql & " @Phone='" & Trim$(tbPhone) & "',"
  Sql = Sql & " @fax='" & Trim$(tbFax) & "',"
  Sql = Sql & " @Cellular='" & Trim$(tbCell) & "',"
  Sql = Sql & " @email='" & Trim$(tbEmail) & "',"
  Sql = Sql & " @cst='" & Trim$(tbCst) & "',"
  If IsNull(dtpCSTDate) Then
    Sql = Sql & " @cstdate=Null,"
  Else
    Sql = Sql & " @cstdate='" & Format(dtpCSTDate, "yyyy-MM-dd") & "',"
  End If
  Sql = Sql & " @lst='" & Trim$(tbLst) & "',"
  If IsNull(dtpLSTDate) Then
    Sql = Sql & " @lstdate=Null,"
  Else
    Sql = Sql & " @lstdate='" & Format(dtpLSTDate, "yyyy-MM-dd") & "',"
  End If
  
  'GST Details
  Sql = Sql & " @GSTRegNo='" & Trim$(txtGSTRegNo) & "',"
  
  If IsNull(dtpGSTDate) Then
    Sql = Sql & " @GSTRegDate=Null,"
  Else
  
    Sql = Sql & " @GSTRegDate='" & Format(dtpGSTDate, "yyyy-MM-dd") & "',"
  End If
  
  Sql = Sql & " @ProvisionalID='" & Trim$(txtProvisionalID) & "',"
  
  'Udyam No. Type-No. & Date
  retVal = Trim$(txtUdyamType.Text)
  retVal = Trim$(txtUdyamNo.Text) & " " & Trim$(retVal)
  Sql = Sql & " @UdyamNo='" & retVal & "',"
  
  If CheckNull(dtpUdyamDate.Value, "") = "" Then
    Sql = Sql & " @UdyamDate=Null,"
  Else
    Sql = Sql & " @UdyamDate='" & Format(dtpUdyamDate.Value, "yyyy-MMM-dd") & "',"
  End If
    
  Sql = Sql & " @areaCode='" & Trim$(txtareacd) & "',"
  Sql = Sql & " @eccno='" & Trim$(txteccno) & "',"
  Sql = Sql & " @Range='" & Trim$(txteccno) & "',"
  Sql = Sql & " @EDRegNo ='" & Trim$(txtexregno) & "',"
  
  If IsNull(dpexregdt) Then
    Sql = Sql & " @edregdate=Null,"
  Else
    Sql = Sql & " @edregdate='" & Format(dpexregdt, "yyyy-MM-dd") & "',"
  End If
  
  Sql = Sql & " @contactperson='" & Trim$(txtContactPerson) & "',"
  
  Sql = Sql & " @Userid='" & gUserID & "',"
  Sql = Sql & " @Branchid='" & gBranchID & "',"
  Sql = Sql & " @Compid='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then GoTo ExitHere

  SavePartyAddress = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function SavePartyMaster() As Boolean
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  
  Try
  SavePartyMaster = False
  
  Sql = "Select PartyCode from FAS_PARTY_Master "
  Sql = Sql & " where PartyCode = " & Val(tbPartyName.Tag) & " and Branchid ='" & gBranchID & "' and Compid='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not snap.EOF Then
    Sql = " Execute FAS_UP_PARTY_MASTER"
  Else
    Sql = " Execute FAS_INS_PARTY_MASTER"
  End If
  
  Sql = Sql & " @PartyCode=" & Val(tbPartyName.Tag) & ","
  Sql = Sql & " @VendorCode='" & Trim$(txtVendorCode.Text) & "',"
  Sql = Sql & " @SuplierCode='" & Trim$(txtsupliercode.Text) & "',"
  Sql = Sql & " @PartyName='" & Trim$(tbPartyName) & "',"
  Sql = Sql & " @AcType='" & GetCode(Trim$(cmbAcType.Text)) & "',"
  Sql = Sql & " @FormCode=" & IIf(CmbFrom.ListIndex >= 0, CmbFrom.ItemData(CmbFrom.ListIndex), "Null") & ","
  Sql = Sql & " @PackingCode=" & IIf(CmbPackFoward.ListIndex >= 0, CmbPackFoward.ItemData(CmbPackFoward.ListIndex), "Null") & ","
  Sql = Sql & " @DeliveryCode=" & IIf(CmbDelIns.ListIndex >= 0, CmbDelIns.ItemData(CmbDelIns.ListIndex), "Null") & ","
  Sql = Sql & " @PaymentCode=" & IIf(CmbTerms.ListIndex >= 0, CmbTerms.ItemData(CmbTerms.ListIndex), "Null") & ","
  Sql = Sql & " @InsuranceCode=" & IIf(CmbInsurance.ListIndex >= 0, CmbInsurance.ItemData(CmbInsurance.ListIndex), "Null") & ","
  Sql = Sql & " @ModeCode=" & IIf(CmbMode.ListIndex >= 0, CmbMode.ItemData(CmbMode.ListIndex), "Null") & ","
  Sql = Sql & " @Freightcode=" & IIf(CmbFreight.ListIndex >= 0, CmbFreight.ItemData(CmbFreight.ListIndex), "Null") & ","
  Sql = Sql & " @IsSubcontractor='" & ChkSub.Value & "',"
  Sql = Sql & " @AliasName='" & Trim$(txtAName) & "',"
  Sql = Sql & " @EffStatus='" & cmbStatus.ListIndex & "',"
  Sql = Sql & " @EffDate='" & Format(IIf(cmbStatus.ListIndex = Val(cmbStatus.Tag), dtpPrevEffDt, CurDate), "yyyy-MM-dd") & "',"
  Sql = Sql & " @CreditDays=" & Val(txtCreditDays) & ","
  Sql = Sql & " @CreditDaysPeriod='" & Trim$(cmbCreditDaysPeriod) & "',"
  Sql = Sql & " @CurrencyCode='" & Trim$(txtCurrency.Tag) & "',"
  Sql = Sql & " @CustomerFormulaCode=" & Val(txtCustomerFormula.Tag) & ","
  Sql = Sql & " @VendorFormulaCode=" & Val(txtVendorFormula.Tag) & ","
  Sql = Sql & " @TinNo='" & Trim$(txtTinNo.Text) & "',"
  Sql = Sql & " @PANNo='" & Trim$(txtPanNo.Text) & "',"
  Sql = Sql & " @Userid='" & gUserID & "',"
  Sql = Sql & " @Branchid='" & gBranchID & "',"
  Sql = Sql & " @Compid='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Function ValidateData(ByVal CurDate As Date) As Boolean
  Dim cmbControl As Control
  
  Try
  ValidateData = False
  
  'City
  If Trim$(tbAdd1.Text) <> "" Then
    If tbCity.Text = "" Then
      sShowMessage "Select the City"
      tbCity.SetFocus
      GoTo ExitHere
    End If
  End If
  
  If dtpCSTDate <> "" Then
    If Trim$(dtpCSTDate) > CurDate Then
      sShowMessage "CST Date should be less than or equal to Current Date"
      dtpCSTDate.SetFocus
      GoTo ExitHere
    End If
  End If
  
  If dtpLSTDate <> "" Then
    If Trim$(dtpLSTDate) > CurDate Then
      sShowMessage "LST Date should be less than or equal to Current Date"
      dtpLSTDate.SetFocus
      GoTo ExitHere
    End If
  End If

  'GST Reg.Date
  If IsGSTRequired = True And dtpGSTDate <> "" Then
    If Trim$(dtpGSTDate) > CurDate Then
      sShowMessage "GST Reg Date should be less than or equal to Current Date"
      dtpGSTDate.SetFocus
      GoTo ExitHere
    End If
  End If
  
  If dpexregdt <> "" Then
    If Trim$(dpexregdt) > CurDate Then
      sShowMessage "Excise Reg. Date should be less than or equal to Current Date"
      dpexregdt.SetFocus
      GoTo ExitHere
    End If
  End If
  
  If Trim$(txtCurrency) = "" Then
    sShowMessage "Please select the Currency"
    txtCurrency.SetFocus
    GoTo ExitHere
  End If
  
  If cmbAcType.Text <> "Creditor" Then
    If Trim$(txtCustomerFormula) = "" Then
      sShowMessage "Please select the Customer Formula"
      txtCustomerFormula.SetFocus
      GoTo ExitHere
    End If
  End If
  
  If cmbAcType.Text <> "Debtor" And Trim$(cmbAcType.Text) = "Party" Then
    If Trim$(txtVendorFormula) = "" Then
      sShowMessage "Please select the Vendor Formula"
      txtVendorFormula.SetFocus
      GoTo ExitHere
    End If
  End If
  
  For Each cmbControl In Controls
    If TypeOf cmbControl Is ComboBox Then
      If cmbControl.ListIndex = -1 Then
        sShowMessage "Please verify all Combo Box option contains only the approved values."
        GoTo ExitHere
      End If
    End If
  Next

  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmbControl = Nothing
End Try
End Function
  
Private Function ValidateVendor() As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Long
    
  Try
  ValidateVendor = ""
  
  Sql = " Select LedgerCode, isActive, (Select GroupName From FAS_GroupMaster "
  Sql = Sql & " Where (GrRoute+GrWidth) = a.LrRoute and CompCode = a.CompCode) as GroupName"
  Sql = Sql & " From FAS_LedgerMaster a"
  Sql = Sql & " Where LedgerName='" & Trim$(tbPartyName.Text) & "' and CompCode = '" & gFasCompCode & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    'Party Name already Exists and Active Status. Pass the LedgerCode as Positive value
    If rs("isActive") = "True" Then
      ValidateVendor = rs("LedgerCode") & "*" & rs("GroupName") 'For Save as PartyCode in PartyMaster
    Else ' Party Name already Exists But InActive Status. Pass the LedgerCode as Negative value
      ValidateVendor = -(rs("LedgerCode")) & "*" & rs("GroupName")
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

