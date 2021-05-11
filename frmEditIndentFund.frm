VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEditIndentFund 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indent For C & F Fund"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtApprovedDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   6630
      TabIndex        =   9
      Top             =   1740
      Width           =   2085
   End
   Begin VB.ComboBox cmbJobNo 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1740
      Width           =   2085
   End
   Begin VB.Frame frameChangeInvoiceType 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Change Invoice Type "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6060
      TabIndex        =   43
      Tag             =   "E"
      Top             =   900
      Width           =   2715
      Begin VB.OptionButton optChangeExport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optChangeImport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1410
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtPayDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4470
      TabIndex        =   8
      Top             =   1740
      Width           =   2085
   End
   Begin VB.ComboBox cmbAccHead 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3060
      Width           =   4275
   End
   Begin VB.CommandButton cmdNewAccHead 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8460
      Picture         =   "frmEditIndentFund.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Add New Name of Item"
      Top             =   3060
      Width           =   285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7410
      TabIndex        =   21
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton cmdAddBuyer 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8460
      Picture         =   "frmEditIndentFund.frx":00F2
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Add New Name of Item"
      Top             =   2400
      Width           =   285
   End
   Begin VB.CommandButton cmdAddExpenses 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Picture         =   "frmEditIndentFund.frx":01E4
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Add New Name of Item"
      Top             =   3780
      Width           =   285
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   7440
      Width           =   1275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      Top             =   4440
      Width           =   1125
   End
   Begin VB.ComboBox cmbExpenses 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3780
      Width           =   4155
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   180
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   4380
      Width           =   7455
   End
   Begin VB.TextBox txtApprovedAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6690
      MaxLength       =   15
      TabIndex        =   18
      Top             =   3780
      Width           =   2145
   End
   Begin VB.TextBox txtIndentAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4710
      MaxLength       =   15
      TabIndex        =   17
      Top             =   3780
      Width           =   1875
   End
   Begin VB.ComboBox cmbBuyerName 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5460
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtLCNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2910
      MaxLength       =   30
      TabIndex        =   11
      Top             =   2400
      Width           =   2505
   End
   Begin VB.TextBox txtInvoiceValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2070
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3060
      Width           =   2025
   End
   Begin VB.TextBox txtInvoiceNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   233
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Frame frameLCType 
      BackColor       =   &H00EAF5F5&
      Caption         =   "L/C Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2565
      TabIndex        =   27
      Tag             =   "E"
      Top             =   900
      Width           =   2265
      Begin VB.OptionButton optBBLC 
         BackColor       =   &H00EAF5F5&
         Caption         =   "BBLC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton optExportLC 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Export LC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame frameInvoice 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Invoice Type "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   285
      TabIndex        =   26
      Tag             =   "E"
      Top             =   900
      Width           =   2235
      Begin VB.OptionButton optImport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1140
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optExport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtFundReqFor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      MaxLength       =   30
      TabIndex        =   10
      Top             =   2400
      Width           =   2595
   End
   Begin VB.TextBox txtJobNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2340
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1740
      Width           =   2085
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   1995
      Left            =   120
      TabIndex        =   24
      Top             =   5010
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   -2147483626
      BackColorBkg    =   15132390
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   48
      Top             =   7200
      Width           =   9015
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indent For C && F Fund"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   47
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   "Approved Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   6630
      TabIndex        =   45
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Job No :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   44
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   4470
      TabIndex        =   42
      Top             =   1500
      Width           =   1485
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Nature Of Expense :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4200
      TabIndex        =   41
      Top             =   2820
      Width           =   1785
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Index           =   11
      Left            =   0
      TabIndex        =   37
      Top             =   7050
      Width           =   9015
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   180
      TabIndex        =   35
      Top             =   4140
      Width           =   1635
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Approved Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   6690
      TabIndex        =   34
      Top             =   3540
      Width           =   1635
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Indent Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   4710
      TabIndex        =   33
      Top             =   3540
      Width           =   1365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars Of Expenses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   180
      TabIndex        =   32
      Top             =   3540
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5460
      TabIndex        =   31
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "L/C  No  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2910
      TabIndex        =   30
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Value :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2070
      TabIndex        =   29
      Top             =   2790
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   2790
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Req. For  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   25
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Job No.  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Index           =   10
      Left            =   -30
      TabIndex        =   36
      Top             =   3450
      Width           =   9105
   End
End
Attribute VB_Name = "frmEditIndentFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytRowNo As Integer

Private Sub cmbAccHead_Click()
    cmbAccHead.Tag = 0
    If Len(cmbAccHead.Text) > 0 Then cmbAccHead.Tag = cmbAccHead.ItemData(cmbAccHead.ListIndex)
    cmbAccHead.ToolTipText = cmbAccHead.Text
End Sub

Private Sub cmbAccHead_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbJobNo_Click()
On Error GoTo x
    Call prcBlank(0)
    If Len(cmbJobNo.Text) = 0 Then Exit Sub
    Call prcLoadData
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub cmbJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAddBuyer_Click()
    flagMark = "INSERT"
    strMark = "BUYER"
    strValue = ""
    
    frmNewBuyer.Show 1
    
    Call prcAddBuyerName(cmbBuyerName)
    If Len(strValue) > 0 Then cmbBuyerName.Text = strValue
    cmbBuyerName.SetFocus
End Sub

Private Sub cmbBuyerName_Click()
    cmbBuyerName.Tag = 0
    If Len(cmbBuyerName.Text) > 0 Then cmbBuyerName.Tag = cmbBuyerName.ItemData(cmbBuyerName.ListIndex)
    cmbBuyerName.ToolTipText = cmbBuyerName.Text
End Sub

Private Sub cmbBuyerName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbExpenses_Click()
    cmbExpenses.Tag = 0
    If Len(cmbExpenses.Text) > 0 Then cmbExpenses.Tag = cmbExpenses.ItemData(cmbExpenses.ListIndex)
    cmbExpenses.ToolTipText = cmbExpenses.Text
End Sub

Private Sub cmbExpenses_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    Dim strItem As String
    If fncBlank(2) = True Then Exit Sub
    
    With flexDetails
        If .Rows > 1 Then
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 1: strItem = .Text
                
                If strItem = cmbExpenses.Text And inti <> bytRowNo Then
                    MsgBox "Adding Duplicate Expense Head.", vbInformation, cnstMsgInfo
                    cmbExpenses.SetFocus
                    Exit Sub
                End If
            Next inti
        End If
    End With
    
    With flexDetails
        If cmdAdd.Caption = "&Add" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        Else
            .Row = bytRowNo
        End If
      
        .Col = 0:      .Text = cmbExpenses.Tag
        .Col = 1:      .Text = cmbExpenses.Text
        .Col = 2:      .Text = txtIndentAmount.Text
        .Col = 3:      .Text = txtApprovedAmount.Text
        .Col = 4:      .Text = txtRemarks.Text
    End With
    cmdAdd.Caption = "&Add"
    
    Call prcBlank(2)
    cmbExpenses.SetFocus
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAddExpenses_Click()
    flagMark = "INSERT"
    strMark = "EXPENSE"
    strValue = ""
    
    frmNewExpense.Show 1
    
    Call prcAddExpenses(cmbExpenses)
    If Len(strValue) > 0 Then cmbExpenses.Text = strValue
    cmbExpenses.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewAccHead_Click()
    flagMark = "INSERT"
    strMark = "ACCHEAD"
    strValue = ""
    
    frmNewAccountsHead.Show 1
    
    Call addAccName(cmbAccHead, 1)
    If Len(strValue) > 0 Then cmbAccHead.Text = strValue
    cmbAccHead.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    If fncBlank(0) = True Then Exit Sub
    If fncDuplicate = True Then Exit Sub
    If MsgBox("Do you want to save data ?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    
    MainConn.BeginTrans
        Dim NewID As Integer
        NewID = fncNewID("FUND")
        
                 
        strRecord = "Insert Into Indent_Fund_Main (CompID, RefID, FundReqFor, JobNo, InvoiceNo, InvoiceValue, LCType, LCNo, BuyerID, accID, InvoiceStatus) " + _
            " Values(" & bytCompID & ", " & NewID & ", '" & txtFundReqFor.Text & "', '" & txtJobNo.Text & "', '" & txtInvoiceNo.Text & "', " & CDbl(txtInvoiceValue.Text) & "," + _
            " '" & frameLCType.Tag & "', '" & txtLCNo.Text & "', " & cmbBuyerName.Tag & ", " & cmbAccHead.Tag & ", '" & frameInvoice.Tag & "');"
        MainComm.CommandText = strRecord
        MainComm.Execute
        
        strRecord = "Insert Into Indent_Fund_Main_bak (CompID, RefID, FundReqFor, JobNo, InvoiceNo, InvoiceValue, LCType, LCNo, BuyerID, accID, InvoiceStatus,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime) " + _
            " Values(" & bytCompID & ", " & NewID & ", '" & txtFundReqFor.Text & "', '" & txtJobNo.Text & "', '" & txtInvoiceNo.Text & "', " & CDbl(txtInvoiceValue.Text) & "," + _
            " '" & frameLCType.Tag & "', '" & txtLCNo.Text & "', " & cmbBuyerName.Tag & ", " & cmbAccHead.Tag & ", '" & frameInvoice.Tag & "',1,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
        MainComm.CommandText = strRecord
        MainComm.Execute
        
        Dim ExpenseID As Integer, IndentAmt As Currency, ApprovedAmt As Currency, Remarks As String
        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0:   ExpenseID = .Text
'                .Col = 1:  Expensename = .text
                .Col = 2:   IndentAmt = IIf(Len(.Text) = 0, "0.00", CDbl(.Text))
                .Col = 3:   ApprovedAmt = IIf(Len(.Text) = 0, "0.00", CDbl(.Text))
                .Col = 4:   Trim (.Text)
                
                strRecord = "Insert Into Indent_Fund_Sub (RefID, ExpensesID, IndentAmt, ApprovedAmt, Remarks, RowNo)"
                strRecord = strRecord + " Values (" & NewID & ", " & ExpenseID & ", " & IndentAmt & ", " & IndentAmt & ", '" & Remarks & "', " & inti & ");"
                MainComm.CommandText = strRecord
                MainComm.Execute
                
                strRecord = "Insert Into Indent_Fund_Sub_bak (RefID, ExpensesID, IndentAmt, ApprovedAmt, Remarks, RowNo,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)"
                strRecord = strRecord + " Values (" & NewID & ", " & ExpenseID & ", " & IndentAmt & ", " & IndentAmt & ", '" & Remarks & "', " & inti & ",1,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
                MainComm.CommandText = strRecord
                MainComm.Execute
     
                 
            Next inti
        End With
    MainConn.CommitTrans
    flexDetails.Row = 1
Exit Sub
x:
MainConn.RollbackTrans
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub flexDetails_DblClick()
On Error Resume Next
    flexDetails_KeyDown 13, 0
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With flexDetails
            If .RowSel = 0 Then Exit Sub
            '.Col = 0:    cmbExpenses.TAG = .Text
            .Col = 1:    cmbExpenses.Text = .Text
            .Col = 2:    txtIndentAmount.Text = .Text
            .Col = 3:    txtApprovedAmount.Text = .Text
            .Col = 4:    txtRemarks.Text = .Text
            
            bytRowNo = .Row
        End With
        cmdAdd.Caption = "&Replace"
        cmbExpenses.SetFocus
    End If

    If KeyCode = vbKeyDelete Then
        With flexDetails
            If .RowSel = 0 Then
                Beep
            Else
                If MsgBox("Are you sure to Delete", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
            End If
        End With
        cmdAdd.Caption = "&Add"
    End If

End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    Call prcGridSetting
    
    Call prcAddBuyerName(cmbBuyerName)
    Call prcAddExpenses(cmbExpenses)
    Call addAccName(cmbAccHead, 1)
    
    Call prcAddJobNo(cmbJobNo, frameInvoice.Tag)
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub optBBLC_Click()
    frameLCType.Tag = "B"
End Sub

Private Sub optBBLC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub


Private Sub optChangeExport_Click()
    frameChangeInvoiceType.Tag = "E"
End Sub

Private Sub optChangeExport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optChangeImport_Click()
frameChangeInvoiceType.Tag = "I"
End Sub

Private Sub optChangeImport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optExport_Click()
    frameInvoice.Tag = "E"
    Call prcBlank(0): Call prcBlank(2)
    Call prcAddJobNo(cmbJobNo, frameInvoice.Tag)
End Sub

Private Sub optExport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optExportLC_Click()
    frameLCType.Tag = "E"
End Sub

Private Sub optExportLC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optImport_Click()
    frameInvoice.Tag = "I"
    Call prcBlank(0): Call prcBlank(2)
    Call prcAddJobNo(cmbJobNo, frameInvoice.Tag)
End Sub

Public Sub prcGridSetting()
    With flexDetails
        .Rows = 1
        .Cols = 5
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignLeftCenter
        
        .Col = 0:   .Text = "Expense ID"
        .Col = 1:   .Text = "Expenses Name"
        .Col = 2:   .Text = "Indent Amount"
        .Col = 3:   .Text = "Aproved Amount"
        .Col = 4:   .Text = "Remarks"
        
        .ColWidth(0) = 0:       .ColWidth(1) = 2800
        .ColWidth(2) = 2000:    .ColWidth(3) = 2000
        .ColWidth(4) = 4000
    End With
End Sub

Private Sub prcInsertData()
On Error GoTo x
    Dim NewID As Integer
    Dim ExpenseID As Integer, ExpensName As String, IndentAmount As Currency, ApprovedAmount As Currency, strRemark As String
    
    NewID = fncNewID("FUND")
    MainConn.BeginTrans
        strRecord = "Insert Into Indent_Fund_Main (CompID, RefID, FundReqFor, JobNo, InvoiceNO, InvoiceValue, LCType, LCNO, BuyerID, FundStatus, InvoiceStatus)"
        strRecord = strRecord + " Values (" & bytCompID & ", " & NewID & ", '" & txtFundReqFor.Text & "', '" & txtJobNo.Text & "', '" & txtInvoiceNo.Text & "', '" & txtInvoiceValue.Text & "', '" & frameInvoice.Tag & "', '" & txtLCNo.Text & "', " & cmbBuyerName.Tag & ", " & frameFund.Tag & ", " & frameInvoice.Tag & ")"
        MainComm.CommandText = strRecord
        MainComm.Execute

        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0:   ExpenseID = CInt(.Text)
                .Col = 1:   ExpenseName = .Text
                .Col = 2:   IndentAmount = CCur(.Text)
                .Col = 3:   ApprovedAmount = CCur(.Text)
                .Col = 4:   Trim (.Text)
                
                strRecord = " Insert  Into Indent_Fund_Sub (RowNo, ExpenseID, IndentAmt, ApprovedAmt, Remarks)"
                strRecord = strRecord + " values (" & inti & ", " & ExpenseID & ", " & IndentAmount & ", " & ApprovedAmount & ", '" & Remarks & "');"
                MainComm.CommandText = strRecord
                MainComm.Execute
            Next inti
        End With
    MainConn.CommitTrans
Exit Sub
x:
MainConn.RollbackTrans
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub optImport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtApprovedAmount_GotFocus()
    Call txtGotFocus(txtApprovedAmount)
End Sub

Private Sub txtApprovedAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtApprovedAmount_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtApprovedAmount_LostFocus()
    If Len(txtApprovedAmount.Text) = 0 Then
        txtApprovedAmount.Text = "0.00"
    Else
        txtApprovedAmount.Text = FormatNumber(txtApprovedAmount.Text, 2)
    End If
End Sub

Private Sub txtApprovedDate_GotFocus()
    Call txtGotFocus(txtApprovedDate)
End Sub

Private Sub txtApprovedDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtApprovedDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtApprovedDate_LostFocus()
    Call FormatDate(txtApprovedDate)
End Sub

Private Sub txtFundReqFor_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFundReqFor_LostFocus()
    txtFundReqFor.Text = Trim(UCase(txtFundReqFor.Text))
End Sub

Private Sub txtIndentAmount_GotFocus()
Call txtGotFocus(txtIndentAmount)
End Sub

Private Sub txtIndentAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtIndentAmount_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtIndentAmount_LostFocus()
    If Len(txtIndentAmount.Text) = 0 Then
        txtIndentAmount.Text = "0.00"
    Else
        txtIndentAmount.Text = FormatNumber(txtIndentAmount.Text, 2)
    End If
End Sub

Private Sub txtInvoiceNo_GotFocus()
    Call txtGotFocus(txtInvoiceNo)
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo.Text = Trim(UCase(txtInvoiceNo.Text))
End Sub

Private Sub txtInvoiceValue_GotFocus()
    Call txtGotFocus(txtInvoiceValue)
End Sub

Private Sub txtInvoiceValue_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtInvoiceValue_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtInvoiceValue_LostFocus()
    If Len(txtInvoiceValue.Text) = 0 Then
        txtInvoiceValue.Text = "0.00"
    Else
        txtInvoiceValue.Text = FormatNumber(txtInvoiceValue.Text, 2)
    End If
End Sub

Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtJobNo_LostFocus()
    txtJobNo.Text = Trim(UCase(txtJobNo.Text))
    If Len(txtInvoiceNo.Text) = 0 Then txtInvoiceNo.Text = txtJobNo.Text
End Sub

Private Sub txtLCNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtLCNo_LostFocus()
    txtLCNo.Text = Trim(UCase(txtLCNo.Text))
End Sub

Private Sub txtPayDate_GotFocus()
    Call txtGotFocus(txtPayDate)
End Sub

Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPayDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtPayDate_LostFocus()
    Call FormatDate(txtPayDate)
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Function fncBlank(param As Integer) As Boolean
fncBlank = True
    If param = 0 Then
        If Len(txtJobNo.Text) = 0 Then
            MsgBox "You should provide job no.", vbCritical, App.Title
            txtJobNo.SetFocus
            Exit Function
        End If
        If Len(txtFundReqFor.Text) = 0 Then
            MsgBox "You should provide fund require for", vbCritical, App.Title
            txtFundReqFor.SetFocus
            Exit Function
        End If
        If Len(txtLCNo.Text) = 0 Then
            MsgBox "You should provide LCNO.", vbCritical, App.Title
            txtLCNo.SetFocus
            Exit Function
        End If
        If Len(cmbBuyerName.Text) = 0 Then
            MsgBox "You should provide buyer name.", vbCritical, App.Title
            cmbBuyerName.SetFocus
            Exit Function
        End If
        If Len(txtInvoiceNo.Text) = 0 Then
            MsgBox "You should provide invoice no.", vbCritical, App.Title
            txtInvoiceNo.SetFocus
            Exit Function
        End If
        If Len(txtInvoiceValue.Text) = 0 Then
            MsgBox "You should provide invoice value.", vbCritical, App.Title
            txtInvoiceValue.SetFocus
            Exit Function
        End If
        If flexDetails.Rows = 1 Then
            MsgBox "No data found to save.", vbCritical, App.Title
            cmbExpenses.SetFocus
            Exit Function
        End If
        
    ElseIf param = 2 Then
        If Len(cmbExpenses.Text) = 0 Then
            MsgBox "You should provide expense head.", vbCritical, App.Title
            cmbExpenses.SetFocus
            Exit Function
        End If
        If Len(txtIndentAmount.Text) = 0 Then
            MsgBox "You should provide Indent Amount.", vbCritical, App.Title
            txtIndentAmount.SetFocus
            Exit Function
        End If
        If Len(txtApprovedAmount.Text) = 0 Then
            MsgBox "You should provide Approved Amount.", vbCritical, App.Title
            txtApprovedAmount.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Private Sub prcBlank(param As Integer)
    If param = 0 Then
'        optExport.Value = True
'        optExportLC.Value = True
'        optChangeExport.Value = True

        txtJobNo.Text = ""
        txtFundReqFor.Text = ""
        txtLCNo.Text = ""
        cmbBuyerName.ListIndex = -1
        txtInvoiceNo.Text = ""
        txtInvoiceValue.Text = ""
        
        flexDetails.Rows = 1
    ElseIf param = 2 Then
        txtIndentAmount.Text = "0.00"
        txtApprovedAmount.Text = "0.00"
        txtRemarks.Text = ""
    End If
End Sub

Private Function fncDuplicate() As Boolean
fncDuplicate = True
    Set R = New ADODB.Recordset
        strRecord = "Select * from Indent_Fund_Main where CompID = " & bytCompID & " and JobNO = '" & Trim(txtJobNo.Text) & "' and InvoiceStatus = '" & frameInvoice.Tag & "';"
        MainComm.CommandText = strRecord
        Set R = MainComm.Execute
        If R.EOF = False And R.BOF = False Then
            MsgBox "Duplicate job no.", vbExclamation, App.Title
            txtJobNo.SetFocus
        Exit Function
        End If
    Set R = Nothing
fncDuplicate = False
End Function

Private Sub txtRemarks_LostFocus()
    txtRemarks.Text = Trim(UCase(txtRemarks.Text))
End Sub

Private Sub prcLoadData()
    Set R = New ADODB.Recordset
    strRecord = "SELECT JobNo, FundReqFor, InvoiceNo, InvoiceValue, LCType, LCNo, BuyerName, AccName, InvoiceStatus, PayDate, AdjustDate, isLock, ExpensesID, ExpenseName, IndentAmt, ApprovedAmt, Remarks, RowNo"
    strRecord = strRecord + " FROM Query_Indent_Fund Where JobNo = '" & cmbJobNo.Text & "' and Compid = " & bytCompID & " order By RowNO;"
    R.Open strRecord, MainConn, adOpenStatic
    
    Call prcBlank(0): Call prcBlank(2)
    
    If R.EOF = False And R.BOF = False Then
        txtJobNo.Text = R("JobNo").Value
        txtPayDate.Text = IIf(Format(R("PayDate").Value, cnstDtFrmtI) = CDate("01-Jan-1900"), "", Format(R("PayDate").Value, cnstDtFrmtI))
        txtApprovedDate.Text = IIf(Format(R("AdjustDate").Value, cnstDtFrmtI) = CDate("01-Jan-1900"), "", Format(R("AdjustDate").Value, cnstDtFrmtI))
        txtFundReqFor.Text = R("FundReqFor").Value
        txtLCNo.Text = R("LCNO").Value
        cmbBuyerName.Text = R("BuyerName").Value
        txtInvoiceNo.Text = R("InvoiceNo").Value
        txtInvoiceValue.Text = FormatNumber(R("InvoiceValue").Value, 2)
        cmbAccHead.Text = R("AccName").Value
        
        Do While Not R.EOF
            
        R.MoveNext
        Loop
    End If
    Set R = Nothing
End Sub
