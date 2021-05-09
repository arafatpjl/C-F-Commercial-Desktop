VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCandFOut 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " C&F Bill [OUT]"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8865
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
   Icon            =   "frmCandFOut.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Change Party && Forwarding Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   5400
      TabIndex        =   36
      Top             =   810
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtChangeForwardDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   38
         Top             =   870
         Width           =   1665
      End
      Begin VB.ComboBox cmbChangeParty 
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   510
         Width           =   3165
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Forward Date :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   40
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtApprovedAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   1530
      TabIndex        =   9
      Top             =   3060
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   5430
      TabIndex        =   12
      Top             =   3060
      Width           =   1665
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4950
      Top             =   1530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      TabIndex        =   34
      Top             =   1470
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbPartyName 
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   3345
   End
   Begin VB.TextBox txtVoucherNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   150
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1740
      Width           =   3315
   End
   Begin VB.TextBox txtBENO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   5310
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2430
      Width           =   1665
   End
   Begin VB.CommandButton cmdAddParty 
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
      Left            =   3510
      Picture         =   "frmCandFOut.frx":3332
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Add New Party Name"
      Top             =   1140
      Width           =   285
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   120
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2430
      Width           =   1455
   End
   Begin VB.TextBox txtBEDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   7020
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2430
      Width           =   1575
   End
   Begin VB.TextBox txtBillDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2430
      Width           =   1575
   End
   Begin VB.TextBox txtForwardDate 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3870
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox txtHAWB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3240
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2430
      Width           =   2025
   End
   Begin VB.TextBox txtBillAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   3060
      Width           =   1365
   End
   Begin VB.TextBox txtAIT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3060
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3060
      Width           =   1365
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   4470
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3060
      Width           =   915
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
      Height          =   345
      Left            =   7380
      TabIndex        =   13
      Top             =   3030
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6930
      TabIndex        =   18
      Top             =   7590
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5730
      TabIndex        =   17
      Top             =   7590
      Width           =   1095
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
      Height          =   345
      Left            =   4530
      TabIndex        =   16
      Top             =   7590
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3330
      TabIndex        =   15
      Top             =   7590
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3720
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "To delete a row press ""Delete"" Key"
      Top             =   3420
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   6562
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   285
      BackColor       =   -2147483633
      ForeColor       =   12582912
      BackColorFixed  =   -2147483632
      ForeColorFixed  =   16776960
      BackColorSel    =   -2147483642
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483630
      GridColorFixed  =   -2147483630
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      FillStyle       =   1
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
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "C && F Bill [Out]"
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
      TabIndex        =   43
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   42
      Top             =   7320
      Width           =   8895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "App. Amount :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   35
      Top             =   2790
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5460
      TabIndex        =   33
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "B/E Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   18
      Left            =   6990
      TabIndex        =   31
      Top             =   2190
      Width           =   1365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "B/E No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   17
      Left            =   5340
      TabIndex        =   30
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vocher No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   150
      TabIndex        =   29
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   1620
      TabIndex        =   28
      Top             =   2190
      Width           =   1365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   150
      TabIndex        =   27
      Top             =   2190
      Width           =   1245
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Index           =   10
      Left            =   150
      TabIndex        =   26
      Top             =   2130
      Width           =   8655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "HAWB :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3240
      TabIndex        =   25
      Top             =   2190
      Width           =   1245
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   24
      Top             =   870
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Forward Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3870
      TabIndex        =   23
      Top             =   870
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   22
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   -30
      TabIndex        =   21
      Top             =   7230
      Width           =   9045
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4470
      TabIndex        =   20
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "AIT :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3090
      TabIndex        =   19
      Top             =   2790
      Width           =   855
   End
End
Attribute VB_Name = "frmCandFOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmprst As Recordset
Dim bytRowNo As Byte

Private Sub cmbChangeParty_Click()
    cmbChangeParty.Tag = 0
    If cmbChangeParty.ListIndex >= 0 Then
        cmbChangeParty.Tag = cmbChangeParty.ItemData(cmbChangeParty.ListIndex)
        cmbChangeParty.ToolTipText = cmbChangeParty.Text
    End If
End Sub

Private Sub cmdAddParty_Click()
    flagMark = "INSERT"
    strMark = "PARTY"
    strValue = ""
    
    frmNewParty.Show 1
    
    Call prcAddParty(cmbPartyName)
    If Len(strValue) > 0 Then cmbPartyName.Text = strValue
    cmbPartyName.SetFocus
End Sub

Private Sub cmdPreview_Click()
    If Len(cmbPartyName.Text) = 0 Then
        MsgBox "You should provide party name.", vbInformation, App.Title
        cmbPartyName.SetFocus
        Exit Sub
    End If
    If Len(txtForwardDate.Text) = 0 Then
        MsgBox "You should provide forward date.", vbInformation, App.Title
        txtForwardDate.SetFocus
        Exit Sub
    End If
    If flexDetails.Rows = 1 Then
        MsgBox "Data not found to display report.", vbInformation, App.Title
        cmbPartyName.SetFocus
        Exit Sub
    End If
    
    With CrystalReport1
        .ReportFileName = ReportPath + "statement_of_c_and_f_bills_voucher.rpt"
        .PrintFileName = ReportPath + "statement_of_c_and_f_bills_voucher.rpt"
        
        .SQLQuery = "Select * from Query_C_And_F_Out Where M_ID = " & flexDetails.Tag & " Order By SLNO;"
        .Destination = crptToWindow
        .PrintReport
    End With
    
    
End Sub

Private Sub txtAIT_LostFocus()
    If Len(txtAIT.Text) = 0 Then
        txtAIT.Text = "0.00"
    Else
        txtAIT.Text = FormatNumber(txtAIT.Text, 2)
    End If
End Sub

Private Sub txtBEDate_GotFocus()
    Call txtGotFocus(txtBEDate)
End Sub

Private Sub txtBEDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtBEDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtBEDate_LostFocus()
    Call FormatDate(txtBEDate)
End Sub

Private Sub txtBillAmount_GotFocus()
    Call txtGotFocus(txtBillAmount)
End Sub

Private Sub txtBillAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtBillAmount_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtBillAmount_LostFocus()
    If Len(txtBillAmount.Text) = 0 Then
        txtBillAmount.Text = "0.00"
    Else
        txtBillAmount.Text = FormatNumber(txtBillAmount.Text, 2)
    End If
End Sub

Private Sub txtApprovedAmount_GotFocus()
    Call txtGotFocus(txtApprovedAmount)
End Sub

Private Sub txtApprovedAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtApprovedAmount_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtApprovedAmount_LostFocus()
    If Len(txtApprovedAmount.Text) = 0 Then
        txtApprovedAmount.Text = "0.00"
    Else
        txtApprovedAmount.Text = FormatNumber(txtApprovedAmount.Text, 2)
    End If
End Sub

Private Sub txtBillDate_GotFocus()
    Call txtGotFocus(txtBillDate)
End Sub

Private Sub txtBillNo_GotFocus()
    Call txtGotFocus(txtBillNo)
End Sub

Private Sub cmbPartyName_Click()
    cmbPartyName.Tag = 0
    If cmbPartyName.ListIndex >= 0 Then
        cmbPartyName.Tag = cmbPartyName.ItemData(cmbPartyName.ListIndex)
        cmbPartyName.ToolTipText = cmbPartyName.Text
    End If
    
    If Len(cmbPartyName.Text) = 0 Or Len(txtForwardDate.Text) = 0 Then
        txtVoucherNo.Text = ""
        flexDetails.Tag = 0: flexDetails.Rows = 1
    Else
        txtVoucherNo.Text = ""
        flexDetails.Tag = 0: flexDetails.Rows = 1
        
        Call cmdFind_Click
    End If
End Sub

Private Sub cmbPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAddNew_Click()
    If cmdAddNew.Caption = "&Edit" Then
        cmdAddNew.Caption = "Add &New"
        cmdSave.Caption = "&Update"
        cmdDelete.Enabled = True
        cmdSave.Enabled = False
    Else
        cmdAddNew.Caption = "&Edit"
        cmdSave.Caption = "&Save"
        cmdDelete.Enabled = False
        cmdSave.Enabled = True
        
        cmbPartyName.SetFocus
    End If
    cmdAdd.Caption = "&Add"
    
    flexDetails.Rows = 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Len(cmbPartyName.Text) = 0 Then
        MsgBox "You should provide party name.", vbCritical, App.Title
        cmbPartyName.SetFocus
        Exit Sub
    End If
    If Len(txtForwardDate.Text) = 0 Then
        MsgBox "You should provide forward date.", vbCritical, App.Title
        txtForwardDate.SetFocus
        Exit Sub
    End If
    
    If flexDetails.Rows = 1 Then
        MsgBox "No record found to delete.", vbCritical, App.Title
        cmbPartyName.SetFocus
    Exit Sub
    End If
    
    If MsgBox("Do you want to delete Bill information.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    intID = fncFindID
    MainConn.BeginTrans
        Mysql = "INSERT INTO C_And_F_Out_Main_bak(M_ID, PartyID, ForwardDate, VoucherNo, Compid, Deleterow, Username,Userpwd, Comname, Entrydate, Entrytime) " _
        & " SELECT  M_ID, PartyID, ForwardDate, VoucherNo, Compid, 2, '" & strUser & "','" & StrPass & "','" & getComName & "','" & Date & "','" & Time & "'" _
        & " FROM  dbo.C_And_F_Out_Main where M_ID = " & intID & ";"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = "Delete from C_And_F_Out_Main where M_ID = " & intID & ";"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = "Insert into C_And_F_Out_Sub_bak(M_ID, BillNo, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TOTALAMOUNT, SLNO, Deleterow, Username,Userpwd, Comname,Entrydate, Entrytime)" _
        & " Select M_ID, BillNo, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TOTALAMOUNT, SLNO, 2, '" & strUser & "','" & StrPass & "','" & getComName & "','" & Date & "','" & Time & "' from C_And_F_Out_Sub Where M_ID = " & intID & ";"
        MainComm.CommandText = Mysql
        MainComm.Execute
    
        Mysql = "Delete from C_And_F_Out_Sub Where M_ID = " & intID & ";"
        MainComm.CommandText = Mysql
        MainComm.Execute
    MainConn.CommitTrans
    
    MsgBox "Bill information successfully deleted.", vbInformation, App.Title

    flexDetails.Rows = 1
    
    cmbPartyName.SetFocus
    
    Set tmprst = Nothing
End Sub

Private Sub cmdFind_Click()
    If Len(cmbPartyName.Text) = 0 Then
        MsgBox "You should provide party name.", vbInformation, App.Title
        cmbPartyName.SetFocus
        Exit Sub
    End If
    If Len(txtForwardDate.Text) = 0 Then
        MsgBox "You should provide forward date.", vbInformation, App.Title
        txtForwardDate.SetFocus
        Exit Sub
    End If
    Call prcLoadData
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    Dim intID As Integer
    Dim BillNo As String, BillDate As String
    Dim HAWB As String, BENO As String, BEDate As String
    Dim BillAmount As Currency, AppAmount As Currency, AIT As Currency, VAT As Currency, Total As Currency

''-------------------------------------
    If fncBlank(2) = True Then Exit Sub
    If flexDetails.Rows = 1 Then
        MsgBox "No record found to save.", vbCritical, App.Title
        cmbPartyName.SetFocus
    Exit Sub
    End If

''------------------- INSERT STATEMENT -----------------------
    If cmdSave.Caption = "&Save" Then
        If MsgBox("Do you want to save Bill information.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
        If fncDuplicate = True Then Exit Sub
        
        intID = fncNewID("CFOUT")
        
        MainConn.BeginTrans
            Mysql = " Insert Into C_And_F_Out_Main(M_ID, PartyID, ForwardDate, VoucherNo, Compid) "
            Mysql = Mysql + " Values (" & intID & ", " & cmbPartyName.Tag & ", '" & Format(txtForwardDate.Text, cnstDtFrmtI) & "', '" & Trim(txtVoucherNo.Text) & "', " & bytCompID & ");"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            Mysql = " Insert Into C_And_F_Out_Main_bak(M_ID, PartyID, ForwardDate, VoucherNo, Compid,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime) "
            Mysql = Mysql + " Values (" & intID & ", " & cmbPartyName.Tag & ", '" & Format(txtForwardDate.Text, cnstDtFrmtI) & "', '" & Trim(txtVoucherNo.Text) & "', " & bytCompID & ",0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            With flexDetails
                For i = 1 To .Rows - 1
                    .Row = i
                    BillNo = "": BillDate = "": HAWB = "": BENO = "": BEDate = "": BillAmount = 0: AIT = 0: VAT = 0
                        
                    .Col = 0: BillNo = .Text
                    .Col = 1: BillDate = .Text
                    .Col = 2: HAWB = .Text
                    .Col = 3: BENO = .Text
                    .Col = 4: BEDate = IIf(Len(.Text) = 0, "01-Jan-1990", .Text)
                    .Col = 5: BillAmount = CDbl(.Text)
                    .Col = 6: AppAmount = CDbl(.Text)
                    .Col = 7: AIT = CDbl(.Text)
                    .Col = 8: VAT = CDbl(.Text)
                    .Col = 9: Total = CDbl(.Text)
                    
                    Mysql = " Insert Into C_And_F_Out_Sub (M_ID, BillNO, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TotalAmount, SLNO)"
                    Mysql = Mysql + " Values ( " & intID & ", '" & BillNo & "', '" & BillDate & "', '" & HAWB & "', '" & BENO & "', '" & BEDate & "', " & BillAmount & ", " & AppAmount & ", " & AIT & ", " & VAT & ", " & Total & ", " & i & ");"
                    MainComm.CommandText = Mysql
                    MainComm.Execute
                    
                    Mysql = " Insert Into C_And_F_Out_Sub_bak (M_ID, BillNO, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TotalAmount, SLNO,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)"
                    Mysql = Mysql + " Values ( " & intID & ", '" & BillNo & "', '" & BillDate & "', '" & HAWB & "', '" & BENO & "', '" & BEDate & "', " & BillAmount & ", " & AppAmount & ", " & AIT & ", " & VAT & ", " & Total & ", " & i & ",0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
                    MainComm.CommandText = Mysql
                    MainComm.Execute
                Next
            End With
            MainConn.CommitTrans
           
        MsgBox "Bill Information Successfully Saved.", vbInformation, App.Title
    Else
''------------------- UPDATE STATEMENT -----------------------
        If MsgBox("Do you want to update bill information.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
        
        MainConn.BeginTrans
            ''intID = flexDetails.Tag 'fncFindID
            intID = fncFindID
            
            ''Mysql = " DELETE BillSub.* FROM BillMain INNER JOIN BillSub ON BillMain.ChallID = BillSub.ChallID WHERE BillMain.ChallID =" & intID & ";"
            Mysql = " DELETE C_And_F_Out_Sub FROM C_And_F_Out_Main INNER JOIN C_And_F_Out_Sub ON C_And_F_Out_Main.M_ID = C_And_F_Out_Sub.M_ID WHERE (((C_And_F_Out_Sub.M_ID)=" & intID & "));"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            Mysql = "Update C_And_F_Out_Main "
            Mysql = Mysql + " Set PartyID = " & cmbChangeParty.Tag & ",  ForwardDate = '" & Format(txtChangeForwardDate.Text, cnstDtFrmtI) & "', VoucherNO = '" & txtVoucherNo.Text & "'"
            Mysql = Mysql + " Where M_ID = " & intID & ";"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            Mysql = "Insert into C_And_F_Out_Main_bak(M_ID , PartyID, ForwardDate, VoucherNo, Compid, Deleterow, Username, Userpwd, Comname, Entrydate, Entrytime) Select M_ID , PartyID, ForwardDate, VoucherNo, Compid,1,'" & strUser & "','" & StrPass & "','" & getComName & "','" & Date & "','" & Time & "' from C_And_F_Out_Main Where M_ID = " & intID & ";"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
                               
            With flexDetails
                For i = 1 To .Rows - 1
                    .Row = i
                    BillNo = "": BillDate = "": HAWB = "": BENO = "": BEDate = "": BillAmount = 0: AIT = 0: VAT = 0
                        
                    .Col = 0: BillNo = .Text
                    .Col = 1: BillDate = .Text
                    .Col = 2: HAWB = .Text
                    .Col = 3: BENO = .Text
                    .Col = 4: BEDate = IIf(Len(.Text) = 0, "01-Jan-1990", .Text)
                    .Col = 5: BillAmount = CDbl(.Text)
                    .Col = 6: AppAmount = CDbl(.Text)
                    .Col = 7: AIT = CDbl(.Text)
                    .Col = 8: VAT = CDbl(.Text)
                    .Col = 9: Total = CDbl(.Text)
                    
                                       
                    Mysql = " Insert Into C_And_F_Out_Sub (M_ID, BillNO, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TotalAmount, SLNO)"
                    Mysql = Mysql + " Values ( " & intID & ", '" & BillNo & "', '" & BillDate & "', '" & HAWB & "', '" & BENO & "', '" & BEDate & "', " & BillAmount & ", " & AppAmount & ", " & AIT & ", " & VAT & ", " & Total & ", " & i & ");"
                    MainComm.CommandText = Mysql
                    MainComm.Execute
                    
                    Mysql = " Insert Into C_And_F_Out_Sub_bak (M_ID, BillNO, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TotalAmount, SLNO,Deleterow,ComName,Username,Userpwd,Entrydate,Entrytime)"
                    Mysql = Mysql + " Values ( " & intID & ", '" & BillNo & "', '" & BillDate & "', '" & HAWB & "', '" & BENO & "', '" & BEDate & "', " & BillAmount & ", " & AppAmount & ", " & AIT & ", " & VAT & ", " & Total & ", " & i & ",1,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
                    MainComm.CommandText = Mysql
                    MainComm.Execute
                Next
            End With
        MainConn.CommitTrans
        MsgBox "Bill Information Successfully Updated.", vbInformation, App.Title
    End If
    
    flexDetails.Rows = 1
    cmbPartyName.SetFocus
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
MainConn.RollbackTrans
End Sub

Private Sub flexDetails_DblClick()
    Call flexDetails_KeyDown(13, 0)
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With flexDetails
            bytRowNo = .Row
            .Col = 0: txtBillNo.Text = .Text
            .Col = 1: txtBillDate.Text = .Text
            .Col = 2: txtHAWB.Text = .Text
            .Col = 3: txtBENO.Text = .Text
            .Col = 4: txtBEDate.Text = .Text
            .Col = 5: txtBillAmount.Text = .Text
            .Col = 6: txtApprovedAmount.Text = .Text
            .Col = 7: txtAIT.Text = .Text
            .Col = 8: txtVAT.Text = .Text
            .Col = 9: txtTotal.Text = .Text
            
            cmdAdd.Caption = "&Replace"
        End With
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
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    Call prcAddParty(cmbPartyName)
    Call prcAddParty(cmbChangeParty)
    Call prcGridSetting
    Call prcConReport(CrystalReport1)
End Sub

Public Function fncBlank(param As Integer) As Boolean
fncBlank = True
    If param = 1 Then
        If Len(txtBillNo.Text) = 0 Then
            MsgBox "You should provide bill no.", vbCritical, App.Title
            txtBillNo.SetFocus
            Exit Function
        End If

        If Len(txtBillDate.Text) = 0 Then
            MsgBox "You should bill date.", vbCritical, App.Title
            txtBillDate.SetFocus
            Exit Function
        End If
    
    ElseIf param = 2 Then
        If Len(cmbPartyName.Text) = 0 Then
            MsgBox "You should provide party name.", vbCritical, App.Title
            cmbPartyName.SetFocus
            Exit Function
        End If
        If Len(txtForwardDate.Text) = 0 Then
            MsgBox "You should provide forward date.", vbCritical, App.Title
            txtForwardDate.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Public Function fncDuplicate() As Boolean
fncDuplicate = True
    Set tmprst = New Recordset
        Mysql = "SELECT Count(*) AS RC FROM C_And_F_Out_Main WHERE PartyID = " & cmbPartyName.Tag & " and Compid = " & bytCompID & " And ForwardDate = '" & Trim(txtForwardDate.Text) & "';"
        MainComm.CommandText = Mysql
        Set R = MainComm.Execute
        If R("RC") > 0 Then
            MsgBox "Duplicate Entry. ", vbCritical, App.Title
            Exit Function
        End If
    Set tmprst = Nothing
fncDuplicate = False
End Function

Public Sub prcLoadData()
    Set tmprst = New Recordset
        cmdSave.Caption = "&Save"
        cmdDelete.Enabled = False
        cmdPreview.Enabled = False
        
        Frame1.Visible = False
    
        Mysql = "SELECT M_ID, PartyID, PartyName, ForwardDate, VoucherNo, BillNo, BillDate, HAWB, BENO, BEDate, BillAmount, ApprovedAmount, AIT, VAT, TotalAmount, SLNO FROM Query_C_And_F_Out WHERE CompID = " & bytCompID & " and PartyID = " & cmbPartyName.Tag & " And  ForwardDate = '" & Format(txtForwardDate.Text, cnstDtFrmtI) & "' order by SLNO;"
        MainComm.CommandText = Mysql
        Set tmprst = MainComm.Execute
            
        If tmprst.EOF = False And tmprst.BOF = False Then
            txtVoucherNo.Text = tmprst("VoucherNO").Value
            flexDetails.Tag = tmprst("M_ID").Value
            
            cmbChangeParty.Text = tmprst("PartyName").Value
            txtChangeForwardDate.Text = Format(tmprst("ForwardDate").Value, "DD-MMM-YYYY")
    
            With flexDetails
                Dim i As Integer
                i = 1
                .Rows = 1
                tmprst.MoveFirst
                Do While Not tmprst.EOF
                    .Rows = .Rows + 1
                    .Row = i
                    
                    .Tag = tmprst("M_ID")
                    
                    .Col = 0: .Text = tmprst("BillNO").Value & ""
                    .Col = 1: .Text = IIf(Format(tmprst("BillDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(tmprst("BillDate").Value, "DD-MMM-YYYY"))
                    ''tmprst("BillDate").Value & ""
                    .Col = 2: .Text = tmprst("HAWB").Value & ""
                    .Col = 3: .Text = tmprst("BENO").Value & ""
                    .Col = 4: .Text = IIf(Format(tmprst("BEDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(tmprst("BEDate").Value, "DD-MMM-YYYY"))
                    .Col = 5: .Text = FormatNumber(tmprst("BillAmount").Value, 2)
                    .Col = 6: .Text = FormatNumber(tmprst("ApprovedAmount").Value, 2)
                    .Col = 7: .Text = FormatNumber(tmprst("AIT").Value, 2)
                    .Col = 8: .Text = FormatNumber(tmprst("VAT").Value, 2)
                    .Col = 9: .Text = FormatNumber(tmprst("TotalAmount").Value, 2)
                i = i + 1
                tmprst.MoveNext
                Loop
            End With
            
            cmdSave.Caption = "&Update"
            cmdDelete.Enabled = True
            cmdPreview.Enabled = True
        End If
    Set tmprst = Nothing
End Sub

Private Sub txtBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtBillDate_LostFocus()
    Call FormatDate(txtBillDate)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtBillNo_LostFocus()
    txtBillNo.Text = UCase(Trim(txtBillNo.Text))
End Sub

Private Sub txtChangeForwardDate_GotFocus()
    Call txtGotFocus(txtChangeForwardDate)
End Sub

Private Sub txtChangeForwardDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtChangeForwardDate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtChangeForwardDate_LostFocus()
    Call FormatDate(txtChangeForwardDate)
End Sub

Private Sub txtForwardDate_GotFocus()
    Call txtGotFocus(txtForwardDate)
End Sub

Private Sub txtForwardDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtForwardDate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtForwardDate_LostFocus()
    Call FormatDate(txtForwardDate)
    
    If Len(cmbPartyName.Text) = 0 Or Len(txtForwardDate.Text) = 0 Then
        txtVoucherNo.Text = ""
        flexDetails.Tag = 0: flexDetails.Rows = 1
    Else
        txtVoucherNo.Text = ""
        flexDetails.Tag = 0: flexDetails.Rows = 1
        
        Call cmdFind_Click
    End If
End Sub

Private Sub txtVat_GotFocus()
    Call txtGotFocus(txtVAT)
End Sub

Private Sub txtVat_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtVat_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtVat_LostFocus()
    If Len(txtVAT.Text) = 0 Then
        txtVAT.Text = "0.00"
    Else
        txtVAT.Text = FormatNumber(txtVAT.Text, 2)
    End If
End Sub

Private Sub txttotal_GotFocus()
    Call txtGotFocus(txtTotal)
End Sub

Private Sub txttotal_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txttotal_LostFocus()
    If Len(txtTotal.Text) = 0 Then
        txtTotal.Text = "0.00"
    Else
        txtTotal.Text = FormatNumber(txtTotal.Text, 2)
    End If
End Sub

Private Sub txtVoucherNo_GotFocus()
    Call txtGotFocus(txtVoucherNo)
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Trim(UCase(txtVoucherNo.Text))
End Sub

Private Sub txtHAWB_GotFocus()
    Call txtGotFocus(txtHAWB)
End Sub

Private Sub txtHAWB_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtHAWB_LostFocus()
    txtHAWB.Text = UCase(Trim(txtHAWB.Text))
End Sub

Private Sub txtbeno_GotFocus()
    Call txtGotFocus(txtBENO)
End Sub

Private Sub txtbeno_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtBEno_LostFocus()
    txtBENO.Text = UCase(Trim(txtBENO.Text))
End Sub

Private Sub txtAIT_GotFocus()
    Call txtGotFocus(txtAIT)
End Sub

Private Sub txtAIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtAIT_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
    If KeyAscii = 46 And InStr(txtAIT.Text, ".") > 0 Then KeyAscii = 0
End Sub

Public Sub prcGridSetting()
    With flexDetails
        .Rows = 1: .Cols = 10
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        
        .ColWidth(0) = 2200 'Bill No
        .ColWidth(1) = 1500 'Bill Date
        .ColWidth(2) = 1600 'HAWB
        .ColWidth(3) = 1400  'B/E No
        .ColWidth(4) = 1400 'B/E Date
        .ColWidth(5) = 1500 'Bill Amount
        .ColWidth(6) = 1500 'Approved Amount
        .ColWidth(7) = 1000 'AIT
        .ColWidth(8) = 1000 'VAT
        .ColWidth(9) = 1500 'Total
        
        .Row = 0
        .Col = 0: .Text = "Bill No"
        .Col = 1: .Text = "Bill Date"
        .Col = 2: .Text = "HAWB"
        .Col = 3: .Text = "B/E No"
        .Col = 4: .Text = "B/E Date"
        .Col = 5: .Text = "Bill Amount"
        .Col = 6: .Text = "App. Amount"
        .Col = 7: .Text = "AIT"
        .Col = 8: .Text = "VAT"
        .Col = 9: .Text = "Total"
    End With
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    Dim BillNo As String
    If fncBlank(1) = True Then Exit Sub
    
    With flexDetails
        If .Rows > 1 Then
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0: BillNo = .Text

                If BillNo = txtBillNo.Text And inti <> bytRowNo Then
                    MsgBox "Adding Duplicate Bill NO", vbInformation, App.Title
                    txtBillNo.SetFocus
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
            cmdAdd.Caption = "&Add"
        End If
        
        .Col = 0: .Text = txtBillNo.Text
        .Col = 1: .Text = txtBillDate.Text
        .Col = 2: .Text = txtHAWB.Text
        .Col = 3: .Text = txtBENO.Text
        .Col = 4: .Text = txtBEDate.Text
        .Col = 5: .Text = FormatNumber(txtBillAmount.Text, 2)
        .Col = 6: .Text = FormatNumber(txtApprovedAmount.Text, 2)
        .Col = 7: .Text = FormatNumber(txtAIT.Text, 2)
        .Col = 8: .Text = FormatNumber(txtVAT.Text, 2)
        .Col = 9: .Text = FormatNumber(txtTotal.Text, 2)
    End With
    
    txtBillNo.Text = "": txtBillDate.Text = "": txtHAWB.Text = ""
    txtBENO.Text = "": txtBEDate.Text = "": txtBillAmount.Text = "0.00": txtApprovedAmount.Text = "0.00"
    txtAIT.Text = "0.00": txtVAT.Text = "0.00": txtTotal.Text = "0.00"
    
    txtBillNo.SetFocus
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Public Function fncFindID() As Integer
    Set tmprst = New Recordset
    Mysql = "SELECT M_ID  FROM Query_C_And_F_Out Where PartyID = " & cmbPartyName.Tag & " and Compid = " & bytCompID & " and ForwardDate = '" & Trim(txtForwardDate.Text) & "';"

    MainComm.CommandText = Mysql
    Set tmprst = MainComm.Execute
    If tmprst.EOF = False And tmprst.BOF = False Then
        fncFindID = tmprst(0).Value
    End If
    Set tmprst = Nothing
End Function

Private Sub prcTotal()
    Dim BA As Double, AA As Double, AIT As Double, VAT As Double
    BA = IIf(Len(txtBillAmount.Text) = 0, 0, txtBillAmount.Text)
    AA = IIf(Len(txtApprovedAmount.Text) = 0, 0, txtApprovedAmount.Text)
    AIT = IIf(Len(txtAIT.Text) = 0, 0, txtAIT.Text)
    VAT = IIf(Len(txtVAT.Text) = 0, 0, txtVAT.Text)
    
    lblTotal.Caption = FormatNumber(BA - (AIT + VAT), 2)
End Sub
