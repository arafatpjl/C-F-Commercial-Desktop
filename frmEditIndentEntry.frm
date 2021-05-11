VERSION 5.00
Begin VB.Form frmEditIndentEntry 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Indent Information"
   ClientHeight    =   8865
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   8910
   ControlBox      =   0   'False
   DrawMode        =   4  'Mask Not Pen
   Icon            =   "frmEditIndentEntry.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   69
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdjUpdate 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit Adjust"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   68
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   67
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   66
      ToolTipText     =   "Update Record"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton CmdFind 
      BackColor       =   &H00C0C000&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   62
      ToolTipText     =   "Update Record"
      Top             =   1620
      Width           =   915
   End
   Begin VB.TextBox txtTotAdjust 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   4845
      Width           =   2085
   End
   Begin VB.TextBox txtAdjDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
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
      Left            =   2475
      TabIndex        =   2
      Top             =   2790
      Width           =   1935
   End
   Begin VB.TextBox txtJobNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
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
      Left            =   2475
      MaxLength       =   30
      TabIndex        =   47
      Top             =   1980
      Width           =   3600
   End
   Begin VB.ComboBox cboJobNo 
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
      ItemData        =   "frmEditIndentEntry.frx":0442
      Left            =   2475
      List            =   "frmEditIndentEntry.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cboJobNo"
      Top             =   1635
      Width           =   3615
   End
   Begin VB.TextBox txtInvNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
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
      Left            =   2475
      MaxLength       =   30
      TabIndex        =   46
      Top             =   2325
      Width           =   3615
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EAF5F5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2805
      Left            =   6660
      TabIndex        =   45
      Top             =   5250
      Width           =   2025
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   180
      TabIndex        =   42
      Top             =   5250
      Width           =   6450
      Begin VB.TextBox txtBalance 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0.00"
         Top             =   570
         Width           =   2025
      End
      Begin VB.TextBox txtIndAmount 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2175
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   225
         Width           =   2025
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance Amount"
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
         Left            =   105
         TabIndex        =   57
         Top             =   615
         Width           =   1935
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1965
         TabIndex        =   56
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1965
         TabIndex        =   44
         Top             =   255
         Width           =   225
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Indent Amount"
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
         Left            =   105
         TabIndex        =   43
         Top             =   255
         Width           =   1875
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAF5F5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   180
      TabIndex        =   33
      Top             =   6300
      Width           =   6450
      Begin VB.ComboBox cboAccNo 
         BackColor       =   &H80000009&
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
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   3585
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1320
         Width           =   3585
      End
      Begin VB.TextBox txtTax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
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
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtVat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
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
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Vat"
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
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax"
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
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Expenses"
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
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1980
         TabIndex        =   37
         Top             =   1000
         Width           =   225
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1980
         TabIndex        =   36
         Top             =   1380
         Width           =   225
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1980
         TabIndex        =   35
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1980
         TabIndex        =   34
         Top             =   620
         Width           =   225
      End
   End
   Begin VB.Frame BFrame 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4500
      TabIndex        =   25
      Top             =   3300
      Width           =   4230
      Begin VB.TextBox txtBAdjAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
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
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox txtChequePayDate 
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
         Left            =   2160
         TabIndex        =   9
         Top             =   255
         Width           =   1935
      End
      Begin VB.TextBox txtChequePayValue 
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
         Left            =   2160
         TabIndex        =   10
         Top             =   652
         Width           =   1935
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Amount"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   120
         TabIndex        =   30
         Top             =   255
         Width           =   1380
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Indent Value"
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
         Left            =   120
         TabIndex        =   29
         Top             =   652
         Width           =   1380
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   28
         Top             =   630
         Width           =   225
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   26
         Top             =   1020
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   180
      TabIndex        =   17
      Top             =   3300
      Width           =   4230
      Begin VB.TextBox txtAdjAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
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
         Height          =   315
         Left            =   2145
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtPayDate 
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
         Left            =   2145
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCashPayValue 
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
         Left            =   2145
         TabIndex        =   18
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label lblAdjCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Amount"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   24
         Top             =   975
         Width           =   225
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   607
         Width           =   225
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Indent Value"
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
         Left            =   120
         TabIndex        =   20
         Top             =   637
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   7110
      TabIndex        =   14
      Top             =   1485
      Width           =   1545
      Begin VB.OptionButton optEditIm 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton optEditEx 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   3060
      TabIndex        =   0
      Top             =   840
      Width           =   2295
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
         Left            =   1230
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
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
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   65
      Top             =   8160
      Width           =   8895
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Indent Information"
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
      TabIndex        =   64
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2100
      TabIndex        =   61
      Top             =   4875
      Width           =   225
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Adjust  Amount"
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
      Left            =   300
      TabIndex        =   60
      Top             =   4890
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Date"
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
      Left            =   255
      TabIndex        =   55
      Top             =   2790
      Width           =   1275
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2205
      TabIndex        =   54
      Top             =   2775
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
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
      Left            =   255
      TabIndex        =   53
      Top             =   2325
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find Job No."
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
      Left            =   255
      TabIndex        =   52
      Top             =   1635
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "New Job No."
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
      Left            =   255
      TabIndex        =   51
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2205
      TabIndex        =   50
      Top             =   2010
      Width           =   225
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2205
      TabIndex        =   49
      Top             =   1635
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2205
      TabIndex        =   48
      Top             =   2400
      Width           =   225
   End
End
Attribute VB_Name = "frmEditIndentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cboJobNo_Click()
    If Len(cboJobNo) = 0 Then Exit Sub
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='I'));"
    Else
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='E'));"
    End If
    R.Open strRecord, MainConn, adOpenStatic
    
    If R.RecordCount > 0 Then
        txtJobNo.Text = R![JobNo]
        txtInvNo.Text = R![INVNo]
        txtPayDate.Text = Format(R![PayDate], cnstDtFrmtI)
        txtCashPayValue.Text = Format(R![IndAmount], "##,##0.00")
        txtChequePayDate.Text = Format(R![BPayDate], cnstDtFrmtI)
        txtChequePayValue.Text = Format(R![BPayValue], "##,##0.00")
        txtIndAmount.Text = Format(R![TotIndValue], "##,##0.00")
        txtAdjDate.Text = IIf(Format(R![AdjDate], cnstDtFrmtI) = "30-Dec-1899", "", Format(R![AdjDate], cnstDtFrmtI))
        txtTax.Text = Format(R![AdjTax], "##,##0.00")
        txtVat.Text = Format(R![AdjVat], "##,##0.00")
        txtAdjAmount.Text = "0.00"
        txtBAdjAmount.Text = "0.00"
        txtBalance.Text = Format(R![TotBAmount], "##,##0.00")
        txtTotAdjust.Text = Format(R![TotAdjust], "##,##0.00")
       
       '' txtBalance.Text = Format(R![IndAmount] - R![AdjAmount], "##,##0.00")
        cboAccNo.Text = R![AccName]
        txtRemarks.Text = R![Remarks]
        If R![Status] = "I" Then optEditIm.Value = True Else optEditEx.Value = True
    End If
    R.Close
    Set R = Nothing
    txtAdjDate.SetFocus
End Sub
Private Sub cboJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub CmdAdjUpdate_Click()
If Len(frmEditIndentEntry.cboJobNo.Text) <> 0 Then
    
    Set R1 = New ADODB.Recordset
    If frmEditIndentEntry.optImport.Value = True Then
        strRecord = "Select * from new_Indent_information Where  CompID=" & bytCompID & " And JobNo = '" & frmEditIndentEntry.cboJobNo.Text & "'And Status='I' ;"
    Else
    strRecord = "Select * from new_Indent_information Where  CompID=" & bytCompID & " And JobNo = '" & frmEditIndentEntry.cboJobNo.Text & "' And Status='E' ;"
    End If
    R1.Open strRecord, MainConn, adOpenStatic
    
    If R1.RecordCount <> 0 Then
    
        frmAdjustmentDetails.txtJobNo.Text = R1![JobNo]
         frmAdjustmentDetails.txtAdjDate.Text = Format(R1![AdjDate], cnstDtFrmtI)
         frmAdjustmentDetails.txtAdjAmount = Format(R1![AdjAmount], "##,##0.00")
         frmAdjustmentDetails.txtBAdjAmount = Format(R1![BAdjAmount], "##,##0.00")
         frmAdjustmentDetails.Show 1
    End If
    R1.Close
    Set R1 = Nothing
    
Else
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If Len(cboJobNo) = 0 Then
        MsgBox "Select Job No", vbInformation, cnstMsgInfo
        cboJobNo.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
                    
        MainComm.CommandText = "Insert into dbo.new_Indent_Information_bak(compID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount, BAdjAmount, TotAdjust, AdjTax, AdjVat, Deleterow, Username,Userpwd, Comname, Entrydate, Entrytime)" _
                                & " SELECT  compID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, 2, '" & strUser & "','" & StrPass & "', '" & getComName & "', '" & Date & "','" & Time & "' FROM dbo.new_Indent_Information WHERE (compID=" & bytCompID & ")" _
                                & " AND (JobNo='" & cboJobNo.Text & "')  ;"
        MainComm.Execute
        
        MainComm.CommandText = "DELETE FROM new_Indent_Information WHERE ((compID=" & bytCompID & ") AND (JobNo='" & cboJobNo.Text & "'));"
        MainComm.Execute
        
        
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    optImport_Click

txtJobNo.Text = "":    txtInvNo.Text = ""
txtPayDate.Text = "":   txtIndAmount.Text = ""
txtAdjDate.Text = "":   txtAdjAmount.Text = ""
txtTax.Text = "":       txtVat.Text = ""
txtBalance.Text = "":   cboAccNo.ListIndex = -1
txtRemarks.Text = "": txtCashPayValue.Text = ""
txtChequePayDate.Text = "": txtChequePayValue.Text = ""
txtBAdjAmount.Text = "": txtTotAdjust.Text = ""
optEditIm.Value = False
optEditEx.Value = False

cboJobNo.SetFocus
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdFind_Click()
If Len(cboJobNo) = 0 Then Exit Sub
Call CheckRecord
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='I'));"
    Else
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='E'));"
    End If
    R.Open strRecord, MainConn, adOpenStatic
    
    If R.RecordCount > 0 Then
        txtJobNo.Text = R![JobNo]
        txtInvNo.Text = R![INVNo]
        txtPayDate.Text = Format(R![PayDate], cnstDtFrmtI)
        txtCashPayValue.Text = Format(R![IndAmount], "##,##0.00")
        txtChequePayDate.Text = Format(R![BPayDate], cnstDtFrmtI)
        txtChequePayValue.Text = Format(R![BPayValue], "##,##0.00")
        txtIndAmount.Text = Format(R![TotIndValue], "##,##0.00")
        txtAdjDate.Text = IIf(Format(R![AdjDate], cnstDtFrmtI) = "30-Dec-1899", "", Format(R![AdjDate], cnstDtFrmtI))
        txtTax.Text = Format(R![AdjTax], "##,##0.00")
        txtVat.Text = Format(R![AdjVat], "##,##0.00")
        txtAdjAmount.Text = "0.00"
        txtBAdjAmount.Text = "0.00"
        txtBalance.Text = Format(R![TotBAmount], "##,##0.00")
        txtTotAdjust.Text = Format(R![TotAdjust], "##,##0.00")
       '' txtBalance.Text = Format(R![IndAmount] - R![AdjAmount], "##,##0.00")
        cboAccNo.Text = R![AccName]
        txtRemarks.Text = R![Remarks]
        If R![Status] = "I" Then optEditIm.Value = True Else optEditEx.Value = True
    End If
    R.Close
    Set R = Nothing
'    txtAdjDate.SetFocus
End Sub
Private Sub cmdSave_Click()
    Dim strNewStatus  As String, strOldStatus As String
    Dim StrTotAdjust As Currency, StrCashAdjust As Currency, StrBankAdjust As Currency
    
    If Len(cboJobNo) = 0 Then
        MsgBox "Select Job No", vbInformation, cnstMsgInfo
        cboJobNo.SetFocus
    Exit Sub
    End If

    If Len(txtJobNo) = 0 Then
        MsgBox "Enter Job No", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
    Exit Sub
    End If

    If Len(txtInvNo) = 0 Then
        MsgBox "Enter Invoice No", vbInformation, cnstMsgInfo
        txtInvNo.SetFocus
    Exit Sub
    End If
    
    If Len(txtAdjDate) = 0 Then
        MsgBox "Enter Adjust Date", vbInformation, cnstMsgInfo
        txtAdjDate.SetFocus
        Exit Sub
    End If
    
    If Len(txtPayDate) = 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtPayDate.SetFocus
    Exit Sub
    End If

    If Len(txtIndAmount) = 0 Then
        MsgBox "Enter Indent Amount", vbInformation, cnstMsgInfo
        txtIndAmount.SetFocus
    Exit Sub
    End If
    
    If CCur(txtIndAmount) = 0 Then
        MsgBox "Invalid Indent Amount", vbInformation, cnstMsgInfo
        txtIndAmount.SetFocus
    Exit Sub
    End If
    
    If Len(cboAccNo) = 0 Then
        MsgBox "Select Head of Account", vbInformation, cnstMsgInfo
        cboAccNo.SetFocus
    Exit Sub
    End If
''Added By Shemul
StrTotAdjust = TotAdjustAmt
StrCashAdjust = TotCashAdjustAmt
StrBankAdjust = TotBankAdjustAmt

    If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
    If Len(txtChequePayValue) = 0 Then txtChequePayValue.Text = "0"
    If optImport.Value = True Then strOldStatus = "I" Else strOldStatus = "E"
    If optEditIm.Value = True Then strNewStatus = "I" Else strNewStatus = "E"

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
    

        MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "',IndAmount=" & CCur(txtCashPayValue.Text) & ",BPayDate='" & Format(IIf(txtChequePayDate.Text <> "", txtChequePayDate.Text, "JAN-01-1900"), cnstDtFrmtI) & "'," _
            & " BPayValue=" & CCur(txtChequePayValue.Text) & ",TotIndValue=" & CCur(txtIndAmount.Text) & ",AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) + StrCashAdjust & ",BAdjAmount=" & CCur(txtBAdjAmount.Text) + StrBankAdjust & ",TotAdjust=" & CCur(txtAdjAmount.Text) + CCur(txtBAdjAmount.Text) + StrTotAdjust & ",TotBAmount=" & CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text) & ",AdjTax=" & CCur(txtTax.Text) & "," _
            & " AdjVat=" & CCur(txtVat.Text) & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "',Status='" & strNewStatus & "' WHERE ((compID=" & bytCompID & ")" _
            & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "'));"
        MainComm.Execute
          
        MainComm.CommandText = "Insert into dbo.new_Indent_Information_bak(compID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount, BAdjAmount, TotAdjust, AdjTax, AdjVat, Deleterow, Username,Userpwd, Comname, Entrydate, Entrytime)" _
                                & " SELECT  compID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, 1, '" & strUser & "','" & StrPass & "', '" & getComName & "', '" & Date & "', '" & Time & "' FROM dbo.new_Indent_Information WHERE (compID=" & bytCompID & ")" _
                                & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "');"
                   
'            MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," _
'            & " IndAmount=" & CCur(txtIndAmount.Text) & ",AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) & ",AdjTax=" & CCur(txtTax.Text) & "," _
'            & " AdjVat=" & CCur(txtVat.Text) & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "',Status='" & strNewStatus & "' WHERE ((compID=" & bytCompID & ")" _
'            & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "'));"
        MainComm.Execute
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    txtJobNo.Text = "":         txtInvNo.Text = ""
    txtPayDate.Text = "":       txtCashPayValue.Text = "0.00"
    txtChequePayDate.Text = "": txtChequePayValue.Text = "0.00"
    txtIndAmount.Text = "0.00": txtAdjAmount.Text = "0.00"
    txtTax.Text = "0.00":       txtVat.Text = "0.00"
    txtBalance.Text = "0.00":   cboAccNo.ListIndex = -1
    txtRemarks.Text = ""
    txtTotAdjust.Text = "0.00"
    optEditIm.Value = False:    optEditEx.Value = False
    
    cboJobNo.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    Call addAccName(cboAccNo)
    Call BFrameDisable
    optImport.Value = True
    txtIndAmount.Text = "0.00"
    txtBalance.Text = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditIndentEntry = Nothing
End Sub
Private Sub optExport_Click()
    optImport_Click
End Sub
Private Sub optImport_Click()
    txtJobNo.Text = "":         txtInvNo.Text = ""
    txtPayDate.Text = "":       txtIndAmount.Text = ""
    txtAdjDate.Text = "":       txtAdjAmount.Text = ""
    txtTax.Text = "":           txtVat.Text = ""
    txtBalance.Text = "":       cboAccNo.ListIndex = -1
    txtRemarks.Text = ""
    
    Set R = New ADODB.Recordset
    If optImport.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='I';"
    Else
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='E';"
    End If
    
    cboJobNo.Clear
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboJobNo.AddItem R![JobNo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub
Private Sub txtAdjAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtAdjAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtAdjAmount_LostFocus()
    txtTotAdjust.Text = "0.00"
    If Len(txtAdjAmount) = 0 Then txtAdjAmount.Text = "0.00"
'    txtAdjAmount.Text = Format(txtAdjAmount.Text, "##,##0.00")
    txtTotAdjust.Text = Format(CCur(TotAdjustAmt) + CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
    txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text), "##,##0.00")
'    txtBalance.Text = Format(findTotal - Val(txtTotAdjust.Text), "##,##0.00")
   End Sub
Private Sub txtAdjDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtAdjDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtAdjDate_LostFocus()
    If Len(txtAdjDate) = 0 Then Exit Sub
    txtAdjDate.Text = FormatDate(txtAdjDate)
    
End Sub

Private Sub txtBAdjAmount_GotFocus()
SendKeys "{Home}+{End}"
End Sub
Private Sub txtBAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtBAdjAmount_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub
Private Sub txtBAdjDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtBAdjDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtBAdjDate_KeyPress(KeyAscii As Integer)
 DateFunc KeyAscii
End Sub

Private Sub txtBAdjAmount_LostFocus()
  txtTotAdjust.Text = "0.00"
    If Len(txtBAdjAmount) = 0 Then txtBAdjAmount.Text = "0.00"
'    txtBAdjAmount.Text = Format(txtBAdjAmount.Text, "##,##0.00")
'    txtTotAdjust.Text = TotAdjustAmt + Format(CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
txtTotAdjust.Text = Format(CCur(TotAdjustAmt) + CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text), "##,##0.00")
End Sub

Private Sub txtBalance_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBalance_LostFocus()
    If Len(txtBalance) = 0 Then txtBalance.Text = "0.00"
    txtBalance.Text = Format(txtBalance.Text, "##,##0.00")
End Sub

Private Sub txtCashPayValue_GotFocus()
 SendKeys "{Home}+{End}"
End Sub
Private Sub txtCashPayValue_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCashPayValue_LostFocus()
    If Len(txtCashPayValue) = 0 Then txtCashPayValue.Text = "0.00"
    txtCashPayValue.Text = Format(txtCashPayValue.Text, "##,##0.00")
End Sub

Private Sub txtChequePayDate_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtChequePayDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChequePayDate_KeyPress(KeyAscii As Integer)
DateFunc KeyAscii
End Sub

Private Sub txtChequePayDate_LostFocus()
    If Len(txtChequePayDate) = 0 Then Exit Sub
    txtChequePayDate.Text = FormatDate(txtChequePayDate)
End Sub

Private Sub txtChequePayValue_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtChequePayValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChequePayValue_KeyPress(KeyAscii As Integer)
CurrFunc KeyAscii
End Sub

Private Sub txtChequePayValue_LostFocus()
    If Len(txtChequePayValue) = 0 Then txtChequePayValue.Text = "0.00"
    txtChequePayValue.Text = Format(txtChequePayValue.Text, "##,##0.00")
End Sub

Private Sub txtIndAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtIndAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtIndAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtIndAmount_LostFocus()
    If Len(txtIndAmount) = 0 Then txtIndAmount.Text = "0.00"
    txtIndAmount.Text = Format(txtIndAmount.Text, "##,##0.00")
   '' txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtAdjAmount.Text), "##,##0.00")
End Sub
Private Sub txtInvNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtInvNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtInvNo_LostFocus()
    If Len(txtInvNo) = 0 Then Exit Sub
    txtInvNo.Text = Trim(UCase(txtInvNo.Text))
End Sub
Private Sub txtJobNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtJobNo_LostFocus()
    If Len(txtJobNo) = 0 Then Exit Sub
    txtJobNo.Text = Trim(UCase(txtJobNo.Text))
    If txtJobNo.Text = cboJobNo.Text Then Exit Sub
  
    Set R = New ADODB.Recordset
    If optEditIm.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    ElseIf optEditEx.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        MsgBox "Duplicate Job No", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
    Exit Sub
    End If
    
    R.Close
    Set R = Nothing
End Sub
Private Sub txtPayDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPayDate_KeyPress(KeyAscii As Integer)
'    DateFunc KeyAscii
End Sub
Private Sub txtPayDate_LostFocus()
    If Len(txtPayDate) = 0 Then Exit Sub
    txtPayDate.Text = FormatDate(txtPayDate)

End Sub
Private Sub txtRemarks_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtTax_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTax_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtTax_LostFocus()
    If Len(txtTax) = 0 Then txtTax.Text = "0.00"
    txtTax.Text = Format(txtTax.Text, "##,##0.00")
End Sub

Private Sub txtTotAdjust_LostFocus()
    If Len(txtTotAdjust) = 0 Then txtTotAdjust.Text = "0.00"
    txtTotAdjust.Text = Format(CCur(txtTotAdjust.Text), "##,##0.00")
End Sub

Private Sub txtVat_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtVat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtVat_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtVat_LostFocus()
    If Len(txtVat) = 0 Then txtVat.Text = "0.00"
    txtVat.Text = Format(txtVat.Text, "##,##0.00")
End Sub
Private Sub BFrameDisable()
txtPayDate.Enabled = False
txtCashPayValue.Enabled = False
txtChequePayDate.Enabled = False
txtChequePayValue.Enabled = False
txtIndAmount.Enabled = False
txtBalance.Enabled = False
End Sub

Function TotAdjustAmt() As Currency
    TotAdjustAmt = 0
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT TotAdjust FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT TotAdjust FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TotAdjust]) Then TotAdjustAmt = R![TotAdjust]
    End If
    R.Close
    Set R = Nothing
End Function

Function TotCashAdjustAmt() As Currency
TotCashAdjustAmt = 0
Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT AdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT AdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![AdjAmount]) Then TotCashAdjustAmt = R![AdjAmount]
    End If
    R.Close
    Set R = Nothing
End Function
Function TotBankAdjustAmt() As Currency
TotBankAdjustAmt = 0
Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT BAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT BAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![BAdjAmount]) Then TotBankAdjustAmt = R![BAdjAmount]
    End If
    R.Close
    Set R = Nothing
End Function



Public Sub CheckRecord()
Set R = New ADODB.Recordset
    If optImport.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If

    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount = 0 Then
        MsgBox "Record Not Found", vbInformation, cnstMsgInfo
        cboJobNo.SetFocus
    cboJobNo.Text = ""
    Exit Sub
    End If

    R.Close
    Set R = Nothing


End Sub
