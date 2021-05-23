VERSION 5.00
Begin VB.Form frmNewIndentEntry 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Indent Information"
   ClientHeight    =   7335
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   6495
   Icon            =   "frmNewIndentEntry.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   2490
      TabIndex        =   42
      Text            =   "0.00"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAF5F5&
      Height          =   1860
      Left            =   4980
      TabIndex        =   41
      Top             =   2400
      Width           =   1215
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
         Left            =   90
         TabIndex        =   10
         Top             =   1335
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C000&
         Caption         =   "&Save"
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
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Save Record"
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdUpdate 
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
         Left            =   90
         TabIndex        =   9
         Top             =   757
         Width           =   1035
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1740
      Left            =   210
      TabIndex        =   29
      Top             =   4800
      Width           =   6060
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
         Left            =   5550
         Picture         =   "frmNewIndentEntry.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Add New Head of Accounts"
         Top             =   930
         Width           =   285
      End
      Begin VB.ComboBox cboAccNo 
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
         Left            =   2310
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   915
         Width           =   3180
      End
      Begin VB.TextBox txtRemarks 
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
         Left            =   2310
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1290
         Width           =   3180
      End
      Begin VB.TextBox txtIndentValue 
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
         Left            =   2310
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   180
         Width           =   2295
      End
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   540
         Width           =   2295
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
         Height          =   315
         Left            =   90
         TabIndex        =   39
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Indent Value"
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
         Top             =   225
         Width           =   1605
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
         Left            =   1995
         TabIndex        =   37
         Top             =   225
         Width           =   225
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
         Left            =   90
         TabIndex        =   36
         Top             =   930
         Width           =   1635
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
         Left            =   1995
         TabIndex        =   35
         Top             =   945
         Width           =   225
      End
      Begin VB.Label Label7 
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
         Left            =   1995
         TabIndex        =   34
         Top             =   1290
         Width           =   225
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance Value"
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
         Left            =   150
         TabIndex        =   33
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label24 
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
         Left            =   1995
         TabIndex        =   32
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   960
      Left            =   210
      TabIndex        =   24
      Top             =   3420
      Width           =   4695
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
         Left            =   2295
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   570
         Width           =   2295
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
         Left            =   2295
         TabIndex        =   4
         Top             =   195
         Width           =   2295
      End
      Begin VB.Label Label19 
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
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label17 
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
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label20 
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
         Left            =   1995
         TabIndex        =   26
         Top             =   540
         Width           =   225
      End
      Begin VB.Label Label18 
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
         Left            =   1995
         TabIndex        =   25
         Top             =   180
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1035
      Left            =   210
      TabIndex        =   19
      Top             =   2250
      Width           =   4710
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
         Left            =   2280
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   600
         Width           =   2295
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
         Left            =   2280
         TabIndex        =   2
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label Label16 
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
         Left            =   1995
         TabIndex        =   23
         Top             =   600
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
         Left            =   1995
         TabIndex        =   22
         Top             =   225
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
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label Label15 
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
         Top             =   585
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdFind 
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
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Save Record"
      Top             =   1470
      Width           =   585
   End
   Begin VB.TextBox txtJobNo 
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
      Left            =   2550
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1470
      Width           =   2295
   End
   Begin VB.TextBox txtInvNo 
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
      Left            =   2550
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1845
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   2490
      TabIndex        =   11
      Top             =   750
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
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   47
      Top             =   6600
      Width           =   8895
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Indent Information"
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
      TabIndex        =   46
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Dated Indent Value"
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
      Left            =   210
      TabIndex        =   44
      Top             =   4530
      Width           =   1935
   End
   Begin VB.Label Label22 
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
      Left            =   2160
      TabIndex        =   43
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Job No."
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
      Left            =   210
      TabIndex        =   17
      Top             =   1470
      Width           =   1125
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
      Left            =   1890
      TabIndex        =   16
      Top             =   1470
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
      Left            =   210
      TabIndex        =   15
      Top             =   1845
      Width           =   1125
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
      Left            =   1890
      TabIndex        =   14
      Top             =   1845
      Width           =   225
   End
End
Attribute VB_Name = "frmNewIndentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdFind_Click()
On Error GoTo ErrorHandler
If Len(txtJobNo) = 0 Then
        MsgBox "Enter Job No", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
    Exit Sub
    End If
    
'    Call CheckRecord

Set R = New ADODB.Recordset
    If optImport.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If

    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount = 0 Then
        MsgBox "Record Not Found", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
    txtJobNo.Text = ""
    txtInvNo.Text = ""
    Exit Sub
    End If

    R.Close
    Set R = Nothing

'''---------
Set R = New ADODB.Recordset
If optImport.Value = True Then
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.TotBAmount," _
              & " new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & txtJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='I'));"
    Else
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.TotBAmount," _
              & " new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & txtJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='E'));"
    End If
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        txtInvNo.Text = R![JobNo]
        txtPayDate.Text = Format(R![PayDate], cnstDtFrmtI)
        txtCashPayValue.Text = Format(R![IndAmount], "##,##0.00")
        txtChequePayDate.Text = Format(R![BPayDate], cnstDtFrmtI)
        txtChequePayValue.Text = Format(R![BPayValue], "##,##0.00")
        txtIndentValue.Text = Format(R![TotIndValue], "##,##0.00")
        txtBalance.Text = Format(R![TotBAmount], "##,##0.00")
        cboAccNo.Text = R![AccName]
        txtRemarks.Text = R![Remarks]

    End If
    
    R.Close
    Set R = Nothing
    
    
   
    CmdUpdate.Enabled = True

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        
    Exit Sub
    End If
End Sub
Private Sub cmdNewAccHead_Click()
    strMark = "ACC"
    frmNewAccHead.Show 1
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    Dim strStatus As String
    Dim StrTotBalance As Currency
    
    
        Set R = New ADODB.Recordset
        If optImport.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
        Else
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

    
    If Len(cboAccNo) = 0 Then
        MsgBox "Select Head of Account", vbInformation, cnstMsgInfo
        cboAccNo.SetFocus
    Exit Sub
    End If

 
'If Len(txtPayDate) = 0 Then
'        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
'        txtPayDate.SetFocus
'    Exit Sub
'    End If
    
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

    If Len(txtPayDate) <> 0 And CCur(txtCashPayValue) = 0 Then
        MsgBox "Enter Indent Amount", vbInformation, cnstMsgInfo
        txtCashPayValue.SetFocus
    Exit Sub
    End If
    
    If Len(txtPayDate) = 0 And CCur(txtCashPayValue) <> 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtPayDate.SetFocus
    Exit Sub
    End If
    
    If Len(txtChequePayDate) <> 0 And CCur(txtChequePayValue) = 0 Then
        MsgBox "Enter Indent Amount", vbInformation, cnstMsgInfo
        txtChequePayValue.SetFocus
    Exit Sub
    End If
    
    If Len(txtChequePayDate) = 0 And CCur(txtChequePayValue) <> 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtChequePayDate.SetFocus
    Exit Sub
    End If
    
If Len(txtPayDate) = 0 And Len(txtChequePayDate) = 0 Then
    MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtPayDate.SetFocus
    Exit Sub
End If


If CCur(txtChequePayValue) = 0 And CCur(txtCashPayValue) = 0 Then
    MsgBox "Enter Indent Amount", vbInformation, cnstMsgInfo
        txtCashPayValue.SetFocus
    Exit Sub
End If

'    If CCur(txtCashPayValue) = 0 Then
'        MsgBox "Invalid Indent Amount", vbInformation, cnstMsgInfo
'        txtCashPayValue.SetFocus
'    Exit Sub
'    End If


StrTotBalance = TotBalance
    If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

    If optImport.Value = True Then strStatus = "I" Else strStatus = "E"
    Screen.MousePointer = vbHourglass

'    Set MainComm = New ADODB.Command
'    MainComm.ActiveConnection = MainConn
'    MainConn.BeginTrans
        
        If Len(txtPayDate) <> 0 And Len(txtIndentValue) <> 0 And Len(txtChequePayDate) = 0 Then
            Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
            
        MainComm.CommandText = "INSERT INTO new_Indent_Information(compID,JobNo,InvNo,PayDate,IndAmount,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," & CCur(txtCashPayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "');"
            MainComm.Execute
        
        MainComm.CommandText = "INSERT INTO new_Indent_Information_bak(compID,JobNo,InvNo,PayDate,IndAmount,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," & CCur(txtCashPayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "',0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainConn.CommitTrans
        End If
        
        If Len(txtChequePayDate) <> 0 And Len(txtChequePayValue) <> 0 And Len(txtPayDate) = 0 Then
            Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
        MainComm.CommandText = "INSERT INTO new_Indent_Information(compID,JobNo,InvNo,BPayDate,BPayValue,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtChequePayDate.Text, cnstDtFrmtI) & "'," & CCur(txtChequePayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "');"
            MainComm.Execute
        MainComm.CommandText = "INSERT INTO new_Indent_Information_bak(compID,JobNo,InvNo,BPayDate,BPayValue,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtChequePayDate.Text, cnstDtFrmtI) & "'," & CCur(txtChequePayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "',0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            MainConn.CommitTrans
        End If
        
        If Len(txtPayDate) <> 0 And Len(txtIndentValue) <> 0 And Len(txtChequePayDate) <> 0 And Len(txtChequePayValue) <> 0 Then
            Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
            MainComm.CommandText = "INSERT INTO new_Indent_Information(compID,JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," & CCur(txtCashPayValue.Text) & ",'" & Format(txtChequePayDate.Text, cnstDtFrmtI) & "'," & CCur(txtChequePayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "');"
            MainComm.Execute
            
            MainComm.CommandText = "INSERT INTO new_Indent_Information_bak(compID,JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotIndValue,AdjAmount,BAdjAmount,TotBAmount,TotAdjust,AdjTax,AdjVat,AccID,Remarks,Status,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," & CCur(txtCashPayValue.Text) & ",'" & Format(txtChequePayDate.Text, cnstDtFrmtI) & "'," & CCur(txtChequePayValue.Text) & "," _
              & " " & CCur(txtIndentValue.Text) & ",0,0," & CCur(txtIndentValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "',0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            MainConn.CommitTrans
        End If


'              MainComm.CommandText = "INSERT INTO new_Indent_Information(compID,JobNo,InvNo,PayDate,IndAmount,AdjAmount,AdjTax,AdjVat,AccID,Remarks,Status)" _
'              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," _
'              & " " & CCur(txtCashPayValue.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "');"
'        MainComm.Execute
'    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

txtJobNo.Text = ""
txtInvNo.Text = ""
txtIndentValue.Text = "0.00"
txtChequePayValue.Text = "0.00"
txtCashPayValue.Text = "0.00"
txtRemarks.Text = ""
txtTotal.Text = Format(findTotal, "##,##0.00")
txtJobNo.SetFocus
Screen.MousePointer = vbDefault

End Sub



Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdUpdate_Click()

If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        If optImport.Value = True Then strStatus = "I" Else strStatus = "E"

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        
         MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "',IndAmount=" & CCur(txtCashPayValue.Text) & ", " _
            & " BPayDate='" & Format(IIf(txtChequePayDate.Text <> "", txtChequePayDate.Text, "JAN-01-1900"), cnstDtFrmtI) & "',BPayValue=" & CCur(txtChequePayValue.Text) & ",TotIndValue=" & CCur(txtIndentValue.Text) & "," _
            & " TotBAmount=" & CCur(txtIndentValue.Text) - TotBalance & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "' WHERE ((compID=" & bytCompID & ")" _
            & " AND (JobNo='" & txtJobNo.Text & "') AND (Status='" & strStatus & "'));"
         MainComm.Execute
         
         MainComm.CommandText = "Insert into new_Indent_Information_bak(CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, Deleterow, UserName,Userpwd, Comname, Entrydate, Entrytime) " _
                                & " SELECT  CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, 1, '" & strUser & "','" & StrPass & "', '" & getComName & "', '" & Date & "', '" & Time & "' FROM  dbo.new_Indent_Information WHERE ((compID=" & bytCompID & ")" _
                                & " AND (JobNo='" & txtJobNo.Text & "') AND (Status='" & strStatus & "'));"
            
            
'            MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "',IndAmount=" & CCur(txtCashPayValue.Text) & ", " _
'            & " BPayDate='" & Format(IIf(txtChequePayDate.Text <> "", txtChequePayDate.Text, "JAN-01-1900"), cnstDtFrmtI) & "',BPayValue=" & CCur(txtChequePayValue.Text) & ",TotIndValue=" & CCur(txtIndentValue.Text) & "," _
'            & " TotBAmount=" & CCur(txtChequePayValue.Text) + TotBalance & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "' WHERE ((compID=" & bytCompID & ")" _
'            & " AND (JobNo='" & txtJobNo.Text & "') AND (Status='" & strStatus & "'));"
            
  
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
    txtIndentValue.Text = "0.00": cboAccNo.ListIndex = -1
    txtRemarks.Text = ""

    txtJobNo.SetFocus
    Screen.MousePointer = vbDefault
    CmdUpdate.Enabled = False
    txtPayDate.Enabled = True
    txtCashPayValue.Enabled = True
End Sub
Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Activate()
    If strMark = "ACC" Then addAccName cboAccNo
End Sub
Private Sub Form_Load()
    Call addAccName(cboAccNo)
    txtIndentValue.Enabled = False
    txtBalance.Enabled = False
    CmdUpdate.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    strMark = ""
    Set frmNewIndentEntry = Nothing
End Sub

Private Sub txtIndentValue_GotFocus()
 SendKeys "{Home}+{End}"
End Sub
Private Sub txtIndentValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtIndentValue_KeyPress(KeyAscii As Integer)
SingleCodeFunc KeyAscii
End Sub
Private Sub txtIndentValue_LostFocus()
    If Len(txtIndentValue) = 0 Then txtIndentValue.Text = "0.00"
    txtIndentValue.Text = Format(txtIndentValue.Text, "##,##0.00")
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
    'txtTotal.Text = ""
    If Len(txtChequePayDate) = 0 Then Exit Sub
        txtChequePayDate.Text = FormatDate(txtChequePayDate)
        varBeginDate = Format(txtChequePayDate.Text, cnstDtFrmtI)
        'txtTotal.Text = Format(findTotal, "##,##0.00")
End Sub



Private Sub txtChequePayValue_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtChequePayValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChequePayValue_KeyPress(KeyAscii As Integer)
SingleCodeFunc KeyAscii
End Sub

Private Sub txtChequePayValue_LostFocus()
If Len(txtChequePayValue) = 0 Then txtChequePayValue.Text = "0.00"
txtIndentValue.Text = Format(CCur(txtCashPayValue.Text) + CCur(txtChequePayValue.Text), "##,##0.00")
txtChequePayValue.Text = Format(txtChequePayValue.Text, "##,##0.00")
txtTotal.Text = Format(CCur(txtIndentValue.Text), "##,#00.00")
End Sub
Private Sub txtCashPayValue_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtCashPayValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtCashPayValue_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtCashPayValue_LostFocus()
If Len(txtCashPayValue) = 0 Then txtCashPayValue.Text = "0.00"

txtIndentValue.Text = Format(CCur(txtCashPayValue.Text) + CCur(txtChequePayValue.Text), "##,##0.00")

txtCashPayValue.Text = Format(txtCashPayValue.Text, "##,##0.00")
txtTotal.Text = Format(CCur(txtIndentValue.Text), "##,#00.00")

End Sub
Private Sub txtInvNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtInvNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPayDate.SetFocus
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
    If KeyCode = 13 Then txtInvNo.SetFocus
End Sub
Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtJobNo_LostFocus()
    If Len(txtJobNo) = 0 Then Exit Sub
    txtJobNo.Text = Trim(UCase(txtJobNo.Text))
  
'    Set R = New ADODB.Recordset
'    If optImport.Value = True Then
'        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
'    Else
'        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
'    End If
'
'    R.Open strRecord, MainConn, adOpenStatic
'    If R.RecordCount > 0 Then
'        MsgBox "Duplicate Job No", vbInformation, cnstMsgInfo
'        txtJobNo.SetFocus
'    Exit Sub
'    End If
'
'    R.Close
'    Set R = Nothing
'
    txtInvNo.Text = txtJobNo.Text
    txtTotal.Text = Format(findTotal, "##,##0.00")
End Sub
Private Sub txtPayDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPayDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtPayDate_LostFocus()
    'txtTotal.Text = ""
    If Len(txtPayDate) = 0 Then Exit Sub
    txtPayDate.Text = FormatDate(txtPayDate)
    varBeginDate = Format(txtPayDate.Text, cnstDtFrmtI)
    'txtTotal.Text = Format(findTotal, "##,##0.00")
End Sub
Private Sub txtRemarks_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSave.SetFocus
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Function findTotal() As Currency
    findTotal = 0
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
       strRecord = "SELECT Sum(TotIndValue) AS TValue FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
       ''strRecord = "SELECT Sum(IndAmount) AS TValue FROM new_Indent_Information WHERE ((PayDate=" & varBeginDate & ") AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT Sum(TotIndValue) AS TValue FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
        ''strRecord = "SELECT Sum(IndAmount) AS TValue FROM new_Indent_Information WHERE ((PayDate=" & varBeginDate & ") AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TValue]) Then findTotal = R![TValue]
    End If
    R.Close
    Set R = Nothing
End Function
Function TotBalance() As Currency
    TotBalance = 0
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT TotAdjust FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT TotAdjust  FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TotAdjust]) Then TotBalance = R![TotAdjust]
    End If
    R.Close
    Set R = Nothing
End Function

'Public Sub CheckRecord()
'Set R = New ADODB.Recordset
'    If optImport.Value = True Then
'        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
'    Else
'        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
'    End If
'
'    R.Open strRecord, MainConn, adOpenStatic
'    If R.RecordCount = 0 Then
'        MsgBox "Record Not Found", vbInformation, cnstMsgInfo
'        txtJobNo.SetFocus
'    txtJobNo.Text = ""
'    txtInvNo.Text = ""
'    Exit Sub
'    End If
'
'    R.Close
'    Set R = Nothing
'
'End Sub


'Function TotBalance() As Currency
'    TotBalance = 0
'    Set R = New ADODB.Recordset
'
'    If optImport.Value = True Then
'        strRecord = "SELECT TotBAmount FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
'    Else
'        strRecord = "SELECT TotBAmount  FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
'    End If
'
'    R.Open strRecord, MainConn, adOpenStatic
'    If R.RecordCount > 0 Then
'        If Not IsNull(R![TotBAmount]) Then TotBalance = R![TotBAmount]
'    End If
'    R.Close
'    Set R = Nothing
'End Function


