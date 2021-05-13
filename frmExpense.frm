VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExpense 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Of The Expenses Head "
   ClientHeight    =   5100
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   4980
   Icon            =   "frmExpense.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
      Enabled         =   0   'False
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
      Left            =   1440
      TabIndex        =   4
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   3645
      Left            =   120
      TabIndex        =   3
      Top             =   750
      Width           =   4695
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BB00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   3270
         Width           =   675
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   2640
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4657
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   255
         BackColor       =   -2147483626
         ForeColor       =   0
         BackColorFixed  =   -2147483632
         ForeColorFixed  =   16776960
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483628
         BackColorBkg    =   15132390
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BB00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entries  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   3240
         Width           =   4395
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BB00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Name of Expenses Head"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   4560
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   8895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "List of the Expenses Head"
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
      TabIndex        =   10
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  strMark = "EXPENSE"
  flagMark = "INSERT"
  strValue = ""
  
  frmNewExpense.Show 1
  
  Call prcLoadData
End Sub
Private Sub cmdDelete_Click()
On Error GoTo x
    If MsgBox("Are you sure to delete" & vbCrLf & "Expense '" & flexDetails.Text & "' ?", vbQuestion + vbYesNo, AppTitle) = vbNo Then Exit Sub
        MainConn.BeginTrans
            MainComm.CommandText = "DELETE FROM new_Expense WHERE ExpenseName='" & flexDetails.Text & "';"
            MainComm.Execute
        MainConn.CommitTrans
    Call prcLoadData
    Call prcGridFormat(Me, flexDetails)
Exit Sub
x:
MsgBox Err.Description, vbCritical, cnstMsgErDB
MainConn.RollbackTrans
Exit Sub
End Sub
Private Sub cmdEdit_Click()
    strMark = "EXPENSE"
    flagMark = "EDIT"
    
    With frmNewExpense
        flexDetails.Col = 0:    .Tag = flexDetails.Text
        flexDetails.Col = 1:
            .txtDist.Text = flexDetails.Text
            .txtDist.Tag = flexDetails.Text
        .Caption = "Edit Expense"
        .Show 1
    End With
    Call prcLoadData
    Call prcGridFormat(Me, flexDetails)
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    
    Call prcGridInitialize
    Call prcLoadData
Exit Sub
x:
MsgBox Err.Description, vbCritical, AppTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExpense = Nothing
End Sub

Private Sub prcLoadData()
On Error GoTo x
    Set R = New ADODB.Recordset
        Dim inti As Integer
        strRecord = "SELECT ExpenseID, ExpenseName FROM new_Expense ORDER BY ExpenseName;"
'        MainComm.CommandText = strRecord
'        Set R = MainComm.Execute
    R.Open strRecord, MainConn, adOpenStatic
            
        inti = 0
        flexDetails.Rows = 0
        If R.EOF = False And R.BOF = False Then
            With flexDetails
                Do While Not R.EOF
                    inti = inti + 1
                    
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    
                    .Col = 0:  .Text = R(0).Value
                    .Col = 1:  .Text = R(1).Value
                R.MoveNext
                Loop
            End With
            txtTotal.Text = inti
        End If
        Call prcGridFormat(Me, flexDetails)
    Set R = Nothing
Exit Sub
x:
MsgBox Err.Description, vbCritical, AppTitle
End Sub

Private Sub prcGridInitialize()
    With flexDetails
        .Row = 0
        .Cols = 2
        
        .ColWidth(0) = 0: .ColWidth(1) = 4000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub
