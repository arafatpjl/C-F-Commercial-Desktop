VERSION 5.00
Begin VB.Form frmNewIndentEntry 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Indent Information"
   ClientHeight    =   3675
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5985
   Icon            =   "frmIndentFund.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5985
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtJobNo 
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
      MaxLength       =   30
      TabIndex        =   0
      Top             =   660
      Width           =   2295
   End
   Begin VB.TextBox txtInvNo 
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
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
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
      Left            =   5490
      Picture         =   "frmIndentFund.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Add New Head of Accounts"
      Top             =   2760
      Width           =   285
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
      Height          =   315
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3180
      Width           =   3585
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
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2760
      Width           =   3285
   End
   Begin VB.TextBox txtIndAmount 
      Alignment       =   2  'Center
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
      TabIndex        =   3
      Top             =   2340
      Width           =   2295
   End
   Begin VB.TextBox txtPayDate 
      Alignment       =   2  'Center
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
      TabIndex        =   2
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   2160
      TabIndex        =   9
      Top             =   30
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   180
         Value           =   -1  'True
         Width           =   975
      End
   End
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
      Left            =   4620
      TabIndex        =   7
      Top             =   1530
      Width           =   1155
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
      Left            =   4620
      TabIndex        =   6
      ToolTipText     =   "Save Record"
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Indent Value"
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
      TabIndex        =   26
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label Label13 
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
      TabIndex        =   25
      Top             =   1920
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
      TabIndex        =   23
      Top             =   660
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
      TabIndex        =   22
      Top             =   660
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
      TabIndex        =   21
      Top             =   1080
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
      TabIndex        =   20
      Top             =   1080
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
      Left            =   1890
      TabIndex        =   18
      Top             =   3180
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
      Left            =   1890
      TabIndex        =   17
      Top             =   2760
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
      Left            =   210
      TabIndex        =   16
      Top             =   2760
      Width           =   1635
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
      Left            =   1890
      TabIndex        =   15
      Top             =   2340
      Width           =   225
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "New Indent Value"
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
      TabIndex        =   14
      Top             =   2340
      Width           =   1605
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
      Left            =   1890
      TabIndex        =   13
      Top             =   1500
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
      Left            =   210
      TabIndex        =   12
      Top             =   1500
      Width           =   1275
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
      Left            =   210
      TabIndex        =   8
      Top             =   3180
      Width           =   855
   End
End
Attribute VB_Name = "frmNewIndentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtRemarks.SetFocus
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdNewAccHead_Click()
    strMark = "ACC"
    frmNewAccHead.Show 1
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
    Dim strStatus As String
    
    If Len(cboAccNo) = 0 Then
        MsgBox "Select Head of Account", vbInformation, cnstMsgInfo
        cboAccNo.SetFocus
    Exit Sub
    End If

    If Len(txtPayDate) = 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtPayDate.SetFocus
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

    If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

    If optImport.Value = True Then strStatus = "I" Else strStatus = "E"
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "INSERT INTO new_Indent_Information(compID,JobNo,InvNo,PayDate,IndAmount,AdjAmount,AdjTax,AdjVat,AccID,Remarks,Status)" _
              & " VALUES(" & bytCompID & ",'" & txtJobNo.Text & "','" & txtInvNo.Text & "','" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," _
              & " " & CCur(txtIndAmount.Text) & ",0,0,0," & findAccID(cboAccNo.Text) & ",'" & txtRemarks.Text & "','" & strStatus & "');"
        MainComm.Execute
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

txtJobNo.Text = ""
txtInvNo.Text = ""
txtIndAmount.Text = ""
txtRemarks.Text = ""
txtTotal.Text = Format(findTotal, "##,##0.00")

txtJobNo.SetFocus
Screen.MousePointer = vbDefault

End Sub
Private Sub Form_Activate()
    If strMark = "ACC" Then addAccName cboAccNo
End Sub
Private Sub Form_Load()
    Call addAccName(cboAccNo)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMark = ""
    Set frmNewIndentEntry = Nothing
End Sub

Private Sub txtIndAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtIndAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cboAccNo.SetFocus
End Sub

Private Sub txtIndAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub

Private Sub txtIndAmount_LostFocus()
    txtIndAmount.Text = Format(txtIndAmount.Text, "##,##0.00")
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

    txtInvNo.Text = txtJobNo.Text
End Sub

Private Sub txtPayDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtIndAmount.SetFocus
End Sub

Private Sub txtPayDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub

Private Sub txtPayDate_LostFocus()
    txtTotal.Text = ""
    
    If Len(txtPayDate) = 0 Then Exit Sub
    
    txtPayDate.Text = FormatDate(txtPayDate)
    varBeginDate = "#" & Format(txtPayDate.Text, cnstDtFrmtI) & "#"
    txtTotal.Text = Format(findTotal, "##,##0.00")
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
        strRecord = "SELECT Sum(IndAmount) AS TValue FROM new_Indent_Information WHERE ((PayDate=" & varBeginDate & ") AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT Sum(IndAmount) AS TValue FROM new_Indent_Information WHERE ((PayDate=" & varBeginDate & ") AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TValue]) Then findTotal = R![TValue]
    End If
    R.Close
    Set R = Nothing
End Function
