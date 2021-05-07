VERSION 5.00
Begin VB.Form frmAdjustmentDetails 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1095
   ClientLeft      =   4410
   ClientTop       =   750
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7050
   Begin VB.TextBox txtJobNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   150
      TabIndex        =   8
      Top             =   330
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EAF5F5&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   630
      Width           =   1005
   End
   Begin VB.CommandButton CmdUpdate 
      BackColor       =   &H00EAF5F5&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2490
      TabIndex        =   4
      Top             =   630
      Width           =   1005
   End
   Begin VB.TextBox txtAdjAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3420
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   330
      Width           =   1545
   End
   Begin VB.TextBox txtBAdjAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   5070
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   330
      Width           =   1545
   End
   Begin VB.TextBox txtAdjDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1770
      TabIndex        =   0
      Top             =   330
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00EAF5F5&
      Caption         =   "JobNo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00EAF5F5&
      Caption         =   "Bank Adjust Value"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5160
      TabIndex        =   7
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00EAF5F5&
      Caption         =   "Cash Adjust Value"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3510
      TabIndex        =   6
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00EAF5F5&
      Caption         =   "Adjustment Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1770
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdjustmentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
'
'If Len(txtJobNo.Text) = 0 Then
'    MsgBox "Empty Job No", vbInformation, cnstMsgInfo
'    txtJobNo.SetFocus
'    Exit Sub
'End If
If Len(txtAdjDate.Text) = 0 Then
     MsgBox "Enter Adjust Date ", vbInformation, cnstMsgInfo
     txtAdjDate.SetFocus
    Exit Sub
End If

If Len(txtAdjDate.Text) <> 0 And Len(txtAdjAmount.Text) = 0 Then
    MsgBox "Enter Cash Adjust Amount", vbInformation, cnstMsgInfo
    txtAdjAmount.SetFocus
    Exit Sub
End If

If Len(txtAdjDate.Text) <> 0 And Len(txtBAdjAmount.Text) = 0 Then
    MsgBox "Enter Bank Adjust Amount", vbInformation, cnstMsgInfo
    txtBAdjAmount.SetFocus
    Exit Sub
End If

Dim StrTotalIndent As Currency
Dim TempTotAdjust As Currency
If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    StrTotalIndent = TotIndentAmt
    TempTotAdjust = Format((CCur(txtAdjAmount.Text) + CCur(txtBAdjAmount.Text)), "##,##0.00")
    StrTotalIndent = Format((StrTotalIndent - TempTotAdjust), "##,##0.00")
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
    If frmEditIndentEntry.optImport.Value = True Then
        MainComm.CommandText = "UPDATE new_Indent_Information SET AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) & ",BAdjAmount=" & CCur(txtBAdjAmount.Text) & ",TotAdjust=" & CCur(txtAdjAmount.Text) + CCur(txtBAdjAmount.Text) & ",TotBAmount=" & StrTotalIndent & " WHERE ((compID=" & bytCompID & ") AND (JobNo='" & txtJobNo.Text & "') AND (Status='I')) ;"
        MainComm.Execute
        MainComm.CommandText = "Insert into new_Indent_Information_bak (CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, Deleterow, UserName, Userpwd, Comname, Entrydate, Entrytime)" _
                            & "SELECT CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, 1, '" & strUser & "', '" & StrPass & "', '" & getComName & "','" & Date & "','" & Time & "' FROM dbo.new_Indent_Information WHERE ((compID=" & bytCompID & ") AND (JobNo='" & txtJobNo.Text & "') AND (Status='I')) ;"
        MainComm.Execute
    Else
        MainComm.CommandText = "UPDATE new_Indent_Information SET AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) & ",BAdjAmount=" & CCur(txtBAdjAmount.Text) & ",TotAdjust=" & CCur(txtAdjAmount.Text) + CCur(txtBAdjAmount.Text) & ",TotBAmount=" & StrTotalIndent & " WHERE ((compID=" & bytCompID & ") AND (JobNo='" & txtJobNo.Text & "') AND (Status='E')) ;"
        MainComm.Execute
        MainComm.CommandText = "Insert into new_Indent_Information_bak (CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, Deleterow, UserName, Userpwd, Comname, Entrydate, Entrytime)" _
                            & "SELECT CompID, JobNo, InvNo, PayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount , TotAdjust, AdjTax, AdjVat, 1, '" & strUser & "', '" & StrPass & "', '" & getComName & "','" & Date & "','" & Time & "' FROM dbo.new_Indent_Information WHERE ((compID=" & bytCompID & ") AND (JobNo='" & txtJobNo.Text & "') AND (Status='E')) ;"
    MainComm.Execute
    End If
    
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If
    txtJobNo.Text = "": txtAdjDate.Text = ""
    txtAdjAmount.Text = "0.00": txtBAdjAmount.Text = "0.00"
    frmEditIndentEntry.txtJobNo.Text = "":  frmEditIndentEntry.txtInvNo.Text = ""
    frmEditIndentEntry.txtAdjDate.Text = "": frmEditIndentEntry.txtPayDate.Text = ""
    frmEditIndentEntry.txtCashPayValue.Text = "0.00": frmEditIndentEntry.txtChequePayDate.Text = ""
    frmEditIndentEntry.txtChequePayValue.Text = "0.00": frmEditIndentEntry.txtTotAdjust.Text = "0.00"
    frmEditIndentEntry.txtIndAmount.Text = "0.00": frmEditIndentEntry.txtBalance.Text = "0.00"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtJobNo.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAdjustmentDetails = Nothing
End Sub


Function TotIndentAmt() As Currency
TotIndentAmt = 0
Set R = New ADODB.Recordset
    
    If frmEditIndentEntry.optImport.Value = True Then
        strRecord = "SELECT TotIndValue FROM new_Indent_Information WHERE ((JobNo='" & frmEditIndentEntry.cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT TotIndValue FROM new_Indent_Information WHERE ((JobNo='" & frmEditIndentEntry.cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TotIndValue]) Then TotIndentAmt = R![TotIndValue]
    End If
    R.Close
    Set R = Nothing
End Function
Private Sub txtAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtAdjAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtAdjAmount_LostFocus()
If Len(txtAdjAmount) = 0 Then txtAdjAmount.Text = "0.00"
    txtAdjAmount.Text = Format(txtAdjAmount.Text, "##,##0.00")
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

Private Sub txtBAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtBAdjAmount_KeyPress(KeyAscii As Integer)
CurrFunc KeyAscii
End Sub

Private Sub txtBAdjAmount_LostFocus()
If Len(txtBAdjAmount) = 0 Then txtBAdjAmount.Text = "0.00"
    txtBAdjAmount.Text = Format(txtBAdjAmount.Text, "##,##0.00")
End Sub
