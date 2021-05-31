VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptAdjustmentSt 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report - Adjustment Statement"
   ClientHeight    =   4305
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   4305
   Icon            =   "frmRptAdjustmentSt.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRest 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show without PSA Job#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1590
      TabIndex        =   15
      Top             =   3150
      Width           =   2025
   End
   Begin VB.CheckBox chkPSA 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show PSA Job # Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1590
      TabIndex        =   14
      Top             =   2820
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   960
      TabIndex        =   4
      Top             =   750
      Width           =   2505
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   960
      TabIndex        =   11
      Top             =   1140
      Width           =   2505
      Begin VB.OptionButton optSummary 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Summary"
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
         Width           =   1155
      End
      Begin VB.OptionButton optDetails 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Details"
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
   Begin VB.TextBox txtToDate 
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
      Left            =   1590
      TabIndex        =   1
      Top             =   2400
      Width           =   1875
   End
   Begin VB.TextBox txtFromDate 
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
      Left            =   1590
      TabIndex        =   0
      Top             =   1890
      Width           =   1875
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
      Left            =   1530
      TabIndex        =   3
      Top             =   3750
      Width           =   1185
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00C0C000&
      Caption         =   "&Show"
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
      Left            =   2850
      TabIndex        =   2
      ToolTipText     =   "Save Record"
      Top             =   3750
      Width           =   1185
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   8895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Statement"
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
      TabIndex        =   18
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   -180
      TabIndex        =   16
      Top             =   3540
      Width           =   6405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   810
      TabIndex        =   10
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label Label1 
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
      Left            =   1320
      TabIndex        =   9
      Top             =   2400
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1890
      Width           =   225
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Date From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1035
   End
End
Attribute VB_Name = "frmRptAdjustmentSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdShow_Click()
    Dim strStatus As String

    If Len(txtFromDate) = 0 Then
        MsgBox "Enter From Adjustment Date", vbInformation, cnstMsgInfo
        txtFromDate.SetFocus
    Exit Sub
    End If

    If Len(txtToDate) = 0 Then
        MsgBox "Enter To Adjustment Date", vbInformation, cnstMsgInfo
        txtToDate.SetFocus
    Exit Sub
    End If

    varBeginDate = Format(txtFromDate.Text, cnstDtFrmtI)
    varEndDate = Format(txtToDate.Text, cnstDtFrmtI)

    If optImport.Value = True Then strStatus = "I" Else strStatus = "E"

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Adjust_Statement;"
        MainComm.Execute

        If optDetails.Value = True Then
            MainComm.CommandText = "INSERT INTO Temp_Adjust_Statement(JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotIndValue,AccName,Remarks,AdjDate,AdjAmount,BAdjAmount,TotAdjust,TotBAmount,AdjTax,AdjVat,compName,compAdd)" _
                & " SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Acc_Head.AccName," _
                & " new_Indent_Information.Remarks,new_Indent_Information.AdjDate,new_Indent_Information.AdjAmount,new_Indent_Information.BAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.TotBAmount,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat," _
                & " Company_Information.Name,Company_Information.Address FROM (Company_Information INNER JOIN new_Indent_Information ON Company_Information.compID = new_Indent_Information.compID)" _
                & " INNER JOIN new_Acc_Head ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.AdjDate) Between '" & varBeginDate & "' And '" & varEndDate & "')" _
                & " AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='" & strStatus & "'));"
            MainComm.Execute

            If chkPSA.Value = 1 And chkRest.Value = 0 Then
                MainComm.CommandText = "DELETE FROM Temp_Adjust_Statement WHERE RIGHT([JobNo],3) <> 'PSA';"
                MainComm.Execute
            ElseIf chkPSA.Value = 0 And chkRest.Value = 1 Then
                MainComm.CommandText = "DELETE FROM Temp_Adjust_Statement WHERE RIGHT([JobNo],3) = 'PSA';"
                MainComm.Execute
            End If
        Else
            If chkPSA.Value = 1 And chkRest.Value = 0 Then
                MainComm.CommandText = "INSERT INTO Temp_Adjust_Statement(AdjDate,TotIndValue,TotAdjust,AdjTax,compName,compAdd)SELECT new_Indent_Information.AdjDate,Sum(new_Indent_Information.TotIndValue) AS TotalIndent," _
                    & " Sum(new_Indent_Information.TotAdjust) AS TotalAdjust,Sum([new_Indent_Information].[TotIndValue]-[new_Indent_Information].[TotAdjust]) AS Balance,Company_Information.Name,Company_Information.Address" _
                    & " FROM new_Indent_Information INNER JOIN Company_Information ON new_Indent_Information.compID = Company_Information.compID WHERE ((RIGHT(new_Indent_Information.JobNo,3) = 'PSA')" _
                    & " And ((new_Indent_Information.Status)='" & strStatus & "') And ((new_Indent_Information.compID)=" & bytCompID & ")) GROUP BY new_Indent_Information.AdjDate,Company_Information.Name," _
                    & " Company_Information.Address HAVING (((new_Indent_Information.AdjDate) Between '" & varBeginDate & "' And '" & varEndDate & "'));"
            ElseIf chkPSA.Value = 0 And chkRest.Value = 1 Then
                MainComm.CommandText = "INSERT INTO Temp_Adjust_Statement(AdjDate,TotIndValue,TotAdjust,AdjTax,compName,compAdd)SELECT new_Indent_Information.AdjDate,Sum(new_Indent_Information.TotIndValue) AS TotalIndent," _
                    & " Sum(new_Indent_Information.TotAdjust) AS TotalAdjust,Sum([new_Indent_Information].[TotIndValue]-[new_Indent_Information].[TotAdjust]) AS Balance,Company_Information.Name,Company_Information.Address" _
                    & " FROM new_Indent_Information INNER JOIN Company_Information ON new_Indent_Information.compID = Company_Information.compID WHERE ((RIGHT(new_Indent_Information.JobNo,3) <> 'PSA')" _
                    & " And ((new_Indent_Information.Status)='" & strStatus & "') And ((new_Indent_Information.compID)=" & bytCompID & ")) GROUP BY new_Indent_Information.AdjDate,Company_Information.Name," _
                    & " Company_Information.Address HAVING (((new_Indent_Information.AdjDate) Between '" & varBeginDate & "' And '" & varEndDate & "'));"
            Else
                MainComm.CommandText = "INSERT INTO Temp_Adjust_Statement(AdjDate,TotIndValue,TotAdjust,AdjTax,compName,compAdd)SELECT new_Indent_Information.AdjDate,Sum(new_Indent_Information.TotIndValue) AS TotalIndent," _
                    & " Sum(new_Indent_Information.TotAdjust) AS TotalAdjust,Sum([new_Indent_Information].[TotIndValue]-[new_Indent_Information].[TotAdjust]) AS Balance,Company_Information.Name,Company_Information.Address" _
                    & " FROM new_Indent_Information INNER JOIN Company_Information ON new_Indent_Information.compID = Company_Information.compID WHERE (((new_Indent_Information.Status)='" & strStatus & "')" _
                    & " And ((new_Indent_Information.compID)=" & bytCompID & ")) GROUP BY new_Indent_Information.AdjDate,Company_Information.Name,Company_Information.Address" _
                    & " HAVING (((new_Indent_Information.AdjDate) Between '" & varBeginDate & "' And '" & varEndDate & "'));"
            End If
            
            
            
            
            MainComm.Execute
        End If

        strRecord = "SELECT AdjDate FROM Temp_Adjust_Statement GROUP BY AdjDate;"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount = 0 Then
            MsgBox "Record not Found", vbInformation, cnstMsgInfo
            Screen.MousePointer = vbDefault
            MainConn.RollbackTrans
        Exit Sub
        End If
        R.Close

        MainComm.CommandText = "UPDATE Temp_Adjust_Statement SET BeginDate='" & varBeginDate & "',EndDate='" & varEndDate & "';"
        MainComm.Execute

        If optDetails.Value = True Then
            MainComm.CommandText = "DELETE FROM Temp_Summary;"
            MainComm.Execute

            MainComm.CommandText = "INSERT INTO Temp_Summary(AccName,TotAdjust) SELECT AccName,Sum(([Temp_Adjust_Statement].[TotAdjust]-([Temp_Adjust_Statement].[AdjTax]+[Temp_Adjust_Statement].[AdjVat]))) AS TAdjAmount From Temp_Adjust_Statement GROUP BY AccName;"
            MainComm.Execute

            MainComm.CommandText = "INSERT INTO Temp_Summary(AccName,TotAdjust) SELECT 'INCOME TAX' AS AccName,Sum(AdjTax) AS TAdjTax From Temp_Adjust_Statement GROUP BY AccName;"
            MainComm.Execute

            MainComm.CommandText = "INSERT INTO Temp_Summary(AccName,TotAdjust) SELECT 'VAT' AS AccName,Sum(AdjVat) AS TAdjVat From Temp_Adjust_Statement GROUP BY AccName;"
            MainComm.Execute

            strRecord = "SELECT Sum(TotIndValue) AS TIndAmount,Sum(TotAdjust) AS TAdjAmount From Temp_Adjust_Statement;"
            R.Open strRecord, MainConn, adOpenStatic
            If R.RecordCount > 0 Then
'
                 MainComm.CommandText = "UPDATE Temp_Summary SET TotIndValue=" & (Val(R![TIndAmount]) - Val(R![TAdjAmount])) & ";"
                MainComm.Execute
            End If
            R.Close
        End If
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    If optDetails.Value = True Then
'    Set Report = Appl.OpenReport(ReportPath + "\adjustment st summary.rpt")
          Set Report = Appl.OpenReport(ReportPath + "\adjustment statement.rpt")
    Else
        Set Report = Appl.OpenReport(ReportPath + "\adjustment st summary.rpt")
    End If
    frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault

'Screen.MousePointer = vbDefault
Set R = Nothing
End Sub

Private Sub Form_Load()
Call prcConReport(CrystalReport1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRptAdjustmentSt = Nothing
End Sub
Private Sub txtFromDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtToDate.SetFocus
End Sub
Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtFromDate_LostFocus()
    If Len(txtFromDate) = 0 Then Exit Sub
    txtFromDate.Text = FormatDate(txtFromDate)
End Sub
Private Sub txtToDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdShow.SetFocus
End Sub
Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtToDate_LostFocus()
    If Len(txtToDate) = 0 Then Exit Sub
    txtToDate.Text = FormatDate(txtToDate)
End Sub
