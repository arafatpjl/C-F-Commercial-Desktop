VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptChallanPaid 
   Caption         =   "Challan Paid Report"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcomid 
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cbocomname 
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
      ItemData        =   "frmRprChallaPaid.frx":0000
      Left            =   2160
      List            =   "frmRprChallaPaid.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
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
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Save Record"
      Top             =   4080
      Width           =   1185
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
      Left            =   3360
      TabIndex        =   8
      Top             =   4080
      Width           =   1185
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1875
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
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   1875
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name."
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
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1365
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
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Paid Report"
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
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Date  :"
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
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   -1080
      TabIndex        =   5
      Top             =   3960
      Width           =   8895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indent Statement"
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
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmRptChallanPaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbocomname_Click()

    Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Company_Information.compID FROM  dbo.Company_Information  where Company_Information.Name='" & cbocomname.Text & "' ;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF


             txtcomid.Text = R![CompID]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub



Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmdShow_Click()
    Dim strStatus As String

    If Len(txtFromDate) = 0 Then
        MsgBox "Enter From Date", vbInformation, cnstMsgInfo
        txtFromDate.SetFocus
    Exit Sub
    End If
    
    If Len(txtToDate) = 0 Then
        MsgBox "Enter To Date", vbInformation, cnstMsgInfo
        txtToDate.SetFocus
    Exit Sub
    End If

    varBeginDate = Format(txtFromDate.Text, cnstDtFrmtQ)
    varEndDate = Format(txtToDate.Text, cnstDtFrmtQ)



    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Challan_STATUS_Report;"
        MainComm.Execute

            MainComm.CommandText = "INSERT INTO Temp_Challan_STATUS_Report(CompName,Sb_No,Date,Amount,Year,STATUS)" _
                & " SELECT Company_Information.Name,Challan_Entry.Sb_No,Challan_Entry.ChallanDate,Challan_Entry.Amount,Challan_Entry.Year,Challan_Entry.Paid" _
                & " FROM (Company_Information INNER JOIN Challan_Entry ON Company_Information.compID = Challan_Entry.compID)" _
                & " WHERE (((Challan_Entry.Entrydate) Between '" & varBeginDate & "' And '" & varEndDate & "') AND ((Challan_Entry.compID)='" & txtcomid.Text & "')" _
                & " );"
            MainComm.Execute


        

    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If


 Set Report = Appl.OpenReport(ReportPath + "\Challan_Status_Report.rpt")
'    End If
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()


Set R = New ADODB.Recordset
    strRecord = "SELECT CompID,Name FROM Company_Information ORDER BY Name;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

             cbocomname.AddItem R![Name]
             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFromDate_LostFocus()
If Len(txtFromDate) = 0 Then Exit Sub
    txtFromDate.Text = FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtToDate_LostFocus()
If Len(txtToDate) = 0 Then Exit Sub
    txtToDate.Text = FormatDate(txtToDate)
End Sub
