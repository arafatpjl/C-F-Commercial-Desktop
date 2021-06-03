VERSION 5.00
Begin VB.Form frmrptInvoiceEntry 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFromDate 
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
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   1875
   End
   Begin VB.TextBox txtToDate 
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
      Left            =   3600
      TabIndex        =   1
      Top             =   1200
      Width           =   1875
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C000&
      Caption         =   "&Preview"
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
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Save Record"
      Top             =   2160
      Width           =   1035
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
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label3 
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
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label Label4 
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
      Left            =   720
      TabIndex        =   8
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Entry Report"
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
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Entry Report"
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
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmrptInvoiceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()

'Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "Truncate Table Temp_Invoice_Report;"
        MainComm.Execute


            MainComm.CommandText = "INSERT INTO Temp_Invoice_Report(CompName,Date,Receipt_No,Invoice_No,Amount)" _
                & " SELECT Company_Information.Name,Invoice_Entry.Date,Invoice_Entry.Receipt_No,Invoice_Entry.Invoice_No,Invoice_Entry.Amount FROM Invoice_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Invoice_Entry.CompID " _
                & " WHERE (Invoice_Entry.Date Between  '" & txtFromDate.Text & "' AND '" & txtToDate.Text & "') AND (Invoice_Entry.DeleteRow<>1) AND Invoice_Entry.CompID=" & bytCompID & "  ;"
            MainComm.Execute
'             R.Close
'            Next inti
            MainConn.CommitTrans

    Set Report = Appl.OpenReport(ReportPath & "Invoice_Report.rpt")
'End If
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault



Screen.MousePointer = vbDefault
End Sub

Private Sub txtFromDate_GotFocus()
    Call txtGotFocus(txtFromDate)
End Sub

Private Sub txtFromDate_LostFocus()
If Len(txtFromDate) = 0 Then Exit Sub
    txtFromDate.Text = FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_GotFocus()
    Call txtGotFocus(txtToDate)
End Sub

Private Sub txtToDate_LostFocus()
If Len(txtToDate) = 0 Then Exit Sub
    txtToDate.Text = FormatDate(txtToDate)
End Sub
Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
