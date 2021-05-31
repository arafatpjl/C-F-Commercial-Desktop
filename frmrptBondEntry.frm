VERSION 5.00
Begin VB.Form frmrptBondEntry 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcomid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox CboInvoice 
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
      ItemData        =   "frmrptBondEntry.frx":0000
      Left            =   2520
      List            =   "frmrptBondEntry.frx":0002
      TabIndex        =   10
      Top             =   1680
      Width           =   2295
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
      ItemData        =   "frmrptBondEntry.frx":0004
      Left            =   2520
      List            =   "frmrptBondEntry.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
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
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Save Record"
      Top             =   5640
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
      Left            =   3840
      TabIndex        =   7
      Top             =   5640
      Width           =   1035
   End
   Begin VB.ListBox listsbNo 
      Appearance      =   0  'Flat
      Columns         =   3
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1785
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   4575
   End
   Begin VB.ComboBox CboToInvoice 
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
      ItemData        =   "frmrptBondEntry.frx":0008
      Left            =   2520
      List            =   "frmrptBondEntry.frx":000A
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdRemoveAll 
      BackColor       =   &H00C0C000&
      Caption         =   "Re&move All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C0C000&
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1950
      TabIndex        =   3
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add to List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3210
      TabIndex        =   2
      Top             =   2880
      Width           =   1185
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox txtyear 
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
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   480
      TabIndex        =   22
      Top             =   2280
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
      Index           =   2
      Left            =   2040
      TabIndex        =   21
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bond Entry Report"
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
      TabIndex        =   20
      Top             =   240
      Width           =   3615
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
      TabIndex        =   19
      Top             =   1080
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
      Index           =   0
      Left            =   2040
      TabIndex        =   18
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   1125
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   5520
      Width           =   6975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   240
      Width           =   3615
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
      Left            =   2040
      TabIndex        =   13
      Top             =   1080
      Width           =   225
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H8000000B&
      Caption         =   "Total : "
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
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   1680
      Width           =   645
   End
End
Attribute VB_Name = "frmrptBondEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbocomname_Click()
CboInvoice.Clear
CboToInvoice.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Bond_Entry.IdNo,dbo.Bond_Entry.Invoice_No,dbo.Company_Information.compID FROM Bond_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Bond_Entry.CompID where Company_Information.Name='" & cbocomname.Text & "' and Year='" & txtyear & "' and Deleterow <>1 ;"
    R.Open strRecord, MainConn, adOpenStatic
    
    
    
    
    
        If R.RecordCount > 0 Then
           Do Until R.EOF

             CboInvoice.AddItem R![Invoice_No]
             CboToInvoice.AddItem R![Invoice_No]
             txtcomid.Text = R![CompID]
             CboInvoice.ItemData(CboInvoice.NewIndex) = R.Fields(0).Value
             CboToInvoice.ItemData(CboToInvoice.NewIndex) = R.Fields(0).Value


             
             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub



Private Sub CboInvoice_Click()
CboInvoice.Tag = 0
    If Len(CboInvoice.Text) > 0 Then CboInvoice.Tag = CboInvoice.ItemData(CboInvoice.ListIndex)
End Sub



Private Sub CboToInvoice_Click()
CboToInvoice.Tag = 0
    If Len(CboToInvoice.Text) > 0 Then CboToInvoice.Tag = CboToInvoice.ItemData(CboToInvoice.ListIndex)
End Sub

Private Sub cmdAdd_Click()
If Len(CboInvoice) = 0 Or Len(CboToInvoice) = 0 Then
        MsgBox "Select SBno", vbInformation, cnstMsgInfo
        CboInvoice.SetFocus
    Exit Sub
    End If

    If CboInvoice.Text = CboToInvoice.Text Then
        listsbNo.AddItem CboInvoice.Text
        txtTotal.Text = CStr(listsbNo.ListCount)
    Exit Sub
    End If

    Set R = New ADODB.Recordset
    
      strRecord = "SELECT Invoice_No FROM Bond_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Bond_Entry.CompID where IdNo Between '" & CboInvoice.Tag & "' And '" & CboToInvoice.Tag & "' AND Company_Information.Name='" & cbocomname.Text & "' and Deleterow <>1 order by IdNo;"
'            & " GROUP BY EmployeeCode HAVING EmployeeCode Between '" & cboEmpCode.Text & "' And '" & cboToCode.Text & "';"
   
    R.Open strRecord, MainConn, adOpenStatic
    If R.EOF = False And R.BOF = False Then
        Do Until R.EOF
            listsbNo.AddItem R![Invoice_No]
            txtTotal.Text = CStr(listsbNo.ListCount)
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
If listsbNo.ListIndex = 0 Then
        MsgBox "Select Invoice No", vbInformation, cnstMsgInfo
        CboInvoice.SetFocus
    Exit Sub
    End If
If Len(txtTotal.Text) = 0 Then Exit Sub
'If txtTotal.Text > 30 Then
'        MsgBox "Total Quantity Existed.Can not Entry More Than 30", vbInformation, cnstMsgInfo
'        txtTotal.SetFocus
'    Exit Sub
'    End If
Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "Truncate Table Temp_Bond_Report;"
        MainComm.Execute
            listsbNo.ListIndex = -1
            For inti = 1 To listsbNo.ListCount
            listsbNo.ListIndex = listsbNo.ListIndex + 1
                    
        

            MainComm.CommandText = "INSERT INTO Temp_Bond_Report(CompName,Invoice_No,Bond_No,BondDate,Bill_No,BillDate)" _
                & " SELECT Company_Information.Name,Bond_Entry.Invoice_No,Bond_Entry.Bond_No,Bond_Entry.BondDate,Bond_Entry.Bill_No,Bond_Entry.BillDate FROM Bond_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Bond_Entry.CompID " _
                & " WHERE (Bond_Entry.Invoice_No= '" & listsbNo.Text & "')AND Year='" & txtyear.Text & "' AND Bond_Entry.DeleteRow<>1  ;"
            MainComm.Execute
            Next inti
            MainConn.CommitTrans
'            If optDetails.Value = True Then
'    CrystalReport1.ReportFileName = ReportPath + "indent statement.rpt"
'Else

   Set Report = Appl.OpenReport(ReportPath + "\Bond_Report.rpt")
'End If
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault

'Screen.MousePointer = vbDefault

Screen.MousePointer = vbDefault
End Sub


Private Sub cmdRemove_Click()
 If listsbNo.ListIndex >= 0 Then
        listsbNo.RemoveItem (listsbNo.ListIndex)
        txtTotal.Text = CStr(listsbNo.ListCount)
    End If
End Sub

Private Sub cmdRemoveAll_Click()
 listsbNo.Clear
    txtTotal.Text = ""
End Sub

Private Sub Form_Load()
'Call prcConReport(CrystalReport1)
txtyear.Text = Year(Date)
'CboSBno.Visible = False
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
