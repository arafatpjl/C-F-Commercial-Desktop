VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptchallanEntry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReportChallanEntry"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5160
      MaxLength       =   30
      TabIndex        =   26
      Top             =   2040
      Width           =   975
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
      Left            =   5160
      TabIndex        =   25
      ToolTipText     =   "Save Record"
      Top             =   6000
      Width           =   1035
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3480
      TabIndex        =   4
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox TxtEmpcode 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
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
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   21
      Top             =   3960
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
      TabIndex        =   19
      Top             =   3360
      Width           =   1185
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
      TabIndex        =   18
      Top             =   3360
      Width           =   1035
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
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox CboToSBno 
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
      ItemData        =   "frmrptchallanEntry.frx":0000
      Left            =   2400
      List            =   "frmrptchallanEntry.frx":0002
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
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
      Left            =   210
      TabIndex        =   14
      Top             =   3960
      Width           =   4455
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
      TabIndex        =   5
      Top             =   6000
      Width           =   1035
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
      TabIndex        =   6
      ToolTipText     =   "Save Record"
      Top             =   6000
      Width           =   1035
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
      ItemData        =   "frmrptchallanEntry.frx":0004
      Left            =   2400
      List            =   "frmrptchallanEntry.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ComboBox CboSBno 
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
      ItemData        =   "frmrptchallanEntry.frx":0008
      Left            =   2400
      List            =   "frmrptchallanEntry.frx":000A
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
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
      Left            =   5400
      TabIndex        =   27
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   23
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empcode."
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
      Left            =   360
      TabIndex        =   22
      Top             =   2880
      Width           =   1125
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
      Left            =   5280
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblTo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   2280
      Width           =   255
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
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   225
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
      Left            =   1680
      TabIndex        =   12
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   -360
      TabIndex        =   10
      Top             =   5880
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S/B No."
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
      TabIndex        =   9
      Top             =   1680
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
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   225
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
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1365
   End
End
Attribute VB_Name = "frmrptchallanEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbocomname_Click()
CboSBno.Clear
CboToSBno.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Challan_Entry.IdNo,dbo.Challan_Entry.Sb_No,dbo.Company_Information.compID FROM Challan_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Challan_Entry.CompID where Company_Information.Name='" & cbocomname.Text & "' and Year='" & txtyear & "' and Deleterow <>1 ;"
    R.Open strRecord, MainConn, adOpenStatic
    
    
    
    
    
        If R.RecordCount > 0 Then
           Do Until R.EOF

             CboSBno.AddItem R![Sb_No]
             CboToSBno.AddItem R![Sb_No]
             txtcomid.Text = R![CompID]
             CboSBno.ItemData(CboSBno.NewIndex) = R.Fields(0).Value
             CboToSBno.ItemData(CboToSBno.NewIndex) = R.Fields(0).Value

             
             
'            If R.EOF = False And R.BOF = False Then
'            Do While Not R.EOF
'            cmb.AddItem R.Fields(1).Value
'            cmb.ItemData(cmb.NewIndex) = R.Fields(0).Value
             
             
             
             
'Public Sub prcAddExpenses(ByRef cmb As ComboBox)
'    Set R = New ADODB.Recordset
'    strRecord = "Select * from New_Expense Order By ExpenseName;"
'    R.Open strRecord, MainConn, adOpenStatic
'
'    cmb.Clear
'    If R.EOF = False And R.BOF = False Then
'        Do While Not R.EOF
'            cmb.AddItem R.Fields(1).Value
'            cmb.ItemData(cmb.NewIndex) = R.Fields(0).Value
'        R.MoveNext
'        Loop
'    End If
'    Set R = Nothing
'End Sub
'
             
'  cmbExpenses.Tag = 0
'    If Len(cmbExpenses.Text) > 0 Then cmbExpenses.Tag = cmbExpenses.ItemData(cmbExpenses.ListIndex)
'    cmbExpenses.ToolTipText = cmbExpenses.Text
             
             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub CboSBno_CLICK()
  CboSBno.Tag = 0
    If Len(CboSBno.Text) > 0 Then CboSBno.Tag = CboSBno.ItemData(CboSBno.ListIndex)

End Sub

Private Sub CboToSBno_CLICK()
CboToSBno.Tag = 0
    If Len(CboToSBno.Text) > 0 Then CboToSBno.Tag = CboToSBno.ItemData(CboToSBno.ListIndex)
End Sub

Private Sub cmdAdd_Click()
If Len(CboSBno) = 0 Or Len(CboToSBno) = 0 Then
        MsgBox "Select SBno", vbInformation, cnstMsgInfo
        CboSBno.SetFocus
    Exit Sub
    End If

    If CboSBno.Text = CboToSBno.Text Then
        listsbNo.AddItem CboSBno.Text
        txtTotal.Text = CStr(listsbNo.ListCount)
    Exit Sub
    End If

    Set R = New ADODB.Recordset
    
      strRecord = "SELECT Sb_No FROM Challan_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Challan_Entry.CompID where IdNo Between '" & CboSBno.Tag & "' And '" & CboToSBno.Tag & "' AND Company_Information.Name='" & cbocomname.Text & "' and Deleterow <>1 order by IdNo;"
'            & " GROUP BY EmployeeCode HAVING EmployeeCode Between '" & cboEmpCode.Text & "' And '" & cboToCode.Text & "';"
   
    R.Open strRecord, MainConn, adOpenStatic
    If R.EOF = False And R.BOF = False Then
        Do Until R.EOF
            listsbNo.AddItem R![Sb_No]
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
        MsgBox "Select SBno", vbInformation, cnstMsgInfo
        CboSBno.SetFocus
    Exit Sub
    End If
If Len(txtTotal.Text) = 0 Then Exit Sub
If txtTotal.Text > 44 Then
        MsgBox "Total Quantity Existed.Can not Entry More Than 44", vbInformation, cnstMsgInfo
        txtTotal.SetFocus
    Exit Sub
    End If
Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Challan_Report;"
        MainComm.Execute
            listsbNo.ListIndex = -1
            For inti = 1 To listsbNo.ListCount
            listsbNo.ListIndex = listsbNo.ListIndex + 1
                    
        
            MainComm.CommandText = "INSERT INTO Temp_Challan_Report(CompName,EmpName,Sb_No,Date,Amount)" _
                & " SELECT Company_Information.Name,'" & TxtName.Text & "',Challan_Entry.Sb_No,Challan_Entry.ChallanDate,Challan_Entry.Amount FROM Challan_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Challan_Entry.CompID " _
                & " WHERE (Challan_Entry.Sb_No= '" & listsbNo.Text & "')AND Year='" & txtyear.Text & "' AND Challan_Entry.DeleteRow<>1  ;"
            MainComm.Execute
            Next inti
            MainConn.CommitTrans


Set Report = Appl.OpenReport(ReportPath + "\Challan_Report.rpt")
'End If
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
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

Private Sub cmdUpdate_Click()
If listsbNo.ListIndex = 0 Then
        MsgBox "Select SBno", vbInformation, cnstMsgInfo
        CboSBno.SetFocus
    Exit Sub
    End If


Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        
            listsbNo.ListIndex = -1
            For inti = 1 To listsbNo.ListCount
            listsbNo.ListIndex = listsbNo.ListIndex + 1
                    
        
'            MainComm.CommandText = "INSERT INTO Temp_Challan_Report(CompName,EmpName,Sb_No,Date,Amount)" _
'                & " SELECT Company_Information.Name,'" & TxtName.Text & "',Challan_Entry.Sb_No,Challan_Entry.ChallanDate,Challan_Entry.Amount FROM Challan_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Challan_Entry.CompID " _
'                & " WHERE (Challan_Entry.Sb_No= '" & listsbNo.Text & "') ;"

MainComm.CommandText = "UPDATE Challan_Entry SET Paid ='OK'," _
                        & "DeleteRow =2,Username= '" & strUser & "', Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "'" _
                        & "where SB_NO='" & listsbNo.Text & "' and Year='" & txtyear.Text & "' AND DeleteRow <> 1;"

            MainComm.Execute
            Next inti
            MsgBox "Transaction  successfully Updated.", vbInformation, cnstMsgInfo
            MainConn.CommitTrans


Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

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

Private Sub TxtEmpcode_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub TxtEmpcode_LostFocus()
On Error GoTo ErrorHandler
If txtcomid <> "" Then
 Set R = New ADODB.Recordset
    strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & TxtEmpcode & "' and mleft=0 and compid=" & txtcomid & "  ;"
    R.Open strRecord, MainConn, adOpenStatic
    TxtName = ""
    If R.RecordCount > 0 Then
        TxtName = R![EmpName]
        
        Else
    MsgBox "Sorry,Check Your Employee Code"
        
    End If
    R.Close
    Set R = Nothing
    End If
ErrorHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, cnstMsgErDB
End Sub
