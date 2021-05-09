VERSION 5.00
Begin VB.Form frmchallanEntry 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Challan Entry"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmchallanEntry.frx":0000
      Left            =   2640
      List            =   "frmchallanEntry.frx":0002
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
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
      Left            =   5160
      MaxLength       =   30
      TabIndex        =   23
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   21
      Text            =   "NEW"
      Top             =   120
      Width           =   975
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
      Left            =   720
      TabIndex        =   20
      ToolTipText     =   "Save Record"
      Top             =   4200
      Width           =   1035
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
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
      Left            =   1920
      TabIndex        =   19
      ToolTipText     =   "Save Record"
      Top             =   4200
      Width           =   1035
   End
   Begin VB.TextBox txtamount 
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
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Text            =   "20"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtSBNo 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1800
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
      ItemData        =   "frmchallanEntry.frx":0004
      Left            =   2640
      List            =   "frmchallanEntry.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
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
      Left            =   3120
      TabIndex        =   14
      ToolTipText     =   "Save Record"
      Top             =   4200
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
      Left            =   4320
      TabIndex        =   13
      Top             =   4200
      Width           =   1035
   End
   Begin VB.TextBox txtDate 
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
      Left            =   2640
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2280
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
      TabIndex        =   22
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount."
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
      TabIndex        =   18
      Top             =   2760
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
      Index           =   1
      Left            =   2040
      TabIndex        =   17
      Top             =   2760
      Width           =   225
   End
   Begin VB.Label Label2 
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
      Left            =   2040
      TabIndex        =   16
      Top             =   1200
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
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1365
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
      Left            =   2040
      TabIndex        =   12
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date."
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
      Left            =   360
      TabIndex        =   11
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
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   225
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
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Challan Entry"
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
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6375
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
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmchallanEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COMID As Integer
Dim SBNo As String
Dim SBId As String
Private Sub CboSBno_CLICK()
'cmdRefresh_Click
Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Company_Information.NAME,dbo.Challan_Entry.IdNo,dbo.Challan_Entry.Amount,dbo.Challan_Entry.ChallanDate FROM Challan_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Challan_Entry.CompID where SB_NO='" & CboSBno.Text & "' AND dbo.Challan_Entry.Deleterow <>1 order by dbo.Challan_Entry.SB_NO;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

              cbocomname.Text = R![Name]
              TxtDate.Text = R![ChallanDate]
              TxtAmount.Text = R![Amount]
              SBId = R![IdNo]
'             txtDate.AddItem R![ChallanDate]
'             txtamount.AddItem R![Amount]
             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub CboSBno_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
        
        MainComm.CommandText = "UPDATE Challan_Entry SET DeleteRow =1,Username= '" & strUser & "'," _
                              & "Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "' " _
                              & "where SB_NO='" & CboSBno.Text & "' and Year='" & txtyear.Text & "' AND DeleteRow <>1;"
        MainComm.Execute
        MsgBox "Transaction  successfully Deleted.", vbInformation, cnstMsgInfo
        cmdRefresh_Click
        CboSBnoLOAD_Click
End Sub

Private Sub cmdEdit_Click()
cmdRefresh_Click
If txtMode.Text = "NEW" Then
CboSBnoLOAD_Click
cbocomname.Enabled = False
cmdSave.Caption = "&Update"
cmdDelete.Enabled = True
cmdedit.Caption = "&New"
txtMode.Text = "Edit"
CboSBno.Visible = True
txtSBNo.Visible = False
TxtAmount.Text = ""
TxtDate.Text = ""
Else
cbocomname.Enabled = True
cmdSave.Caption = "&Save"
cmdDelete.Enabled = False
cmdedit.Caption = "&Edit"
CboSBno.Visible = False
txtSBNo.Visible = True
txtMode.Text = "NEW"
End If
End Sub

Private Sub txtamount_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub TxtDate_LostFocus()
 If Len(TxtDate) = 0 Then Exit Sub
    TxtDate.Text = FormatDate(TxtDate)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtyear.Text = Year(Date)
CboSBno.Visible = False
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

Private Sub cbocomname_Click()
Set R = New ADODB.Recordset
    strRecord = "SELECT CompID FROM Company_Information WHERE NAME='" & cbocomname.Text & "'  ORDER BY compid;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
'           txtAddress.Text = UCase(R![CompAdd])
             COMID = R![CompID]

        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub cmdSave_Click()
    If Len(cbocomname) = 0 Then
        MsgBox "Select  Company Name", vbInformation, cnstMsgInfo
        cbocomname.SetFocus
    Exit Sub
    End If

 
    
    If Len(TxtDate) = 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        TxtDate.SetFocus
    Exit Sub
    End If
    If Len(TxtAmount) = 0 Then
        MsgBox "Enter Amount", vbInformation, cnstMsgInfo
        TxtAmount.SetFocus
    Exit Sub
    End If
     If txtMode.Text = "NEW" Then
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
     Else
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
'         Exit Sub
     End If
     
'    On Error GoTo ErrorHandler
'    Screen.MousePointer = vbHourglass
If cmdSave.Caption = "&Save" Then
   If Len(txtSBNo) = 0 Then
        MsgBox " Enter  SBNo", vbInformation, cnstMsgInfo
        txtSBNo.SetFocus
    Exit Sub
    End If
    SBNo = "C- " & txtSBNo.Text
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
            
            strRecord = "SELECT Sb_No FROM Challan_Entry WHERE Sb_No='" & SBNo & "' and Year='" & txtyear.Text & "' AND DeleteRow<>1 ;"

                       R.Open strRecord, MainConn, adOpenStatic

                        If R.RecordCount > 0 Then

      MsgBox "Duplicate Sb-No", vbInformation, cnstMsgInfo
      txtSBNo.SetFocus
      Exit Sub
      End If
'
            MainConn.BeginTrans
        MainComm.CommandText = "INSERT INTO Challan_Entry(CompID,Sb_No,Amount,ChallanDate,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & COMID & ",'" & SBNo & "','" & TxtAmount.Text & "','" & Format(TxtDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
        MainComm.CommandText = "INSERT INTO Challan_Entry_Bak(CompID,Sb_No,Amount,ChallanDate,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & COMID & ",'" & SBNo & "','" & TxtAmount.Text & "','" & Format(TxtDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainConn.CommitTrans
        MsgBox "Transaction  successfully completed.", vbInformation, cnstMsgInfo
Else
If Len(CboSBno) = 0 Then
        MsgBox " Enter  SBNo", vbInformation, cnstMsgInfo
        txtSBNo.SetFocus
    Exit Sub
    End If
  Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
'         MainComm.CommandText = "UPDATE Bond_Entry SET Invoice_No='" & cboInvno.Text & "',Bond_No=" & txtGbond.Text & ",BondDate='" & Format(BondDate.Text, cnstDtFrmtI) & "'," _
'                               & " Bill_No=" & Sbill.Text & ",BillDate='" & Format(BillDate.Text, cnstDtFrmtI) & "',DeleteRow =2,Username= '" & strUser & "', Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "'" _
'                               & "where IdNo=" & invid & " and Year='" & txtyear.Text & "' AND DeleteRow <> 1;"
   
        MainComm.CommandText = "UPDATE Challan_Entry SET SB_NO='" & CboSBno.Text & "',Amount=" & TxtAmount.Text & ",ChallanDate='" & Format(TxtDate.Text, cnstDtFrmtI) & "'," _
                               & "DeleteRow =2,Username= '" & strUser & "', Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "'" _
                               & "where IdNo=" & SBId & " and Year='" & txtyear.Text & "' AND DeleteRow <> 1;"
        MainComm.Execute
        MainConn.CommitTrans
        CboSBno.Text = ""
        MsgBox "Transaction  successfully Updated.", vbInformation, cnstMsgInfo
        
   End If
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbCritical, cnstMsgErDB
'        Screen.MousePointer = vbDefault
'        MainConn.RollbackTrans
'        Exit Sub
'    End If
        cmdRefresh_Click
'        txtamount.Text = ""
End Sub

Private Sub cmdRefresh_Click()
    txtSBNo.Text = ""
'    txtDate.Text = ""
'    cbocomname.ListIndex = -1
    CboSBno.ListIndex = -1
'    txtamount.Text = ""
'    cbocomname.Text = ""
'    CboSBno.Text = ""

End Sub
Private Sub CboSBnoLOAD_Click()
CboSBno.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT Sb_No FROM Challan_Entry where  Year='" & txtyear.Text & "' AND Deleterow <>1 order by IdNo;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

             CboSBno.AddItem R![Sb_No]
             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing

End Sub

Private Sub txtSBNo_LostFocus()
'SBNo = "C- " & txtSBNo.Text
Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
            SBNo = "C- " & txtSBNo.Text
            strRecord = "SELECT Sb_No FROM Challan_Entry WHERE Sb_No='" & SBNo & "' and Year='" & txtyear.Text & "' AND DeleteRow<>1  ;"

                       R.Open strRecord, MainConn, adOpenStatic
   
                        If R.RecordCount > 0 Then
     
      MsgBox "Duplicate Sb-No", vbInformation, cnstMsgInfo
      txtSBNo.SetFocus
      Exit Sub
      End If
End Sub
