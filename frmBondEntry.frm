VERSION 5.00
Begin VB.Form frmBondEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   LinkTopic       =   "Bond Entry"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboInvno 
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
      ItemData        =   "frmBondEntry.frx":0000
      Left            =   2400
      List            =   "frmBondEntry.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox BondDate 
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
      Left            =   2400
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox BillDate 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtinvoice 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Sbill 
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
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
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
      Left            =   4440
      TabIndex        =   15
      Top             =   4440
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
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Save Record"
      Top             =   4440
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
      ItemData        =   "frmBondEntry.frx":0004
      Left            =   2400
      List            =   "frmBondEntry.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtGbond 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
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
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Save Record"
      Top             =   4440
      Width           =   1035
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
      Left            =   840
      TabIndex        =   12
      ToolTipText     =   "Save Record"
      Top             =   4440
      Width           =   1035
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
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "NEW"
      Top             =   120
      Width           =   975
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
      TabIndex        =   7
      Top             =   2520
      Width           =   975
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
      TabIndex        =   22
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label Label1 
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
      Index           =   4
      Left            =   360
      TabIndex        =   24
      Top             =   1680
      Width           =   1125
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
      Left            =   2040
      TabIndex        =   21
      Top             =   2640
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G.Bond No."
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
      Index           =   3
      Left            =   360
      TabIndex        =   27
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Bond Information"
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
      TabIndex        =   20
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   6495
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
      TabIndex        =   17
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S Bill No."
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
      TabIndex        =   25
      Top             =   2160
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
      TabIndex        =   16
      Top             =   2160
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bond Date."
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
      TabIndex        =   28
      Top             =   3600
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
      Left            =   2040
      TabIndex        =   14
      Top             =   3120
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
      Left            =   360
      TabIndex        =   23
      Top             =   1080
      Width           =   1365
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
      TabIndex        =   13
      Top             =   1080
      Width           =   225
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
      TabIndex        =   11
      Top             =   3600
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " S Bill Date"
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
      TabIndex        =   26
      Top             =   2640
      Width           =   1125
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
      TabIndex        =   10
      Top             =   2160
      Width           =   645
   End
End
Attribute VB_Name = "frmBondEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COMID As Integer
Dim INVNo As String
Dim invid As Integer

Private Sub Form_Load()
txtyear.Text = Year(Date)
cboInvno.Visible = False
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
    txtGbond.Text = ""
    BondDate.Text = ""
    Sbill.Text = ""
    BillDate.Text = ""
    txtinvoice.Text = ""
    cboInvno.Clear
Set R = New ADODB.Recordset
    strRecord = "SELECT CompID FROM Company_Information WHERE NAME='" & cbocomname.Text & "'  ORDER BY compid;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

             COMID = R![CompID]

        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
  If txtMode.Text = "Edit" Then
  cboInvno.Clear

    Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Bond_Entry.IdNo,dbo.Bond_Entry.Invoice_No FROM Bond_Entry INNER JOIN dbo.Company_Information ON dbo.Company_Information.compID = dbo.Bond_Entry.CompID where Company_Information.Name='" & cbocomname.Text & "' and Year='" & txtyear & "' and Deleterow <>1 ;"
    R.Open strRecord, MainConn, adOpenStatic
    
    
        If R.RecordCount > 0 Then
           Do Until R.EOF

             cboInvno.AddItem R![Invoice_No]
   
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
    
   End If
End Sub

Private Sub cboInvno_Click()
Set R = New ADODB.Recordset
    strRecord = "SELECT dbo.Bond_Entry.IdNo,dbo.Bond_Entry.Bond_No,dbo.Bond_Entry.BondDate,dbo.Bond_Entry.Bill_No,dbo.Bond_Entry.BillDate FROM Bond_Entry where  Invoice_No='" & cboInvno.Text & "' AND dbo.Bond_Entry.Deleterow <>1 order by dbo.Bond_Entry.Invoice_No;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

              txtGbond.Text = R![Bond_No]
              BondDate.Text = R![BondDate]
              Sbill.Text = R![Bill_No]
              BillDate.Text = R![BillDate]
              invid = R![IdNo]

             
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
     Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset

            MainConn.BeginTrans
        
        MainComm.CommandText = "UPDATE Bond_Entry SET DeleteRow =1,Username= '" & strUser & "'," _
                              & "Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "' " _
                              & "where Invoice_No='" & cboInvno.Text & "' and Year='" & txtyear.Text & "' AND DeleteRow <>1;"
        MainComm.Execute
        MainConn.CommitTrans
        MsgBox "Transaction  successfully Deleted.", vbInformation, cnstMsgInfo
   txtGbond.Text = ""
    BondDate.Text = ""
    Sbill.Text = ""
    BillDate.Text = ""
    txtinvoice.Text = ""
    cbocomname_Click
End Sub

Private Sub cmdEdit_Click()
cmdRefresh_Click
If txtMode.Text = "NEW" Then

cbocomname.Enabled = True
cmdSave.Caption = "&Update"
cmdDelete.Enabled = True
cmdedit.Caption = "&New"
txtMode.Text = "Edit"
cboInvno.Visible = True
txtinvoice.Visible = False

Else
cbocomname.Enabled = True
cmdSave.Caption = "&Save"
cmdDelete.Enabled = False
cmdedit.Caption = "&Edit"
cboInvno.Visible = False
txtinvoice.Visible = True
txtMode.Text = "NEW"
End If

End Sub

Private Sub cmdSave_Click()

If Len(cbocomname) = 0 Then
        MsgBox "Select  Company Name", vbInformation, cnstMsgInfo
        cbocomname.SetFocus
    Exit Sub
    End If


     If Len(txtGbond) = 0 Then
        MsgBox "Enter Bond No", vbInformation, cnstMsgInfo
        txtGbond.SetFocus
    Exit Sub
    End If
    
    If Len(BondDate) = 0 Then
        MsgBox "Enter Date of Bond", vbInformation, cnstMsgInfo
        BondDate.SetFocus
    Exit Sub
    End If
    
    If Len(Sbill) = 0 Then
        MsgBox "Enter Bill no", vbInformation, cnstMsgInfo
        Sbill.SetFocus
    Exit Sub
    End If
    
    If Len(BillDate) = 0 Then
        MsgBox "Enter Date of Bill", vbInformation, cnstMsgInfo
        BillDate.SetFocus
    Exit Sub
    End If
    
     If txtMode.Text = "NEW" Then
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
     Else
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        

     End If

If cmdSave.Caption = "&Save" Then
   If Len(txtinvoice) = 0 Then
        MsgBox " Enter  Invoice No", vbInformation, cnstMsgInfo
        txtinvoice.SetFocus
    Exit Sub
    End If
If cbocomname.Text = "PACIFIC JEANS LIMITED" Then
INVNo = "PJL" & txtinvoice.Text
ElseIf cbocomname.Text = "UNIVERSAL JEANS LIMITED" Then
INVNo = "UJL" & txtinvoice.Text
ElseIf cbocomname.Text = "JEANS 2000 LTD." Then
INVNo = "J2L" & txtinvoice.Text
ElseIf cbocomname.Text = "PACIFIC ACCESSORIES LIMITED" Then
INVNo = "PAL" & txtinvoice.Text
ElseIf cbocomname.Text = "H.U.ACCESSORIES LTD." Then
INVNo = "HUAL" & txtinvoice.Text
ElseIf cbocomname.Text = "NHT FASHIONS LTD." Then
INVNo = "NHTF" & txtinvoice.Text
End If

Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
'            SBNo = "C- " & txtSBNo.Text
            strRecord = "SELECT  Invoice_No FROM Bond_Entry WHERE Invoice_No='" & INVNo & "' AND deleterow=0;"

                       R.Open strRecord, MainConn, adOpenStatic
   
       If R.RecordCount > 0 Then
     
      MsgBox "Duplicate Invoice_No", vbInformation, cnstMsgInfo
      txtinvoice.SetFocus
      Exit Sub
       End If

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset

            MainConn.BeginTrans
        MainComm.CommandText = "INSERT INTO Bond_Entry(CompID,Invoice_No,Bond_No,BondDate,Bill_No,BillDate,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & COMID & ",'" & INVNo & "','" & txtGbond.Text & "','" & Format(BondDate.Text, cnstDtFrmtI) & "','" & Sbill.Text & "','" & Format(BillDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
        MainComm.CommandText = "INSERT INTO Bond_Entry_bak(CompID,Invoice_No,Bond_No,BondDate,Bill_No,BillDate,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & COMID & ",'" & INVNo & "','" & txtGbond.Text & "','" & Format(BondDate.Text, cnstDtFrmtI) & "','" & Sbill.Text & "','" & Format(BillDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainConn.CommitTrans
        MsgBox "Transaction  successfully completed.", vbInformation, cnstMsgInfo
    
    
Else

If Len(cboInvno) = 0 Then
        MsgBox " Enter  Invoice", vbInformation, cnstMsgInfo
        cboInvno.SetFocus
    Exit Sub
End If

'Set MainComm = New ADODB.Command
'    MainComm.ActiveConnection = MainConn
'    Set R = New ADODB.Recordset
''            SBNo = "C- " & txtSBNo.Text
'            strRecord = "SELECT  Invoice_No FROM Bond_Entry WHERE Invoice_No='" & cboInvno.Text & "' AND Bill_No='" & Sbill.Text & "' AND Bond_No='" & txtGbond.Text & "' and Year='" & txtyear.Text & "' ;"
'
'             R.Open strRecord, MainConn, adOpenStatic
'
'       If R.RecordCount > 0 Then
'
'      MsgBox "Duplicate Invoice_No", vbInformation, cnstMsgInfo
''      txtinvoice.SetFocus
'      Exit Sub
'       End If
       
  Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
            
        MainComm.CommandText = "UPDATE Bond_Entry SET Invoice_No='" & cboInvno.Text & "',Bond_No=" & txtGbond.Text & ",BondDate='" & Format(BondDate.Text, cnstDtFrmtI) & "'," _
                               & " Bill_No=" & Sbill.Text & ",BillDate='" & Format(BillDate.Text, cnstDtFrmtI) & "',DeleteRow =2,Username= '" & strUser & "', Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "'" _
                               & "where IdNo=" & invid & " and Year='" & txtyear.Text & "' AND DeleteRow <> 1;"
        MainComm.Execute
        cboInvno.Text = ""
        BondDate.Text = ""
        BillDate.Text = ""
        MainConn.CommitTrans
        MsgBox "Transaction  successfully Updated.", vbInformation, cnstMsgInfo
'     txtGbond.Text = ""
'     BondDate.Text = ""
'     Sbill.Text = ""
'     BillDate.Text = ""
'     txtinvoice.Text = ""
        
   End If

    txtGbond.Text = ""
    Sbill.Text = ""
    txtinvoice.Text = ""

End Sub

Private Sub cmdRefresh_Click()
    txtGbond.Text = ""
    BondDate.Text = ""
    Sbill.Text = ""
    BillDate.Text = ""
    txtinvoice.Text = ""
    cbocomname.Text = ""
    cboInvno.Text = ""
    cboInvno.Clear

End Sub
Private Sub cbocomname_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
Private Sub cboInvno_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
Private Sub txtinvoice_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
Private Sub Sbill_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
Private Sub BillDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
Private Sub BondDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtGbond_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub


Private Sub BillDate_LostFocus()
If Len(BillDate) = 0 Then Exit Sub
    BillDate.Text = FormatDate(BillDate)
End Sub

Private Sub BondDate_LostFocus()
If Len(BondDate) = 0 Then Exit Sub
    BondDate.Text = FormatDate(BondDate)
End Sub

'Private Sub txtinvoice_LostFocus()
'Set MainComm = New ADODB.Command
'    MainComm.ActiveConnection = MainConn
'    Set R = New ADODB.Recordset
'            SBNo = "C- " & txtSBNo.Text
'            strRecord = "SELECT Sb_No FROM Challan_Entry WHERE Sb_No='" & SBNo & "' and Year='" & txtyear.Text & "' ;"
'
'                       R.Open strRecord, MainConn, adOpenStatic
'
'       If R.RecordCount > 0 Then
'
'      MsgBox "Duplicate Sb-No", vbInformation, cnstMsgInfo
'      txtSBNo.SetFocus
'      Exit Sub
'      End If
'End Sub
