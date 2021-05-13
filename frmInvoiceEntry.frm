VERSION 5.00
Begin VB.Form frmInvoiceEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   LinkTopic       =   "Bond Entry"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboInvoice 
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
      ItemData        =   "frmInvoiceEntry.frx":0000
      Left            =   2400
      List            =   "frmInvoiceEntry.frx":0002
      TabIndex        =   23
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox TxtCargoReceipt 
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
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TxtDate 
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
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox TxtAmount 
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
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox TxtInv 
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
      TabIndex        =   0
      Top             =   1440
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
      TabIndex        =   10
      Top             =   3840
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
      TabIndex        =   4
      ToolTipText     =   "Save Record"
      Top             =   3840
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
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Save Record"
      Top             =   3840
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
      TabIndex        =   9
      ToolTipText     =   "Save Record"
      Top             =   3840
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
      TabIndex        =   22
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
      Left            =   5160
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1800
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
      TabIndex        =   17
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo Receipt No."
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
      TabIndex        =   18
      Top             =   1920
      Width           =   1605
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
      TabIndex        =   16
      Top             =   2400
      Width           =   225
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Invoice Information"
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
      TabIndex        =   15
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   3600
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
      TabIndex        =   12
      Top             =   120
      Width           =   2175
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
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   1440
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
      TabIndex        =   11
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Date."
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
      TabIndex        =   21
      Top             =   2880
      Width           =   1605
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
      TabIndex        =   8
      Top             =   2880
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      TabIndex        =   20
      Top             =   2400
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
      Left            =   5400
      TabIndex        =   7
      Top             =   1440
      Width           =   645
   End
End
Attribute VB_Name = "frmInvoiceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INVNo As String
Private Sub CboInvoice_Click()
Set R = New ADODB.Recordset
    strRecord = "SELECT Receipt_No,Invoice_No,Amount,Date FROM Invoice_Entry where CompID=" & bytCompID & " AND Invoice_No ='" & cboInvoice.Text & "'AND dbo.Invoice_Entry.Deleterow <>1 order by Receipt_No;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF

'              cbocomname.Text = R![Receipt_No]
              TxtCargoReceipt.Text = R![Receipt_No]
              TxtAmount.Text = R![Amount]
              TxtDate = R![Date]
              INVNo = R![Invoice_No]

        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub Form_Load()
txtyear.Text = Year(Date)
cmdDelete.Enabled = False
'CboCargoReceiptNo.Visible = False
cboInvoice.Visible = False


End Sub

Private Sub cbocomname_Click()
    TxtInv.Text = ""
    TxtAmount.Text = ""
    TxtDate.Text = ""

'    CboCargoReceiptNo.Clear
    
  If txtMode.Text = "Edit" Then
  cboInvoice.Clear

    Set R = New ADODB.Recordset
    strRecord = "SELECT Invoice_No FROM Invoice_Entry where CompID=" & bytCompID & " AND dbo.Invoice_Entry.Deleterow <>1 order by Receipt_No;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
              
         cboInvoice.AddItem R![Invoice_No]
             
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
    
    
        
        MainComm.CommandText = "UPDATE Invoice_Entry SET DeleteRow =1,Username= '" & strUser & "'," _
                              & "Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "' " _
                              & "where Invoice_No='" & cboInvoice.Text & "' and Year='" & txtyear.Text & "' AND DeleteRow <>1;"
        MainComm.Execute
        MsgBox "Transaction  successfully Deleted.", vbInformation, cnstMsgInfo
   TxtCargoReceipt.Text = ""
    TxtInv.Text = ""
    TxtAmount.Text = ""
    TxtDate.Text = ""
'    txtinvoice.Text = ""
    cbocomname_Click
End Sub

Private Sub cmdEdit_Click()
cmdRefresh_Click
If txtMode.Text = "NEW" Then

'cbocomname.Enabled = True
cmdSave.Caption = "&Update"
cmdDelete.Enabled = True
cmdedit.Caption = "&New"
txtMode.Text = "Edit"
'CboCargoReceiptNo.Visible = True
'TxtCargoReceipt.Visible = False
cboInvoice.Visible = True
cbocomname_Click

Else
'cbocomname.Enabled = True
cmdSave.Caption = "&Save"
cmdDelete.Enabled = False
cmdedit.Caption = "&Edit"
'CboCargoReceiptNo.Visible = False
cboInvoice.Visible = False
TxtCargoReceipt.Visible = True
txtMode.Text = "NEW"
End If

End Sub

Private Sub cmdSave_Click()

    If Len(TxtCargoReceipt) = 0 Then
         MsgBox "Enter Cargo Receipt No", vbInformation, cnstMsgInfo
        TxtCargoReceipt.SetFocus
    Exit Sub
    End If
    
    If Len(TxtAmount) = 0 Then
        MsgBox "Enter Amount", vbInformation, cnstMsgInfo
        TxtAmount.SetFocus
    Exit Sub
    End If
    
    If Len(TxtDate) = 0 Then
        MsgBox "Enter Date", vbInformation, cnstMsgInfo
        TxtDate.SetFocus
    Exit Sub
    End If
    
   
     If txtMode.Text = "NEW" Then
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
     Else
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        

     End If

If cmdSave.Caption = "&Save" Then




     If Len(TxtInv) = 0 Then
        MsgBox "Enter Invoice No", vbInformation, cnstMsgInfo
        TxtInv.SetFocus
    Exit Sub
    End If
    
    
   

Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
'            SBNo = "C- " & txtSBNo.Text
            strRecord = "SELECT  Invoice_No FROM Invoice_Entry WHERE  Invoice_No= '" & TxtInv & "' and Year='" & txtyear.Text & "' ;"

                       R.Open strRecord, MainConn, adOpenStatic
   
       If R.RecordCount > 0 Then
     
      MsgBox "Duplicate Receipt No", vbInformation, cnstMsgInfo
      TxtCargoReceipt.SetFocus
      Exit Sub
       End If

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset

            MainConn.BeginTrans
        MainComm.CommandText = "INSERT INTO Invoice_Entry(CompID,Receipt_No,Invoice_No,Amount,Date,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & bytCompID & ",'" & TxtCargoReceipt & "','" & TxtInv.Text & "','" & TxtAmount.Text & "','" & Format(TxtDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
        MainComm.CommandText = "INSERT INTO Invoice_Entry_Bak(CompID,Receipt_No,Invoice_No,Amount,Date,Year,Deleterow,Comname,Username,Userpwd,Entrydate,Entrytime)" _
              & " VALUES(" & bytCompID & ",'" & TxtCargoReceipt & "','" & TxtInv.Text & "','" & TxtAmount.Text & "','" & Format(TxtDate.Text, cnstDtFrmtI) & "'," & txtyear.Text & "," _
              & " 0,'" & getComName & "','" & strUser & "','" & StrPass & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainConn.CommitTrans
        MsgBox "Transaction  successfully completed.", vbInformation, cnstMsgInfo
    
    
Else

If Len(cboInvoice) = 0 Then
        MsgBox " Enter  Invoice", vbInformation, cnstMsgInfo
        cboInvoice.SetFocus
    Exit Sub
End If


       
  Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn
            MainConn.BeginTrans
            
        MainComm.CommandText = "UPDATE Invoice_Entry SET Receipt_No='" & TxtCargoReceipt.Text & "',Invoice_No='" & cboInvoice.Text & "',Date='" & Format(TxtDate.Text, cnstDtFrmtI) & "'," _
                               & " Amount=" & TxtAmount.Text & ",DeleteRow =2,Username= '" & strUser & "', Comname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "'" _
                               & "where Invoice_No='" & INVNo & "' and Year='" & txtyear.Text & "' AND DeleteRow <> 1;"
        MainComm.Execute
        TxtCargoReceipt.Text = ""
        TxtAmount.Text = ""
        TxtDate.Text = ""
        cbocomname_Click
        MainConn.CommitTrans
        MsgBox "Transaction  successfully Updated.", vbInformation, cnstMsgInfo

        
   End If

    TxtCargoReceipt.Text = ""
    TxtInv.Text = ""
    TxtAmount.Text = ""
    TxtDate.Text = ""

End Sub

Private Sub cmdRefresh_Click()
    TxtCargoReceipt.Text = ""
    TxtInv.Text = ""
    TxtAmount.Text = ""
    TxtDate.Text = ""
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



Private Sub txtamount_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub

Private Sub TxtCargoReceipt_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub
Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub

Private Sub TxtDate_LostFocus()
 If Len(TxtDate) = 0 Then Exit Sub
    TxtDate.Text = FormatDate(TxtDate)


End Sub

Private Sub TxtInv_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub
