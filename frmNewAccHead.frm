VERSION 5.00
Begin VB.Form frmNewAccHead 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New or Edit Head of Account"
   ClientHeight    =   2775
   ClientLeft      =   4260
   ClientTop       =   3270
   ClientWidth     =   3525
   Icon            =   "frmNewAccHead.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEditAccName 
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
      Left            =   180
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cboAccName 
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
      Left            =   180
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   555
      TabIndex        =   3
      ToolTipText     =   "Edit/Delete Record"
      Top             =   2280
      Width           =   1335
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
      Left            =   1980
      TabIndex        =   2
      ToolTipText     =   "Save/Update Record"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtAccName 
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
      Left            =   180
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1260
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   8895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Head of Account"
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
      TabIndex        =   7
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Expenses :"
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
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   1935
   End
End
Attribute VB_Name = "frmNewAccHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccName_Click()
    If Len(cboAccName) = 0 Then Exit Sub
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT AccName FROM new_Acc_Head WHERE AccName='" & cboAccName.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then txtEditAccName.Text = R![AccName]
    
    R.Close
    Set R = Nothing
End Sub
Private Sub cboAccName_LostFocus()
    cboAccName.Text = Trim(UCase(cboAccName.Text))
    If Len(cboAccName) > 50 Then
        MsgBox "Maximum size is '50'", vbInformation, cnstMsgInfo
        cboAccName.SetFocus
    End If
End Sub
Private Sub cmdEdit_Click()
On Error GoTo ErrorHandler
    If cmdEdit.Caption = "&Delete" Then
        If Len(cboAccName) = 0 Then
            MsgBox "Select Name of Accounts", vbInformation, cnstMsgInfo
            cboAccName.SetFocus
        Exit Sub
        End If
    
        If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
      
        MainComm.CommandText = "DELETE FROM new_Acc_Head WHERE AccName='" & cboAccName.Text & "';"
        MainComm.Execute
      
        addAccName cboAccName
        cboAccName.SetFocus
    Else
        txtAccName.Visible = False
        cboAccName.Top = txtAccName.Top
        cboAccName.Visible = True
        cmdSave.Caption = "&Update"
        cmdEdit.Caption = "&Delete"
    Exit Sub
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        If Err.Number = -2147467259 Then
            MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
        Else
            MsgBox Err.Description, vbCritical, cnstMsgErDB
        End If
    End If
End Sub
Private Sub cmdSave_Click()
    Dim lngLastID As Integer

    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    If cmdSave.Caption = "&Save" Then
        If Len(txtAccName) = 0 Then
            MsgBox "Enter Name of Accounts", vbInformation, cnstMsgInfo
            txtAccName.SetFocus
        Exit Sub
        End If

        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
        lngLastID = 1
        strRecord = "SELECT Max(AccID) AS LastID FROM new_Acc_Head HAVING Max(AccID) Is Not Null;"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
        R.Close

        MainConn.BeginTrans
            MainComm.CommandText = "INSERT INTO new_Acc_Head(AccID,AccName) VALUES(" & lngLastID & ",'" & txtAccName.Text & "');"
            MainComm.Execute
        MainConn.CommitTrans
  
        txtAccName.Text = ""
        txtAccName.SetFocus
    Else
        If Len(cboAccName) = 0 Then
            MsgBox "Select Name of Accounts", vbInformation, cnstMsgInfo
            cboAccName.SetFocus
        Exit Sub
        End If

        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
        If cboAccName.Text <> txtEditAccName.Text Then
            strRecord = "SELECT AccName FROM new_Acc_Head WHERE AccName='" & cboAccName.Text & "';"
            R.Open strRecord, MainConn, adOpenStatic
            If R.RecordCount > 0 Then
                MsgBox "Duplicate Name of Accounts", vbInformation, cnstMsgInfo
                cboAccName.SetFocus
            Exit Sub
            End If
            R.Close
        End If
  
        MainConn.BeginTrans
            MainComm.CommandText = "UPDATE new_Acc_Head SET AccName='" & cboAccName.Text & "' WHERE AccName='" & txtEditAccName.Text & "';"
            MainComm.Execute
        MainConn.CommitTrans
        
        cboAccName.SetFocus
    End If
    Set R = Nothing
    addAccName cboAccName
End Sub

Private Sub Form_Load()
    Call addAccName(cboAccName)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmNewAccHead = Nothing
End Sub
Private Sub txtAccName_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtAccName_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtAccName_LostFocus()
    If Len(txtAccName) = 0 Then Exit Sub
    txtAccName.Text = Trim(UCase(txtAccName.Text))
  
    Set R = New ADODB.Recordset
    strRecord = "SELECT AccName FROM new_Acc_Head WHERE AccName='" & txtAccName.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
  
    If R.RecordCount > 0 Then
        MsgBox "Duplicate Name of Accounts", vbInformation, cnstMsgInfo
        txtAccName.SetFocus
    End If
    R.Close
    Set R = Nothing
End Sub
