VERSION 5.00
Begin VB.Form frmCreateCompany 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a new company"
   ClientHeight    =   4455
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   6180
   ControlBox      =   0   'False
   Icon            =   "frmCreateCompany.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4080
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "C&reate"
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
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Create New Company"
      Top             =   3270
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Set Company Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1155
      Left            =   600
      TabIndex        =   8
      Top             =   2940
      Width           =   3375
      Begin VB.TextBox txtPassword 
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   1905
      End
      Begin VB.TextBox txtVerrify 
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   690
         Width           =   1905
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
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
         Left            =   180
         TabIndex        =   11
         Top             =   690
         Width           =   705
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1050
         TabIndex        =   10
         Top             =   300
         Width           =   255
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   1050
         TabIndex        =   9
         Top             =   690
         Width           =   255
      End
   End
   Begin VB.TextBox txtAddress 
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
      ForeColor       =   &H00800080&
      Height          =   855
      Left            =   600
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1980
      Width           =   4815
   End
   Begin VB.TextBox txtName 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   600
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1290
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -60
      TabIndex        =   13
      Top             =   870
      Width           =   6405
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name :"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF5B4&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmCreateCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    Dim strPassword As String, bytLastID As Byte
    
    If Len(txtName) = 0 Then
        MsgBox "Enter Name of Company", vbInformation, cnstMsgInfo
        txtName.SetFocus
    Exit Sub
    End If

    If Len(txtAddress) = 0 Then
        MsgBox "Enter Company Address", vbInformation, cnstMsgInfo
        txtAddress.SetFocus
    Exit Sub
    End If

    If Len(txtPassword) = 0 Then
        If MsgBox("You don't set any Password." + vbCrLf + "Are you sure to continue !", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then
            txtPassword.SetFocus
        Exit Sub
        End If
    Else
        If txtVerrify.Text <> txtPassword.Text Then
            MsgBox "Invalid Varrify Password", vbInformation, cnstMsgInfo
            txtVerrify.SetFocus
        Exit Sub
        End If
    End If

    If Len(txtPassword) <> 0 Then strPassword = EncryptIt(txtPassword.Text, 11) Else strPassword = ""

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

        Set R = New ADODB.Recordset
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn

        bytLastID = 1
        strRecord = "SELECT Max(compID) AS LastID FROM Company_Information HAVING ((Max(compID) Is Not Null));"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then bytLastID = bytLastID + R![LastID]
        R.Close

        MainConn.BeginTrans
            MainComm.CommandText = "INSERT INTO Company_Information(compID,compName,compAdd,PWord)" _
                & " VALUES(" & bytLastID & ",'" & txtName.Text & "','" & txtAddress.Text & "','" & strPassword & "');"
            MainComm.Execute
        MainConn.CommitTrans

        bytCompID = bytLastID
        frmMDIMain.Caption = txtName.Text

'        frmMDIMain.mnuCompInformation.Visible = False
        frmMDIMain.mnuDataEntry.Visible = True
        frmMDIMain.mnuDataEdit.Visible = True
        frmMDIMain.mnuReport.Visible = True
        frmMDIMain.mnuOption.Visible = True
        frmMDIMain.mnuCloseCompany.Visible = True
        frmMDIMain.mnuNext4.Visible = True

        Unload Me

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    Set R = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateCompany = Nothing
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 Or KeyAscii = 39) Then KeyAscii = 0
End Sub
Private Sub txtAddress_LostFocus()
    txtAddress.Text = Trim(UCase(txtAddress.Text))
End Sub
Private Sub txtName_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtName_LostFocus()
    If Len(txtName) = 0 Then Exit Sub
    txtName.Text = Trim(UCase(txtName.Text))

    On Error GoTo ErrorHandler
        Set R = New ADODB.Recordset
        strRecord = "SELECT compName FROM Company_Information WHERE compName='" & txtName.Text & "';"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            MsgBox "Duplicate Name of Company", vbInformation, cnstMsgInfo
            txtName.SetFocus
        Exit Sub
        End If
        R.Close

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error"
        End
    End If
    Set R = Nothing
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtPassword_LostFocus()
    txtPassword.Text = Trim(txtPassword.Text)
End Sub
Private Sub txtVerrify_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtVerrify_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub txtVerrify_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtVerrify_LostFocus()
    If Len(txtPassword) = 0 And Len(txtVerrify) <> 0 Then
        MsgBox "Enter Password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
    Else
        txtVerrify.Text = Trim(txtVerrify.Text)
    End If
End Sub
