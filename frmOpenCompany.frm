VERSION 5.00
Begin VB.Form frmOpenCompany 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Company"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmOpenCompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
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
         Left            =   1800
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "100"
         Top             =   240
         Width           =   1935
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
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   255
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
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
      Left            =   1500
      TabIndex        =   3
      Top             =   3300
      Width           =   1575
   End
   Begin VB.ComboBox cboCompany 
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0C000&
      Caption         =   "&Open Company"
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
      TabIndex        =   2
      Top             =   3300
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   3210
      Width           =   6405
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   0
      Left            =   -180
      TabIndex        =   8
      Top             =   870
      Width           =   6405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Company"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF5B4&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   -120
      TabIndex        =   9
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmOpenCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCompany_Click()
    If Len(cboCompany) = 0 Then Exit Sub

    Set R = New ADODB.Recordset
    
    strRecord = "SELECT PWord FROM Company_Information WHERE compName='" & cboCompany.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 And R![PWord] <> vbNullString Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
    R.Close
    
    Set R = Nothing
End Sub
Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Frame1.Visible = True Then txtPassword.SetFocus Else cmdOpen.SetFocus
    End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOpen_Click()
    Dim strPWord As String

    If Len(cboCompany) = 0 Then
        MsgBox "Select name of company", vbInformation, cnstMsgInfo
        cboCompany.SetFocus
    Exit Sub
    End If
    
    If Frame1.Visible = True And Len(txtPassword) = 0 Then
        MsgBox "Enter password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
    Exit Sub
    End If

    strPWord = EncryptIt(txtPassword.Text, 11)

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set R = New ADODB.Recordset
    If Len(txtPassword) <> 0 Then
        strRecord = "SELECT * FROM Company_Information WHERE compName='" & cboCompany.Text & "'AND PWord='" & strPWord & "';"
    Else
        strRecord = "SELECT * FROM Company_Information WHERE compName='" & cboCompany.Text & "';"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount = 0 Then
        MsgBox "Invalid password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
        Screen.MousePointer = vbDefault
    Exit Sub
    End If

    bytCompID = R![CompID]
    frmMDIMain.Caption = frmMDIMain.Caption & " - " & R![compName]
    
    
'    frmMDIMain.mnuCompInformation.Visible = False
    frmMDIMain.mnuHouseKeeping.Visible = True
    frmMDIMain.mnuDataEntry.Visible = True
    frmMDIMain.mnuDataEdit.Visible = True
    frmMDIMain.mnuReport.Visible = True
    frmMDIMain.mnuCloseCompany.Visible = True
    frmMDIMain.mnuNext4.Visible = True

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
    Exit Sub
    End If

    Screen.MousePointer = vbDefault
    Set R = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Set R = New ADODB.Recordset
    
    strRecord = "SELECT compName FROM Company_Information;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboCompany.AddItem R![compName]
        R.MoveNext
        Loop
    End If
    R.Close
    
    Set R = Nothing
    cboCompany.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmOpenCompany = Nothing
End Sub
Private Sub txtPassword_GotFocus()
    Call txtGotFocus(txtPassword)
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen.SetFocus
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub
Private Sub txtPassword_LostFocus()
    txtPassword.Text = Trim(txtPassword.Text)
End Sub
