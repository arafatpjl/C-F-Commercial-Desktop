VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log In"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Change Password"
      Height          =   285
      Left            =   1290
      TabIndex        =   7
      Top             =   2700
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
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
      Left            =   3780
      TabIndex        =   2
      Top             =   3270
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   2430
      TabIndex        =   3
      Top             =   3270
      Width           =   1245
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2250
      Width           =   3075
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
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
      Height          =   345
      Left            =   1020
      TabIndex        =   0
      Text            =   "Admin"
      Top             =   1470
      Width           =   3075
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   870
      Width           =   6075
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF5B4&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   -120
      TabIndex        =   9
      Top             =   0
      Width           =   6075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -990
      TabIndex        =   6
      Top             =   3210
      Width           =   6225
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   5
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   4
      Top             =   1170
      Width           =   1065
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim valchk As Byte

Private Sub Check1_Click()
On Error Resume Next
    If Check1.Value = 1 Then
        If Trim(txtUserName.Text) = "" Then
            MsgBox "You should provide user name.", vbInformation, App.Title
            txtUserName.SetFocus
            Check1.Value = 0
            Exit Sub
        End If
        
        Mysql = "Select * from Login Where UserName='" & Trim(txtUserName.Text) & "' "
        Set tmprst = MainConn.Execute(Mysql)
        
        If tmprst.EOF = True And tmprst.BOF = True Then
            MsgBox "Invalid user name.", vbCritical, App.Title
            txtUserName.Text = ""
            txtUserName.SetFocus
            Check1.Value = 0
            Exit Sub
        End If
        strUser = txtUserName.Text
        frmChangePassword.Show 1
        
    End If
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdLogin_Click()
On Error GoTo x
    Dim tmprst As New Recordset
    
    If Trim(txtUserName.Text) = "" Then
        MsgBox "You should provide User Name.", vbInformation, App.Title
        txtUserName.SetFocus
        Exit Sub
    End If
    
    Mysql = "Select * from Login Where UserName='" & Trim(txtUserName.Text) & "' and Password='" & Trim(txtPassword.Text) & "'"
    Set tmprst = MainConn.Execute(Mysql)
    
    If tmprst.EOF = True And tmprst.BOF = True Then
        MsgBox "Invalid user name or password.", vbCritical, App.Title
        txtPassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    End If
    
    frmMain.mnuLogin.Visible = False
    frmMain.mnuCompanyInfo.Visible = True
    Unload Me
    
    Set tmprst = Nothing
Exit Sub
x:
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
'    SetWindowPos hwnd, HWND_TOPMOST, Left / 15, Top / 15, Width / 15, Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

    valchk = 1
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUserName_GotFocus()
On Error Resume Next
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
    If valchk = 1 Then Call txtUserName_KeyDown(13, 0)
    valchk = 0
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
