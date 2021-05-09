VERSION 5.00
Begin VB.Form frmCreateModifyUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   1
      Top             =   1125
      Width           =   2925
   End
   Begin VB.TextBox txtOldPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2925
   End
   Begin VB.TextBox txtVerPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2745
      Width           =   2925
   End
   Begin VB.TextBox txtNewPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2190
      Width           =   2925
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   525
      Left            =   3195
      TabIndex        =   6
      Top             =   3600
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Label lblOldPassName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password      :"
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
      Height          =   195
      Left            =   420
      TabIndex        =   14
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label lblVerifyPassName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password :"
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
      Height          =   195
      Left            =   420
      TabIndex        =   13
      Top             =   2745
      Width           =   1440
   End
   Begin VB.Label lblNewPassName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password    :"
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
      Height          =   195
      Left            =   420
      TabIndex        =   12
      Top             =   2190
      Width           =   1425
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name           :"
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
      Height          =   195
      Left            =   420
      TabIndex        =   11
      Top             =   1125
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -120
      TabIndex        =   10
      Top             =   3390
      Width           =   6015
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   3435
      Width           =   5835
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   5850
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set New Password....."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   2730
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmCreateModifyUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempUserID As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
If Len(txtOldPass.Text) = 0 Then
    MsgBox "Type Old Password.", vbInformation, cnstMsgInfo
    txtOldPass.SetFocus
    Exit Sub
End If
If Len(txtNewPass.Text) = 0 Then
        MsgBox "Type new password..", vbInformation, cnstMsgInfo
        txtNewPass.SetFocus
        Exit Sub
End If
If Len(txtVerPass.Text) = 0 Then
        MsgBox "Type Verify password..", vbInformation, cnstMsgInfo
        txtVerPass.SetFocus
        Exit Sub
End If
If txtNewPass.Text <> txtVerPass.Text Then
        MsgBox "New Password does not match with Verify Password", vbInformation, cnstMsgInfo
        txtNewPass.SetFocus
        Exit Sub
End If
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans
MainComm.CommandText = "Update Sys_User_Name Set UserPWord='" & EncryptIt(txtVerPass.Text, 11) & "' Where UserID=" & TempUserID & " "
MainComm.Execute
MsgBox "New Password Set Successfully..", vbOKOnly, cnstMsgQ
MainConn.CommitTrans
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Err.Number = 0
    Exit Sub
    End If
txtUserName.Text = ""
txtOldPass.Text = ""
txtNewPass.Text = ""
txtVerPass.Text = ""
Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtNewPass_GotFocus()
    Call txtGotFocus(txtNewPass)
End Sub

Private Sub txtNewPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtNewPass_LostFocus()
    txtNewPass.Text = Trim(txtNewPass.Text)
End Sub

Private Sub txtOldPass_GotFocus()
    Call txtGotFocus(txtOldPass)
End Sub

Private Sub txtOldPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtOldPass_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtOldPass_LostFocus()
    txtOldPass.Text = Trim(txtOldPass.Text)
    If Len(txtOldPass.Text) > 0 Then
    Call prcCheckPass(txtOldPass, txtUserName)
    End If
End Sub

Private Sub txtUserName_GotFocus()
Call txtGotFocus(txtUserName)
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtUserName_LostFocus()
     txtUserName.Text = Trim(txtUserName.Text)
End Sub

Private Sub txtVerPass_GotFocus()
    Call txtGotFocus(txtVerPass)
End Sub

Private Sub txtVerPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtVerPass_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub
Private Sub Form_Load()
    Call prcMakeCenter(Me)
'    Call addUserName(cboAddUser)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateModifyUser = Nothing
    Unload Me
    frmUserLogin.Refresh
End Sub
Private Sub txtVerPass_LostFocus()
    txtVerPass.Text = Trim(txtVerPass)
End Sub
Public Sub prcCheckPass(txtpass As TextBox, txtUser As TextBox)
Set R = New ADODB.Recordset
    strRecord = "Select UserID,UserName,UserPWord from Sys_User_Name Where UserName='" & txtUser.Text & "' And UserPWord='" & EncryptIt(txtpass.Text, 11) & "'"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount = 0 Then
        MsgBox "You have given wrong password..Try Again..", vbCritical, "Error"
        txtOldPass.SetFocus
        Exit Sub
     Else
     TempUserID = R![Userid]
    End If
    R.Close
    Set R = Nothing
End Sub


