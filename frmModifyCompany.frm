VERSION 5.00
Begin VB.Form frmModifyCompany 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Company"
   ClientHeight    =   3810
   ClientLeft      =   2685
   ClientTop       =   3255
   ClientWidth     =   5160
   Icon            =   "frmModifyCompany.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "Clo&se"
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
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   3
      Top             =   3330
      Width           =   1425
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   3480
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   3330
      Width           =   1545
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
      Height          =   1095
      Left            =   420
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2010
      Width           =   4335
   End
   Begin VB.TextBox txtName 
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
      Left            =   420
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1290
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -60
      TabIndex        =   7
      Top             =   870
      Width           =   6075
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   6075
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
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
      Left            =   420
      TabIndex        =   5
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
      Left            =   420
      TabIndex        =   4
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF5B4&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "frmModifyCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrorHandler
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
    
    If MsgBox("Are You sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
        MainComm.CommandText = "UPDATE Company_Information SET compName='" & txtName.Text & "',compAdd='" & txtAddress.Text & "' WHERE compID=" & findCompID & ";"
        MainComm.Execute
    MainConn.CommitTrans
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        MainConn.RollbackTrans
    Exit Sub
    End If
    
    MsgBox "Successfully Updated", vbInformation, cnstMsgInfo

    frmMDIMain.Caption = txtName.Text
    Unload Me
End Sub
Private Sub Form_Load()
    Set R = New ADODB.Recordset
    
    strRecord = "SELECT Name,Address FROM Company_Information WHERE COMPID=" & findCompID & ";"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        txtName.Text = R![Name]
        txtAddress.Text = R![Address]
    End If
    
    R.Close
    Set R = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmModifyCompany = Nothing
End Sub
Private Sub txtAddress_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 Or KeyAscii = 39) Then KeyAscii = 0
End Sub
Private Sub txtAddress_LostFocus()
    txtAddress.Text = Trim(txtAddress.Text)
End Sub
Private Sub txtName_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtName_LostFocus()
    txtName.Text = Trim(txtName.Text)
End Sub
