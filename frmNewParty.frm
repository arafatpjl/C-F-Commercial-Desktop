VERSION 5.00
Begin VB.Form frmNewParty 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New Party Name"
   ClientHeight    =   2760
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   3825
   Icon            =   "frmNewParty.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   3585
      Begin VB.TextBox txtDist 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   150
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Name Of Party  :"
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
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancel"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
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
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Create New Company"
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   8895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Party Name"
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
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmNewParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()
On Error GoTo x
    Dim NewID As Integer
    
    If Len(txtDist) = 0 Then
        MsgBox "Enter name of the party.", vbInformation, AppTitle
        txtDist.SetFocus
    Exit Sub
    End If

        If fncDuplicate = True Then Exit Sub
        MainConn.BeginTrans
            If flagMark = "INSERT" Then
                NewID = fncNewID("PARTY")
                MainComm.CommandText = "INSERT INTO New_Party( PartyID, PartyName) VALUES(" & NewID & ", '" & txtDist.Text & "');"
                MainComm.Execute
            Else
                MainComm.CommandText = "UPDATE New_Party SET PartyName = '" & txtDist.Text & "' WHERE PartyID=" & Tag & ";"
                MainComm.Execute
            End If
        MainConn.CommitTrans
    strValue = txtDist.Text
    Unload Me
Exit Sub
x:
    MsgBox Err.Description, vbCritical, cnstMsgErDB
    MainConn.RollbackTrans
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNewParty = Nothing
End Sub

Private Sub txtDist_GotFocus()
    Call txtGotFocus(txtDist)
End Sub

Private Sub txtDist_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDist_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub

Private Sub txtDist_LostFocus()
    txtDist.Text = UCase(Trim(txtDist.Text))
End Sub

Private Function fncDuplicate() As Boolean
fncDuplicate = True
    Set R = New ADODB.Recordset
    If flagMark = "INSERT" Or (Trim(txtDist.Text) = Trim(txtDist.Tag)) Then
        strRecord = "SELECT Count(*) FROM new_Party WHERE (PartyName='" & txtDist.Text & "');"
        MainComm.CommandText = strRecord
        Set R = MainComm.Execute
        If R.Fields(0) > 0 Then
            MsgBox "Name of the party '" & txtDist.Text & "' already exist." & vbCrLf & "Change the name of the party.", vbInformation, AppTitle
            txtDist.SetFocus
        Exit Function
        End If
    End If
    Set R = Nothing
fncDuplicate = False
End Function
