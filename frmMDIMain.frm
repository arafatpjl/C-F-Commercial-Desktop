VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808000&
   Caption         =   "C & F Indent and Adjustment System"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "PictureClip1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Banner1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.TextBox txtUserPassword 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   180
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblcomname 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   615
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   11895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   15255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C && F                            Commercial"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3495
         Left            =   2880
         TabIndex        =   3
         Top             =   3240
         Width           =   11535
      End
   End
   Begin VB.Menu mnuHouseKeeping 
      Caption         =   "     &House Keeping"
      Begin VB.Menu mnuNewAccountsHead 
         Caption         =   "New &Accounts Head"
      End
      Begin VB.Menu mnuNewExpenses 
         Caption         =   "New &Expenses"
      End
      Begin VB.Menu mnuNewParty 
         Caption         =   "New &Party"
      End
      Begin VB.Menu mnuNewBuyer 
         Caption         =   "New &Buyer"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "     Data E&ntry"
      Begin VB.Menu mnuNewIndentInfo 
         Caption         =   "New Indent Information"
      End
      Begin VB.Menu mnunext02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewIndentFund 
         Caption         =   "New Indent Fund"
      End
      Begin VB.Menu mnunewBond 
         Caption         =   "New Bond Entry"
      End
      Begin VB.Menu mnunewInvoice 
         Caption         =   "New Invoice Entry"
      End
      Begin VB.Menu mnunewchallan 
         Caption         =   "New Challan Entry"
      End
      Begin VB.Menu mnuCandFBillOut 
         Caption         =   "C&&F Bill [OUT]"
      End
   End
   Begin VB.Menu mnuDataEdit 
      Caption         =   "   Data E&dit"
      Begin VB.Menu mnuEditIndentInfo 
         Caption         =   "Edit Indent Information"
      End
      Begin VB.Menu mnunext03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditIndentFund 
         Caption         =   "Edit Indent Fund"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "   Data &Report   "
      Begin VB.Menu mnuRptIndentSt 
         Caption         =   "Indent Statement"
      End
      Begin VB.Menu mnuRptAdjustSt 
         Caption         =   "Adjustment Statement"
      End
      Begin VB.Menu mnuNext2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptClearBill 
         Caption         =   "Clearing Bill Register"
      End
      Begin VB.Menu mnuRptForwardBill 
         Caption         =   "Forwarding Bill Register"
      End
      Begin VB.Menu mnustatementofcandfbills 
         Caption         =   "Statement of C and F Bills"
      End
      Begin VB.Menu mnustatementofforwardingdate 
         Caption         =   "Statement of C and F Bills Forwarding Date"
      End
      Begin VB.Menu Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInvoiceEntry 
         Caption         =   "Invoice Entry Report"
      End
      Begin VB.Menu mnurptBondEntry 
         Caption         =   "Bond Entry Report"
      End
      Begin VB.Menu mnurptchallanEntry 
         Caption         =   "Challan Entry Report"
      End
      Begin VB.Menu mnuChallanpaid 
         Caption         =   "Challan Paid Report"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "   Op&tion   "
      Begin VB.Menu mnuModifyComp 
         Caption         =   "Modify Company"
      End
      Begin VB.Menu mnuDeleteComp 
         Caption         =   "Delete Company"
      End
      Begin VB.Menu mnuNext3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateUser 
         Caption         =   "Create or Modify User"
      End
      Begin VB.Menu mnupermission 
         Caption         =   "Permission"
         Begin VB.Menu mnumenupermission 
            Caption         =   "Menu Permission"
         End
         Begin VB.Menu Mnunext04 
            Caption         =   "-"
         End
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "   E&xit   "
      Begin VB.Menu mnuCloseCompany 
         Caption         =   "Close Company"
      End
      Begin VB.Menu mnuNext4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Finally Exit"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Resize()
    Banner1.Height = Me.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set frmMDIMain = Nothing
End Sub
Private Sub mnuAbout_Click()
    ''frmAbout.Show 1
End Sub

Private Sub mnuCandFBillOut_Click()
    frmCandFOut.Show 1, Me
End Sub

Private Sub mnuChallanpaid_Click()
frmRptChallanPaid.Show 1
End Sub

Private Sub mnuCloseCompany_Click()
'  frmSplash.Show
End Sub
Private Sub mnuCreateUser_Click()
'  frmCreateUser.Show 1
End Sub
Private Sub mnuDeleteComp_Click()
'  frmDeleteCompany.Show 1
End Sub

Private Sub mnuEditIndentFund_Click()
    frmEditIndentFund.Show 1, Me
End Sub

Private Sub mnuEditIndentInfo_Click()
  frmEditIndentEntry.Show 1
End Sub

Private Sub MnuInvoiceEntry_Click()
frmrptInvoiceEntry.Show 1
End Sub

Private Sub mnumenupermission_Click()
frmPermissionMenu.Show 1
End Sub

Private Sub mnuModifyComp_Click()
  frmModifyCompany.Show 1
End Sub

Private Sub mnuNewAccountsHead_Click()
    frmAccountsHead.Show 1, Me
End Sub

Private Sub mnunewBond_Click()
frmBondEntry.Show 1
End Sub

Private Sub mnuNewBuyer_Click()
    frmBuyer.Show 1, Me
End Sub

Private Sub mnuNewComp_Click()
  frmCreateCompany.Show 1
End Sub

Private Sub mnunewchallan_Click()
frmchallanEntry.Show 1
End Sub

Private Sub mnuNewExpenses_Click()
    frmExpense.Show 1, Me
End Sub

Private Sub mnuNewIndentFund_Click()
    frmNewIndentFund.Show 1, Me
End Sub

Private Sub mnuNewIndentInfo_Click()
  frmNewIndentEntry.Show 1
End Sub

Private Sub mnunewInvoice_Click()
frmInvoiceEntry.Show 1
End Sub

Private Sub mnuNewParty_Click()
    frmParty.Show 1, Me
End Sub

Private Sub mnuOpenComp_Click()
  frmOpenCompany.Show 1
End Sub
Private Sub mnuQuit_Click()
  If MsgBox("Are you sure to Exit", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  End
End Sub
Private Sub mnuRptAdjustSt_Click()
  frmRptAdjustmentSt.Show 1
End Sub

Private Sub mnurptBondEntry_Click()
frmrptBondEntry.Show 1
End Sub

Private Sub mnurptchallanEntry_Click()
frmrptchallanEntry.Show 1
End Sub

Private Sub mnuRptClearBill_Click()
  frmRptClearBillRegister.Show 1
End Sub
Private Sub mnuRptForwardBill_Click()
  frmRptFordBillRegister.Show 1
End Sub
Private Sub mnuRptIndentSt_Click()
'  frmRptIndentStatement.Show 1
End Sub

Private Sub mnustatementofcandfbills_Click()
    frmRptStatementOfCandFBills.Show 1
End Sub

Private Sub mnustatementofforwardingdate_Click()
    frmRptStatementForwarderDate.Show 1
End Sub
