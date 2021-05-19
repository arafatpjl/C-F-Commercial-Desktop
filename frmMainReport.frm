VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmMainReport 
   ClientHeight    =   10950
   ClientLeft      =   75
   ClientTop       =   480
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleMode       =   0  'User
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPRINTER 
      DownPicture     =   "frmMainReport.frx":0000
      Height          =   375
      Left            =   120
      Picture         =   "frmMainReport.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRVIEWER 
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      lastProp        =   600
      _cx             =   26908
      _cy             =   18865
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmMainReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDPRINTER_Click()
Report.PrinterSetup (hwnd)
End Sub


Private Sub Form_Load()

Screen.MousePointer = vbDefault

  ' MAIN REPORT LOGIN UPDATE

 SetReportLocation Report


  ' SUB REPORT LOGIN UPDATE
  
 Dim CrxSubRep As CRAXDRT.Report

 
   For i = 1 To Report.Sections.Count
        For ii = 1 To Report.Sections(i).ReportObjects.Count
            If Report.Sections(i).ReportObjects(ii).Kind = crSubreportObject Then
                Set CrxSubRep = Report.OpenSubreport(Report.Sections(i).ReportObjects(ii).SubreportName)
                SetReportLocation CrxSubRep
                  
            End If
        Next ii
    Next
    
 
 
End Sub

Private Sub SetReportLocation(ByRef Report As CRAXDRT.Report)
   

Dim mytable As CRAXDRT.DatabaseTable

 Dim sServer As String
    Dim sDatabase As String
    Dim bTrusted As String
    Dim sUserName As String
    Dim sPassword As String
    Dim nTimeout As Long
       
        sServer = WSName
        sDatabase = WDBName
        bTrusted = False
        sUserName = "SA"
        sPassword = "SOFT2ooo"
        nTimeout = 50



For Each mytable In Report.Database.Tables
             
   


Select Case mytable.DllName
    Case "crdb_ado.dll"  'OLE DB
'    crdb_odbc.dll
        With mytable.ConnectionProperties
            .DeleteAll
            .Add "Database Type", "OLE DB (ADO)"
            .Add "Provider", "SQLOLEDB"
            .Add "Data Source", sServer
            .Add "Initial Catalog", sDatabase
            .Add "Integrated Security", bTrusted
            .Add "User ID", sUserName
            .Add "Password", sPassword
            .Add "OLE DB Services", -1
            .Add "General Timeout", nTimeout
        End With
        
    Case Else ' ODBC
    If mytable.TestConnectivity = False Then
  mytable.ConnectBufferString = "Connection String=DRIVER=SQL Server;Database=" & sDatabase & ";Server=" & sServer & ";Trusted_Connection=0;UseDSNProperties=-1;USER ID=" & sUserName & ";Password=" & sPassword & ""
  End If
  
    End Select
    
    
    ' IF not set data from Record set
    
If mytable.Location <> "ado" Then
mytable.Location = WDBName & ".dbo." & mytable.Location
    End If
    
    
Next mytable



End Sub

