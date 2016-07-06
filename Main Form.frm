VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{FE1D1F8B-EC4B-11D3-B06C-00500427A693}#1.1#0"; "vbalLBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "103.24%"
   ClientHeight    =   15855
   ClientLeft      =   105
   ClientTop       =   -450
   ClientWidth     =   28800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   15850
   ScaleMode       =   0  'User
   ScaleWidth      =   28805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerLoggerConnect 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6840
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock sockClient 
      Left            =   6000
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockMain 
      Index           =   0
      Left            =   5400
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerScreenSaver 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5160
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   7080
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Interval        =   20000
      Left            =   6120
      Top             =   840
   End
   Begin VB.CommandButton RefreshCMD 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   475
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   5160
      Top             =   840
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   2
      Left            =   20955
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   14655
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   2
      Left            =   20955
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":05A0
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   13725
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   1
      Left            =   25770
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":0B4C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   14655
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   1
      Left            =   25770
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   13725
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageDownBtn 
      Height          =   615
      Index           =   0
      Left            =   13410
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":1698
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   14580
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PageUpBtn 
      Height          =   615
      Index           =   0
      Left            =   13410
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":1C38
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   13650
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   2
      Left            =   19380
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":21E4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   14100
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   2
      Left            =   18660
      Picture         =   "Main Form.frx":25CC
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   2
      Left            =   20115
      Picture         =   "Main Form.frx":29C1
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   2
      Left            =   17700
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":2DBE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   2
      Left            =   16260
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":31E3
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   2
      Left            =   16980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":3606
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   14700
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   2
      Left            =   16980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":3A40
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   13515
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   1
      Left            =   24165
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":3E73
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   14130
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   1
      Left            =   23445
      Picture         =   "Main Form.frx":425B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   1
      Left            =   24900
      Picture         =   "Main Form.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   1
      Left            =   22485
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":4A4D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   1
      Left            =   21045
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":4E72
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   14130
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   1
      Left            =   21765
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5295
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   14730
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   1
      Left            =   21765
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":56CF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   13530
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   4125
      Left            =   15150
      TabIndex        =   11
      Top             =   1380
      Width           =   2970
      _cx             =   5239
      _cy             =   7276
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
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
      LocaleID        =   1033
   End
   Begin VB.CommandButton PanUpBtn 
      Height          =   615
      Index           =   0
      Left            =   9375
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5B02
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   13515
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanDownBtn 
      Height          =   615
      Index           =   0
      Left            =   9375
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":5F35
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   14700
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanLeftBtn 
      Height          =   615
      Index           =   0
      Left            =   8655
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":636F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton PanRightBtn 
      Height          =   615
      Index           =   0
      Left            =   10095
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Main Form.frx":6792
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomInBtn 
      Height          =   615
      Index           =   0
      Left            =   12510
      Picture         =   "Main Form.frx":6BB7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomOutBtn 
      Height          =   615
      Index           =   0
      Left            =   11055
      Picture         =   "Main Form.frx":6FB4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   14100
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton ZoomReturn 
      Height          =   615
      Index           =   0
      Left            =   11775
      MaskColor       =   &H80000013&
      Picture         =   "Main Form.frx":73A9
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   14100
      Width           =   735
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   0
      Left            =   16080
      TabIndex        =   3
      Top             =   1800
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   16
      Size            =   13236
      Images          =   "Main Form.frx":7791
      Version         =   131072
      KeyCount        =   3
      Keys            =   "ÿÿ"
   End
   Begin vbalLbar6.vbalListBar ListBar 
      Height          =   14175
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   25003
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox PartSelectCombo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Main Form.frx":AB65
      Left            =   960
      List            =   "Main Form.frx":AB6C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   15360
      Width           =   28800
      _ExtentX        =   50800
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   43842
            MinWidth        =   2646
            Text            =   "DNC Status"
            TextSave        =   "DNC Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   4313
            MinWidth        =   4304
            TextSave        =   "9:33 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "8/2/2013"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   1
      Left            =   9540
      TabIndex        =   12
      Top             =   4455
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF 
      Height          =   7065
      Index           =   2
      Left            =   14400
      TabIndex        =   33
      Top             =   2760
      Width           =   6720
      _cx             =   5080
      _cy             =   5080
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer Crviewer2 
      Height          =   4125
      Left            =   11160
      TabIndex        =   37
      Top             =   480
      Width           =   2970
      _cx             =   5239
      _cy             =   7276
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
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
      LocaleID        =   1033
   End
   Begin VB.CommandButton cmdSetCnc 
      Caption         =   "PICK CNC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton DocumentList 
      Caption         =   "Document List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label lblEff 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "100.0%"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   19800
      TabIndex        =   38
      Top             =   14520
      Width           =   3600
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   13215
      Left            =   5160
      TabIndex        =   34
      Top             =   1800
      Width           =   3375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5953
      _cy             =   23310
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KeyToSend As String
Public rsScreenSaver As New ADODB.Recordset
Public clientConnected As Boolean
Public waitingToSend As Boolean
Public CompleteEffReceived As Boolean
Dim workingPanel As String




Public bT1Enable As Boolean
Public bT2Enable As Boolean
Public bT3Enable As Boolean
Public bT4Enable As Boolean
Public bT5Enable As Boolean

Private intSockCnt As Integer

Const LeftViewer = 0
Const RightViewer = 1
Const LargeViewer = 2





Private Sub Command2_Click()
    Pieces.Show vbModal, Me
    If newCnc <> 0 Then
        If IsNull(sockClient.RemoteHost) Or sockClient.RemoteHost <> sockClient.LocalIP Then
             sockClient.RemoteHost = sockClient.LocalIP
        End If
        If sockClient.RemotePort <> 12346 Then
             sockClient.RemotePort = 12346
        End If
        sockClient.Close
        MainForm.StatusBar1.Panels(1).Text = "Attempting to connect to logger...If connection not established in a few minutes then try rebooting computer."
        TimerLoggerConnect.Enabled = True
        sockClient.Connect
        waitingToSend = True
    End If

End Sub

Private Sub cmdSetCnc_Click()
    Pieces.Show vbModal, Me
    If newCnc <> 0 Then
        If IsNull(sockClient.RemoteHost) Or sockClient.RemoteHost <> sockClient.LocalIP Then
             sockClient.RemoteHost = sockClient.LocalIP
        End If
        If sockClient.RemotePort <> 12346 Then
             sockClient.RemotePort = 12346
        End If
        sockClient.Close
        MainForm.StatusBar1.Panels(1).Text = "Attempting to connect to logger...If connection not established in a few minutes then try rebooting computer."
        TimerLoggerConnect.Enabled = True
        sockClient.Connect
        waitingToSend = True
    End If
End Sub

Private Sub sockClient_Connect()
    TimerLoggerConnect.Enabled = False
    If waitingToSend = True Then
        sockClient.SendData str(newCnc)
       MainForm.StatusBar1.Panels(1).Text = "Connection to logger is established...Waiting for server broadcast..."
    End If
        
End Sub

Private Sub sockMain_Close(Index As Integer)
    Unload sockMain(Index)
'    Unload sockMain(intSockCnt)
End Sub
'     If sockMain(Index).State = sckConnected Then
'      sockMain(Index).Close
'     End If

'End Sub

Private Sub TimerLoggerConnect_Timer()
    TimerLoggerConnect.Enabled = False
    MainForm.StatusBar1.Panels(1).Text = "Logger DOWN...Please reboot computer.....Logger DOWN...Please reboot computer.....Logger DOWN...Please reboot computer....."
End Sub


Private Sub Form_Terminate()
  For intCnt = 1 To intSockCnt
     If sockMain(intCnt).State = sckConnected Then
      sockMain(intCnt).Close
     End If
   Next intCnt
    If sockClient.State = sckConnected Then
     sockClient.Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  For intCnt = 1 To intSockCnt
     If sockMain(intCnt).State = sckConnected Then
      sockMain(intCnt).Close
     End If
   Next intCnt
    If sockClient.State = sckConnected Then
     sockClient.Close
    End If


End Sub





Private Sub sockMain_ConnectionRequest(Index As Integer, _
   ByVal requestID As Long)
   
   If intSockCnt = 5 Then
      intSockCnt = 0
   End If
   intSockCnt = intSockCnt + 1
   Load sockMain(intSockCnt)
   sockMain(intSockCnt).Accept requestID
End Sub

'Private Sub sockMain_ConnectionRequest(ByVal requestID As Long)
'   If sockMain.State <> sckClosed Then
'      sockMain.Close  ' This should probably be closed when program is terminated.
'   End If
    
'   sockMain.Accept requestID
    
'   txtStatus.Text = txtStatus.Text & _
 '     "Accepted connection from: " & _
  '    sockMain.RemoteHostIP & vbCrLf
'End Sub
Private Function updateLastCycleTime(lastCncCycle As Double, lastUpdate As Double) As String
    On Error GoTo DataErr
 '       lastUpdate = shiftEff.Item("lastUpdate")
 '       lastCncCycle = 1363345200000# '7 am on 3/15/2013
'       lastCncCycle = 1363172400000# '7 am on 3/13/2013
'        lastCncCycle = 1363363500000# ' 12:5
'        lastCncCycle = shiftEff.Item("lastCncCycle")
        timeDiffInMills = lastUpdate - lastCncCycle
        timeDiffInMins = timeDiffInMills / 1000 / 60
        days = Int(timeDiffInMins / 60 / 24)
        timeLeftInMins = timeDiffInMins - (days * 60 * 24)
        hours = Int(timeLeftInMins / 60)
        mins = Int(timeLeftInMins - (hours * 60))
        Panel = Panel + "   /   Last cycle: "
        If days > 0 Then
            Panel = Panel + str(days) + " days,"
        End If
        If days > 0 Or hours > 0 Then
            Panel = Panel + str(hours) + " hours"
        End If
        If days > 0 Or hours > 0 Or mins > 0 Then
            If days > 0 Or hours > 0 Then
                Panel = Panel + " and " + str(mins) + " mins ago"
            Else
                Panel = Panel + " " + str(mins) + " mins ago"
            End If
        End If
        updateLastCycleTime = Panel

'        If days > 0 Then
'            panel = panel + "  "
'        ElseIf hours > 0 Then
'            panel = panel + "  "
'        ElseIf mins > 0 Then
'            panel = panel + "  "
'        End If
DataErr:
End Function
Private Sub updatePanel(lastCncCycle As Double, lastUpdate As Double)
    On Error GoTo DataErr

        If (eff < 100) Then
          lblEff.BackColor = vbRed
        Else
          lblEff.BackColor = vbGreen
        End If
        lblEff.Caption = Trim(eff) + "%"
        
        Dim Panel As String
        Panel = "     Part# " + shiftPartNumber
        Panel = Panel + "   /   " + jobDescription
        Panel = Panel + "   /   Count: " + partCount
        Panel = Panel + updateLastCycleTime(lastCncCycle, lastUpdate)
        If cell <> "0" Then
            Panel = Panel + "   /   cell# " + cell
        Else
            Panel = Panel + "   /   cnc# " + shiftCnc
        End If
        Panel = Panel + "    "
        MainForm.StatusBar1.Panels(1).Text = Panel
        If screenSaverActive = True Then
            frmScreenSaver.ShiftEffEvent
            frmBackGround.ShiftEffEvent
            screenSaverActive = False
        End If
DataErr:
        
End Sub


Private Sub sockMain_DataArrival(Index As Integer, _
   ByVal bytesTotal As Long)
    On Error GoTo DataErr
    
   Dim strData As String
   Dim strEff As String
   Dim shiftEffString As String
   Dim shiftEff As Object
   Dim strWorking As String
   
   
    Dim intCnt As Integer
    Dim strTemp As String
    Dim bFirstPacket As Boolean
    bFirstPacket = False
    
    sockMain(Index).GetData strData, vbString
    workingPanel = workingPanel + strData
    If (InStr(strData, "}") <> 0) Then
        Set shiftEff = JSON.parse(workingPanel)
        eff = shiftEff.Item("eff")
        cell = shiftEff.Item("cell")
        shiftCnc = shiftEff.Item("cnc")
        shiftPartNumber = shiftEff.Item("partNumber")
        descrL = UCase(Left(shiftEff.Item("jobDescription"), 1))
        descrR = LCase(Right(shiftEff.Item("jobDescription"), Len(shiftEff.Item("jobDescription")) - 1))
        jobDescription = descrL + descrR
        partCount = str(shiftEff.Item("partCount"))
        lastCncCycle = shiftEff.Item("lastCncCycle")
'        lastCncCycle = 1363345200000# '7 am on 3/15/2013
        lastUpdate = CDbl(shiftEff.Item("lastUpdate"))
        lastUpdateTickCount = GetTickCount
        lastUpdateTime = Time
'       lastCncCycle = 1363172400000# '7 am on 3/13/2013
'        lastCncCycle = 1363363500000# ' 12:5
'        lastCncCycle = shiftEff.Item("lastCncCycle")
        updatePanel lastCncCycle, lastUpdate
        workingPanel = ""
        updateLastCycleEnabled = True
   End If
   Exit Sub
DataErr:
    workingPanel = ""
    updateLastCycleEnabled = True
End Sub

Private Sub RefreshCMD_Click()
    On Error GoTo Reconnect
    Dim pn As Integer

    pn = Me.PartSelectCombo.ListIndex
    MainForm.ListBar.Bars.Clear
    PopulateCategories
    Dim sqlrs As ADODB.Recordset
    Set sqlrs = New ADODB.Recordset
    MainForm.PartSelectCombo.Clear
    
    sqlrs.Open "SELECT DISTINCT PARTNUMBER FROM [DOCUMENT PARTNUMBERS] ORDER BY PARTNUMBER ASC", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        MainForm.PartSelectCombo.AddItem Trim(sqlrs.Fields("PARTNUMBER"))
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    Set sqlrs = Nothing
    IdleTime = Now
    MainForm.PartSelectCombo.ListIndex = pn
    PartSelectCombo_Click
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Private Sub DocumentList_Click()
    If Timer5.Enabled Then
        Exit Sub
    End If
    Timer5.Enabled = True
On Error GoTo Reconnect
    ResetViewers
    ShowLeftControls
    ShowEffCenter
    HideRightControls
    HideCenterControls
    
    DocListView = True
    craxReport2.DiscardSavedData
    craxReport2.ParameterFields.GetItemByName("PartNumber").ClearCurrentValueAndRange
    craxReport2.ParameterFields.GetItemByName("PartNumber").AddCurrentValue PartSelectCombo.Text
    MainForm.Crviewer2.ViewReport
    MainForm.Crviewer2.Refresh
    MainForm.Crviewer2.Zoom 80
    CrystalZoom2 = 80
    MainForm.Crviewer2.Left = 5055
    MainForm.Crviewer2.Visible = True
    IdleTime = Now
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Private Sub Form_Load()
    IdleTime = Now
    Init
    TimerScreenSaver.Enabled = True
    screenSaverActive = False
    
    IdleTime = Now
    intSockCnt = 0
    waitingToSend = False
    workingPanel = ""
    newCnc = 0
    lastUpdate = 0#
    lastCncCycle = 0#
    updateLastCycleEnabled = False
    'test code start


    
End Sub

Private Sub LargeComboBtn_Click()
    Cl.ShowDropDownCombo PartSelectCombo
    IdleTime = Now
End Sub

'*****************************************************************************
Private Sub ListBar_ItemClick(Item As vbalLbar6.cListBarItem, Bar As vbalLbar6.cListBar)
'   ARGUMENTS:
'     RETURNS:
'   CALLED BY:
'       CALLS:
' DESCRIPTION:
'*****************************************************************************
    If Timer5.Enabled Then
        Exit Sub
    End If
    Timer5.Enabled = True
    On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT TYPE] WHERE RTRIM(LTRIM(DOCUMENTDESC)) LIKE '" + Trim(Left(Trim(Bar.Caption), InStr(Trim(Bar.Caption), "(") - 1)) + "'", SQLConn, adOpenKeyset, adLockReadOnly
        
'                                                                                                                       PCC cac001 11-23-09
'x      If sqlrs.RecordCount < 1 Then
'x              ViewDocument Str(Val(Right(Item.Key, Len(Item.Key) - 1))), Item.IconIndex, False
'x      Else
'x              ViewDocument Str(Val(Right(Item.Key, Len(Item.Key) - 1))), Item.IconIndex, sqlrs.Fields("LARGEFORMAT")
'x      End If

        '       NOTE:   Need to check file type and set ViewerType (Item.IconIndex) for pdf/mpeg
        '                       This is currently handled in ViewDocument, but should be done somewhere else - Chuck Collatz

        '=========================================================================
        '   Send the whole filename (extension striped in function "ViewDocument")
        '=========================================================================
        If (sqlrs.RecordCount < 1) Then
                ViewDocument Item.key, Item.IconIndex, False
        Else
                ViewDocument Item.key, Item.IconIndex, sqlrs.Fields("LARGEFORMAT")
        End If
    
    IdleTime = Now
    Exit Sub

'-----------------------------------------------------------------------------
Reconnect:
'-----------------------------------------------------------------------------
    ReconnectForm.Show
    MakeConnections
End Sub 'ListBar_ItemClick

Private Sub PageDownBtn_Click(Index As Integer)
    KeyToSend = ("{PGDN}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PageUpBtn_Click(Index As Integer)
    KeyToSend = ("{PGUP}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanDownBtn_Click(Index As Integer)
    KeyToSend = ("{DOWN}")
        If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanLeftBtn_Click(Index As Integer)
    KeyToSend = ("{LEFT}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanRightBtn_Click(Index As Integer)
    KeyToSend = ("{RIGHT}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PanUpBtn_Click(Index As Integer)
    KeyToSend = ("{UP}")
    If DocListView Then
        Me.Crviewer2.SetFocus
        Timer1.Enabled = True
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).SetFocus
        Case "acropdf1"
            Me.AcroPDF(1).SetFocus
        Case "crviewer1"
            Me.CRViewer1.SetFocus
        End Select
        Timer1.Enabled = True
    Case LargeViewer
        Me.AcroPDF(2).SetFocus
        Timer1.Enabled = True
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub PartSelectCombo_Click()
    ClearDocuments
    PopulateDocuments (Trim(PartSelectCombo.Text))
    DocumentList.Enabled = True
    IdleTime = Now
End Sub

Private Sub Timer1_Timer()
    SendKeys (KeyToSend)
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    SendKeys ("^h")
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    If IdleTime < DateAdd("n", -20, Now) Then
        RefreshCMD_Click
        IdleTime = Now
    End If
End Sub

Private Sub Timer5_Timer()
    Timer5.Enabled = False
End Sub


Private Sub TimerScreenSaver_Timer()
'Dim dt As Date
 On Error GoTo Reconnect
 Dim pn As Integer
 If Timer1.Enabled Then
    Exit Sub
 End If
 If Timer2.Enabled Then
    Exit Sub
 End If
 If Timer4.Enabled Then
    Exit Sub
 End If
 If Timer5.Enabled Then
    Exit Sub
 End If


 If rsScreenSaver.State = adStateOpen Then
    rsScreenSaver.Close
 End If


 rsScreenSaver.Open "Select [Document Master].filename, [Document Master].DocumentTitle " & _
 " from [Document Master] inner join [Document PartNumbers] " & _
 " on  [Document Master].DocumentId = [Document PartNumbers].DocumentId " & _
 " Where [Document Master].DocumentType = '3' and [Document PartNumbers].PartNumber = '" & MainForm.PartSelectCombo.Text & "'", SQLConn, adOpenKeyset, adLockReadOnly
    'Order by [Document PartNumbers].PartNumber 2838257
 
 If ((rsScreenSaver.RecordCount > 0) And DateDiff("n", IdleTime, Now) > 15) And (MainForm.PartSelectCombo.ListIndex <> -1) Then
' If ((rsScreenSaver.RecordCount > 0) And DateDiff("n", IdleTime, Now) > 15) And (MainForm.PartSelectCombo.ListIndex <> -1) Then
' If ((rsScreenSaver.RecordCount > 0) And (MainForm.PartSelectCombo.ListIndex <> -1)) Then

   TimerScreenSaver.Enabled = False
   rsScreenSaver.Close

   frmBackGround.Show vbModal, Me
   IdleTime = Now 'Don't do refresh command while screensaver is starting up
   TimerScreenSaver.Enabled = True
 Else
    Dim millsToAdd As Double
    If updateLastCycleEnabled Then
        minsToAdd = Format(Time - lastUpdateTime, "hh:mm:ss")
        ' get time diff
        
        millsToAdd = GetTickCount - lastUpdateTickCount
        lastUpdateTickCount = GetTickCount
        lastUpdateTime = Time
        lastUpdate = lastUpdate + millsToAdd
        updatePanel lastCncCycle, lastUpdate
        rsScreenSaver.Close
    End If
 End If


 
 Exit Sub
 
 

'-----------------------------------------------------------------------------
Reconnect:
'-----------------------------------------------------------------------------
    ReconnectForm.Show
    MakeConnections


 'if Now - IdleTime
End Sub

Private Sub ZoomInBtn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = CrystalZoom2 + 15
        Me.Crviewer2.Zoom (CrystalZoom2)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom + 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom + 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom + 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom + 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom + 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom + 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case LargeViewer
        AcrobatLargeZoom = AcrobatLargeZoom + 15
        Me.AcroPDF(2).setZoom (AcrobatLargeZoom)
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub ZoomOutBtn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = CrystalZoom2 - 15
        Me.Crviewer2.Zoom (CrystalZoom2)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom - 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom - 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom - 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Acrobat0Zoom = Acrobat0Zoom - 15
            Me.AcroPDF(0).setZoom (Acrobat0Zoom)
        Case "acropdf1"
            Acrobat1Zoom = Acrobat1Zoom - 15
            Me.AcroPDF(1).setZoom (Acrobat1Zoom)
        Case "crviewer1"
            CrystalZoom = CrystalZoom - 15
            Me.CRViewer1.Zoom (CrystalZoom)
        End Select
    Case LargeViewer
        AcrobatLargeZoom = AcrobatLargeZoom - 15
        Me.AcroPDF(2).setZoom (AcrobatLargeZoom)
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub ZoomReturn_Click(Index As Integer)
    If DocListView Then
        CrystalZoom2 = 80
        Me.Crviewer2.Zoom (80)
    Else
    Select Case Index
    Case LeftViewer
        Select Case LCase(Trim(LeftView))
        Case "acropdf0"
            Me.AcroPDF(0).setLayoutMode "OneColumn"
            Me.AcroPDF(0).setView "Fit"
        Case "acropdf1"
            Me.AcroPDF(1).setLayoutMode "OneColumn"
            Me.AcroPDF(1).setView "Fit"
        Case "crviewer1"
            CrystalZoom = 80
            Me.CRViewer1.Zoom (80)
        End Select
    Case RightViewer
        Select Case LCase(Trim(RightView))
        Case "acropdf0"
            Me.AcroPDF(0).setLayoutMode "OneColumn"
            Me.AcroPDF(0).setView "Fit"
        Case "acropdf1"
            Me.AcroPDF(1).setLayoutMode "OneColumn"
            Me.AcroPDF(1).setView "Fit"
        Case "crviewer1"
            CrystalZoom = 80
            Me.CRViewer1.Zoom (80)
        End Select
    Case LargeViewer
            Me.AcroPDF(2).setLayoutMode "OneColumn"
            Me.AcroPDF(2).setView "Fit"
    End Select
    End If
    IdleTime = Now
End Sub

Private Sub Timer4_Timer()
    Timer4.Enabled = False
  '  DoEvents moved to makeConnections because we want the connection screen to show up
    MakeConnections
End Sub

