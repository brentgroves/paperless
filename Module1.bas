Attribute VB_Name = "Module1"
Public SQLConn As ADODB.Connection
Public SQLConnTOOLLIST As ADODB.Connection
Public Const Acrobat = 2
Public Const Crystal = 1
Public Const MPEG = 3
Public craxReport As New CRAXDRT.Report
Public craxApp As New CRAXDRT.Application
Public craxReport2 As New CRAXDRT.Report
Public craxApp2 As New CRAXDRT.Application
Public RightView As String
Public LeftView As String
Public CrystalZoom As Integer
Public CrystalZoom2 As Integer
Public Acrobat0Zoom As Integer
Public Acrobat1Zoom As Integer
Public AcrobatLargeZoom As Integer
Public IdleTime As Date
Public DocListView As Boolean
Public dsDataSource As String
Public newCnc As Integer
Public screenSaverActive As Boolean
Public lastUpdate As Double
Public lastUpdateTime As Date
Public lastUpdateTickCount As Long

Public lastCncCycle As Double
Public updateLastCycleEnabled As Boolean
Public shiftPartNumber As String
Public jobDescription As String
Public partCount As String
Public cell As String
Public shiftCnc As String
Public eff As Double



Global Const SITE = 0  ' 0 for Indiana, 1 for Alabama

Public Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As String) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
    Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40

Public Sub Init()
        DisplayTaskBar (True)
  '      DisplayTaskBar (False)
        MainForm.sockMain(0).Close
        MainForm.sockMain(0).LocalPort = "12345"
        MainForm.sockMain(0).Listen

        
        MainForm.Top = 0
        MainForm.Left = 0
'        MainForm.Width = 28850
'        MainForm.Height = 15850
        MainForm.StatusBar1.Panels(1).Text = ""
 '       MainForm.StatusBar1.Panels(2).Text = "Panel 2"
        MainForm.StatusBar1.Panels(2).Width = 1000
        MainForm.StatusBar1.Panels(3).Width = 2000
   '     MainForm.StatusBar1.Panels(3).Width = 3000
        
        MainForm.ListBar.ImageList = MainForm.vbalImageList1
        MakeConnections
        SetOriginalViewerPositions
        HideLeftControls
        HideRightControls
        HideCenterControls
        ShowEffCenter
        
End Sub

Public Sub MakeConnections()
    On Error GoTo Retry
    DoEvents ' Make sure screen gets updated before attempting to open databases
    Set SQLConn = New ADODB.Connection
    Set SQLConnTOOLLIST = New ADODB.Connection
    If 1 = SITE Then  ' Hartselle
        SQLConn.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=busche document management;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
        SQLConnTOOLLIST.Open "Provider=sqloledb;" & _
               "Data Source=hartselle-sql;" & _
               "Initial Catalog=BUSCHE TOOLLIST;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
    Else ' Indiana
        SQLConn.Open "Provider=sqloledb;" & _
           "Data Source=busche-sql;" & _
           "Initial Catalog=busche document management;" & _
           "User Id=sa;" & _
           "Password=buschecnc1"
        SQLConnTOOLLIST.Open "Provider=sqloledb;" & _
               "Data Source=busche-sql;" & _
               "Initial Catalog=BUSCHE TOOLLIST;" & _
               "User Id=sa;" & _
               "Password=buschecnc1"
    End If
    
    LoadPartNumbers
    
    ClearDocuments
    InitializeReport
    ReconnectForm.Hide
    DoEvents ' Make sure reconnect screen gets hidden
    Exit Sub
Retry:
    MainForm.Timer4.Enabled = True
End Sub

Public Sub LoadPartNumbers()
On Error GoTo Reconnect
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
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub DisplayTaskBar(Visible As Boolean)
    Dim hWnd As Long
    hWnd = FindWindow("Shell_TrayWnd", "")
    If Visible Then
       Call ShowWindow(hWnd, SW_SHOW)
    Else
       Call ShowWindow(hWnd, SW_HIDE)
    End If
    Call EnableWindow(hWnd, Visible)
End Sub

Public Sub PopulateDocuments(PartNumber As String)
On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Dim TEMP
    Dim IsPDF As Long
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT MASTER] INNER JOIN [DOCUMENT PARTNUMBERS] ON [DOCUMENT MASTER].DOCUMENTID = [DOCUMENT PARTNUMBERS].DOCUMENTID WHERE [PARTNUMBER] = '" + PartNumber + "' AND ACTIVE = 1 AND GLOBALDOC = 0 ORDER BY DOCUMENTTITLE", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        If Right(sqlrs.Fields("FILENAME"), 3) = "pdf" Then
            IsPDF = 3
        Else
            IsPDF = 2
        End If
        TEMP = MainForm.ListBar.Bars("A" + Trim(str(sqlrs.Fields("DOCUMENTTYPE")))).Items.Add("A" + Trim(sqlrs.Fields("FILENAME")), , Trim(sqlrs.Fields("DOCUMENTTITLE")), IsPDF)
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    sqlrs.Open "SELECT * FROM [DOCUMENT MASTER] WHERE ACTIVE = 1 AND GLOBALDOC = 1 ORDER BY DOCUMENTTITLE", SQLConn, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        If Right(sqlrs.Fields("FILENAME"), 3) = "pdf" Then
            IsPDF = 3
        Else
            IsPDF = 2
        End If
        TEMP = MainForm.ListBar.Bars("A" + Trim(str(sqlrs.Fields("DOCUMENTTYPE")))).Items.Add("A" + Trim(sqlrs.Fields("FILENAME")), , Trim(sqlrs.Fields("DOCUMENTTITLE")), IsPDF)
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    sqlrs.Open "SELECT * FROM [TOOLLIST MASTER] INNER JOIN [TOOLLIST PARTNUMBERS] ON [TOOLLIST MASTER].PROCESSID = [TOOLLIST PARTNUMBERS].PROCESSID WHERE [PARTNUMBERS] = '" + PartNumber + "' AND (([REVOFPROCESSID] = 0 AND [REVINPROCESS] = 0) OR ([REVOFPROCESSID] <> 0 AND [REVINPROCESS] <> 0) OR ([REVOFPROCESSID] = 0 AND [REVINPROCESS] <> 0))", SQLConnTOOLLIST, adOpenKeyset, adLockReadOnly
    While Not sqlrs.EOF
        TEMP = MainForm.ListBar.Bars("TOOLLIST").Items.Add("A" + Trim(str(sqlrs.Fields("PROCESSID"))), , Trim(sqlrs.Fields("OPERATIONDESCRIPTION")), 1)
  
        sqlrs.MoveNext
    Wend
    Set sqlrs = Nothing
    Dim i
    i = 0
    MainForm.ListBar.Bars("TOOLLIST").Caption = MainForm.ListBar.Bars("TOOLLIST").Caption + "  (" + Trim(str(MainForm.ListBar.Bars("TOOLLIST").Items.Count)) + " Docs)"
    For i = 0 To 200
        On Error Resume Next
        MainForm.ListBar.Bars("A" + Trim(str(i))).Caption = MainForm.ListBar.Bars("A" + Trim(str(i))).Caption + "  (" + Trim(str(MainForm.ListBar.Bars("A" + Trim(str(i))).Items.Count)) + " Docs)"
    Next
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub ClearDocuments()
    MainForm.ListBar.Bars.Clear
    PopulateCategories
End Sub

'*****************************************************************************
'       NOTE:   Changed the 1st parameter type                                  PCC cac001 11-23-09
'ublic Sub ViewDocument(DocumentID As Integer, ViewerType As Integer, Landscape As Boolean)
Public Sub ViewDocument(DocumentID As String, ViewerType As Integer, Landscape As Boolean)
'   ARGUMENTS:
'     RETURNS:
'   CALLED BY:
'       CALLS:
' DESCRIPTION:
'*****************************************************************************
    On Error GoTo Reconnect

    Dim doc_fname   As String
    ' The Tool List Crystal Report is expecting an integer parameter so convert DocumentId to string type
    Dim intDocumentID As Integer
    
        '=========================================================================
        '
        '=========================================================================
    ResetViewers

'Start of additions                                                                                     PCC cac001 11-23-09
    '=========================================================================
    '   Temp fix to catch video files because ViewerType is not set
    '   Just look for the ".mpg" extension
    '
    '   NOTE:   The correct solution should be to set/use ViewerType of MPEG
    '=========================================================================
    doc_fname = str(Val(Right(DocumentID, Len(DocumentID) - 1)))

    '                                     3 = get right 3 extension chars
    '                                         2 = convert to lower case
        If (StrComp(StrConv(Right(DocumentID, 3), 2), "mpg", vbTextCompare) = 0) Then
        '                                                                                                               =0 means strings compare OK
    
'               ResetViewers   Allready done above
                HideLeftControls
                HideRightControls
'xxx    ShowCenterControls
                HideCenterControls
                ShowEffVideo
'                HideEff

                MainForm.WindowsMediaPlayer1.settings.autoStart = False
                MainForm.WindowsMediaPlayer1.stretchToFit = True
                MainForm.WindowsMediaPlayer1.Visible = True

                MainForm.WindowsMediaPlayer1.Left = 5060
                MainForm.WindowsMediaPlayer1.Top = 0
            '    MainForm.WindowsMediaPlayer1.Height = 13400
                MainForm.WindowsMediaPlayer1.Height = 15335
                MainForm.WindowsMediaPlayer1.Width = 23745
                
                If 1 = SITE Then  ' Hartselle
                    MainForm.WindowsMediaPlayer1.URL = "\\hartselle-public\documentstorage\" + Trim(doc_fname) + ".mpg"
                Else
                    MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(doc_fname) + ".mpg"
                End If
                RightView = ""

                Exit Sub
    End If
'End of additions                                                                                       PCC cac001 11-23-09

    '=========================================================================
    '
    '=========================================================================
    If craxReport.ReportTitle <> "Busche Tool List" Then
        If 1 = SITE Then  ' Hartselle
            Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
        Else
            Set craxReport = craxApp.OpenReport("\\buschesv2\public\Report Files\toollist.rpt")
        End If
        craxReport.DiscardSavedData
        craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
        craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (0)
        MainForm.CRViewer1.ReportSource = craxReport
        MainForm.CRViewer1.Zoom 80
        CrystalZoom = 80
    End If
    
    '=========================================================================
    '
    '=========================================================================
    If Landscape = True Then
    ' debug here
        HideLeftControls
        HideRightControls
        'move eff to left

        ShowEffLeft
        ShowCenterControls
        If ViewerType = Acrobat Then
'xxx        MainForm.AcroPDF(2).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".pdf"
            If 1 = SITE Then  ' Hartselle
                MainForm.AcroPDF(2).LoadFile "\\hartselle-public\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
            Else
                MainForm.AcroPDF(2).LoadFile "\\busche-sql\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
            End If
            MainForm.AcroPDF(2).Visible = True
            MainForm.AcroPDF(2).setShowScrollbars (True)
            HideNavigationPanel (2)
        ElseIf ViewerType = MPEG Then
            'XXX Should not get here! (until ViewerType is set correctly) - Chuck Collatz
            MainForm.WindowsMediaPlayer1.Visible = False 'True
'           MainForm.WindowsMediaPlayer1.URL = Trim(Str(documenid)) + ".mpg"
'xxx        MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".mpg"
            'MainForm.WindowsMediaPlayer1.URL = "\\busche-sql\documentstorage\" + Trim(Str(doc_fname)) + ".mpg"
            'MainForm.WindowsMediaPlayer1.play
            'MainForm.WindowsMediaPlayer1.stretchToFit = True
            RightView = ""
        End If
    Else
        Select Case LCase(Trim(RightView))
                                                        '-------------------------------------------------
            Case "acropdf0"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                ShowEffCenter
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(1).Left = 5055 + 11875
                    MainForm.AcroPDF(1).Visible = True
'xxx                MainForm.AcroPDF(1).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".pdf"
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(1).LoadFile "\\hartselle-public\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(1).LoadFile "\\busche-sql\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(1).setShowToolbar (False)
                    MainForm.AcroPDF(1).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "1"
                    MainForm.AcroPDF(0).Left = 5055
                    MainForm.AcroPDF(0).Visible = True
                    HideNavigationPanel (1)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)

                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    LeftView = RightView
                    RightView = MainForm.CRViewer1.Name
                    MainForm.AcroPDF(0).Left = 5055
                    MainForm.AcroPDF(0).Visible = True
                End If
                                                        '-------------------------------------------------
            Case "acropdf1"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                ShowEffCenter
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Left = 5055 + 11875
                    MainForm.AcroPDF(0).Visible = True
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    End If
'xxx                MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".pdf"
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    MainForm.AcroPDF(1).Left = 5055
                    MainForm.AcroPDF(1).Visible = True
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    LeftView = RightView
                    RightView = MainForm.CRViewer1.Name
                    MainForm.AcroPDF(1).Left = 5055
                    MainForm.AcroPDF(1).Visible = True
                End If
                                                        '-------------------------------------------------
            Case "crviewer1"
                                                        '-------------------------------------------------
                HideCenterControls
                ShowLeftControls
                ShowRightControls
                ShowEffCenter
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Left = 5055 + 11875
                    MainForm.AcroPDF(0).Visible = True
'xxx                MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".pdf"
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    Else
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    End If
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    LeftView = RightView
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    MainForm.CRViewer1.Left = 5055
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    RightView = MainForm.CRViewer1.Name
                    HideLeftControls
                End If
                                                        '-------------------------------------------------
            Case ""
                'debug here                                        '-------------------------------------------------
                HideCenterControls
                HideLeftControls
                ShowRightControls
                ShowEffCenter
                If ViewerType = Acrobat Then
                    MainForm.AcroPDF(0).Visible = True
                    If 1 = SITE Then  ' Hartselle
                        MainForm.AcroPDF(0).LoadFile "\\hartselle-public\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    Else
'                        MainForm.AcroPDF(0).LoadFile "c:\" + Trim(str(doc_fname)) + ".pdf"
                        MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(str(doc_fname)) + ".pdf"
                    End If
'xxx                MainForm.AcroPDF(0).LoadFile "\\busche-sql\documentstorage\" + Trim(Str(DocumentID)) + ".pdf"
                    MainForm.AcroPDF(0).setShowToolbar (False)
                    MainForm.AcroPDF(0).setShowScrollbars (True)
                    RightView = MainForm.AcroPDF(0).Name + "0"
                    HideNavigationPanel (0)
                End If
                If ViewerType = Crystal Then
                    MainForm.CRViewer1.Left = 5055 + 11875
                    intDocumentID = Val(doc_fname)
                    
                    craxReport.DiscardSavedData
                    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
'xxx                craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (DocumentID)
                    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (intDocumentID)
                    
                    MainForm.CRViewer1.ViewReport
                    MainForm.CRViewer1.Zoom 80
                    RightView = MainForm.CRViewer1.Name
                End If
            End Select
     End If
     
     Exit Sub
     
'=============================================================================
Reconnect:
'=============================================================================
    ReconnectForm.Show
    MakeConnections
End Sub 'ViewDocument

Public Sub PopulateCategories()
On Error GoTo Reconnect
    Dim sqlrs As ADODB.Recordset
    Dim i As Integer
    Set sqlrs = New ADODB.Recordset
    sqlrs.Open "SELECT * FROM [DOCUMENT TYPE] ORDER BY DOCUMENTDESC ASC", SQLConn, adOpenKeyset, adLockReadOnly
    MainForm.ListBar.Bars.Add "TOOLLIST", , "Tool List"
    While Not sqlrs.EOF
        MainForm.ListBar.Bars.Add "A" + Trim(str(sqlrs.Fields("DocumentTypeID"))), , Trim(sqlrs.Fields("DocumentDesc"))
        sqlrs.MoveNext
    Wend
    sqlrs.Close
    Set sqlrs = Nothing
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub InitializeReport()
On Error GoTo Reconnect
'GoTo irend
 '   Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
    
    If 1 = SITE Then  ' Hartselle
        Set craxReport = craxApp.OpenReport("\\hartselle-public\Shared\Public\Report Files\toollist.rpt")
    Else ' Indiana
        Set craxReport = craxApp.OpenReport("\\buschesv2\public\Report Files\toollist.rpt")
    End If
 
'    Set craxReport = craxApp.OpenReport("P:\Report Files\toollist.rpt")
'\\buschesv2\public\Report Files
    craxReport.DiscardSavedData
    craxReport.ParameterFields.GetItemByName("ProcessID").ClearCurrentValueAndRange
    craxReport.ParameterFields.GetItemByName("ProcessID").AddCurrentValue (0)
    MainForm.CRViewer1.ReportSource = craxReport
    MainForm.CRViewer1.ViewReport
    MainForm.CRViewer1.Zoom 80
    CrystalZoom = 80

    If 1 = SITE Then  ' Hartselle
        Set craxReport2 = craxApp2.OpenReport("\\hartselle-public\Shared\Public\Report Files\Document List.rpt")
    Else
        Set craxReport2 = craxApp2.OpenReport("\\buschesv2\public\Report Files\Document List.rpt")
    End If
 
    
    craxReport2.DiscardSavedData
    craxReport2.ParameterFields.GetItemByName("PartNumber").ClearCurrentValueAndRange
    craxReport2.ParameterFields.GetItemByName("PartNumber").AddCurrentValue ("")
    MainForm.Crviewer2.ReportSource = craxReport2
    MainForm.Crviewer2.ViewReport
    MainForm.Crviewer2.Zoom 80
    CrystalZoom2 = 80
irend:
    Exit Sub
Reconnect:
    ReconnectForm.Show
    MakeConnections
End Sub

Public Sub SetOriginalViewerPositions()
    MainForm.Crviewer2.Left = -12000
    MainForm.Crviewer2.Top = 0
    MainForm.Crviewer2.Height = 13400
    MainForm.Crviewer2.Width = 11875
    MainForm.CRViewer1.Left = -12000
    MainForm.CRViewer1.Top = 0
    MainForm.CRViewer1.Height = 13400
    MainForm.CRViewer1.Width = 11875
    MainForm.AcroPDF(0).Left = 5055 + 11875
    MainForm.AcroPDF(0).Top = 0
    MainForm.AcroPDF(0).Width = 11875
    MainForm.AcroPDF(0).Height = 13400
    MainForm.AcroPDF(1).Left = 5055 + 11875
    MainForm.AcroPDF(1).Top = 0
    MainForm.AcroPDF(1).Width = 11875
    MainForm.AcroPDF(1).Height = 13400
    MainForm.AcroPDF(2).Left = 5055
    MainForm.AcroPDF(2).Top = 0
    MainForm.AcroPDF(2).Width = 23750
    MainForm.AcroPDF(2).Height = 13400
    
    MainForm.WindowsMediaPlayer1.Left = 5055
    MainForm.WindowsMediaPlayer1.Top = 0
    MainForm.WindowsMediaPlayer1.Height = 13400
    MainForm.WindowsMediaPlayer1.Width = 23750
    MainForm.WindowsMediaPlayer1.Visible = False
    
    MainForm.AcroPDF(0).Visible = False
    MainForm.AcroPDF(1).Visible = False
    MainForm.AcroPDF(2).Visible = False
    AcrobatLeftZoom = 80
    AcrobatRightZoom = 80
    AcrobatLargeZoom = 100
End Sub

Public Sub ShowLeftControls()
    With MainForm
            .PanDownBtn(0).Visible = True
            .PanLeftBtn(0).Visible = True
            .PanUpBtn(0).Visible = True
            .PanRightBtn(0).Visible = True
            .ZoomInBtn(0).Visible = True
            .ZoomOutBtn(0).Visible = True
            .ZoomReturn(0).Visible = True
            .PageDownBtn(0).Visible = True
            .PageUpBtn(0).Visible = True
    End With
End Sub

Public Sub ShowRightControls()
    With MainForm
            .PanDownBtn(1).Visible = True
            .PanLeftBtn(1).Visible = True
            .PanUpBtn(1).Visible = True
            .PanRightBtn(1).Visible = True
            .ZoomInBtn(1).Visible = True
            .ZoomOutBtn(1).Visible = True
            .ZoomReturn(1).Visible = True
            .PageDownBtn(1).Visible = True
            .PageUpBtn(1).Visible = True
    End With
End Sub
Public Sub ShowEffCenter()
   With MainForm
      .lblEff.Left = 15722
      .lblEff.Top = 13915
      .lblEff.Visible = True
   End With
End Sub
Public Sub ShowEffVideo()
   With MainForm
      '.lblEff.Left = 5040
'      .lblEff.Left = 13442
      '.lblEff.Left = 24124
'      .lblEff.Left = 23164
      '.lblEff.Left = 21243
      .lblEff.Left = 19803


    

'      .lblEff.Top = 12955
'      .lblEff.Top = 13195
'      .lblEff.Top = 13395
'      .lblEff.Top = 14515
      .lblEff.Top = 14815
      .lblEff.Visible = True
   End With
End Sub

Public Sub ShowEffLeft()
   With MainForm
      .lblEff.Top = 13915
      .lblEff.Left = 11041
      .lblEff.Visible = True
   End With
End Sub

Public Sub HideLeftControls()
    With MainForm
            .PanDownBtn(0).Visible = False
            .PanLeftBtn(0).Visible = False
            .PanUpBtn(0).Visible = False
            .PanRightBtn(0).Visible = False
            .ZoomInBtn(0).Visible = False
            .ZoomOutBtn(0).Visible = False
            .ZoomReturn(0).Visible = False
            .PageDownBtn(0).Visible = False
            .PageUpBtn(0).Visible = False
    End With
End Sub

Public Sub HideRightControls()
    With MainForm
            .PanDownBtn(1).Visible = False
            .PanLeftBtn(1).Visible = False
            .PanUpBtn(1).Visible = False
            .PanRightBtn(1).Visible = False
            .ZoomInBtn(1).Visible = False
            .ZoomOutBtn(1).Visible = False
            .ZoomReturn(1).Visible = False
            .PageDownBtn(1).Visible = False
            .PageUpBtn(1).Visible = False
    End With
End Sub

Public Sub ShowCenterControls()
    With MainForm
            .PanDownBtn(2).Visible = True
            .PanLeftBtn(2).Visible = True
            .PanUpBtn(2).Visible = True
            .PanRightBtn(2).Visible = True
            .ZoomInBtn(2).Visible = True
            .ZoomOutBtn(2).Visible = True
            .ZoomReturn(2).Visible = True
            .PageDownBtn(2).Visible = True
            .PageUpBtn(2).Visible = True
    End With
End Sub

Public Sub HideCenterControls()
    With MainForm
            .PanDownBtn(2).Visible = False
            .PanLeftBtn(2).Visible = False
            .PanUpBtn(2).Visible = False
            .PanRightBtn(2).Visible = False
            .ZoomInBtn(2).Visible = False
            .ZoomOutBtn(2).Visible = False
            .ZoomReturn(2).Visible = False
            .PageDownBtn(2).Visible = False
            .PageUpBtn(2).Visible = False
    End With
End Sub

Public Sub HideEff()
     With MainForm
            .lblEff.Visible = False
     End With
End Sub

Public Sub ResetViewers()
    DocListView = False
    MainForm.CRViewer1.Left = -12000
    MainForm.Crviewer2.Left = -12000
    MainForm.AcroPDF(0).Left = 5055 + 11875
    MainForm.AcroPDF(1).Left = 5055 + 11875
    
'PCC cac001 11-23-09
    '=========================================================================
        '       Move the media player out of the visiable display
    '   Since it left a shadow area where the PDF should display
        '       it is moved off the visiable area of the display when not in use
    '=========================================================================
    MainForm.WindowsMediaPlayer1.Left = 5055
    MainForm.WindowsMediaPlayer1.Top = -13500 '0
    MainForm.WindowsMediaPlayer1.Height = 13400
    MainForm.WindowsMediaPlayer1.Width = 23750
    MainForm.WindowsMediaPlayer1.Visible = False
    
    MainForm.AcroPDF(0).Visible = False
    MainForm.AcroPDF(1).Visible = False
    MainForm.AcroPDF(2).Visible = False
End Sub

Public Sub HideNavigationPanel(Viewer As Integer)
    MainForm.AcroPDF(Viewer).setView "FIT"
    MainForm.AcroPDF(Viewer).setShowToolbar (True)
    MainForm.AcroPDF(Viewer).SetFocus
    MainForm.Timer2.Enabled = True
    While MainForm.Timer2.Enabled = True
        DoEvents
    Wend
    MainForm.AcroPDF(Viewer).setShowToolbar (False)
End Sub
