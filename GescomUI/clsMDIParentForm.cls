VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMDIParentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 0

'''Public Enum AppToolBarCmd
'''    tbcGoFirst = 0
'''    tbcGoPrevious = 1
'''    tbcGoNext = 2
'''    tbcGoLast = 3
'''    tbcLastUsed = 4
'''    tbcSearch = 5
'''    tbcAdd = 6
'''    tbcDelete = 7
'''    tbcUpdate = 8
'''    tbcCancelUpdate = 9
'''    tbcPrintReport = 10
'''    tbcCloseForm = 11
'''    tbcCopy = 12
'''    tbcCut = 13
'''    tbcPaste = 14
'''    tbcHelp = 15
'''    tbcExecuteProcess = 16
'''    tbcExecuteQuery = 17
'''    tbcClearFilter = 18
'''    tbcCheckAll = 19
'''    tbcUncheckAll = 20
'''    tbcShowAllCheckedItems = 21
'''    tbcDuplicate = 22
'''    tbcExportTo = 23
'''    tbcExportToADTG = 24
'''    tbcExportToXML = 25
'''    tbcExportToExcel = 26
'''    tbcExportToCSV = 27
'''    tbcSetRestrictor = 28
'''    tbcRefresh = 29
'''End Enum
'''
'''Public Enum fwUpdateMode
'''    umNavigation = 0
'''    umAddNew = 1
'''    umUpdate = 2
'''    umRetrive = 3
'''End Enum
'''
'''Public Enum fwToolBarEvents
'''    tbeButtonClick = 0
'''    tbeButtonDropDown = 1
'''    tbeMenuClick = 2
'''End Enum
'''
'''Public Enum FieldSecurity
'''    fsAllRights = 0
'''    fsHidden
'''    fsLocked
'''End Enum
'''
'''Public Enum enuProgramTypes
'''    ptyUserControl = 0
'''    ptyWebPage = 1
'''End Enum

''''Private mtlbToolBar As clsFwToolBar
''''Private mfrmActiveChildForm As IEmbeddedControl
Private mfrmMain As Object
Private mlngUpdateMode As Integer
Private mcolPropertiesBag As Collection
Private mcolEmbeddedControls As Collection
Private mcolReports As Collection
'Perfil del usuario
Private mrstUserMenu As Recordset
Private mrstUserPrograms As Recordset
Private mrstUserSpecialPrograms As Recordset
Private mrstUserReports As Recordset
Private mstrComputerName As String
Private mlngGrupoUsuario As Integer
Private mstrEnterpriseName As String
Private mstrDataBase As String
Private mstrServerName As String
Private mstrBackgroundImage As String
Private mstrUserHomePage As String
Private mblnLoadWeb As Boolean
Private mstrIDDefaultProgram As String
Private mstrUserLanguageID As String
''''Private mstrLicenseInfo As String
Private mblnRegisterIdentityWasRetrive As Boolean
Private mstrAdminUser As String
Private mstrAppServerName As String
Private mstrOpenningProgramName As String
Private mrstActions As Recordset

Public Property Get GlobalPropertiesBag() As Collection
    Set GlobalPropertiesBag = mcolPropertiesBag
End Property

Public Property Get ChildForms() As Collection
    Set ChildForms = mcolEmbeddedControls
End Property

Public Function UserName() As String

    If Len(mstrUsuario) = 0 Then
''''        mstrUsuario = GetWinUserName   'OJOOJO
    End If
    UserName = mstrUsuario
End Function

Public Function UserGroup() As Integer
    UserGroup = mlngGrupoUsuario
End Function

Public Function UserBackgroundImage() As String
    UserBackgroundImage = mstrBackgroundImage
End Function

Public Function UserHomePage() As String
    UserHomePage = mstrUserHomePage
End Function

Public Function UserLoadHomePage() As Boolean
    UserLoadHomePage = mblnLoadWeb
End Function

Public Function UserDefaultProgram() As String
    UserDefaultProgram = mstrIDDefaultProgram
End Function

Public Function UserLanguage() As String
    UserLanguage = mstrUserLanguage
End Function

Public Function UserLanguageID() As String
    UserLanguageID = mstrUserLanguageID
End Function

Public Function ComputerName() As String
    
    If Len(mstrComputerName) = 0 Then
''''        mstrComputerName = GetWinComputerName 'OJOOJO
    End If
    ComputerName = mstrComputerName
End Function

Public Function IsFormOpen(strDocName As String) As Boolean
    Dim ctlEmbeddedControl As IEmbeddedControl
    
    IsFormOpen = False
    For Each ctlEmbeddedControl In mcolEmbeddedControls
        If ctlEmbeddedControl.Name = strDocName Then
            IsFormOpen = True
            Exit For
        End If
    Next ctlEmbeddedControl
End Function

''''Public Property Get ActiveDoc() As IEmbeddedControl
''''    Set ActiveDoc = mfrmActiveChildForm
''''End Property

''''Public Property Set ActiveDoc(ctlChildControl As IEmbeddedControl)
''''    Dim strDocName As String
''''
''''    On Error GoTo TratarError
''''
''''    Set mfrmActiveChildForm = ctlChildControl
''''    If ctlChildControl Is Nothing Then
''''        UpdateMode = umNavigation
''''        If Not mtlbToolBar Is Nothing Then mtlbToolBar.Clear
''''        'Pendiente: eliminar menú que no sea predeterminado
''''    End If
''''    Exit Property
''''TratarError:
''''    GenerateError Err.Number, Err.Description, "clsMDIParentForm.ActiveDoc", Title
''''End Property
''''
Public Sub OpenForm(strDocID As String, _
                    strDocName As String, _
                    Optional strCriteria As String, _
                    Optional Parameters As String, _
                    Optional strDocFormTitle As String, _
                    Optional strParentDoc As String)
    
    Dim frmD As Object
    Dim objEmbedded As IEmbeddedControl
    Dim strDocProgID As String
    Dim i As Long
    Dim rstProgramInfo As Recordset
    
    On Error GoTo NoControl
    
    If IsFormOpen(strDocName) Then
        'Cuando el formulario a mostrar tiene dependencias se mostrará el último de la lista
'        Do While mcolEmbeddedControls(strDocName).ChildDoc <> ""
'            strDocName = mcolEmbeddedControls(strDocName).ChildDoc
'        Loop
'        mcolEmbeddedControls(strDocName).OpenDoc strCriteria, Parameters
        mcolEmbeddedControls(strDocName).RefreshDoc strCriteria, Parameters
        mcolEmbeddedControls(strDocName).MDIChildForm.ZOrder 0
    Else
        Screen.MousePointer = vbHourglass
        'Recuperar el ProgID según el ID del objeto y actualizar versiones.
        Set rstProgramInfo = GetProgram(strDocID)
        If Not rstProgramInfo Is Nothing Then
            If Not IsEmptyValue(rstProgramInfo("DescPrograma")) Then
                strDocProgID = rstProgramInfo("DescPrograma")
            End If
            If Len(strDocFormTitle) = 0 Then
                strDocFormTitle = Nz(rstProgramInfo("TraduccionForm"), rstProgramInfo("FormCaption"))
            End If
            If Nz(rstProgramInfo("Tipo"), ptyUserControl) = ptyUserControl Then
                Set frmD = mfrmMain.NewDocViewerForm
                Set objEmbedded = frmD.OpenForm(strDocProgID, strDocName, strCriteria, Parameters, strDocFormTitle)
                If Not objEmbedded Is Nothing Then
                    mcolEmbeddedControls.Add objEmbedded, strDocName
                    objEmbedded.ParentDoc = strParentDoc
                    If Len(strParentDoc) > 0 Then
                        On Error Resume Next
                        mcolEmbeddedControls(strParentDoc).ChildDoc = strDocName
                        On Error GoTo 0
                    End If
                    frmD.Show
                Else
                    Set frmD = Nothing
                End If
            Else
                Set frmD = mfrmMain.NewHTMLViewerForm
                Set objEmbedded = frmD
                If Not objEmbedded Is Nothing Then
                    objEmbedded.Name = strDocName
                    objEmbedded.ProgID = strDocProgID
                    objEmbedded.Caption = strDocFormTitle
                    mcolEmbeddedControls.Add objEmbedded, strDocName
                    objEmbedded.ParentDoc = strParentDoc
                    If Len(strParentDoc) > 0 Then
                        On Error Resume Next
                        mcolEmbeddedControls(strParentDoc).ChildDoc = strDocName
                        On Error GoTo 0
                    End If
                    frmD.GoToHomePage strDocProgID
                    frmD.Show
                End If
            End If
        Else
            GenerateError 55306, vbCritical + vbOKOnly, Title, strDocID
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
NoControl:
    Screen.MousePointer = vbDefault
    GenerateError Err.Number, Err.Description, "clsMDIParentForm.OpenForm", Title
End Sub

Public Sub CloseForm(strDocName As String)
    Dim frmForm As Object
    
    On Error Resume Next
    Set frmForm = mcolEmbeddedControls(strDocName).MDIChildForm
    If Not frmForm Is Nothing Then
        Unload frmForm
    End If
End Sub
''''
''''Public Sub OpenGridPreview(strDocName As String, grdData As Object, Optional strCaption As String, Optional strParentDoc As String)
''''    Dim frmD As Object
''''    Dim objEmbedded As IEmbeddedControl
''''
''''On Error GoTo TratarError
''''
''''    If Not grdData Is Nothing Then
''''        Set frmD = mfrmMain.NewGridViewerForm
''''        Set objEmbedded = frmD
''''        If Not objEmbedded Is Nothing Then
''''            objEmbedded.Name = strDocName
''''            objEmbedded.Caption = strCaption
''''            mcolEmbeddedControls.Add objEmbedded, strDocName
''''            objEmbedded.ParentDoc = strParentDoc
''''            Set frmD.DataGrid = grdData
''''            grdData.PrintPreview frmD.grdPreview
''''            frmD.Show
''''            If Len(strParentDoc) > 0 Then
''''                On Error Resume Next
''''                mcolEmbeddedControls(strParentDoc).ChildDoc = strDocName
''''                On Error GoTo TratarError
''''            End If
''''        End If
''''    End If
''''
''''    Exit Sub
''''
''''TratarError:
''''    'MSG="Ha ocurrido un error al obtener información de su impresora."
''''    '     "Asegúrese de que la impresora está configurada correctamente."
''''    MsgBox TraslateWord(55374), vbCritical, Me.Title
''''End Sub
''''''''
''''''''Public Sub OpenReport(strReportName As String, Optional rcsDataSource As Recordset = Nothing, _
''''''''                      Optional strCaption As String, Optional rptReportType As fwReportType = 0, _
''''''''                      Optional strReportSelectionCriteria As String, _
''''''''                      Optional strReportFormulas As String, _
''''''''                      Optional strReportParameters As String, _
''''''''                      Optional strPrinterName As String, Optional strReportID As String)
''''''''    Dim objReportChildLess As clsReport
''''''''    Dim objEmbeddedCtrl As IEmbeddedControl
''''''''    Dim frmViewer As Object
''''''''    Dim rstReportInfo As Recordset
''''''''    Dim Cancel As Boolean
''''''''    Dim objItemSubreport As clsItemsubreport
''''''''    Dim lngPos As Long
''''''''    Dim strSubreportName As String
''''''''
''''''''    Dim lngWindowState As Long
''''''''    Dim lngDestination As Long
''''''''    Dim blnPromptUser As Boolean
''''''''    Dim lngNumberOfCopy As Long
''''''''    Dim strDiskFileName As String
''''''''    Dim strMailToList As String
''''''''    Dim strFileName As String
''''''''    Dim strFormulasString As String
''''''''    Dim strParametersString As String
''''''''    Dim strKey As String
''''''''
''''''''    On Error GoTo TratarError
''''''''    If Len(strReportName) > 0 Or Len(strReportID) > 0 Then
''''''''        If Len(strReportName) > 0 Then
''''''''            lngPos = InStrRev(strReportName, cnFldNameSep)
''''''''            If lngPos = 0 Then strReportName = strReportName & cnRptFileExt
''''''''        End If
''''''''        Set rstReportInfo = GetReport(strReportName, strReportID)
''''''''        If Not rstReportInfo Is Nothing Then
''''''''            If Len(strReportName) > 0 Then
''''''''                strFileName = ExtractExtension(strReportName)
''''''''            Else
''''''''                strReportName = Nz(rstReportInfo("NombInforme"), vbNullString)
''''''''                strFileName = ExtractExtension(strReportName)
''''''''            End If
''''''''            If Len(strReportID) = 0 Then
''''''''                strReportID = Nz(rstReportInfo("IDInforme"), vbNullString)
''''''''            End If
''''''''            If Len(strPrinterName) = 0 Then
''''''''                strPrinterName = Nz(rstReportInfo("Impresora"), vbNullString)
''''''''            End If
''''''''            If rptReportType = 0 Then
''''''''                rptReportType = Nz(rstReportInfo("IdTipoInforme"), rptCustom)
''''''''            End If
''''''''            If Len(strCaption) = 0 Then
''''''''                If Not IsEmptyValue(rstReportInfo("Traduccion")) Then
''''''''                    strCaption = rstReportInfo("Traduccion")
''''''''                ElseIf Not IsEmptyValue(rstReportInfo("DescInforme")) Then
''''''''                    strCaption = rstReportInfo("DescInforme")
''''''''                End If
''''''''            End If
''''''''            If Len(strReportID) > 0 And Len(strReportName) > 0 Then
''''''''                Set objReportChildLess = New clsReport
''''''''                If Not objReportChildLess Is Nothing Then
''''''''                    Set objReportChildLess.ActiveChildForm = mfrmActiveChildForm
''''''''                    objReportChildLess.FileName = Path & cnReportPath & cnSlashSeparator & strReportName
''''''''                    If Len(strReportID) > 0 Then objReportChildLess.Name = strReportID
''''''''                    objReportChildLess.ReportType = rptReportType
''''''''                    objReportChildLess.PrinterName = strPrinterName
''''''''                    Set objReportChildLess.DataSource = rcsDataSource
''''''''                    objReportChildLess.SelectionFormula = strReportSelectionCriteria
''''''''                    SetFormulasValues objReportChildLess.Formulas, strReportFormulas
''''''''                    SetParametersValues objReportChildLess.Parameters, strReportParameters
'''''''''                    Set frmViewer = mfrmMain.NewReportViewerForm
'''''''''                    If Not frmViewer Is Nothing Then
'''''''''                        Set objReportChildLess.FormViewer = frmViewer
'''''''''                        Set objReportChildLess.Viewer = frmViewer.CRViewer1
''''''''                        If Len(strCaption) > 0 Then
''''''''                            objReportChildLess.FormViewerCaption = strCaption
''''''''                        End If
'''''''''                        Set frmViewer = Nothing
'''''''''                    End If
'''''''''                    mcolReports.Add objReportChildLess
''''''''
''''''''                    If Not mfrmActiveChildForm Is Nothing Then
''''''''                        lngWindowState = objReportChildLess.WindowState
''''''''                        lngDestination = objReportChildLess.Destination
''''''''                        blnPromptUser = objReportChildLess.PromptUser
''''''''                        lngNumberOfCopy = objReportChildLess.NumberOfCopies
''''''''                        strDiskFileName = objReportChildLess.DiskFileName
''''''''                        strMailToList = objReportChildLess.MailToList
''''''''
''''''''                        mfrmActiveChildForm.GetReportOptions strFileName, Cancel, lngWindowState, lngDestination, blnPromptUser, lngNumberOfCopy, strDiskFileName, strMailToList
''''''''                        If Cancel Then Exit Sub
''''''''
''''''''                        objReportChildLess.WindowState = lngWindowState
''''''''                        objReportChildLess.Destination = lngDestination
''''''''                        objReportChildLess.PromptUser = blnPromptUser
''''''''                        objReportChildLess.NumberOfCopies = lngNumberOfCopy
''''''''                        objReportChildLess.DiskFileName = strDiskFileName
''''''''                        objReportChildLess.MailToList = strMailToList
''''''''
''''''''                        mfrmActiveChildForm.SetReportProperties objReportChildLess, Cancel
''''''''                        If Cancel Then Exit Sub
''''''''                        strFormulasString = mfrmActiveChildForm.GetReportFormulas(strFileName, Cancel, True) '+++
''''''''                        If Len(strFormulasString) > 0 Then '+++
''''''''                            SetFormulasValues objReportChildLess.Formulas, strFormulasString '+++
''''''''                        End If
''''''''                        If Cancel Then Exit Sub
''''''''                        strParametersString = mfrmActiveChildForm.GetReportParameters(strFileName, Cancel, True) '+++
''''''''                        If Len(strParametersString) > 0 Then '+++
''''''''                            SetParametersValues objReportChildLess.Parameters, strParametersString '+++
''''''''                        End If
''''''''                        If Cancel Then Exit Sub
''''''''                        Select Case objReportChildLess.ReportType
''''''''                            Case rptCurrentRecord
''''''''                            Case rptAssistantTable
''''''''                            Case rptRecordset
''''''''                                If objReportChildLess.DataSource Is Nothing Then Exit Sub
''''''''                            Case rptSelectionCriteria
''''''''                                If Len(Trim$(objReportChildLess.SelectionFormula)) = 0 Then
''''''''                                    objReportChildLess.SelectionFormula = mfrmActiveChildForm.GetReportSelectionCriteria(objReportChildLess.Alias, Cancel, True)
''''''''                                End If
''''''''                            Case rptFilterLess
''''''''                            Case Else
''''''''                                If Len(Trim$(objReportChildLess.SelectionFormula)) = 0 Then
''''''''                                    objReportChildLess.SelectionFormula = mfrmActiveChildForm.GetReportSelectionCriteria(objReportChildLess.Alias, Cancel, True)
''''''''                                End If
''''''''                                If objReportChildLess.DataSource Is Nothing Then     ' Si se va a cargar el reporte a partir de un recordset
''''''''                                    Set objReportChildLess.DataSource = mfrmActiveChildForm.GetReportDataSource(strFileName, Cancel, True)
''''''''                                End If
''''''''                        End Select
''''''''                        If Cancel Then Exit Sub
''''''''                        If Not objReportChildLess.Subreports Is Nothing Then
''''''''                            For Each objItemSubreport In objReportChildLess.Subreports
''''''''                                strSubreportName = objItemSubreport.Name
''''''''                                strFormulasString = mfrmActiveChildForm.GetReportFormulas(strSubreportName, Cancel, False) '+++
''''''''                                If Len(strFormulasString) > 0 Then '
''''''''                                    SetFormulasValues objItemSubreport.Formulas, strFormulasString  '+++
''''''''                                End If
''''''''                                strParametersString = mfrmActiveChildForm.GetReportParameters(strSubreportName, Cancel, False) '+++
''''''''                                If Len(strParametersString) > 0 Then
''''''''                                    SetParametersValues objItemSubreport.Parameters, strParametersString  '+++
''''''''                                End If
''''''''
''''''''                                If Len(Trim$(objItemSubreport.SelectionFormula)) = 0 Then
''''''''                                    objItemSubreport.SelectionFormula = mfrmActiveChildForm.GetReportSelectionCriteria(objItemSubreport.Alias, Cancel, False)
''''''''                                End If
''''''''                                If objItemSubreport.DataSource Is Nothing Then Set objItemSubreport.DataSource = mfrmActiveChildForm.GetReportDataSource(strSubreportName, Cancel, False)
''''''''                                Set objItemSubreport = Nothing
''''''''                            Next
''''''''                        End If
''''''''                    End If
''''''''                    objReportChildLess.OpenReport
''''''''
''''''''                    If Not objReportChildLess.FormViewer Is Nothing Then
''''''''                        If Not mfrmActiveChildForm Is Nothing Then
''''''''                            strKey = mfrmActiveChildForm.Name & objReportChildLess.Name & (mcolReports.Count + 1)
''''''''                        Else
''''''''                            strKey = objReportChildLess.Name & (mcolReports.Count + 1)
''''''''                        End If
''''''''                        mcolReports.Add objReportChildLess, strKey
''''''''                        Set objEmbeddedCtrl = objReportChildLess.FormViewer
''''''''                        objEmbeddedCtrl.Name = strKey
''''''''                        Set objEmbeddedCtrl = Nothing
''''''''                    End If
''''''''                End If
''''''''                Set objReportChildLess = Nothing
''''''''            End If
''''''''        End If
''''''''    End If
''''''''    Exit Sub
''''''''TratarError:
''''''''    Set objReportChildLess = Nothing
''''''''    GenerateError Err.Number, Err.Description, "clsMDIParentForm.OpenReport", Title
''''''''End Sub

Private Function IsReportID(strReportName As String) As Boolean
    IsReportID = Left$(strReportName, 1) = "<" And Right$(strReportName, 1) = ">"
End Function

''''Private Sub SetFormulasValues(colFormulas As clsColFormulas, strFormulas As String)
''''    Dim arrFormulas() As String
''''    Dim arrFormula() As String
''''    Dim objFormula As clsItemFormula
''''    Dim i As Long
''''
''''    If Not colFormulas Is Nothing Then
''''        If colFormulas.Count > 0 Then
''''            arrFormulas = Split(strFormulas, cnListSeparator)
''''            For i = 0 To UBound(arrFormulas)
''''                arrFormula = Split(arrFormulas(i), cnValuesSeparator)
''''                On Error Resume Next
''''                Set objFormula = colFormulas.Item(Trim$(arrFormula(0)))
''''                objFormula.Value = arrFormula(1)
''''            Next i
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub SetParametersValues(colParameters As clsColParameters, strParameters As String)
''''    Dim arrParametros() As String
''''    Dim arrParametro() As String
''''    Dim objParameter As clsItemParameter
''''    Dim i As Long
''''
''''    If Not colParameters Is Nothing Then
''''        If colParameters.Count > 0 Then
''''            arrParametros = Split(strParameters, cnListSeparator)
''''            For i = 0 To UBound(arrParametros)
''''                arrParametro = Split(arrParametros(i), cnValuesSeparator)
''''                On Error Resume Next
''''                Set objParameter = colParameters.Item(Trim$(arrParametro(0)))
''''                objParameter.Value = arrParametro(1)
''''            Next i
''''        End If
''''    End If
''''End Sub
''''

''''Public Sub CloseReport(ByVal strReportName As String)
''''    Dim objReport As clsReport
''''    Dim i As Long
''''
''''    strReportName = Path & cnReportPath & cnSlashSeparator & IIf(InStr(UCase(strReportName), UCase(".rpt")) > 0, "", ".rpt")
''''    For i = mcolReports.Count To 1 Step -1
''''        Set objReport = mcolReports.Item(i)
''''        If objReport.FileName = strReportName Then
''''            objReport.CloseReport
''''            mcolReports.Remove i
''''            Exit For
''''        End If
''''        Set objReport = Nothing
''''    Next i
''''End Sub
''''
''''Public Property Get Toolbar() As clsFwToolBar
''''    Set Toolbar = mtlbToolBar
''''End Property
''''
Public Property Get ParentForm() As Object
    Set ParentForm = mfrmMain
End Property

Public Property Set ParentForm(frmMDIFrame As Object)
    Set mfrmMain = frmMDIFrame
End Property
''''
''''Public Sub ProcessToolBarEvents(intCmdTag As Integer, intToolBarEvent As Integer, Optional lngButtonIndex As Long = -1)
''''    If Not mfrmActiveChildForm Is Nothing Then
''''        Select Case intCmdTag
''''            'Procesar comandos predeterminados
''''            Case tbcCloseForm
''''                CloseForm mfrmActiveChildForm.Name
''''            Case Else
''''                mfrmActiveChildForm.ProcessToolBarEvents intCmdTag, intToolBarEvent, lngButtonIndex
''''        End Select
''''    End If
''''End Sub

Public Property Let UpdateMode(intState As Integer)
    mlngUpdateMode = intState
    With mfrmMain.sbStatusBar
        Select Case intState
            Case umNavigation
                .Panels("pnlMode").Text = ""
                .Panels("pnlMode").Picture = Nothing
            Case umAddNew
                .Panels("pnlMode").Text = "Nuevo Registro"
                .Panels("pnlMode").Picture = mfrmMain.ilstStatusBar.ListImages(1).Picture
            Case umUpdate
                .Panels("pnlMode").Text = "Actualizando Registro"
                .Panels("pnlMode").Picture = mfrmMain.ilstStatusBar.ListImages(2).Picture
        End Select
    End With
End Property

Public Function Path() As String
    Path = mfrmMain.AppPath
End Function

Public Function Title() As String
    If Not mfrmMain Is Nothing Then
        Title = mfrmMain.AppTitle
    Else
        Title = App.Title
    End If
End Function

Public Function ExeName() As String
    ExeName = mfrmMain.AppExeName
End Function

Public Function CompanyName() As String
    CompanyName = mfrmMain.AppCompanyName
End Function

Public Function Version() As String
    Version = mfrmMain.AppVersion
End Function

''''Public Function ServerName() As String
''''    GetDataBase
''''    ServerName = mstrServerName
''''End Function

''''Public Function DataBaseName() As String
''''    GetDataBase
''''    DataBaseName = mstrDataBase
''''End Function
''''
''''Public Function EnterpriseName() As String
''''    GetDataBase
''''    EnterpriseName = mstrEnterpriseName
''''End Function
''''
''''Private Sub GetDataBase()
''''    Dim objApp As NvNegocio10.Aplicacion
''''    Dim strDataBase As String
''''
''''    Set objApp = New NvNegocio10.Aplicacion
''''    strDataBase = objApp.GetUserDataBase
''''    Set objApp = Nothing
''''    mstrDataBase = GetPropertyValue(strDataBase, "Database")
''''    mstrEnterpriseName = GetPropertyValue(strDataBase, "Enterprise")
''''    mstrServerName = GetPropertyValue(strDataBase, "Server")
''''End Sub

Public Sub StatusMsg(strStateMsg As String, Optional picPicture As StdPicture = Nothing)
    With mfrmMain.sbStatusBar
        Set .Panels("pnlStatus").Picture = picPicture
        .Panels("pnlStatus").Text = strStateMsg
    End With
End Sub

Public Function AddMenuItem(strCaption As String, strKey As String, Optional strParentMenu As String, Optional lngCmdTag As Long, Optional ByVal blnChecked As Boolean = False, Optional ByVal blnEnabled As Boolean = True) As Object
    Set AddMenuItem = mfrmMain.AddMenuItem(strCaption, strKey, strParentMenu, lngCmdTag, blnChecked, blnEnabled)
End Function

Public Function GetMenuItem(strKey As String, Optional strParentMenu As String) As Object
    Set GetMenuItem = mfrmMain.GetMenuItem(strKey, strParentMenu)
End Function

Private Sub RemoveMenuItem(strKey As String, Optional strParentMenu As String)
    mfrmMain.RemoveMenuItem strKey, strParentMenu
End Sub

Public Sub RemoveAllMenuItems()
    mfrmMain.RemoveAllMenuItems
End Sub

Private Sub Class_Initialize()
    Set mcolPropertiesBag = New Collection
    Set mcolEmbeddedControls = New Collection
    Set mcolReports = New Collection
''''    Set mtlbToolBar = New clsFwToolBar
End Sub

Public Sub ShutDown()
    Set mcolPropertiesBag = Nothing
    Set mcolEmbeddedControls = Nothing
''''    If Not mcolReports Is Nothing Then
''''        UnloadChildFormReports
''''        Set mcolReports = Nothing
''''    End If
''''    If Not mtlbToolBar Is Nothing Then
''''        mtlbToolBar.ShutDown
''''        Set mtlbToolBar = Nothing
''''    End If
''''    Set mfrmActiveChildForm = Nothing
End Sub

Private Sub Class_Terminate()
    ShutDown
End Sub

''''Public Function LicenseInfo() As String
''''    LicenseInfo = mstrLicenseInfo
''''End Function
''''
''''Public Function UserSettings() As String
''''    If Len(mstrUserSetting) = 0 Then
''''        GetUserSettings
''''    End If
''''    UserSettings = mstrUserSetting
''''End Function
''''
Public Function GetUserSettings() As Boolean
''''    Dim objApp As Aplicacion
    Dim blnContinue As Boolean
    Dim strMessage As String
''''    Dim udtPrgms As udtPrograms
    
    GetUserSettings = False
    mstrServerName = vbNullString
    mstrDataBase = vbNullString
    mstrEnterpriseName = vbNullString
    mstrUsuario = vbNullString
    mlngGrupoUsuario = 0
    mstrUserLanguage = vbNullString
    mstrUserLanguageID = vbNullString
    mstrBackgroundImage = vbNullString
    mblnLoadWeb = False
    mstrUserHomePage = vbNullString
    mstrIDDefaultProgram = vbNullString
''''    mstrLicenseInfo = vbNullString

    On Error Resume Next
    blnContinue = True
''''    Set objApp = New Aplicacion
''''    If Err.Number <> 0 Then
''''        'MSG="No se ha podido crear el objeto Aplicación." & vbCrLf & _
''''        '    "Causa: Es posible que el componente no esté correctamente registrado." & vbCrLf & _
''''        '    "Solución: La clase Aplicación pertenece a la librería NvNegocio.dll, asegurese que dicho" & vbCrLf & _
''''        '    "componente, en el caso del servidor, esté registrado en el COM+ y, en el caso de los puestos clientes," & vbCrLf & _
''''        '    "esté registrado su correspondiente Proxy."
''''        GenerateMessage 55309, vbCritical + vbOKOnly, Title
''''        Exit Function
''''    End If
''''    'Crear sesión de usuario
''''    If Not objApp.OpenSession(GetWinComputerName, GetWinUserName) Then
'''''''        MsgBox "No se ha podido iniciar la sesión." & vbCrLf & _
'''''''               IIf(Err.Number <> 0, "Ha ocurrido el sgte error : " & Err.Description, vbNullString), vbCritical, App.Title
''''        MsgBox TraslateWord(55416) & vbCrLf & TraslateWord(55417) & " :" & vbCrLf & TraslateWord(55418) & vbCrLf & _
''''               IIf(Err.Number <> 0, ParseFormatString(TraslateWord(55439), Err.Description), vbNullString), vbCritical + vbOKOnly, Title
''''        Exit Function
''''    End If
''''    mstrLicenseInfo = objApp.GetLicenseInfo
''''    If Err.Number <> 0 Then
'''''''        MsgBox "Ha ocurrido el sgte error :" & Err.Description & ".", vbCritical, App.Title
''''        GenerateMessage 55439, vbCritical + vbOKOnly, Title, Err.Description
''''        Exit Function
''''    End If
    On Error GoTo TratarError
''''    mstrUserSetting = objApp.GetInitProperties
''''    On Error Resume Next
''''    If Len(mstrUserSetting) > 0 Then
''''        'Obtener el usuario
''''        mstrUsuario = GetPropertyValue(mstrUserSetting, "IDUsuario")
''''        If Len(mstrUsuario) = 0 Then
''''            'MSG="Usuario desconocido."
''''            strMessage = TraslateWord(55316)
''''            blnContinue = False
''''        End If
''''        'Obtener nombre del servidor actual del usuario
''''        mstrServerName = GetPropertyValue(mstrUserSetting, "Server")
''''        If Len(mstrServerName) = 0 Then
''''            'MSG="Servidor de datos desconocido."
''''            strMessage = strMessage & vbCrLf & TraslateWord(55317)
''''            blnContinue = False
''''        End If
''''        'Obtener base de datos actual del usuario
''''        mstrDataBase = GetPropertyValue(mstrUserSetting, "Database")
''''        If Len(mstrDataBase) = 0 Then
''''            'MSG="Base de datos desconocida."
''''            strMessage = strMessage & vbCrLf & TraslateWord(55318)
''''            blnContinue = False
''''        End If
''''        'Obtener nombre de la empresa
''''        mstrEnterpriseName = GetPropertyValue(mstrUserSetting, "Enterprise")
''''        If Len(mstrEnterpriseName) = 0 Then
''''            'MSG="Empresa desconocida."
''''            strMessage = strMessage & vbCrLf & TraslateWord(55319)
''''            blnContinue = False
''''        End If
''''        mlngGrupoUsuario = CInt(GetPropertyValue(mstrUserSetting, "IDGrupo"))
''''        If Err.Number <> 0 Or mlngGrupoUsuario <= 0 Then
''''            'MSG="No se ha encontrado el grupo administrativo del usuario o no es un grupo válido."
''''            strMessage = strMessage & vbCrLf & TraslateWord(55320)
''''            blnContinue = False
''''        End If
''''        mstrUserLanguage = GetPropertyValue(mstrUserSetting, "Languaje")
''''        mstrUserLanguageID = GetPropertyValue(mstrUserSetting, "IDIdioma")
''''        If Len(mstrUserLanguage) = 0 Or Len(mstrUserLanguageID) = 0 Then
''''            'MSG="El idioma del usuario no es correcto."
''''            strMessage = strMessage & vbCrLf & TraslateWord(55321)
''''            blnContinue = False
''''        End If
''''        mstrBackgroundImage = GetPropertyValue(mstrUserSetting, "Imagen")
''''        mblnLoadWeb = CBool(GetPropertyValue(mstrUserSetting, "CargarWeb"))
''''        mstrUserHomePage = GetPropertyValue(mstrUserSetting, "URL")
''''        mstrIDDefaultProgram = GetPropertyValue(mstrUserSetting, "DefaultProgram")
''''        ComputerName
''''        If blnContinue Then
''''            On Error GoTo TratarError
''''            'Obtener el perfil(programas, informes, entidades, pagina de inicio) del usuario
''''            'Obtener todos los programas del usuario
''''            udtPrgms = objApp.GetUserPrograms(mstrUsuario)
''''            Set mrstUserMenu = udtPrgms.Menu
''''            Set mrstUserPrograms = udtPrgms.Programs
''''            Set mrstUserSpecialPrograms = udtPrgms.SpecialPrograms
                Set mrstUserPrograms = GescomUserPrograms()
''''            'Obtener todos los reportes del usuario
''''            Set mrstUserReports = objApp.GetUserReport(vbNullString)
''''            'Obtener todos los enumerados de la aplicación
''''            Enums
''''            'Obtener las acciones globales
''''            Set mrstActions = objApp.GetActions
''''            On Error Resume Next
''''        End If
''''    Else
''''        'MSG ="No se ha encontrado la información correspondiente al usuario: "
''''        strMessage = ParseFormatString(TraslateWord(55315), GetWinUserName)
''''        blnContinue = False
''''    End If
''''    If Not blnContinue Then
''''        'MSG = "El perfil de usuario obtenido es el siguiente:"
''''        '      "Es posible que el usuario no se haya dado de alta como usuario de eXpertis® o que sus datos no sean correctos."
''''        '      "Si tiene permisos de administración entre en la configuración y revise sus datos"
''''        '      "contacte con su administrador de sistemas para solucionar este problema."
''''        MsgBox strMessage & vbCrLf & vbCrLf & _
''''                TraslateWord(55310) & vbCrLf & _
''''                Nz(mstrUserSetting, "(Vacio)") & vbCrLf & _
''''                TraslateWord(55311) & vbCrLf & _
''''                TraslateWord(55312) & ", " & TraslateWord(55313) & ", " & TraslateWord(55314), vbCritical + vbOKOnly, Title
''''    End If
    GetUserSettings = blnContinue
    
fin:
    Exit Function
TratarError:
    GenerateError Err.Number, Err.Description, "clsMDIParentForm.GetUserSettings", Title
    Resume fin
    Resume
End Function

Private Function GescomUserPrograms() As Recordset
Dim rstProgramas As Recordset

    Set rstProgramas = New Recordset
    
    rstProgramas.Fields.Append "IDPrograma", adVarChar, 10
    ''''IDMenu, Orden,
    rstProgramas.Fields.Append "Caption", adVarChar, 255
    rstProgramas.Fields.Append "FormCaption", adVarChar, 255
    ''''Imagen
    rstProgramas.Fields.Append "DescPrograma", adVarChar, 255
    rstProgramas.Fields.Append "TypeLib", adGUID
    rstProgramas.Fields.Append "Tipo", adInteger
    ''''Filename, Version, Fecha, Traduccion, TraduccionForm
    
    rstProgramas.Open
    rstProgramas.AddNew
    rstProgramas.Fields("IDPrograma") = "uscPrendas"
    rstProgramas.Fields("Caption") = "Mantenimiento de tipos de prendas"
    rstProgramas.Fields("FormCaption") = "Mantenimiento de prendas"
    rstProgramas.Fields("DescPrograma") = "MntoPrendas.uscPrendas"
    rstProgramas.Fields("Tipo") = ptyUserControl
    rstProgramas.Update
    
    Set GescomUserPrograms = rstProgramas
    Set rstProgramas = Nothing
    
End Function

Public Function UserPrograms() As Recordset
    Set UserPrograms = mrstUserPrograms
End Function

Public Function UserMenu() As Recordset
    Set UserMenu = mrstUserMenu
End Function
Public Function UserSpecialPrograms() As Recordset
    Set UserSpecialPrograms = mrstUserSpecialPrograms
End Function

Public Function UserReports() As Recordset
    Set UserReports = mrstUserReports
End Function

Public Function GetProgram(ByRef strControlID As String) As Recordset
    Dim strProgID As String
'''    Dim strClsID As String
    Dim strTypeLib As String
    Dim strFileName As String
    Dim strVersion As String
    Dim dtFileDate As Date
    Dim rstProgramInfo As Recordset
    
    Set rstProgramInfo = GetProgramInfo(strControlID)
    If Not rstProgramInfo Is Nothing Then
        If Nz(rstProgramInfo("Tipo"), ptyUserControl) = ptyUserControl Then
            If Not IsEmptyValue(rstProgramInfo("DescPrograma")) Then
                strProgID = rstProgramInfo("DescPrograma")
            End If
'''            If Not IsEmptyValue(rstProgramInfo("ClsID")) Then
'''                strClsID = rstProgramInfo("ClsID")
'''            End If
            If Not IsEmptyValue(rstProgramInfo("TypeLib")) Then
                strTypeLib = rstProgramInfo("TypeLib")
            End If
            If Not IsEmptyValue(rstProgramInfo("FileName")) Then
                strFileName = rstProgramInfo("FileName")
            End If
            If Not IsEmptyValue(rstProgramInfo("Version")) Then
                strVersion = rstProgramInfo("Version")
            End If
            If Not IsEmptyValue(rstProgramInfo("Fecha")) Then
                dtFileDate = rstProgramInfo("Fecha")
            End If
''''            If Not CheckAndDownLoadComponent(strFileName, strTypeLib, strVersion, dtFileDate) Then
''''                Set rstProgramInfo = Nothing
''''            End If
        End If
    End If
    Set GetProgram = rstProgramInfo
End Function
''''
''''Public Sub CheckAndDownLoadEngineComponents()
''''    Dim strTypeLib As String
''''    Dim strFileName As String
''''    Dim strVersion As String
''''    Dim dtFileDate As Date
''''    Dim blnProxy As Boolean
''''    Dim rstComponentInfo As Recordset
''''    Dim objApp As Aplicacion
''''
''''    Set objApp = New Aplicacion
''''    Set rstComponentInfo = objApp.GetEngineRepositorInfo
''''    Set objApp = Nothing
''''    If Not rstComponentInfo Is Nothing Then
''''        If rstComponentInfo.RecordCount > 0 Then
''''            While Not rstComponentInfo.EOF
''''                If Not IsEmptyValue(rstComponentInfo("TypeLib")) Then
''''                    strTypeLib = rstComponentInfo("TypeLib")
''''                End If
''''                If Not IsEmptyValue(rstComponentInfo("FileName")) Then
''''                    strFileName = rstComponentInfo("FileName")
''''                End If
''''                If Not IsEmptyValue(rstComponentInfo("Version")) Then
''''                    strVersion = rstComponentInfo("Version")
''''                End If
''''                If Not IsEmptyValue(rstComponentInfo("Fecha")) Then
''''                    dtFileDate = rstComponentInfo("Fecha")
''''                End If
''''                If Not IsEmptyValue(rstComponentInfo("Proxy")) Then
''''                    blnProxy = rstComponentInfo("Proxy")
''''                End If
''''                CheckAndDownLoadComponent strFileName, strTypeLib, strVersion, dtFileDate, blnProxy
''''                rstComponentInfo.MoveNext
''''            Wend
''''        End If
''''    End If
''''End Sub
''''
''''Private Function GetProxyCurFileName(strFileName As String) As String
''''    Dim strFileNameTmp As String
''''    Dim strFileNameExt As String
''''
''''    GetProxyCurFileName = strFileName
''''    If Not FileExists(strFileName) Then
''''        strFileNameExt = FileExtension(strFileName)
''''        If strFileNameExt = "TLB" Then
''''            strFileNameTmp = ExtractExtension(strFileName) & ".DLL"
''''        Else
''''            strFileNameTmp = ExtractExtension(strFileName) & ".TLB"
''''        End If
''''        If FileExists(strFileNameTmp) Then GetProxyCurFileName = strFileNameTmp
''''    End If
''''End Function
''''
''''Public Function CheckAndDownLoadComponent(strFileName As String, strTypeLib As String, strVersion As String, dtFileDate As Date, Optional blnProxy As Boolean = False) As Boolean
''''    Dim blnInstall As Boolean
''''    Dim blnError As Boolean
''''    Dim strFileNameTmp As String
''''
''''    GetRegisterIdentity
''''    If Len(strTypeLib) > 0 And Len(strFileName) > 0 Then
''''        strFileName = Path & cnObjectPath & cnSlashSeparator & strFileName
''''        strFileNameTmp = strFileName
''''        If blnProxy Then
''''            strFileNameTmp = GetProxyCurFileName(strFileName)
''''        End If
''''        blnInstall = False
''''        If FileExists(strFileNameTmp) Then
''''            'Comprobar que tenemos la última versión del objeto a cargar
''''            'Comprobamos que coincidan los TypeLib GUID y la fecha del archivo
''''            If Not TypeLibIsRegistered(strTypeLib, strVersion) Then
''''                'MSG="La versión actual de strFileName no está actualizada."
''''                '    "¿Desea actualizar ahora?"
''''                blnInstall = GenerateMessage(55322, vbInformation + vbYesNo, Title, strFileNameTmp) = vbYes
''''            Else
''''                '***Chequeo de la fecha de modificación***
''''                'Se descartarán los segundos y milisegundos para evitar los problemas
''''                'de aproximación que realiza Windows en sistemas de archivos FAT.
''''                If Not blnProxy Then
''''                    If dtFileDate <> 0 Then
''''                        If Not CheckFileDateTime(strFileNameTmp, dtFileDate) Then
''''                            blnInstall = GenerateMessage(55322, vbInformation + vbYesNo, Title, strFileNameTmp) = vbYes
''''                        End If
''''                    End If
''''                End If
''''            End If
''''        Else
''''            'MSG="El objeto """ & strFileName & """ no existe."
''''            '    "Se procederá con la instalación."
''''            GenerateMessage 55323, vbInformation + vbOKOnly, Title, strFileNameTmp
''''            blnInstall = True
''''        End If
''''        If blnInstall Then
''''            blnError = Not InstallObject(strFileName, strTypeLib, , blnProxy)
''''            If blnError Then
''''                'MSG="No se ha podido instalar el objeto """ & strFileName & """."
''''                GenerateMessage 55324, vbInformation + vbOKOnly, Title, strFileName
''''            Else
''''                'MSG="El objeto """ & strFileName & """ se ha instalado con éxito."
''''                GenerateMessage 55325, vbInformation + vbOKOnly, Title, strFileName
''''            End If
''''        End If
''''    End If
''''    CheckAndDownLoadComponent = Not blnError
''''End Function

Public Function GetProgramInfo(strControlID As String) As Recordset
    Dim i As Long
    Dim strCriteria As String
    
    Set GetProgramInfo = Nothing
    If Not mrstUserPrograms Is Nothing Then
        mrstUserPrograms.Filter = vbNullString
        i = InStr(1, strControlID, ".", vbTextCompare)
        If i = 0 Then
            strCriteria = "IDPrograma='" & strControlID & "'"
        Else
            strCriteria = "DescPrograma='" & strControlID & "'"
        End If
        mrstUserPrograms.Find strCriteria, , adSearchForward, 1
        If Not mrstUserPrograms.BOF And Not mrstUserPrograms.EOF Then
            Set GetProgramInfo = mrstUserPrograms
        End If
    End If
End Function
''''
''''Public Function GetReport(ByRef strReportName As String, Optional strReportID As String) As Recordset
''''    Dim strTypeLib As String
''''    Dim strFileName As String
''''    Dim blnInstall As Boolean
''''    Dim blnError As Boolean
''''    Dim rstReportInfo As Recordset
''''
''''    If Len(strReportID) > 0 Then
''''        Set rstReportInfo = GetReportInfo(vbNullString, strReportID)
''''    Else
''''        Set rstReportInfo = GetReportInfo(strReportName)
''''    End If
''''    If Not rstReportInfo Is Nothing Then
''''        If Not IsEmptyValue(rstReportInfo("TypeLib")) Then
''''            strTypeLib = rstReportInfo("TypeLib")
''''        End If
''''        If Not IsEmptyValue(rstReportInfo("NombInforme")) Then
''''            strFileName = rstReportInfo("NombInforme")
''''        End If
''''        blnError = False
''''        If Len(strTypeLib) > 0 And Len(strFileName) > 0 Then
''''            strFileName = Path & cnReportPath & cnSlashSeparator & strFileName
''''            blnInstall = False
''''            If FileExists(strFileName) Then
''''                'Comprobar que tenemos la última versión del objeto a cargar
''''                'Comprobamos que coincida la fecha de actualización
''''                If Not IsEmptyValue(rstReportInfo("Fecha")) Then
''''                    If Not CheckFileDateTime(strFileName, rstReportInfo("Fecha")) Then
''''                        If GenerateMessage(55322, vbInformation + vbYesNo, Title, strFileName) = vbYes Then
''''                            blnInstall = True
''''                        End If
''''                    End If
''''                End If
''''            Else
''''                GenerateMessage 55323, vbInformation + vbOKOnly, Title, strFileName
''''                blnInstall = True
''''            End If
''''            If blnInstall Then
''''                blnError = Not InstallObject(strFileName, strTypeLib, False)
''''                If blnError Then
''''                    GenerateMessage 55324, vbInformation + vbOKOnly, Title, strFileName
''''                Else
''''                    GenerateMessage 55325, vbInformation + vbOKOnly, Title, strFileName
''''                End If
''''            End If
''''        End If
'''''        If blnError Then Set rstReportInfo = Nothing
''''    End If
''''    Set GetReport = rstReportInfo
''''End Function

Public Function GetReportInfo(strReportName As String, Optional strReportID As String) As Recordset
    Set GetReportInfo = Nothing
    If Not mrstUserReports Is Nothing Then
        mrstUserReports.Filter = vbNullString
        If Len(strReportID) > 0 Then
            mrstUserReports.Find "IDInforme='" & strReportID & "'", , adSearchForward, 1
        Else
            mrstUserReports.Find "NombInforme='" & strReportName & "'", , adSearchForward, 1
        End If
        If Not mrstUserReports.BOF And Not mrstUserReports.EOF Then
            Set GetReportInfo = mrstUserReports
        End If
    End If
End Function
''''
''''Public Sub UnloadChildFormReports(Optional strDocName As String)
''''    Dim objReport As clsReport
''''    Dim i As Long
''''    Dim objEmbeddedCtrl As IEmbeddedControl
''''    Static blnCanReentry As Boolean
''''
''''    If Not blnCanReentry Then
''''        blnCanReentry = True
''''        For i = mcolReports.Count To 1 Step -1
''''            Set objReport = mcolReports.Item(i)
''''            Set objEmbeddedCtrl = objReport.FormViewer
''''    ''        If Len(strDocName) = 0 Or UCase(objReport.ActiveChildForm.Name)= UCase(strDocName) Then
''''            If Len(strDocName) = 0 Or InStr(UCase(objEmbeddedCtrl.Name), UCase(strDocName)) Then
''''                objReport.CloseReport
''''                mcolReports.Remove i
''''            End If
''''            Set objEmbeddedCtrl = Nothing
''''            Set objReport = Nothing
''''        Next i
''''        blnCanReentry = False
''''    End If
''''End Sub
''''
''''Public Function InstallObject(strFileName As String, strTypeLib As String, Optional blnCOMObject As Boolean = True, Optional blnProxy As Boolean = False) As Boolean
''''    Dim blnIsFormLoad As Boolean
''''    Dim arrAdminUser() As String
''''    Dim strUserName As String
''''    Dim strPassword As String
''''    Dim strDomain As String
''''    Dim strRegisterPath As String
''''
''''    On Error GoTo TratarError
''''
''''    InstallObject = False
''''    If Len(strTypeLib) > 0 Then
''''        strRegisterPath = Path & cnObjectPath & cnSlashSeparator & "eXpertisRegister.exe"
''''        If Not FileExists(strRegisterPath) Then
''''            GenerateMessage 55464, vbCritical + vbOKOnly, Title, strRegisterPath
''''            InstallObject = False
''''            Exit Function
''''        End If
''''        frmProgress.Show
''''        blnIsFormLoad = True
''''        If Len(mstrAdminUser) > 0 Then
''''            arrAdminUser = Split(mstrAdminUser, ";")
''''            On Error Resume Next
''''            strUserName = Mid$(arrAdminUser(0), InStr(arrAdminUser(0), "=") + 1)
''''            strPassword = Mid$(arrAdminUser(1), InStr(arrAdminUser(1), "=") + 1)
''''            strDomain = Mid$(arrAdminUser(2), InStr(arrAdminUser(2), "=") + 1)
''''            On Error GoTo TratarError
''''        End If
''''        If blnCOMObject Then
''''            'MSG="Desregistrando el componente " & strFileName
''''            frmProgress.ProgressTitle ParseFormatString(TraslateWord(55326), strFileName)
''''            DoEvents
''''            'Desregistrar el objeto actual
''''            If blnProxy Then
''''                RunAsUser strUserName, strPassword, strDomain, strRegisterPath & " -u " & strFileName & " " & mstrAppServerName, Path
''''            Else
''''                RunAsUser strUserName, strPassword, strDomain, strRegisterPath & " -u " & strFileName, Path
''''            End If
''''        End If
''''        'Descargar del Repositorio
''''        'MSG="Descargando el componente " & strFileName
''''        frmProgress.ProgressTitle ParseFormatString(TraslateWord(55327), strFileName)
''''        frmProgress.PlayProgress
''''        DoEvents
''''        InstallObject = DownLoadObject(strFileName, strTypeLib)
''''        frmProgress.StopProgress
''''        DoEvents
''''        If blnCOMObject And InstallObject Then
''''            'Registrar el nuevo objeto
''''            'MSG="Registrando el componente " & strFileName
''''            frmProgress.ProgressTitle ParseFormatString(TraslateWord(55328), strFileName)
''''            DoEvents
''''            If blnProxy Then
''''                InstallObject = RunAsUser(strUserName, strPassword, strDomain, strRegisterPath & " -r " & strFileName & " " & mstrAppServerName, Path) = 0
''''            Else
''''                InstallObject = RunAsUser(strUserName, strPassword, strDomain, strRegisterPath & " -r " & strFileName, Path) = 0
''''            End If
''''        End If
''''    End If
''''fin:
''''    If blnIsFormLoad Then Unload frmProgress
''''    Exit Function
''''TratarError:
''''    GenerateError Err.Number, Err.Description, "clsMDIParentForm.InstallObject", Title
''''    Resume fin
''''End Function
''''
''''Private Function DownLoadObject(ByVal strFileName As String, ByVal strTypeLib As String) As Boolean
''''    Dim rstProgramRepositorInfo As Recordset
''''    Dim bytBuffer() As Byte
''''    Dim FwApp As Aplicacion
''''    Dim strBackFileName As String
''''    Dim lngPosition As String
''''
''''    On Error GoTo TratarError
''''
''''    DownLoadObject = False
''''    If Len(strTypeLib) > 0 Then
''''        'Eliminar los {}
''''        lngPosition = InStr(1, strTypeLib, "{")
''''        If lngPosition > 0 Then strTypeLib = Mid$(strTypeLib, lngPosition + 1)
''''        lngPosition = InStrRev(strTypeLib, "}")
''''        If lngPosition > 0 Then strTypeLib = Left$(strTypeLib, lngPosition - 1)
''''        Set FwApp = New Aplicacion
''''        Set rstProgramRepositorInfo = FwApp.GetProgramRepositorInfo(strTypeLib)
''''        Set FwApp = Nothing
''''        If Not rstProgramRepositorInfo Is Nothing Then
''''            If rstProgramRepositorInfo.RecordCount > 0 Then
''''                If Not IsEmptyValue(rstProgramRepositorInfo("FileName")) Then
''''                    strBackFileName = ExtractExtension(strFileName) & ".back"
''''                    'Eliminar copia si existe
''''                    If FileExists(strBackFileName) Then Kill strBackFileName
''''                    'Renombrar original si existe(hacer copia)
''''                    If FileExists(strFileName) Then Name strFileName As strBackFileName
''''                    bytBuffer = rstProgramRepositorInfo("Imagen").GetChunk(rstProgramRepositorInfo("Longitud"))
''''                    Open strFileName For Binary Access Write Lock Read Write As #1
''''                    Put #1, , bytBuffer
''''                    Close #1
''''                    SetFileDateTime strFileName, rstProgramRepositorInfo("Fecha")
''''                    'Eliminar copia si existe
''''                    If FileExists(strBackFileName) Then Kill strBackFileName
''''                    DownLoadObject = True
''''                End If
''''            End If
''''        End If
''''    End If
''''    Exit Function
''''TratarError:
''''    'En caso de error restaurar la copia
''''    On Error Resume Next
''''    If Len(strFileName) > 0 And FileExists(strFileName) Then Kill strFileName
''''    If Len(strBackFileName) > 0 And FileExists(strBackFileName) Then Name strBackFileName As strFileName
''''    On Error GoTo 0
''''    Set rstProgramRepositorInfo = Nothing
''''End Function

Public Function CommonDialog() As Object
    Set CommonDialog = mfrmMain.CommonDialog
End Function
'''
'''Private Sub GetRegisterIdentity()
'''    Dim objApp As Aplicacion
'''
'''    If Not mblnRegisterIdentityWasRetrive Then
'''        mblnRegisterIdentityWasRetrive = True
'''        Set objApp = New Aplicacion
'''        If Not objApp Is Nothing Then
'''            mstrAdminUser = objApp.GetAdminUser
'''            mstrAppServerName = objApp.GetAppServerName
'''            Set objApp = Nothing
'''        End If
'''    End If
'''End Sub

Public Property Get OpenningProgramName() As String
    OpenningProgramName = mstrOpenningProgramName
End Property

Public Property Let OpenningProgramName(ByVal vNewValue As String)
    mstrOpenningProgramName = vNewValue
End Property

Public Property Get GlobalActions() As Recordset
    Set GlobalActions = mrstActions
End Property
