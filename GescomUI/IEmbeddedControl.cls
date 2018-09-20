VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEmbeddedControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum AppToolBarCmd
    tbcNone = 0
    tbcGoFirst
    tbcGoPrevious
    tbcGoNext
    tbcGoLast
    tbcLastUsed
    tbcSearch
    tbcAdd
    tbcDelete
    tbcUpdate
    tbcCancelUpdate
    tbcPrintReport
    tbcCloseForm
    tbcCopy
    tbcCut
    tbcPaste
    tbcHelp
    tbcExecuteProcess
    tbcExecuteQuery
    tbcClearFilter
    tbcCheckAll
    tbcUncheckAll
    tbcShowAllCheckedItems
    tbcDuplicate
    tbcExportTo
    tbcExportToADTG
    tbcExportToXML
    tbcExportToExcel
    tbcExportToCSV
    tbcSetRestrictor
    tbcRefresh
    tbcGotoRecord
End Enum

Public Enum fwUpdateMode
    umNavigation = 0
    umAddNew = 1
    umUpdate = 2
    umRetrive = 3
End Enum

Public Enum fwToolBarEvents
    tbeButtonClick = 0
    tbeButtonDropDown = 1
    tbeMenuClick = 2
End Enum

Public Enum FieldSecurity
    fsAllRights = 0
    fsHidden
    fsLocked
End Enum

Public Enum enuProgramTypes
    ptyUserControl = 0
    ptyWebPage = 1
End Enum

Public Property Get ProgID() As String
End Property

Public Property Let ProgID(strProgID As String)
End Property

Public Property Get Name() As String
End Property

Public Property Let Name(strName As String)
End Property

Public Property Get MDIChildForm() As Object
End Property

Public Property Set MDIChildForm(NewForm As Object)
End Property

Public Property Get Enabled() As Boolean
End Property

Public Property Let Enabled(blnValue As Boolean)
End Property

Public Sub OpenDoc(Optional strCriteria As String = "", Optional Parameters As String = "")
End Sub

Public Sub QueryCloseDoc(blnCancel As Boolean, lngCloseMode As Long)
End Sub

Public Sub CloseDoc(blnCancel As Boolean)
End Sub

Public Sub ActivateDoc()
End Sub

Public Sub DesactivateDoc()
End Sub

Public Sub RefreshDoc(Optional strCriteria As String = "", Optional Parameters As String = "")
End Sub

Public Property Get Caption() As String
End Property

Public Property Let Caption(strProgID As String)
End Property

Public Property Get LocaleID() As String
End Property

Public Property Let LocaleID(strLocaleID As String)
End Property

''''Public Property Get MaintenaceObject() As IMntoDoc
''''End Property
''''
''''Public Property Set MaintenaceObject(objMntoDoc As IMntoDoc)
''''End Property
''''
Public Sub ProcessToolBarEvents(intCmdTag As Integer, intToolBarEvent As Integer, Optional lngButtonIndex As Long = -1)
End Sub

Public Property Get Controls() As Object
End Property

Public Property Set Controls(ctrControls As Object)
End Property

Public Function GetReportFormulas(strReportName As String, Cancel As Boolean, ByVal blnFiredByReport As Boolean) As String
End Function

Public Function GetReportParameters(strReportName As String, Cancel As Boolean, ByVal blnFiredByReport As Boolean) As String
End Function

Public Function GetReportSelectionCriteria(strReportName As String, Cancel As Boolean, ByVal blnFiredByReport As Boolean) As String
End Function

Public Function GetReportDataSource(strReportName As String, Cancel As Boolean, ByVal blnFiredByReport As Boolean) As Recordset
End Function

Public Sub GetReportOptions(strReportName As String, Cancel As Boolean, lngWindowState As Long, lngDestination As Long, blnPromptUser As Boolean, lngNumberOfCopy As Long, strDiskFileName As String, strMailToList As String)
End Sub

''''Public Sub SetReportProperties(ByRef objReport As clsReport, Cancel As Boolean)
''''End Sub
''''
Public Property Get LastUsedDescFields() As String
End Property

Public Property Let LastUsedDescFields(strLastUsedDescFields As String)
End Property

Public Property Get ParentDoc() As String
End Property

Public Property Let ParentDoc(strParentDoc As String)
End Property

Public Property Get ChildDoc() As String
End Property

Public Property Let ChildDoc(strChildDoc As String)
End Property

Public Sub ShutDown()
End Sub

Public Property Get Object() As Object
End Property

Public Property Set Object(objUserControl As Object)
End Property

Public Property Get AutoResize() As Boolean
End Property

Public Property Let AutoResize(blnValue As Boolean)
End Property

''''Public Property Get ProgramDescriptor() As clsProgramDescriptor
''''End Property
''''
