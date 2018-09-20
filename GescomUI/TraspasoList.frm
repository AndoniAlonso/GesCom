VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form TraspasoList 
   Caption         =   "Lista de traspasos de artículos entre almacenes"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TraspasoList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epAlmacenOrigen 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   873
      End
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
   Begin GridEX20.GridEX jgrdItems 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      AutomaticSort   =   -1  'True
      ScrollToolTips  =   -1  'True
      RecordNavigatorString=   "Registro:|de"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ShowEmptyFields =   0   'False
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Ninguno"
      MultiSelect     =   -1  'True
      HeaderStyle     =   2
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      GroupByBoxInfoText=   "Arrastre una columna aquí para agrupar."
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   6
      FormatStyle(1)  =   "TraspasoList.frx":08CA
      FormatStyle(2)  =   "TraspasoList.frx":09F2
      FormatStyle(3)  =   "TraspasoList.frx":0AA2
      FormatStyle(4)  =   "TraspasoList.frx":0B56
      FormatStyle(5)  =   "TraspasoList.frx":0C2E
      FormatStyle(6)  =   "TraspasoList.frx":0CE6
      ImageCount      =   0
      PrinterProperties=   "TraspasoList.frx":0DC6
   End
End
Attribute VB_Name = "TraspasoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmTraspaso As TraspasoEdit
'Private frmTraspasoAutomatico As AlbaranAutomaticoEdit
Private objTraspaso As Traspaso
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
Private strLayout As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim jcoltemp As JSColumn
    Dim vlSituaciones As JSValueList

    
    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next

    With jgrdItems
        ' Propiedades de la 1ª columna, campo clave.
        .Columns("TraspasoID").ColumnType = jgexIcon
        '.Columns("TraspasoID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("TraspasoID").FetchIcon = True
        .Columns("TraspasoID").Caption = vbNullString
        .Columns("TraspasoID").Visible = True
        .Columns("TraspasoID").ColPosition = 1
        .Columns("TraspasoID").Width = 330
        
        FormatoJColumn .Columns("NombreAlmacenOrigen"), 2, "Origen", , ColumnSize(15)
        FormatoJColumn .Columns("NombreAlmacenDestino"), 3, "Destino", , ColumnSize(15)
        FormatoJColumn .Columns("FechaAlta"), 4, "Fecha", , ColumnSize(12)
        FormatoJColumn .Columns("Situacion"), 5, "Situación", , ColumnSize(8)
        FormatoJColumn .Columns("FechaTransito"), 6, "Fecha envío", , ColumnSize(12)
        FormatoJColumn .Columns("FechaRecepcion"), 7, "Fecha recepción", , ColumnSize(12)
        FormatoJColumn .Columns("Cantidad"), 8, "Cantidad", True, ColumnSize(6)
        jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter
        
        Set jcoltemp = jgrdItems.Columns("Situacion")
        jcoltemp.HasValueList = True
        Set vlSituaciones = jcoltemp.ValueList
        vlSituaciones.Add "0 ", "Pendiente de enviar", 1
        vlSituaciones.Add "1 ", "Enviado", 2
        vlSituaciones.Add "2 ", "Recepcionado", 3
        
        jcoltemp.ColumnType = jgexTextOnly
        
        Dim jscolAgrupar As JSColumn
        Set jscolAgrupar = .Columns.Add(, jgexText, jgexEditNone)
        jscolAgrupar.AggregateFunction = jgexCount
        jscolAgrupar.Width = 0
        jscolAgrupar.Caption = "(Agrupar todos)"
        jscolAgrupar.GroupEmptyStringCaption = "(Agrupar todos)"

        .Groups.Add jscolAgrupar.Index, jgexSortAscending
        
    End With
    
    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout

End Sub



Private Sub Form_Load()
    Dim objButton As Button
    
    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable
    
    LoadImages Me.tlbHerramientas
    
    ' Añadimos los botones especificos de esta opción:
    ' - Imprimir los traspasos.
    ' - Enviar un traspaso.
    ' - Recepcionar un traspaso.
    ' - EnviarYRecepcionar un traspaso. (directamente sin pasar por almacen en tránsito).

'    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
'    Set objButton = Me.tlbHerramientas.Buttons.Add(, "EnviarTraspaso", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("EnviarTraspaso").Key)
'    objButton.ToolTipText = "Marcar el traspaso como enviado"
'
'    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
'    Set objButton = Me.tlbHerramientas.Buttons.Add(, "RecepcionarTraspaso", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("RecepcionarTraspaso").Key)
'    objButton.ToolTipText = "Marcar el traspaso como recepcionado"
'
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "EnviarYRecepcionarTraspaso", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("RecepcionarTraspaso").Key)
    objButton.ToolTipText = "Marcar el traspaso como enviado Y recepcionado"

    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epAlmacenOrigen.Initialize 1, "Almacenes", "AlmacenID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.Terminal.AlmacenID, 0
    epAlmacenOrigen.LoadControl "Almacen"
    
'    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
'    epTemporada.LoadControl "Temporada"
    
'    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
'    epCliente.LoadControl "Cliente"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("Traspaso").Picture
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("EnviarTraspaso").Picture
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("RecepcionarTraspaso").Picture
    
    strLayout = vbNullString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub jgrdItems_DblClick()
    
    Call EditSelected
    
End Sub

Private Sub jgrdItems_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    
    mlngColumn = Column.Index
    'txtQuickSearch.ToolTipText = "Búsqueda rápida en " & Column.Caption

End Sub

Private Sub jgrdItems_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    Dim Rs As Recordset

    If ColIndex = 1 Then
        Set Rs = jgrdItems.ADORecordset
        Rs.Bookmark = RowBookmark
        Select Case Rs.Fields("Situacion")
        Case enuTraspasoSituacionAlta
            IconIndex = 1
        Case enuTraspasoSituacionEnTransito
            IconIndex = 2
        Case enuTraspasoSituacionRecepcionado
            IconIndex = 3
        Case Else
            IconIndex = 0
        End Select
    End If

End Sub

Private Sub jgrdItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    ElseIf KeyCode = 116 Then   ' Tecla F5
        UpdateListView
    ElseIf KeyCode = 114 Then   ' Tecla F3
        QuickSearch
    End If
        
End Sub

Private Sub jgrdItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        jgrdItems.Enabled = False
        jgrdItems.Enabled = True
    End If
    
End Sub

Public Sub DeleteSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objTraspaso = New Traspaso
                    objTraspaso.Load mlngID
                    objTraspaso.BeginEdit
                    objTraspaso.Delete
                    objTraspaso.ApplyEdit
                    Set objTraspaso = Nothing
                End If
            End If
        Next
        UpdateListView SentenciaSQL
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    Dim strClausulaWhere As String

    On Error GoTo ErrorManager
    
'    If epAlmacenOrigen.ClausulaWhere <> vbNullString Then
'        If strClausulaWhere = vbNullString Then
'            strClausulaWhere = epAlmacenOrigen.ClausulaWhere
'        Else
'            strClausulaWhere = strClausulaWhere & " AND AlmacenOrigenID = " & epAlmacenOrigen.SelectedKey
'        End If
'    End If
    
    Set objRecordList = New RecordList
    
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from vTraspasos", strClausulaWhere)
    Set objRecordList = Nothing
    
    strLayout = jgrdItems.LayoutString
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If jgrdItems.SelectedItems.Count >= 5 Then
        Respuesta = MostrarMensaje(MSG_OPEN)
        
        If Respuesta = vbYes Then
            EditItems
        End If
    Else
        EditItems
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditItems()
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmTraspaso = New TraspasoEdit
                Set objTraspaso = New Traspaso
                objTraspaso.Load mlngID
                frmTraspaso.Component objTraspaso
                frmTraspaso.Show
                Set objTraspaso = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmTraspaso = New TraspasoEdit
    Set objTraspaso = New Traspaso
    frmTraspaso.Component objTraspaso
    frmTraspaso.Show
    Set objTraspaso = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   If View = jgexCard Then jgrdItems.View = jgexCard
   If View = 3 Then jgrdItems.View = jgexTable
   
   
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = True
   
End Function

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Nuevo"
            NewObject
        Case Is = "Abrir"
            EditSelected
        Case Is = "Imprimir"
            Imprimir
        Case Is = "Eliminar"
            DeleteSelected
        Case Is = "Actualizar"
            UpdateListView SentenciaSQL
        Case Is = "Buscar"
            ResultSearch
        Case Is = "IconosGrandes"
            SetListViewStyle (lvwIcon)
        Case Is = "IconosPequeños"
            SetListViewStyle (lvwSmallIcon)
        Case Is = "Lista"
            SetListViewStyle (lvwList)
        Case Is = "Detalle"
            SetListViewStyle (lvwReport)
        Case Is = "QuickSearch"
            QuickSearch
        Case Is = "GroupBy"
            frmGroupBy.GroupGrid jgrdItems
        Case Is = "ShowFields"
            frmShowfields.ShowFields jgrdItems
        Case Is = "Sort"
            frmSort.SortGrid jgrdItems
        Case Is = "Cerrar"
            Unload Me
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
'        Case Is = "EnviarTraspaso"
'            EnviarTraspaso
'        Case Is = "RecepcionarTraspaso"
'            RecepcionarTraspaso
        Case Is = "EnviarYRecepcionarTraspaso"
            EnviarYRecepcionarTraspaso
    End Select
        
End Sub

Private Sub Form_Resize()

    GridEX_Resize jgrdItems, Me, frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vTraspasos"
    frmBusqueda.Component mobjBusqueda
    frmBusqueda.Show vbModal
    
    If frmBusqueda.mflgAplicarFiltro Then
        Set mobjBusqueda = frmBusqueda.Consulta
        SentenciaSQL = frmBusqueda.SentenciaSQL
        UpdateListView (SentenciaSQL)
    ElseIf frmBusqueda.lvwConsultaItems.ListItems.Count = 0 Then
        SentenciaSQL = vbNullString
    End If
    
    Unload frmBusqueda

End Sub

Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de albaranes de venta"
        .FooterString(jgexHFLeft) = Now
        .BottomMargin = 720
        .TopMargin = 720
        .LeftMargin = 720
        .RightMargin = 720
        .TransparentBackground = True
        'Right footer is set in the BeforePrintPage to indicate page number
    End With
    Load frmPreview
    frmPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    jgrdItems.PrintPreview frmPreview.grPrev, False  '(Check4.value = vbChecked)
    frmPreview.Show vbModal
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    
End Sub

Private Sub jgrdItems_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & vbCrLf & " de " & nPages
    
End Sub

'Public Sub PrintSelected()
'Dim Respuesta As VbMsgBoxResult
'Dim objPrintAlbaran As PrintAlbaran
'Dim frmPrintOptions As frmPrint
'Dim simTemp As JSSelectedItem
'Dim RowData As JSRowData
'
'    On Error GoTo ErrorManager
'
'    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
'
'    ' aquí hay que avisar de si realmente queremos imprimir los documentos
'    Respuesta = MostrarMensaje(MSG_DOCUMENTO)
'
'    If Respuesta = vbYes Then
'        Set frmPrintOptions = New frmPrint
'        frmPrintOptions.Flags = ShowCopies_po + ShowPrinter_po
'        frmPrintOptions.Copies = 1
'        frmPrintOptions.Show vbModal
'        ' salir de la opcion si no pulsa "imprimir"
'        If Not frmPrintOptions.PrintDoc Then
'            Unload frmPrintOptions
'            Set frmPrintOptions = Nothing
'            Exit Sub
'        End If
'
'        For Each simTemp In jgrdItems.SelectedItems
'            If simTemp.RowType = jgexRowTypeRecord Then
'                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
'                mlngID = RowData(1)
'                If mlngID > 0 Then
'                    Set objTraspaso = New Traspaso
'                    Set objPrintAlbaran = New PrintAlbaran
'
'                    objTraspaso.Load mlngID
'
'                    objPrintAlbaran.PrinterNumber = frmPrintOptions.PrinterNumber
'                    objPrintAlbaran.Copies = frmPrintOptions.Copies
'                    objPrintAlbaran.Component objTraspaso
'
'                    objPrintAlbaran.PrintObject
'
'                    Set objPrintAlbaran = Nothing
'                    Set objTraspaso = Nothing
'                End If
'            End If
'        Next
'        Unload frmPrintOptions
'        Set frmPrintOptions = Nothing
'    End If
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub

'Private Sub EnviarTraspaso()
'    Dim Respuesta As VbMsgBoxResult
'
'    On Error GoTo ErrorManager
'
'    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
'
'    ' aquí hay que avisar de si realmente queremos abrirlos todos
'    ' si el número es mayor que 5
'    If jgrdItems.SelectedItems.Count >= 5 Then
'        Respuesta = MostrarMensaje(MSG_OPEN)
'
'        If Respuesta = vbYes Then
'            EnviarTraspasos
'        End If
'    Else
'        EnviarTraspasos
'    End If
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
'Private Sub EnviarTraspasos()
'    Dim simTemp As JSSelectedItem
'    Dim RowData As JSRowData
'
'    For Each simTemp In jgrdItems.SelectedItems
'        If simTemp.RowType = jgexRowTypeRecord Then
'            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
'            mlngID = RowData(1)
'            If mlngID > 0 Then
'                Set objTraspaso = New Traspaso
'                objTraspaso.Load mlngID
'                objTraspaso.Enviar
'                Set objTraspaso = Nothing
'            End If
'        End If
'    Next
'
'    UpdateListView
'
'End Sub
'
'Private Sub RecepcionarTraspaso()
'    Dim Respuesta As VbMsgBoxResult
'
'    On Error GoTo ErrorManager
'
'    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
'
'    ' aquí hay que avisar de si realmente queremos abrirlos todos
'    ' si el número es mayor que 5
'    If jgrdItems.SelectedItems.Count >= 5 Then
'        Respuesta = MostrarMensaje(MSG_OPEN)
'
'        If Respuesta = vbYes Then
'            RecepcionarTraspasos
'        End If
'    Else
'        RecepcionarTraspasos
'    End If
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
'Private Sub RecepcionarTraspasos()
'    Dim simTemp As JSSelectedItem
'    Dim RowData As JSRowData
'
'    For Each simTemp In jgrdItems.SelectedItems
'        If simTemp.RowType = jgexRowTypeRecord Then
'            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
'            mlngID = RowData(1)
'            If mlngID > 0 Then
'                Set objTraspaso = New Traspaso
'                objTraspaso.Load mlngID
'                objTraspaso.Recepcionar
'                Set objTraspaso = Nothing
'            End If
'        End If
'    Next
'
'    UpdateListView
'
'End Sub
'
Private Sub EnviarYRecepcionarTraspaso()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub

    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If jgrdItems.SelectedItems.Count >= 5 Then
        Respuesta = MostrarMensaje(MSG_OPEN)

        If Respuesta = vbYes Then
            EnviarYRecepcionarTraspasos
        End If
    Else
        EnviarYRecepcionarTraspasos
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub EnviarYRecepcionarTraspasos()
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData

    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set objTraspaso = New Traspaso
                objTraspaso.Load mlngID
                objTraspaso.EnviarYRecepcionar
                Set objTraspaso = Nothing
            End If
        End If
    Next

    UpdateListView

End Sub

Private Sub jgrdItems_ColButtonClick(ByVal ColIndex As Integer)
Dim objCliente As Cliente
Dim frmClienteEdit As ClienteEdit
Dim rdCurrent As JSRowData
Dim lngClienteID As Long

    If ColIndex = jgrdItems.Columns("NombreCliente").Index Then
        Set objCliente = New Cliente
        Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
        lngClienteID = rdCurrent.Value(jgrdItems.Columns("ClienteID").Index)
        objCliente.Load lngClienteID
        
        Set frmClienteEdit = New ClienteEdit
        frmClienteEdit.Component objCliente
        frmClienteEdit.Show
        
        Set objCliente = Nothing
        Set frmClienteEdit = Nothing
    End If
    
End Sub


