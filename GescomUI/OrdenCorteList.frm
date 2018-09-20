VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OrdenCorteList 
   Caption         =   "Lista de órdenes de corte"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   Icon            =   "OrdenCorteList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwItems 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "OrdenCorteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmOrdenCorte As OrdenCorteEdit
Private objOrdenCorte As OrdenCorte
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
'Private strLayout As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim itmList As ListItem
    Dim dblCantidad As Double
    
    dblCantidad = 0
    While Not mrsRecordList.EOF
        Set itmList = _
            lvwItems.ListItems.Add(Key:= _
            Format$(mrsRecordList("OrdenCorteID")) & " K")

        With itmList
            .Text = FormatoCantidad(mrsRecordList("Numero"))
            .SubItems(1) = FormatoFecha(mrsRecordList("Fecha"))
            .SubItems(2) = Trim(mrsRecordList("Nombre") & vbNullString)
            .SubItems(3) = Trim(mrsRecordList("Observaciones") & vbNullString)
            .SubItems(4) = IIf(mrsRecordList("FechaCorte") = "0:00:00", vbNullString, FormatoFecha(mrsRecordList("FechaCorte")))
            .SubItems(5) = FormatoCantidad(mrsRecordList("Cantidad"))
            dblCantidad = dblCantidad + mrsRecordList("Cantidad")
            
            .Icon = GescomMain.mglIconosGrandes.ListImages("Corte").Key
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Corte").Key
        End With

        mrsRecordList.MoveNext
    Wend
   
    Set itmList = _
        lvwItems.ListItems.Add(Key:="0 KTOTAL")
    
    With itmList
        .Text = "TOTAL"
        .SubItems(5) = FormatoCantidad(dblCantidad)
    End With
    

End Sub

Private Sub Form_Load()
    Dim objButton As Button

    Me.Move 0, 0
    lvwItems.ColumnHeaders.Add , , "Número", ColumnSize(10)
    lvwItems.ColumnHeaders.Add , , "Fecha", ColumnSize(10)
    lvwItems.ColumnHeaders.Add , , "Artículo", ColumnSize(20)
    lvwItems.ColumnHeaders.Add , , "Observaciones", ColumnSize(20)
    lvwItems.ColumnHeaders.Add , , "Fecha Corte", ColumnSize(10)
    lvwItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(10), vbRightJustify
    
    lvwItems.Icons = GescomMain.mglIconosGrandes
    lvwItems.SmallIcons = GescomMain.mglIconosPequeños
    
    LoadImages Me.tlbHerramientas
    
    ' Añadimos los botones especificos de esta opción:
    ' - Actualizar la orden de corte.
    ' - Imprimir etiquetas de órdenes de corte.
    ' - Consulta de prevision de etiquetas.
    ' - Imprimir documento orden de corte
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Corte", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Corte").Key)
    objButton.ToolTipText = "Actualizar los materiales y artículos de las órdenes de corte"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Etiquetas", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Etiqueta").Key)
    objButton.ToolTipText = "Etiquetar órdenes de corte"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "PrevEtiquetas", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrevEtiqueta").Key)
    objButton.ToolTipText = "Prevision de etiquetas"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Documento", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("PrintDocument").Key)
    objButton.ToolTipText = "Imprimir las órdenes de corte seleccionadas"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
End Sub

''''???? para liberar memoria ?
Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub lvwItems_DblClick()
    
    Call EditSelected
    
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    End If
        
End Sub

Private Sub lvwItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        lvwItems.Enabled = False
        lvwItems.Enabled = True
    End If
    
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    
    ListView_ColumnClick lvwItems, ColumnHeader
    mlngColumn = ColumnHeader.Index
       
End Sub

Public Sub DeleteSelected()
    
    On Error GoTo ErrorManager
   
    Dim i As Integer
    Dim Respuesta As VbMsgBoxResult
    
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For i = lvwItems.ListItems.Count To 1 Step -1
            If lvwItems.ListItems(i).Selected = True Then
                mlngID = Val(lvwItems.ListItems(i).Key)
                If mlngID > 0 Then
                    Set objOrdenCorte = New OrdenCorte
                    objOrdenCorte.Load mlngID, GescomMain.objParametro.Moneda
                    objOrdenCorte.BeginEdit GescomMain.objParametro.Moneda
                    objOrdenCorte.Delete
                    objOrdenCorte.ApplyEdit
                    Set objOrdenCorte = Nothing
                    lvwItems.ListItems.Remove (i)
                End If
            End If
        Next i
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    
    On Error GoTo ErrorManager
    
    lvwItems.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from vOrdenesCorte", _
                        "TemporadaID = " & GescomMain.objParametro.TemporadaActualID & " AND " & _
                        "EmpresaID = " & GescomMain.objParametro.EmpresaActualID & _
                        IIf(strWhere = vbNullString, vbNullString, " AND " & strWhere))
    Set objRecordList = Nothing
    
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
   
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If NumeroSeleccionados(lvwItems) >= 5 Then
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

    Dim i As Integer
    
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmOrdenCorte = New OrdenCorteEdit
                Set objOrdenCorte = New OrdenCorte
                objOrdenCorte.Load mlngID, GescomMain.objParametro.Moneda
                frmOrdenCorte.Component objOrdenCorte
                frmOrdenCorte.Show
                Set objOrdenCorte = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub ActualizarSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos actualizar
    Respuesta = MostrarMensaje(MSG_ACTUALIZAR_ORDEN)
        
    Screen.MousePointer = vbHourglass
    
    If Respuesta = vbYes Then ActualizarOrdenCorte
    
    ' aquí hay que avisar de que la actualización ha ido OK
    Respuesta = MostrarMensaje(MSG_PROCESO_OK)
    
    UpdateListView SentenciaSQL
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub ActualizarOrdenCorte()
    Dim i As Integer
    
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmOrdenCorte = New OrdenCorteEdit
                Set objOrdenCorte = New OrdenCorte
                objOrdenCorte.Load mlngID, GescomMain.objParametro.Moneda
                objOrdenCorte.BeginEdit (GescomMain.objParametro.Moneda)
                objOrdenCorte.actualizar
                objOrdenCorte.ApplyEdit
                Set objOrdenCorte = Nothing
            End If
        End If
    Next i

End Sub

Public Sub NewObject()

    Set frmOrdenCorte = New OrdenCorteEdit
    Set objOrdenCorte = New OrdenCorte
    frmOrdenCorte.Component objOrdenCorte
    frmOrdenCorte.Show
    Set objOrdenCorte = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   
    lvwItems.View = View
   
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
        Case Is = "Cerrar"
            Unload Me
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
        Case Is = "Corte"
            ActualizarSelected
        Case Is = "Etiquetas"
            EtiquetarSelected
        Case Is = "PrevEtiquetas"
            PrevEtiquetas
        Case Is = "Documento"
            PrintSelected
    End Select
        
End Sub

Private Sub Form_Resize()

    ListView_Resize lvwItems, Me

End Sub

Public Sub QuickSearch()
    
    ListviewQuickSearch lvwItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vOrdenesCorte"
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

Public Sub EtiquetarSelected()
    Dim objEtiquetas As Etiquetas
    Dim frmEtiquetas As EtiquetasEdit
    Dim objEtiqueta As Etiqueta
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim i As Integer
  
    On Error GoTo ErrorManager
    
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    Set objEtiquetas = New Etiquetas
    Set frmEtiquetas = New EtiquetasEdit
  
    objEtiquetas.BeginEdit
    
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmOrdenCorte = New OrdenCorteEdit
                Set objOrdenCorte = New OrdenCorte
                objOrdenCorte.Load mlngID, GescomMain.objParametro.Moneda
                
                For Each objOrdenCorteItem In objOrdenCorte.OrdenCorteItems
                    Set objEtiqueta = objEtiquetas.Add
                    objEtiqueta.BeginEdit
                    objEtiqueta.TemporadaID = GescomMain.objParametro.TemporadaActualID
                    objEtiqueta.ArticuloColorID = objOrdenCorteItem.ArticuloColorID
                    objEtiqueta.CantidadT36 = objOrdenCorteItem.CantidadT36
                    objEtiqueta.CantidadT38 = objOrdenCorteItem.CantidadT38
                    objEtiqueta.CantidadT40 = objOrdenCorteItem.CantidadT40
                    objEtiqueta.CantidadT42 = objOrdenCorteItem.CantidadT42
                    objEtiqueta.CantidadT44 = objOrdenCorteItem.CantidadT44
                    objEtiqueta.CantidadT46 = objOrdenCorteItem.CantidadT46
                    objEtiqueta.CantidadT48 = objOrdenCorteItem.CantidadT48
                    objEtiqueta.CantidadT50 = objOrdenCorteItem.CantidadT50
                    objEtiqueta.CantidadT52 = objOrdenCorteItem.CantidadT52
                    objEtiqueta.CantidadT54 = objOrdenCorteItem.CantidadT54
                    objEtiqueta.CantidadT56 = objOrdenCorteItem.CantidadT56
                    objEtiqueta.ApplyEdit
                
                Next
                    
                Set objOrdenCorte = Nothing
                
            End If
        End If
    Next i
    
    objEtiquetas.ApplyEdit
    
    frmEtiquetas.Component objEtiquetas
    frmEtiquetas.Show
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub Imprimir()
    Dim objItem As ListItem
    Dim objPrintClass As PrintClass
    Dim frmPrintOptions As frmPrint
    
    On Error GoTo ErrorManager
    
    Set frmPrintOptions = New frmPrint
    frmPrintOptions.Flags = ShowCopies_po + ShowPrinter_po
    frmPrintOptions.Copies = 1
    frmPrintOptions.Show vbModal
    ' salir de la opcion si no pulsa "imprimir"
    If Not frmPrintOptions.PrintDoc Then
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
        Exit Sub
    End If
        
    Set objPrintClass = New PrintClass
    objPrintClass.PrinterNumber = frmPrintOptions.PrinterNumber
    objPrintClass.Copies = frmPrintOptions.Copies
    
    objPrintClass.Titulo = "Listado de órdenes de corte de la temporada " & GescomMain.objParametro.TemporadaActual
    
    objPrintClass.Columnas = lvwItems.ColumnHeaders
    
    For Each objItem In lvwItems.ListItems
        objPrintClass.Item = objItem
    Next
    objPrintClass.EndDoc

    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
    Set objPrintClass = Nothing
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

' Prevision de etiquetas a partir de ordenes de corte.
Public Sub PrevEtiquetas()
    Dim frmList As OLAPQueryList
    Dim objOLAPQuery As OLAPQuery

    Set frmList = New OLAPQueryList
    Set objOLAPQuery = New OLAPQuery
    objOLAPQuery.Load QRY_PrevEtiqueta, mobjBusqueda.ClausulaWhere
    With frmList
        .Component objOLAPQuery
        .Show vbModal
  
    End With

    Set objOLAPQuery = Nothing

End Sub

Public Sub PrintSelected()
    Dim i As Integer
    Dim Respuesta As VbMsgBoxResult
    Dim objPrintOrdenCorte As PrintOrdenCorte
    Dim frmPrintOptions As frmPrint
    
    On Error GoTo ErrorManager
   
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos imprimir los documentos
    Respuesta = MostrarMensaje(MSG_DOCUMENTO)
    
    If Respuesta = vbYes Then
        Set frmPrintOptions = New frmPrint
        frmPrintOptions.Flags = ShowCopies_po + ShowPrinter_po
        frmPrintOptions.Copies = 1
        frmPrintOptions.Show vbModal
        ' salir de la opcion si no pulsa "imprimir"
        If Not frmPrintOptions.PrintDoc Then
            Unload frmPrintOptions
            Set frmPrintOptions = Nothing
            Exit Sub
        End If
            
        For i = 1 To lvwItems.ListItems.Count
            If lvwItems.ListItems(i).Selected = True Then
                mlngID = Val(lvwItems.ListItems(i).Key)
                If mlngID > 0 Then
                    Set objOrdenCorte = New OrdenCorte
                    Set objPrintOrdenCorte = New PrintOrdenCorte
                    objOrdenCorte.Load mlngID, GescomMain.objParametro.Moneda
                    
                    objPrintOrdenCorte.PrinterNumber = frmPrintOptions.PrinterNumber
                    objPrintOrdenCorte.Copies = frmPrintOptions.Copies
                    objPrintOrdenCorte.Component objOrdenCorte
                    
                    objPrintOrdenCorte.PrintObject
                    
                    Set objPrintOrdenCorte = Nothing
                    Set objOrdenCorte = Nothing
                End If
            End If
        Next i
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

