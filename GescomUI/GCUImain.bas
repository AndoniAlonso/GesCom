Attribute VB_Name = "GCUIMain"
Option Explicit

Private Type RECT
    Izq As Long
    Sup As Long
    Der As Long
    Inf As Long
End Type

Public Const SPI_GETWORKAREA = 48

Public Enum MsgTypes
    MSG_OPEN = 0
    MSG_DELETE = 1
    MSG_MODIFY = 2
    MSG_ACTUALIZAR_ORDEN = 3
    MSG_COBRAR = 4
    MSG_GENERAR_REMESA = 5
    MSG_DOCUMENTO = 6
    MSG_MODIF_ARTICULO = 7
    MSG_RECALCULAR_ARTICULO = 8
    MSG_RECALCULAR_VENTA = 9
    MSG_CREAR_ARTICULO = 10
    MSG_FACTURARALBARAN = 11
    MSG_RECALCULARVENCIMIENTOS = 12
    MSG_SALIRSINRECALCULAR = 13
    MSG_ACTUALIZAR_PRECIOS = 14
    MSG_CONTABILIZAR = 15
    MSG_VOLVER_A_CONTABILIZAR = 16
    MSG_CONTABILIZAR_OK = 17
    MSG_PROCESO_OK = 18
    MSG_DESCONTABILIZAR = 19
    MSG_DELETEFACTURA = 20
    MSG_EXPORTTOEXCEL = 21
    MSG_RECALCULAR_PVP = 22
End Enum

Public Enum enFormatoJSColumn
    enFormatoTexto = 0
    enFormatoCantidad = 1
    enFormatoFecha = 2
    enFormatoHora = 3
    enFormatoImporte = 4
    enFormatoPorcentaje = 5
End Enum

' Enumerado de opciones de impresión
Public Enum qePrintOptionFlags
  ShowPrinter_po = 1
  ShowCopies_po = 2
End Enum

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000


' esta rutina muestra el mensaje generico deseado
Public Function MostrarMensaje(MsgType As MsgTypes) As VbMsgBoxResult
    Dim Respuesta As VbMsgBoxResult
    
    Select Case MsgType
        Case MSG_OPEN ' Abrir
            Respuesta = MsgBox("Abrir un número elevado de registros puede " _
                & " necesitar altos recursos del sistema." & vbCrLf & " ¿Proceder? " _
                , vbQuestion + vbYesNo, "Confirmar apertura de registros")
            MostrarMensaje = Respuesta
            
        Case MSG_DELETE ' Eliminar
            Respuesta = MsgBox("¿Eliminar los registros seleccionados?" _
                , vbQuestion + vbYesNo, "Confirmar eliminación")
            MostrarMensaje = Respuesta
            
        Case MSG_MODIFY 'Cerrar con modificaciones
            Respuesta = MsgBox("Se han realizado modificaciones en los datos." _
                & vbCrLf & "¿Realmente desea salir?", vbQuestion + vbYesNo, "Modificación en los datos")
            MostrarMensaje = Respuesta
            
        Case MSG_ACTUALIZAR_ORDEN 'Realizar la actualización en almacenes de objetos (ordenes de corte, etc)
            Respuesta = MsgBox("¿Desea actualizar las ordenes de corte seleccionadas?" _
                & vbCrLf & "Esta opcion realiza movimientos de stock de artículos y materiales!", vbQuestion + vbYesNo, "Confirmar actualización de órdenes de corte")
            MostrarMensaje = Respuesta
            
        Case MSG_COBRAR  'Marcar los cobros y pagos pendientes como cobrados/pagados
            Respuesta = MsgBox("¿Desea marcar los cobros/pagos seleccionados como ya cobrados/pagados?" _
                & vbCrLf & "Esta opcion no se puede deshacer.", vbQuestion + vbYesNo, "Confirmar marcado de cobros/pagos pendientes")
            MostrarMensaje = Respuesta
            
        Case MSG_GENERAR_REMESA 'Generar el fichero de remesas formato CSB58
            Respuesta = MsgBox("¿Desea generar el fichero de remesas?" _
                , vbQuestion + vbYesNo, "Confirmar generación de remesas formato CSB58")
            MostrarMensaje = Respuesta
            
        Case MSG_DOCUMENTO 'Imprimir el documento
            Respuesta = MsgBox("¿Desea imprimir los documentos seleccionados?" _
                , vbQuestion + vbYesNo, "Confirmar impresión de documentos")
            MostrarMensaje = Respuesta
        
        Case MSG_MODIF_ARTICULO 'Modificar datos que pueden implicar cambios en el calculo del coste de articulos
            Respuesta = MsgBox("Las modificaciones realizadas pueden implicar cambios en el calculo del coste de artículos existentes." & vbCrLf & _
                "Para recalcular los precios de coste utilice la opción de recálculo en el menú de artículos.", vbInformation + vbOKOnly, "Datos del coste de artículos modificados")
            MostrarMensaje = Respuesta
        
        Case MSG_RECALCULAR_ARTICULO 'Recalcular datos de coste de articulos.
            Respuesta = MsgBox("¿Desea recalcular los precios de coste de los artículos seleccionados?" _
                , vbQuestion + vbYesNo, "Recalcular precio de coste de artículos")
            MostrarMensaje = Respuesta
            
        Case MSG_RECALCULAR_VENTA 'Recalcular datos de venta de articulos.
            Respuesta = MsgBox("¿Desea recalcular el precio de venta de los artículos seleccionados?" _
                , vbQuestion + vbYesNo, "Recalcular precio de venta de artículos")
            MostrarMensaje = Respuesta
            
        Case MSG_FACTURARALBARAN 'Facturar los albaranes seleccionados.
            Respuesta = MsgBox("¿Desea generar las facturas de los albaranes seleccionados?" _
                , vbQuestion + vbYesNo, "Generar facturas desde albaranes")
            MostrarMensaje = Respuesta
            
        Case MSG_RECALCULARVENCIMIENTOS 'Recalcular los vencimientos de una factura
            Respuesta = MsgBox("¿Recalcular los vencimientos de la factura?" _
                , vbQuestion + vbYesNo, "Recalcular vencimientos")
            MostrarMensaje = Respuesta
            
        Case MSG_SALIRSINRECALCULAR 'Salir sin recalcular los vencimientos de una factura
            Respuesta = MsgBox("El importe de la factura no concuerda con el de los vencimientos ¿Desea salir sin recalcular los vencimientos de la factura?" _
                , vbQuestion + vbYesNo, "Salir sin recalcular vencimientos")
            MostrarMensaje = Respuesta
        
        Case MSG_CREAR_ARTICULO 'Preguntar si hay que crear el articulo que no existe a partir de los componentes
            Respuesta = MsgBox("El artículo no existe, ¿desea crearlo?" _
                , vbQuestion + vbYesNo, "Crear artículo")
            MostrarMensaje = Respuesta
            
        Case MSG_ACTUALIZAR_PRECIOS 'Preguntar si hay que actualizar el precio de venta del articulo en los pedidos seleccionados
            Respuesta = MsgBox("¿Desea actualizar los precios de venta de artículos en los pedidos seleccionados?" _
                , vbQuestion + vbYesNo, "Actualizar precios de venta en artículos")
            MostrarMensaje = Respuesta
            
        Case MSG_CONTABILIZAR 'Preguntar si hay que contabilizar la operacion dada
            Respuesta = MsgBox("¿Desea contabilizar las operaciones seleccionadas?" _
                , vbQuestion + vbYesNo, "Contabilización")
            MostrarMensaje = Respuesta
            
        Case MSG_VOLVER_A_CONTABILIZAR 'Preguntar si hay que forzar la contabilizacion de una operacion dada
            Respuesta = MsgBox("Algunas de las operaciones seleccionadas ya fueron contabilizadas" & vbCrLf & _
                               "¿Desea volver a generar la contabilidad de las operaciones seleccionadas?" _
                , vbQuestion + vbYesNoCancel + vbDefaultButton2, "Volver a contabilizar operaciones ya contabilizadas")
            MostrarMensaje = Respuesta
            
        Case MSG_CONTABILIZAR_OK 'Avisar de que la contabilización ha ido OK
            Respuesta = MsgBox("La contabilización se ha realizado correctamente" _
                , vbInformation + vbOKOnly, "Contabilización")
            MostrarMensaje = Respuesta
            
        Case MSG_PROCESO_OK 'Avisar de que el proceso ha ido OK
            Respuesta = MsgBox("El proceso se ha realizado correctamente" _
                , vbInformation + vbOKOnly, "Actualización")
            MostrarMensaje = Respuesta
            
        Case MSG_DESCONTABILIZAR 'Preguntar si hay que DESContabilizar la operacion dada
            Respuesta = MsgBox("¿Desea DESContabilizar las operaciones seleccionadas?" _
                & vbCrLf & "La descontabilización implica que se deben eliminar manualmente los asientos de la aplicación y/o de Contawin" _
                , vbQuestion + vbYesNo, "DESContabilización")
            MostrarMensaje = Respuesta
            
        Case MSG_DELETEFACTURA 'Preguntar si está de acuerdo en borra la factura complementaria
            Respuesta = MsgBox("Se eliminará también la factura asociada con ésta en la otra empresa." _
                & vbCrLf & " ¿Desea continuar?" _
                , vbQuestion + vbYesNo, "Eliminar factura complementaria")
            MostrarMensaje = Respuesta
        
        Case MSG_EXPORTTOEXCEL
            Respuesta = MsgBox("Se exportarán todos los registros de la lista a una hoja Excel" _
                & vbCrLf & " ¿Desea continuar?" _
                , vbQuestion + vbYesNo, "Exportar datos a Excel")
            MostrarMensaje = Respuesta
        
        Case MSG_RECALCULAR_PVP 'Recalcular datos de PVP de articulos.
            Respuesta = MsgBox("¿Desea recalcular el PVP de los artículos seleccionados?" _
                , vbQuestion + vbYesNo, "Recalcular precio de PVP de artículos")
            MostrarMensaje = Respuesta
            
        Case Else
            Err.Raise vbObjectError + 1001, "Tipo de mensaje genérico erróneo."
            
    End Select
    
End Function

' esta funcion devuelve el número de elementos seleccionados
' de un listview
Public Function NumeroSeleccionados(lvwSelection As ListView) As Integer
    Dim i As Integer
    Dim Total As Integer
    
    Total = 0
    
    For i = 1 To lvwSelection.ListItems.Count
        If lvwSelection.ListItems(i).Selected Then
            Total = Total + 1
        End If
    Next
    
    NumeroSeleccionados = Total
    
End Function
Public Sub LoadImages(BarraHerramientas As Toolbar)
    Dim objButton As Button
    
    With BarraHerramientas
        .ImageList = GescomMain.mglIconosPequeños
        Set objButton = .Buttons.Add(, "Nuevo", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Nuevo").Key)
        objButton.ToolTipText = "Crear un nuevo objeto"
        Set objButton = .Buttons.Add(, "Abrir", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Abrir").Key)
        objButton.ToolTipText = "Editar"
        Set objButton = .Buttons.Add(, "Imprimir", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Imprimir").Key)
        objButton.ToolTipText = "Imprimir"
        Set objButton = .Buttons.Add(, , , tbrSeparator)
        Set objButton = .Buttons.Add(, "Eliminar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Eliminar").Key)
        objButton.ToolTipText = "Borrar"
        Set objButton = .Buttons.Add(, "Actualizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Actualizar").Key)
        objButton.ToolTipText = "Refrescar la vista de objetos"
        Set objButton = .Buttons.Add(, "Buscar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Buscar").Key)
        objButton.ToolTipText = "Consultas"
        Set objButton = .Buttons.Add(, , , tbrSeparator)
        'Set objButton = .Buttons.Add(, "IconosGrandes", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("IconosGrandes").Key)
        'objButton.ToolTipText = "Iconos Grandes"
        Set objButton = .Buttons.Add(, "IconosPequeños", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("IconosPequeños").Key)
        objButton.ToolTipText = "Iconos pequeños"
        'Set objButton = .Buttons.Add(, "Lista", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Lista").Key)
        'objButton.ToolTipText = "Lista de iconos"
        Set objButton = .Buttons.Add(, "Detalle", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Detalle").Key)
        objButton.ToolTipText = "Detalle"
        objButton.Value = tbrPressed
        Set objButton = .Buttons.Add(, , , tbrSeparator)
        Set objButton = .Buttons.Add(, "QuickSearch", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Buscar").Key)
        objButton.ToolTipText = "Búsqueda rápida"
        Set objButton = .Buttons.Add(, "GroupBy", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("GroupBy").Key)
        objButton.ToolTipText = "Agrupar"
        Set objButton = .Buttons.Add(, "ShowFields", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Propiedades").Key)
        objButton.ToolTipText = "Lista de campos"
        Set objButton = .Buttons.Add(, "Sort", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Ordenar").Key)
        objButton.ToolTipText = "Ordenar los campos"
                
        Set objButton = .Buttons.Add(, , , tbrSeparator)
        Set objButton = .Buttons.Add(, "ExportToExcel", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Excel").Key)
        objButton.ToolTipText = "Exportar a Excel"
        
        Set objButton = .Buttons.Add(, , , tbrSeparator)
        Set objButton = .Buttons.Add(, "Cerrar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Cerrar").Key)
        objButton.ToolTipText = "Cerrar"
                
    End With
    
Set objButton = Nothing
End Sub

Public Sub SelTextBox(Text As TextBox)

    With Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Public Sub TextChange(Ctl As TextBox, Obj As Object, Prop As String)
    Dim lngPos As Long
  
    On Error GoTo INPUTERR
    CallByName Obj, Prop, VbLet, Ctl.Text
    Exit Sub
  
INPUTERR:
    Beep
  
    lngPos = Ctl.SelStart
    Ctl = CallByName(Obj, Prop, VbGet)
  
    ' Se supone que da error al introducir el último caracter --> lo seleccionamos
    If lngPos > 0 Then _
        Ctl.SelStart = lngPos - 1
    If lngPos = 0 Then _
        SelTextBox Ctl
        
End Sub

Public Function TextLostFocus(Ctl As TextBox, Obj As Object, Prop As String)
  
    TextLostFocus = CallByName(Obj, Prop, VbGet)

End Function

Public Sub LoadCombo(Combo As ComboBox, List As TextList)
    Dim vntItem As Variant
  
    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Public Sub LoadComboCampos(Combo As ComboBox, List As Collection)
    Dim vntItem As ConsultaCampo
  
    Set vntItem = New ConsultaCampo
  
    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem.Alias
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Set vntItem = Nothing

End Sub

' Devuelve el tamaño asignado a un dato en una columna de un ListView
Public Function ColumnSize(Value As Integer) As Integer
    Dim ColSize As Integer

    ColSize = 140 * Value
    If ColSize < 200 Then ColSize = 400
    If ColSize > 6000 Then ColSize = 6000
  
    ColumnSize = ColSize

End Function

' Búsqueda rápida en los campos de un listview
Public Sub ListviewQuickSearch(lvwItems As ListView, nColumn As Integer)
    Dim strBusqueda As String
    Dim strDato As String
    Dim i As Integer

    strBusqueda = InputBox("Buscar el dato:", "Búsqueda rápida en " & lvwItems.ColumnHeaders(nColumn).Text)
    
    For i = 1 To lvwItems.ListItems.Count
        lvwItems.ListItems(i).Selected = False
    Next i
    
    For i = 1 To lvwItems.ListItems.Count
        If nColumn = 1 Then
            strDato = lvwItems.ListItems(i).Text
        Else
            strDato = lvwItems.ListItems(i).SubItems(nColumn - 1)
        End If
        If InStr(1, strDato, strBusqueda, vbTextCompare) Then
            lvwItems.ListItems(i).EnsureVisible
            Set lvwItems.SelectedItem = lvwItems.ListItems(i)
            lvwItems.SetFocus
            Exit Sub
        End If
    Next i
    
End Sub

' Búsqueda rápida en los campos de un Janus
Public Sub JanusQuickSearch(jgrdGridEX As GridEX, nColumn As Integer)
    Dim strBusqueda As String
'    Dim strDato As String
'    Dim i As Integer

    strBusqueda = InputBox("Buscar el dato:", "Búsqueda rápida en " & jgrdGridEX.Columns(nColumn).Caption)
    
    jgrdGridEX.Find nColumn, jgexContains, strBusqueda
    
End Sub

' Funcion general de manejo de errores
Public Sub ManageErrors(NombreFormulario As String)
    Dim Result As VbMsgBoxResult
    
    If Screen.MousePointer <> vbDefault Then Screen.MousePointer = vbDefault
    
    If Err.Number >= vbObjectError + 1001 Then
       Result = MsgBox("Error:" & Err.Source & vbCrLf & Err.Description, vbCritical + vbOKOnly, NombreFormulario)
    Else
       Result = MsgBox("Error:" & Err.Number & "-" & Err.Description & vbCrLf & Err.Source, vbCritical + vbOKOnly, NombreFormulario)
    End If
        
End Sub

' Salir de la ejecucion del programa como resultado de un error critico
Public Sub TerminateProgram()
    Dim Result As VbMsgBoxResult

    Result = MsgBox("Error grave, se finaliza la ejecucion del programa:", _
        vbCritical + vbOKOnly)
    
    Unload GescomMain
    End
    
End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
         SendMessage LV.hWnd, LVM_FIRST + 30, c.Index - 1, -1
        Next
    Else
        SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    
    LV.Refresh
    
End Sub

Public Sub ListView_Resize(lvwItems As ListView, frmForm As Form, Optional frmFiltro As Frame)
    
    If frmFiltro Is Nothing Then
        lvwItems.Move 120, 480, Abs(frmForm.Width - 360), Abs(frmForm.Height - 960)
    Else
        lvwItems.Move 120, 480, Abs(frmForm.Width - 360), Abs(frmForm.Height - 960 - frmFiltro.Height)
        frmFiltro.Move 120, 480 + lvwItems.Height, Abs(frmForm.Width - 360)
    End If
    

End Sub

Public Sub GridEX_Resize(ByRef jgrdItems As GridEX, ByRef frmForm As Form, Optional ByRef frmFiltro As Frame)
    
    If frmFiltro Is Nothing Then
        jgrdItems.Move 120, 480, Abs(frmForm.Width - 360), Abs(frmForm.Height - 960)
    Else
        jgrdItems.Move 120, 480, Abs(frmForm.Width - 360), Abs(frmForm.Height - 960 - frmFiltro.Height)
        frmFiltro.Move 120, 480 + jgrdItems.Height, Abs(frmForm.Width - 360)
    End If

End Sub

Public Sub ListView_ColumnClick(lvwSorted As ListView, ColumnHeader As ColumnHeader)
    
    ' Cuando se hace clic en un objeto ColumnHeader, el
    ' control ListView se ordena por los subelementos de
    ' esa columna.
    If lvwSorted.SortKey = ColumnHeader.Index - 1 Then
        If lvwSorted.SortOrder = lvwAscending Then
            lvwSorted.SortOrder = lvwDescending
        Else
            lvwSorted.SortOrder = lvwAscending
        End If
    End If
    
    ' Establece el SortKey como el Index del ColumnHeader - 1
    lvwSorted.SortKey = ColumnHeader.Index - 1
    
    ' Asigna a Sorted el valor True para ordenar la lista.
    lvwSorted.Sorted = True
    
End Sub

Public Function ShowDialogSave(Titulo As String, Extension As String, NombreFichero As String, _
                                Filtro As String) As VbMsgBoxResult

    On Error GoTo ErrorManager
    With GescomMain.dlgFileSave
        .DialogTitle = Titulo
        .DefaultExt = Extension
        .FileName = NombreFichero
        .Filter = Filtro
        .Flags = cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        
        .CancelError = True
        .ShowSave
    End With
    
    ShowDialogSave = vbOK
    Exit Function

ErrorManager:
    ShowDialogSave = vbCancel
    
End Function

Public Sub SmartComboKeyPress(KeyAscii As Integer, ByRef mintSelStart As Integer, ByRef cboCombo As ComboBox)
    Dim lCnt       As Long 'Generic long counter
    Dim lMax       As Long
    Dim sComboItem As String
    Dim sComboText As String 'Text currently in combobox
    Dim sText      As String 'Text after keypressed

    With cboCombo
        lMax = .ListCount - 1
        sComboText = .Text
        sText = Left(sComboText, mintSelStart) & Chr(KeyAscii)
        
        KeyAscii = 0 'Reset key pressed
        
        For lCnt = 0 To lMax
            sComboItem = .List(lCnt)
            
            If UCase(sText) = UCase(Left(sComboItem, _
                                         Len(sText))) Then
                .ListIndex = lCnt
                .Text = sComboItem
                .SelStart = Len(sText)
                .SelLength = Len(sComboItem) - (Len(sText))
                
                Exit For
            End If
        Next 'lCnt
    End With
End Sub



 
'*----------------------------------------------------------*
'* Name       : DisableX                                    *
'*----------------------------------------------------------*
'* Purpose    : Disables the close button ('X') on form.    *
'*----------------------------------------------------------*
'* Parameters : frm    Required. Form to disable 'X'-button *
'*----------------------------------------------------------*
'* Description: This function disables the X-button on a    *
'*            : form, to keep the user from closing a form  *
'*            : that way, but keeps the min & max buttons.  *
'*----------------------------------------------------------*
Public Sub DisableX(frm As Form)
    Dim hMenu As Long, nCount As Long
    
    'Get handle to system menu
    hMenu = GetSystemMenu(frm.hWnd, 0)
    
    'Get number of items in menu
    nCount = GetMenuItemCount(hMenu)
    
    If frm.MDIChild Then
        'Remove last item from system menu (item 'Close' for MDI windows)
        Call RemoveMenu(hMenu, nCount - 3, MF_DISABLED Or MF_BYPOSITION)
    Else
        'Remove last item from system menu (last item is 'Close')
        Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
    End If
    
    'Redraw menu
    DrawMenuBar frm.hWnd
    
    'Position on top
    frm.Top = 0
    
End Sub

Public Sub FormatoJColumn(jsCol As JSColumn, intPosicion As Integer, strEncabezado As String, Optional bolTotalizar As Boolean = False, Optional intWidth As Integer = -1, Optional Formato As enFormatoJSColumn)

    With jsCol
        .Visible = True
        .ColPosition = intPosicion
        .Caption = strEncabezado
        If bolTotalizar Then
            If Formato = enFormatoPorcentaje Then
                .AggregateFunction = jgexValueCount '= jgexAvg
                .Format = "###,##0.00\%"
                .TotalRowFormat = "###,##0.00\%"
                .TextAlignment = jgexAlignRight
            Else
                .AggregateFunction = jgexSum
                .Format = "###,###,##0.00"
                .TotalRowFormat = "###,###,##0.00"
                .TextAlignment = jgexAlignRight
            End If
        Else
            .AggregateFunction = jgexAggregateNone
        End If
        If intWidth < 0 Then
            .AutoSize
        Else
            .Width = intWidth
        End If
        .EditType = jgexEditNone
        Select Case Formato
        Case enFormatoCantidad
            .Format = "###,###,##0"
             If bolTotalizar Then
                .TotalRowFormat = "###,###,##0"
             End If

        Case enFormatoImporte
            .Format = "###,###,##0.00"
             If bolTotalizar Then
                .TotalRowFormat = "###,###,##0.00"
             End If
        Case enFormatoPorcentaje
            .Format = "###,##0.00\%"
             'If bolTotalizar Then
             '   .TotalRowFormat = "###,###,##0.00"
             'End If
        End Select
        
    End With

End Sub

Public Sub CentrarForma(frm As Form)
Dim R As RECT
Dim lRes As Long
Dim lAncho As Long
Dim lLargo As Long
    With frm
        If .WindowState = vbMinimized Or .WindowState = vbMaximized Then Exit Sub
    End With
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, R, 0)
    If lRes Then
        With R
            .Izq = Screen.TwipsPerPixelX * .Izq
            .Sup = Screen.TwipsPerPixelY * .Sup
            .Der = Screen.TwipsPerPixelX * .Der
            .Inf = Screen.TwipsPerPixelY * .Inf
            lAncho = .Der - .Izq
            lLargo = .Inf - .Sup
            frm.Move .Izq + (lAncho - frm.Width) \ 2, .Sup + (lLargo - frm.Height) \ 2
        End With
    End If
End Sub


Public Sub ExportRSToExcel(rstFilterResult As ADOR.Recordset)
    Dim xlsApp As Object 'Excel.Application
    Dim xlsWrk As Object 'Excel.Workbook
    Dim xlsSheet As Object 'Excel.Workbook.Sheet
    
    Dim i As Long, j As Long
    
    On Error GoTo TratarError:
    If Not rstFilterResult Is Nothing Then
        If rstFilterResult.State <> adStateClosed Then
            If rstFilterResult.RecordCount > 0 Then
                Screen.MousePointer = vbHourglass
                Set xlsApp = CreateObject("Excel.Application")
                If Not xlsApp Is Nothing Then
                    Set xlsWrk = xlsApp.Workbooks.Add
                    Set xlsSheet = xlsWrk.Sheets.Add
                    For j = 1 To rstFilterResult.Fields.Count
                        xlsSheet.Cells(1, j) = rstFilterResult(j - 1).Name
                    Next j
                    i = 2
                    rstFilterResult.MoveFirst
                    While Not rstFilterResult.EOF
                        For j = 1 To rstFilterResult.Fields.Count
                            xlsSheet.Cells(i, j) = rstFilterResult(j - 1)
                        Next j
                        rstFilterResult.MoveNext
                        i = i + 1
                    Wend
                    xlsApp.Visible = True
                    Set xlsSheet = Nothing
                    Set xlsWrk = Nothing
                    Set xlsApp = Nothing
                End If
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
    Exit Sub
TratarError:
    Set xlsSheet = Nothing
    Set xlsWrk = Nothing
    Set xlsApp = Nothing
    MsgBox "Ha ocurrido el siguiente error al cargar MS Excel:" & vbCr & Err.Description, vbCritical + vbOKOnly, "Exportar datos a MS Excel"
End Sub

Public Sub ExportRecordList(rcsRecordList As ADOR.Recordset)
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo TratarError:
    
    Respuesta = MostrarMensaje(MSG_EXPORTTOEXCEL)
    If Respuesta <> vbYes Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    ExportRSToExcel rcsRecordList
    
    Screen.MousePointer = vbDefault
    Exit Sub
TratarError:
    Screen.MousePointer = vbDefault
End Sub
