VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form OrdenCorteEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de corte "
   ClientHeight    =   6744
   ClientLeft      =   2976
   ClientTop       =   2892
   ClientWidth     =   10332
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OrdenCorteEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6744
   ScaleWidth      =   10332
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "&Pedidos..."
      Height          =   375
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la órden de corte"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox cboArticulo 
         Height          =   315
         Left            =   3840
         TabIndex        =   6
         Text            =   "cboArticulo"
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox chkOrdenCortada 
         Caption         =   "Orden de corte actualizada"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtFechaCorte 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   2
         Top             =   340
         Width           =   975
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1425
         Width           =   5655
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   700
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         Format          =   57606145
         CurrentDate     =   36938
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de corte"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1095
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   7800
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas de la órden de corte"
      Height          =   3855
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   9855
      Begin VB.TextBox txtTotalCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8520
         TabIndex        =   18
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7440
         TabIndex        =   17
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwOrdenCorteItems 
         Height          =   2895
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   9375
         _ExtentX        =   16531
         _ExtentY        =   5101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total prendas "
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3375
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   22
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   21
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "OrdenCorteEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintArticuloSelStart As Integer

Private WithEvents mobjOrdenCorte As OrdenCorte
Attribute mobjOrdenCorte.VB_VarHelpID = -1

Public Sub Component(OrdenCorteObject As OrdenCorte)

    Set mobjOrdenCorte = OrdenCorteObject

End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjOrdenCorte.ApplyEdit
    mobjOrdenCorte.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    If mobjOrdenCorte.Numero = GescomMain.objParametro.ObjEmpresaActual.OrdenCorte Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaOrdenCorte
    End If
  
    mobjOrdenCorte.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjOrdenCorte.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceOrdenCorte (mobjOrdenCorte.Numero)
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub cboArticulo_Click()

    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    mobjOrdenCorte.Articulo = cboArticulo.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub dtpFecha_Change()
    
    mobjOrdenCorte.Fecha = dtpFecha.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjOrdenCorte
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Orden de Corte [(nuevo)]"

        Else
            Caption = "Orden de Corte [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        txtFechaCorte = .FechaCorte
        chkOrdenCortada = IIf(.OrdenCortada, vbChecked, vbUnchecked)
        txtObservaciones = .Observaciones
        'txtNombre = .Nombre
        
        .BeginEdit GescomMain.objParametro.Moneda
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaOrdenCorte
            txtNumero = .Numero
       
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
        
        LoadCombo cboArticulo, .Articulos
        cboArticulo.Text = .Articulo
        
    End With
    
    lvwOrdenCorteItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwOrdenCorteItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwOrdenCorteItems.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwOrdenCorteItems.ColumnHeaders.Add , , "36", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "38", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "40", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "42", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "44", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "46", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "48", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "50", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "52", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "54", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "56", ColumnSize(4), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "Cant.", ColumnSize(5), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "Pedido", ColumnSize(6), vbRightJustify
    lvwOrdenCorteItems.ColumnHeaders.Add , , "Cliente", ColumnSize(15)
    LoadOrdenCorteItems
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwOrdenCorteItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjOrdenCorte_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub


Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjOrdenCorte, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjOrdenCorte, "Observaciones")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjOrdenCorte, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjOrdenCorte, "Numero")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

' a partir de aqui -----> child

'Private Sub cmdAdd_Click()
'
'    Dim frmOrdenCorteItem As OrdenCorteItemEdit
'
'    On Error GoTo ErrorManager
'    Set frmOrdenCorteItem = New OrdenCorteItemEdit
'    frmOrdenCorteItem.Component mobjOrdenCorte.OrdenCorteItems.Add
'    frmOrdenCorteItem.Show vbModal
'    LoadOrdenCorteItems
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'    Exit Sub
'
'End Sub

Private Sub cmdEdit_Click()
    Dim frmOrdenCorteItem As OrdenCorteItemEdit
  
    On Error GoTo ErrorManager
    
    Set frmOrdenCorteItem = New OrdenCorteItemEdit
    frmOrdenCorteItem.Component _
        mobjOrdenCorte.OrdenCorteItems(Val(lvwOrdenCorteItems.SelectedItem.Key))
    frmOrdenCorteItem.Show vbModal
    LoadOrdenCorteItems
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    mobjOrdenCorte.OrdenCorteItems.Remove Val(lvwOrdenCorteItems.SelectedItem.Key)
    LoadOrdenCorteItems
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadOrdenCorteItems()
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwOrdenCorteItems.ListItems.Clear
    For lngIndex = 1 To mobjOrdenCorte.OrdenCorteItems.Count
        Set itmList = lvwOrdenCorteItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objOrdenCorteItem = mobjOrdenCorte.OrdenCorteItems(lngIndex)

        With itmList
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objOrdenCorteItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = objOrdenCorteItem.Descripcion
            .SubItems(2) = objOrdenCorteItem.CantidadT36
            .SubItems(3) = objOrdenCorteItem.CantidadT38
            .SubItems(4) = objOrdenCorteItem.CantidadT40
            .SubItems(5) = objOrdenCorteItem.CantidadT42
            .SubItems(6) = objOrdenCorteItem.CantidadT44
            .SubItems(7) = objOrdenCorteItem.CantidadT46
            .SubItems(8) = objOrdenCorteItem.CantidadT48
            .SubItems(9) = objOrdenCorteItem.CantidadT50
            .SubItems(10) = objOrdenCorteItem.CantidadT52
            .SubItems(11) = objOrdenCorteItem.CantidadT54
            .SubItems(12) = objOrdenCorteItem.CantidadT56
            .SubItems(13) = objOrdenCorteItem.Cantidad
            .SubItems(14) = objOrdenCorteItem.Numero
            .SubItems(15) = objOrdenCorteItem.Cliente
        End With

    Next
    
    txtTotalCantidad = mobjOrdenCorte.OrdenCorteItems.Cantidad
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdPedidos_Click()
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
  
    On Error GoTo ErrorManager
  
    ' No hacer nada si no se ha seleccionado un articulo
    If mobjOrdenCorte.ArticuloID = 0 Then Exit Sub
    
    Set frmPicker = New PickerList
  
    frmPicker.LoadData "vPedidoVentaCorte", mobjOrdenCorte.ArticuloID, _
                                            mobjOrdenCorte.EmpresaID, _
                                            mobjOrdenCorte.TemporadaID
    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        ' Primero hay que comprobar que no está ya seleccionado anteriormente
        If Not DocumentoSeleccionado(objPickerItemDisplay.DocumentoID) Then _
            OrdenCorteDesdePedido (objPickerItemDisplay.DocumentoID)
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
      
    LoadOrdenCorteItems
  
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub OrdenCorteDesdePedido(PedidoItemID As Long)
    Dim objOrdenCorteItem As OrdenCorteItem
    Dim objPedidoItem As PedidoVentaItem
    'Dim objArticuloColor As Articulocolor

    Set objPedidoItem = New PedidoVentaItem
    'Set objArticuloColor = New Articulocolor
    
    objPedidoItem.Load PedidoItemID, GescomMain.objParametro.Moneda
    'objArticuloColor.Load objPedidoItem.ArticuloColorID, GescomMain.objParametro.Moneda
    
    'If Not mobjOrdenCorte.ArticuloID Then
    '    mobjOrdenCorte.ArticuloID = objArticuloColor.objArticulo.ArticuloID
    'End If
    ' Hay que asegurarse de que TODAS las lineas de ordenes de corte sean del mismo articulo.
    ' en caso contrario NO hay que añadirla a la coleccion
    'If mobjOrdenCorte.ArticuloID <> objArticuloColor.objArticulo.ArticuloID Then
    '    Exit Sub
    'End If
    
    Set objOrdenCorteItem = mobjOrdenCorte.OrdenCorteItems.Add
    
    With objOrdenCorteItem
        .BeginEdit GescomMain.objParametro.Moneda
        .OrdenDesdePedido PedidoItemID
        .ApplyEdit
    End With
   
    'Set objArticuloColor = Nothing
    Set objOrdenCorteItem = Nothing
    Set objPedidoItem = Nothing
    
End Sub

Private Function DocumentoSeleccionado(DocumentoID As Long) As Boolean
    Dim objOrdenCorteItem As OrdenCorteItem

' Se trata de buscar si existe alguna referencia de ese documento en alguna linea de
' órdenes de corte y es nueva (no se ha actualizado).
    For Each objOrdenCorteItem In mobjOrdenCorte.OrdenCorteItems
        If objOrdenCorteItem.IsNew And _
           objOrdenCorteItem.PedidoVentaItemID = DocumentoID Then
           DocumentoSeleccionado = True
           Exit Function
        End If
    Next
    
    DocumentoSeleccionado = False
    
    Set objOrdenCorteItem = Nothing
End Function

Private Sub cboArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintArticuloSelStart = cboArticulo.SelStart
End Sub

Private Sub cboArticulo_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintArticuloSelStart, cboArticulo
    
End Sub


