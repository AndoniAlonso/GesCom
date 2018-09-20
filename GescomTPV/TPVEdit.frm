VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TPVEdit 
   BackColor       =   &H80000005&
   Caption         =   "Tickets de Venta"
   ClientHeight    =   5580
   ClientLeft      =   2985
   ClientTop       =   2910
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TPVEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   11430
   Begin VB.CommandButton cmdDevolucion 
      BackColor       =   &H00C99497&
      Caption         =   "&Devolución"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C99497&
      Caption         =   "&Editar"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C99497&
      Caption         =   "El&iminar"
      Height          =   375
      Left            =   10320
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtTotalBruto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ListBox lstIncidencias 
      BackColor       =   &H00E7CDCD&
      ForeColor       =   &H00591E1E&
      Height          =   840
      Left            =   960
      TabIndex        =   10
      Top             =   3720
      Width           =   10335
   End
   Begin VB.CommandButton cmdClearIncidencias 
      BackColor       =   &H00C99497&
      Caption         =   "Borrar incidencia"
      Height          =   612
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   852
   End
   Begin VB.TextBox txtBarCode 
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E7CDCD&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   405
      Left            =   9480
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C99497&
      Caption         =   "&Cancelar TICKET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C99497&
      Caption         =   "&Fin TICKET"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwAlbaranVentaItems 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   5840414
      BackColor       =   15191501
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Incidencias:"
      ForeColor       =   &H00591E1E&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "Código de barras:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00591E1E&
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   0
      Width           =   1110
   End
End
Attribute VB_Name = "TPVEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mobjPedidosPendientes As PickerItems

Private WithEvents mobjAlbaranVenta As AlbaranVenta
Attribute mobjAlbaranVenta.VB_VarHelpID = -1

Private mResize As clsResize

Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject

End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager
    
    ' Si hay lineas preguntamos para ver si está seguro
    If mobjAlbaranVenta.AlbaranVentaItems.Count > 0 Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta <> vbYes Then
            Exit Sub
        End If
    End If

    If mobjAlbaranVenta.Numero = TPVMain.Parametro.ObjEmpresaActual.AlbaranVentas Then
        TPVMain.Parametro.ObjEmpresaActual.DecrementaAlbaranVentas
    End If
  
    mobjAlbaranVenta.CancelEdit
    Set mobjAlbaranVenta = New AlbaranVenta
    Form_Load
    
    'Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdClearIncidencias_Click()
    lstIncidencias.Clear
End Sub

Private Sub cmdDevolucion_Click()

    mobjAlbaranVenta.EsDevolucion = Not mobjAlbaranVenta.EsDevolucion
    
    CambiarColorFormulario

End Sub

Private Sub CambiarColorFormulario()
    Dim objControl As Control
    
    For Each objControl In Me.Controls
'        if type of objcontrol is
'        If objControl.ForeColor = &H591E1E Then objControl.ForeColor = vbRed
'        If objControl.ForeColor = vbRed Then objControl.ForeColor = &H591E1E
        
    Next
          

End Sub

Private Sub cmdOK_Click()
    Dim blnNuevoAlbaran As Boolean
    Dim frmTPVGeneralEdit As TPVGeneralEdit
    
    
    On Error GoTo ErrorManager

    Set frmTPVGeneralEdit = New TPVGeneralEdit
    frmTPVGeneralEdit.Component mobjAlbaranVenta
    frmTPVGeneralEdit.Show vbModal
    Set frmTPVGeneralEdit = Nothing
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    blnNuevoAlbaran = mobjAlbaranVenta.IsNew
    
    
    Dim Buffer() As Byte
    Buffer = mobjAlbaranVenta.GetSuperState
    Set mobjAlbaranVenta = Nothing
    Set mobjAlbaranVenta = New AlbaranVenta
    mobjAlbaranVenta.SetSuperState Buffer
    
    mobjAlbaranVenta.AltaTicketTPV
    TPVMain.Parametro.ObjEmpresaActual.EstableceAlbaranVentas (mobjAlbaranVenta.Numero)
    'If blnNuevoAlbaran Then ImprimirTicketVenta
    
    
    MsgBox mobjAlbaranVenta.ImpresionTicket, , "Imprimir Ticket"
    
    
    'Unload Me
    Set mobjAlbaranVenta = Nothing
    Set mobjAlbaranVenta = New AlbaranVenta
    Form_Load
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub ImprimirTicketVenta()  'ojoojo tratar el posible error localmente
    Dim objPrintTicket As PrintTicket
    
    Set objPrintTicket = New PrintTicket
    objPrintTicket.Component mobjAlbaranVenta
    objPrintTicket.PrintObject
    Set objPrintTicket = Nothing

End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Activate()

    txtBarCode.SetFocus

End Sub

Private Sub Form_Load()

    'DisableX Me
    
    mflgLoading = True
    With mobjAlbaranVenta
        EnableOK .IsValid And .AlbaranVentaItems.Count > 0
        
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        txtTotalBruto = FormatoMoneda(.TotalBruto, TPVMain.Parametro.Moneda)
        
        .BeginEdit
        
        If .IsNew Then
            .Numero = TPVMain.Parametro.ObjEmpresaActual.IncrementaAlbaranVentas
            txtNumero = .Numero
       
            .TemporadaID = TPVMain.Parametro.TemporadaActualID
            .EmpresaID = TPVMain.Parametro.EmpresaActualID
            .TerminalID = TPVMain.Terminal.TerminalID
            .CentroGestionID = TPVMain.Terminal.CentroGestionID
            .AlmacenID = TPVMain.Terminal.AlmacenID
            .AsignarClientePredeterminado
        
        End If
        Caption = "Ticket de Venta [" & .Cliente & "]"
    
    End With
    
    lvwAlbaranVentaItems.SmallIcons = TPVMain.mglIconosPequeños
    lvwAlbaranVentaItems.ColumnHeaders.Clear
    lvwAlbaranVentaItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(30)
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Talla", ColumnSize(6), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Precio", ColumnSize(14), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Descuento", ColumnSize(12), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(11), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Importe", ColumnSize(14), vbRightJustify
    LoadAlbaranVentaItems
    
    lstIncidencias.Clear
  
    mflgLoading = False

    Set mResize = New clsResize
    mResize.Init Me
    
    Me.WindowState = 2 'vbext_ws_Maximize
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid 'And mobjAlbaranVenta.AlbaranVentaItems.Count > 0
'    cmdApply.Enabled = flgValid
    
End Sub

Private Sub Form_Resize()
   mResize.FormResize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjPedidosPendientes = Nothing

End Sub

Private Sub lvwAlbaranVentaItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjAlbaranVenta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub


'OJOOJO refactorizar esto que es muy largo.
Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim objAlbaranVentaItem  As AlbaranVentaItem
    Dim strCodigo As String
    Dim lngCodigo As Long
    Dim strIncidencia As String
    
    On Error GoTo ErrorManager

    If KeyCode = 13 Then
        strCodigo = txtBarCode.Text
        txtBarCode.Text = vbNullString
        ' Validar que sea una cantidad numerica
        If Not IsNumeric(strCodigo) Then Exit Sub
        ' Validar que no sea un numero demasiado grande que provoque desbordamiento
        If Len(strCodigo) > 9 Then Exit Sub

        ' Validar que tenga al menos información de talla + articulo
        lngCodigo = CLng(strCodigo)
        If lngCodigo < 100 Then
            lstIncidencias.AddItem "Falta información del código de artículo!, " & strCodigo
            Exit Sub
        End If
        
        strIncidencia = mobjAlbaranVenta.AlbaranItemTPV(strCodigo)
        If strIncidencia <> vbNullString Then
            lstIncidencias.AddItem strIncidencia & " Código:" & strCodigo
        End If
        
        Beep
        LoadAlbaranVentaItems
        txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, TPVMain.Parametro.Moneda)
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjAlbaranVenta, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjAlbaranVenta, "Numero")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdEdit_Click()
    Dim frmTPVItemEdit As TPVItemEdit
    
    On Error GoTo ErrorManager
    If lvwAlbaranVentaItems.ListItems.Count = 0 Then Exit Sub
  
    Set frmTPVItemEdit = New TPVItemEdit
    frmTPVItemEdit.Component _
        mobjAlbaranVenta.AlbaranVentaItems(Val(lvwAlbaranVentaItems.SelectedItem.Key))
    frmTPVItemEdit.Show vbModal
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, TPVMain.Parametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
    
    If lvwAlbaranVentaItems.ListItems.Count = 0 Then Exit Sub
        
    mobjAlbaranVenta.AlbaranVentaItems.Remove Val(lvwAlbaranVentaItems.SelectedItem.Key)
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, TPVMain.Parametro.Moneda)

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadAlbaranVentaItems()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwAlbaranVentaItems.ListItems.Clear
    For lngIndex = 1 To mobjAlbaranVenta.AlbaranVentaItems.Count
        Set itmList = lvwAlbaranVentaItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objAlbaranVentaItem = mobjAlbaranVenta.AlbaranVentaItems(lngIndex)

        With itmList
            .SmallIcon = TPVMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objAlbaranVentaItem.IsDeleted Then .SmallIcon = TPVMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = IIf(objAlbaranVentaItem.ArticuloColorID, Trim(objAlbaranVentaItem.ArticuloColor), Trim(objAlbaranVentaItem.Descripcion))
            
            Select Case True
            Case objAlbaranVentaItem.CantidadT36 <> 0
                .SubItems(2) = "36"
            Case objAlbaranVentaItem.CantidadT38 <> 0
                .SubItems(2) = "38"
            Case objAlbaranVentaItem.CantidadT40 <> 0
                .SubItems(2) = "40"
            Case objAlbaranVentaItem.CantidadT42 <> 0
                .SubItems(2) = "42"
            Case objAlbaranVentaItem.CantidadT44 <> 0
                .SubItems(2) = "44"
            Case objAlbaranVentaItem.CantidadT46 <> 0
                .SubItems(2) = "46"
            Case objAlbaranVentaItem.CantidadT48 <> 0
                .SubItems(2) = "48"
            Case objAlbaranVentaItem.CantidadT50 <> 0
                .SubItems(2) = "50"
            Case objAlbaranVentaItem.CantidadT52 <> 0
                .SubItems(2) = "52"
            Case objAlbaranVentaItem.CantidadT54 <> 0
                .SubItems(2) = "54"
            Case objAlbaranVentaItem.CantidadT56 <> 0
                .SubItems(2) = "56"
            Case Else
            ''ojoojo: devolver error
            End Select
            .SubItems(3) = FormatoMoneda(objAlbaranVentaItem.PrecioVenta, TPVMain.Parametro.Moneda, False)
            .SubItems(4) = objAlbaranVentaItem.Descuento
            .SubItems(5) = objAlbaranVentaItem.Cantidad
            .SubItems(6) = FormatoMoneda(objAlbaranVentaItem.BRUTO, TPVMain.Parametro.Moneda, False)
        End With

    Next
    EnableOK mobjAlbaranVenta.IsValid And mobjAlbaranVenta.AlbaranVentaItems.Count > 0
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

'
'Private Sub ImprimirAlbaranVenta()
'Dim Respuesta As VbMsgBoxResult
'Dim objPrintAlbaran As PrintAlbaran
'Dim frmPrintOptions As frmPrint
'
'    On Error GoTo ErrorManager
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
'        Set objPrintAlbaran = New PrintAlbaran
'        objPrintAlbaran.PrinterNumber = frmPrintOptions.PrinterNumber
'        objPrintAlbaran.Copies = frmPrintOptions.Copies
'        objPrintAlbaran.Component mobjAlbaranVenta
'
'        objPrintAlbaran.PrintObject
'
'        Set objPrintAlbaran = Nothing
'
'        Unload frmPrintOptions
'        Set frmPrintOptions = Nothing
'    End If
'    Exit Sub
'
'ErrorManager:
'    Unload frmPrintOptions
'    ManageErrors (Me.Caption)
'End Sub
Private Sub txtTotalBruto_Change()

End Sub
