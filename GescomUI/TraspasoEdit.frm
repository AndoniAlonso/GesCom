VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TraspasoEdit 
   BackColor       =   &H80000005&
   Caption         =   "Traspaso de artículos almacen/tiendas"
   ClientHeight    =   5580
   ClientLeft      =   2985
   ClientTop       =   2910
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TraspasoEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   11385
   Begin VB.ComboBox cboAlmacenDestino 
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
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Text            =   "cboAlmacenDestino"
      Top             =   120
      Width           =   4455
   End
   Begin VB.ComboBox cboAlmacenOrigen 
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
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Text            =   "cboAlmacenOrigen"
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C99497&
      Caption         =   "&Editar"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C99497&
      Caption         =   "El&iminar"
      Height          =   375
      Left            =   10320
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtCantidad 
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ListBox lstIncidencias 
      BackColor       =   &H00E7CDCD&
      ForeColor       =   &H00591E1E&
      Height          =   840
      Left            =   0
      TabIndex        =   12
      Top             =   4680
      Width           =   6735
   End
   Begin VB.CommandButton cmdClearIncidencias 
      BackColor       =   &H00C99497&
      Caption         =   "Borrar incidencia"
      Height          =   612
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
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
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C99497&
      Caption         =   "&Cancelar TRASPASO"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C99497&
      Caption         =   "&Fin TRASPASO"
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
      Left            =   7800
      MaskColor       =   &H00591E1E&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwTraspasoItems 
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   1200
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Destino"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Origen"
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
      Top             =   120
      Width           =   945
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
      TabIndex        =   7
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "Incidencias:"
      ForeColor       =   &H00591E1E&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4320
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
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "TraspasoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintAlmacenOrigenSelStart As Integer
Private mintAlmacenDestinoSelStart As Integer

Private WithEvents mobjTraspaso As Traspaso
Attribute mobjTraspaso.VB_VarHelpID = -1

Private mResize As clsResize

Public Sub Component(TraspasoObject As Traspaso)

    Set mobjTraspaso = TraspasoObject

End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager
    
    ' Si hay lineas preguntamos para ver si está seguro
    If mobjTraspaso.TraspasoItems.Count > 0 Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta <> vbYes Then
            Exit Sub
        End If
    End If

    mobjTraspaso.CancelEdit
    Set mobjTraspaso = New Traspaso
    Form_Load
    
    'Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdClearIncidencias_Click()
    lstIncidencias.Clear
End Sub

Private Sub cmdOK_Click()
    'Dim blnNuevoAlbaran As Boolean
    
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjTraspaso.ApplyEdit
    
    Set mobjTraspaso = Nothing
    
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Activate()

    txtBarCode.SetFocus

End Sub

Private Sub Form_Load()

    'DisableX Me
    
    mflgLoading = True
    With mobjTraspaso
        EnableOK .IsValid And .TraspasoItems.Count > 0
        
        LoadCombo cboAlmacenOrigen, mobjTraspaso.AlmacenesOrigen
        cboAlmacenOrigen.Text = mobjTraspaso.AlmacenOrigen
    
        LoadCombo cboAlmacenDestino, mobjTraspaso.AlmacenesDestino
        cboAlmacenDestino.Text = mobjTraspaso.AlmacenDestino
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCantidad = .Cantidad
        
        .BeginEdit
        
        If .IsNew Then
            Caption = "Traspaso de artículos entre almacén/tiendas [(nuevo)]"
        Else
            Caption = "Traspaso de artículos entre almacén/tiendas [" & .AlmacenOrigen & "]"
        End If
    
    End With
    
    lvwTraspasoItems.SmallIcons = GescomMain.mglIconosPequeños
    lvwTraspasoItems.ColumnHeaders.Clear
    lvwTraspasoItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwTraspasoItems.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(30)
    lvwTraspasoItems.ColumnHeaders.Add , , "Talla", ColumnSize(6), vbRightJustify
    lvwTraspasoItems.ColumnHeaders.Add , , "Precio", ColumnSize(14), vbRightJustify
    lvwTraspasoItems.ColumnHeaders.Add , , "Descuento", ColumnSize(12), vbRightJustify
    lvwTraspasoItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(11), vbRightJustify
    lvwTraspasoItems.ColumnHeaders.Add , , "Importe", ColumnSize(14), vbRightJustify
    LoadTraspasoItems
    
    lstIncidencias.Clear
  
    mflgLoading = False

    Set mResize = New clsResize
    mResize.Init Me
    
    Me.WindowState = 2 'vbext_ws_Maximize
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid 'And mobjTraspaso.TraspasoItems.Count > 0
'    cmdApply.Enabled = flgValid
    
End Sub

Private Sub Form_Resize()
   mResize.FormResize Me
End Sub

Private Sub lvwTraspasoItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

'Private Sub mfrmCapturaCodigo_FinCaptura()
'    Me.Enabled = True
'End Sub

Private Sub mobjTraspaso_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub


'OJOOJO refactorizar esto que es muy largo.
Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim objTraspasoItem  As TraspasoItem
    Dim strCodigo As String
    Dim lngCodigo As Long
    Dim intTalla As Integer
    Dim lngArticuloColorID As Long
    Dim intOrdenTalla As Integer
    Dim objArticuloColor As ArticuloColor
    Dim strArticuloColor As String
    Dim strIncidencia As String
    Dim strNombreArticuloColor As String
    
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
        
        ' Validar que la información de talla sea correcta
        intTalla = CInt(Left(strCodigo, 2))
        If intTalla Mod 2 <> 0 Then
            lstIncidencias.AddItem "Talla errónea! (" & intTalla & "), en el código " & strCodigo
            Exit Sub
        End If
        If intTalla > 56 Or intTalla < 36 Then
            lstIncidencias.AddItem "Talla errónea! (" & intTalla & "), en el código " & strCodigo
            Exit Sub
        End If
        
        
        intOrdenTalla = (intTalla - 36) / 2
        
        ' Cargar el artículo, etc
        Set objArticuloColor = New ArticuloColor
        lngArticuloColorID = CLng(Right(strCodigo, Len(strCodigo) - 2))
        objArticuloColor.Load lngArticuloColorID, "EUR"
        strArticuloColor = objArticuloColor.Nombre
        strNombreArticuloColor = objArticuloColor.Articulo + " " + objArticuloColor.NombreColor
        'lngTemporadaArticulo = objArticuloColor.TemporadaID
        Set objArticuloColor = Nothing
        
        strIncidencia = mobjTraspaso.TraspasoItemCodigoBarras(intTalla, lngArticuloColorID)
        If strIncidencia <> vbNullString Then
            lstIncidencias.AddItem strIncidencia & " Código:" & strCodigo
        End If
        
        Beep
        LoadTraspasoItems
        txtCantidad = mobjTraspaso.Cantidad
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

' a partir de aqui -----> child

'Private Sub cmdAdd_Click()
'    Dim frmTraspasoItemEdit As TraspasoItemEdit
'
'    On Error GoTo ErrorManager
'    Set frmTraspasoItemEdit = New TraspasoItemEdit
'    frmTraspasoItemEdit.Component mobjTraspaso.TraspasoItems.Add
'    frmTraspasoItemEdit.Show vbModal
'    LoadTraspasoItems
'    txtCantidad = FormatoMoneda(mobjTraspaso.TotalBruto, GescomMain.Parametro.Moneda)
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub

Private Sub cmdEdit_Click()
    Dim frmTraspasoItemEdit As TraspasoItemEdit
    
    On Error GoTo ErrorManager
    If lvwTraspasoItems.ListItems.Count = 0 Then Exit Sub
  
    Set frmTraspasoItemEdit = New TraspasoItemEdit
    frmTraspasoItemEdit.Component _
        mobjTraspaso.TraspasoItems(Val(lvwTraspasoItems.SelectedItem.Key))
    frmTraspasoItemEdit.Show vbModal
    LoadTraspasoItems
    txtCantidad = mobjTraspaso.Cantidad
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
    
    If lvwTraspasoItems.ListItems.Count = 0 Then Exit Sub
        
    mobjTraspaso.TraspasoItems.Remove Val(lvwTraspasoItems.SelectedItem.Key)
    LoadTraspasoItems
    txtCantidad = mobjTraspaso.Cantidad

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadTraspasoItems()
    Dim objTraspasoItem As TraspasoItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwTraspasoItems.ListItems.Clear
    For lngIndex = 1 To mobjTraspaso.TraspasoItems.Count
        Set itmList = lvwTraspasoItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objTraspasoItem = mobjTraspaso.TraspasoItems(lngIndex)

        With itmList
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objTraspasoItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = objTraspasoItem.ArticuloColorID ' IIf(objTraspasoItem.ArticuloColorID, Trim(objTraspasoItem.ArticuloColor), Trim(objTraspasoItem.Descripcion))
            
            Select Case True
            Case objTraspasoItem.CantidadT36 <> 0
                .SubItems(2) = "36"
            Case objTraspasoItem.CantidadT38 <> 0
                .SubItems(2) = "38"
            Case objTraspasoItem.CantidadT40 <> 0
                .SubItems(2) = "40"
            Case objTraspasoItem.CantidadT42 <> 0
                .SubItems(2) = "42"
            Case objTraspasoItem.CantidadT44 <> 0
                .SubItems(2) = "44"
            Case objTraspasoItem.CantidadT46 <> 0
                .SubItems(2) = "46"
            Case objTraspasoItem.CantidadT48 <> 0
                .SubItems(2) = "48"
            Case objTraspasoItem.CantidadT50 <> 0
                .SubItems(2) = "50"
            Case objTraspasoItem.CantidadT52 <> 0
                .SubItems(2) = "52"
            Case objTraspasoItem.CantidadT54 <> 0
                .SubItems(2) = "54"
            Case objTraspasoItem.CantidadT56 <> 0
                .SubItems(2) = "56"
            Case Else
                Err.Raise vbObjectError + 1001, "TraspasoEdit LoadTraspasoItems", "No hay cantidad en ninguna talla!"
            End Select
            '.SubItems(3) = FormatoMoneda(objTraspasoItem.PrecioVenta, GescomMain.Parametro.Moneda, False)
            '.SubItems(4) = objTraspasoItem.Descuento
            .SubItems(5) = objTraspasoItem.Cantidad
            '.SubItems(6) = FormatoMoneda(objTraspasoItem.Bruto, GescomMain.Parametro.Moneda, False)
        End With

    Next
    EnableOK mobjTraspaso.IsValid And mobjTraspaso.TraspasoItems.Count > 0
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboAlmacenOrigen_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    
    mobjTraspaso.AlmacenOrigen = cboAlmacenOrigen.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboAlmacenOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintAlmacenOrigenSelStart = cboAlmacenOrigen.SelStart
End Sub

Private Sub cboAlmacenOrigen_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintAlmacenOrigenSelStart, cboAlmacenOrigen
    
End Sub

Private Sub cboAlmacenDestino_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    
    mobjTraspaso.AlmacenDestino = cboAlmacenDestino.Text
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboAlmacenDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintAlmacenDestinoSelStart = cboAlmacenDestino.SelStart
End Sub

Private Sub cboAlmacenDestino_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintAlmacenDestinoSelStart, cboAlmacenDestino
    
End Sub


