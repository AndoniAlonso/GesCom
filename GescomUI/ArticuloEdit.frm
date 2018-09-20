VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ArticuloEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   4920
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ArticuloEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvwArticulo 
      Height          =   4215
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7435
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Artículo"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox cboTallaje 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Text            =   "cboTallaje"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboPrenda 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cboPrenda"
         Top             =   340
         Width           =   2535
      End
      Begin VB.ComboBox cboModelo 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Text            =   "cboModelo"
         Top             =   700
         Width           =   2535
      End
      Begin VB.ComboBox cboSerie 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Text            =   "cboSerie"
         Top             =   1060
         Width           =   2535
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tallaje"
         Height          =   195
         Left            =   4320
         TabIndex        =   7
         Top             =   735
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de prenda"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9480
      TabIndex        =   36
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   35
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   34
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stocks"
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   7215
      Begin VB.TextBox txtSuReferencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Precios"
         Height          =   1815
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txtPrecioCompra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPrecioVentaPublico 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   30
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtPrecioVenta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   28
            Top             =   1065
            Width           =   1455
         End
         Begin VB.TextBox txtPrecioCoste 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   340
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Precio de compra"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "PVP"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1455
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Precio de coste"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Precio de venta"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.TextBox txtLoteEconomico 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   1780
         Width           =   1455
      End
      Begin VB.TextBox txtStockMaximo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   1420
         Width           =   1455
      End
      Begin VB.TextBox txtStockMinimo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   1060
         Width           =   1455
      End
      Begin VB.TextBox txtStockPendiente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   700
         Width           =   1455
      End
      Begin VB.TextBox txtStockActual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   340
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Su referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   2175
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Stock pendiente"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Lote económico"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Stock máximo"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Stock mínimo"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stock actual"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "ArticuloEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean
Private mflgAutorized As Boolean

Private mintPrendaSelStart As Integer
Private mintModeloSelStart As Integer
Private mintSerieSelStart As Integer
Private mintTallajeSelStart As Integer

Private WithEvents mobjArticulo As Articulo
Attribute mobjArticulo.VB_VarHelpID = -1

Public Sub Component(Articulobject As Articulo)

    Set mobjArticulo = Articulobject

End Sub

Private Sub cboPrenda_Click()
  
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjArticulo.Prenda = cboPrenda.Text
    
    If cboPrenda.Text <> "(Seleccionar uno)" _
        And cboModelo.Text <> "(Seleccionar uno)" _
        And cboSerie.Text <> "(Seleccionar uno)" Then
        
        mobjArticulo.CalcularPrecioCoste
        
        mflgAutorized = True
        txtNombre = mobjArticulo.Nombre
        txtPrecioCoste = mobjArticulo.PrecioCoste
        txtPrecioCompra = mobjArticulo.PrecioCompra
        txtPrecioVenta = mobjArticulo.PrecioVenta
        txtPrecioVentaPublico = mobjArticulo.PrecioVentaPublico
        mflgAutorized = False
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboModelo_Click()
  
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjArticulo.Modelo = cboModelo.Text
    
    If cboPrenda.Text <> "(Seleccionar uno)" _
        And cboModelo.Text <> "(Seleccionar uno)" _
        And cboSerie.Text <> "(Seleccionar uno)" Then
        
        mobjArticulo.CalcularPrecioCoste
        
        mflgAutorized = True
        txtNombre = mobjArticulo.Nombre
        txtPrecioCoste = mobjArticulo.PrecioCoste
        txtPrecioCompra = mobjArticulo.PrecioCompra
        txtPrecioVenta = mobjArticulo.PrecioVenta
        txtPrecioVentaPublico = mobjArticulo.PrecioVentaPublico
        mflgAutorized = False
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboSerie_Click()
  
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjArticulo.Serie = cboSerie.Text
    
    If cboPrenda.Text <> "(Seleccionar uno)" _
        And cboModelo.Text <> "(Seleccionar uno)" _
        And cboSerie.Text <> "(Seleccionar uno)" Then
        
        mobjArticulo.CalcularPrecioCoste
        
        mflgAutorized = True
        txtNombre = mobjArticulo.Nombre
        txtPrecioCoste = mobjArticulo.PrecioCoste
        txtPrecioCompra = mobjArticulo.PrecioCompra
        txtPrecioVenta = mobjArticulo.PrecioVenta
        txtPrecioVentaPublico = mobjArticulo.PrecioVentaPublico
        mflgAutorized = False
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()
    Dim bolComboLocked As Boolean
    
    On Error GoTo ErrorManager

    mobjArticulo.ApplyEdit
    txtNombre = mobjArticulo.Nombre
    txtPrecioCoste = mobjArticulo.PrecioCoste
    txtPrecioCompra = mobjArticulo.PrecioCompra
    txtPrecioVenta = mobjArticulo.PrecioVenta
    txtPrecioVentaPublico = mobjArticulo.PrecioVentaPublico
    
    ' si no es un articulo nuevo no dejo modificar los combo(prenda,modelo,serie)
    bolComboLocked = Not mobjArticulo.IsNew
    cboPrenda.Locked = bolComboLocked
    cboModelo.Locked = bolComboLocked
    cboSerie.Locked = bolComboLocked
    'cboTallaje.Locked = bolComboLocked
        
    mobjArticulo.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjArticulo.IsDirty And Not mobjArticulo.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjArticulo.CancelEdit
            Unload Me
        End If
    Else
        mobjArticulo.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjArticulo.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub


Private Sub Form_Load()
    Dim bolComboLocked As Boolean

    DisableX Me
    
    tvwArticulo.ImageList = GescomMain.mglIconosPequeños
    
    mflgLoading = True
    With mobjArticulo
        EnableOK .IsValid

        If .IsNew Then
            Caption = "Artículo [(nuevo)]"
            ' Asigno la temporada antes de cargar los "combo"
            .TemporadaID = GescomMain.objParametro.TemporadaActualID

        Else
            Caption = "Artículo [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre.Text = .Nombre
    
        LoadCombo cboPrenda, .Prendas
        cboPrenda.Text = .Prenda
        
        LoadCombo cboModelo, .Modelos
        cboModelo.Text = .Modelo
    
        LoadCombo cboSerie, .Series
        cboSerie.Text = .Serie
    
        txtStockActual.Text = .StockActual
        txtStockPendiente.Text = .StockPendiente
        txtStockMinimo.Text = .StockMinimo
        txtStockMaximo.Text = .StockMaximo
        txtLoteEconomico.Text = .LoteEconomico
        txtSuReferencia.Text = .SuReferencia
        
        .BeginEdit
    
        If .IsNew Then
            .AsignarTallajePredeterminado
        End If
        
        LoadCombo cboTallaje, .Tallajes
        cboTallaje.Text = .Tallaje
    
        txtPrecioCoste.Text = .PrecioCoste
        txtPrecioCompra.Text = .PrecioCompra
        txtPrecioVenta.Text = .PrecioVenta
        txtPrecioVentaPublico.Text = .PrecioVentaPublico
        
        ' si no es un articulo nuevo no dejo modificar los combo(prenda,modelo,serie)
        bolComboLocked = Not mobjArticulo.IsNew
        cboPrenda.Locked = bolComboLocked
        cboModelo.Locked = bolComboLocked
        cboSerie.Locked = bolComboLocked
        'cboTallaje.Locked = bolComboLocked
        
        ' De momento dejo el campo su referencia como no editable
        txtSuReferencia.Locked = True
        
        ' Muestro la ficha de coste de artículos.
        FichaArticulo
    
    End With
  
    mflgLoading = False

End Sub

Private Sub FichaArticulo()
    Dim ndArticuloVenta As Node
    Dim ndArticuloPVP As Node
    Dim ndArticuloCoste As Node
    Dim ndAdministracion As Node
    Dim ndPrenda As Node
    Dim ndPrendaItem As Node
    Dim ndSerie As Node
    Dim dblPrecioSerie As Double
    Dim ndModelo As Node
    Dim objMaterial As Material
    Dim objprenda As Prenda
    
    ' Evidentemente no se hace nada si no hay composiciones
    If mobjArticulo.objprenda Is Nothing Or _
       mobjArticulo.objModelo Is Nothing Or _
       mobjArticulo.objSerie Is Nothing Then Exit Sub
    
    ' El material de la serie debe estar definido
    If mobjArticulo.objSerie.objMaterial(GescomMain.objParametro.Moneda) Is Nothing Then _
        Exit Sub
    
    Set objprenda = mobjArticulo.objprenda
    With tvwArticulo
        .Sorted = True
        .LabelEdit = False
        .Nodes.Clear
        Set ndArticuloVenta = .Nodes.Add()
        ndArticuloVenta.Text = "Precio venta :" & FormatoMoneda(mobjArticulo.PrecioVenta, GescomMain.objParametro.Moneda) & _
                          " (" & mobjArticulo.objModelo.Beneficio & "% de beneficio)"
        ndArticuloVenta.Image = "Articulo"
        
        Set ndArticuloPVP = .Nodes.Add()
        Set ndArticuloPVP.Parent = ndArticuloVenta
        ndArticuloPVP.Text = "PVP:" & FormatoMoneda(mobjArticulo.PrecioVentaPublico, GescomMain.objParametro.Moneda) & _
                          " (" & mobjArticulo.objModelo.BeneficioPVP & "% de beneficio PVP)"
        ndArticuloPVP.Image = "Articulo"
        Set ndArticuloPVP = Nothing
        
        Set ndArticuloCoste = .Nodes.Add()
        Set ndArticuloCoste.Parent = ndArticuloVenta
        ndArticuloCoste = "Precio coste :" & FormatoMoneda(mobjArticulo.PrecioCoste, GescomMain.objParametro.Moneda)
        ndArticuloCoste.Image = "Articulo"
        
        Set ndAdministracion = .Nodes.Add()
        Set ndAdministracion.Parent = ndArticuloCoste
        ndAdministracion = "Costes administracion: " & objprenda.Administracion & "%"
        ndAdministracion.Image = "Prenda"
        
        Set ndPrenda = .Nodes.Add()
        Set ndPrenda.Parent = ndArticuloCoste
        ndPrenda = "Costes prenda: " & FormatoMoneda(objprenda.PrecioCoste, GescomMain.objParametro.Moneda)
        ndPrenda.Image = "Prenda"
        
        Set ndPrendaItem = .Nodes.Add()
        Set ndPrendaItem.Parent = ndPrenda
        ndPrendaItem = "Cartón: " & FormatoMoneda(objprenda.Carton, GescomMain.objParametro.Moneda)
        ndPrendaItem.Image = "Prenda"
        
        Set ndPrendaItem = .Nodes.Add()
        Set ndPrendaItem.Parent = ndPrenda
        ndPrendaItem = "Etiqueta: " & FormatoMoneda(objprenda.Etiqueta, GescomMain.objParametro.Moneda)
        ndPrendaItem.Image = "Prenda"
        
        Set ndPrendaItem = .Nodes.Add()
        Set ndPrendaItem.Parent = ndPrenda
        ndPrendaItem = "Percha: " & FormatoMoneda(objprenda.percha, GescomMain.objParametro.Moneda)
        ndPrendaItem.Image = "Prenda"
        
        Set ndPrendaItem = .Nodes.Add()
        Set ndPrendaItem.Parent = ndPrenda
        ndPrendaItem = "Plancha: " & FormatoMoneda(objprenda.Plancha, GescomMain.objParametro.Moneda)
        ndPrendaItem.Image = "Prenda"
        
        Set ndPrendaItem = .Nodes.Add()
        Set ndPrendaItem.Parent = ndPrenda
        ndPrendaItem = "Transporte: " & FormatoMoneda(objprenda.transporte, GescomMain.objParametro.Moneda)
        ndPrendaItem.Image = "Prenda"
        
        ' Cargamos el material de la serie
        Set objMaterial = mobjArticulo.objSerie.objMaterial(GescomMain.objParametro.Moneda)
        dblPrecioSerie = (objMaterial.PrecioCoste * _
                         mobjArticulo.objModelo.CantidadTela * _
                         objMaterial.AnchuraEstandar) / _
                         objMaterial.AnchuraTela
        dblPrecioSerie = Round(dblPrecioSerie, 2)
                         
        Set ndSerie = .Nodes.Add()
        Set ndSerie.Parent = ndArticuloCoste
        ndSerie = "Serie: " & _
            FormatoMoneda(dblPrecioSerie, GescomMain.objParametro.Moneda) & _
            "(" & objMaterial.Codigo & " " & _
            CStr(Round((mobjArticulo.objModelo.CantidadTela * _
                  objMaterial.AnchuraEstandar) / _
                  objMaterial.AnchuraTela, 2)) _
            & objMaterial.UnidadMedida & _
            "*" & objMaterial.PrecioCoste & ")"
        ndSerie.Image = "Serie"
        Set objMaterial = Nothing
        Set objprenda = Nothing
        
        
        Set ndModelo = .Nodes.Add()
        Set ndModelo.Parent = ndArticuloCoste
        ndModelo = "Modelo: " & _
            FormatoMoneda(mobjArticulo.objModelo.PrecioCoste, GescomMain.objParametro.Moneda)
        ndModelo.Image = "Modelo"
        
    End With
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjArticulo_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtLoteEconomico_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtLoteEconomico
        
End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtPrecioCoste_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCoste
        
End Sub

Private Sub txtPrecioCompra_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCompra
        
End Sub

Private Sub txtPrecioVenta_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioVenta
        
End Sub

Private Sub txtPrecioVentaPublico_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioVentaPublico
        
End Sub

Private Sub txtStockMaximo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockMaximo
        
End Sub

Private Sub txtStockMinimo_Change()

    If Not mflgLoading Then _
        TextChange txtStockMinimo, mobjArticulo, "StockMinimo"

End Sub

Private Sub txtStockMinimo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockMinimo
        
End Sub

Private Sub txtStockMinimo_LostFocus()

    txtStockMinimo = TextLostFocus(txtStockMinimo, mobjArticulo, "StockMinimo")

End Sub

Private Sub txtStockMaximo_Change()

    If Not mflgLoading Then _
        TextChange txtStockMaximo, mobjArticulo, "StockMaximo"

End Sub

Private Sub txtStockMaximo_LostFocus()

    txtStockMaximo = TextLostFocus(txtStockMaximo, mobjArticulo, "StockMaximo")

End Sub

Private Sub txtLoteEconomico_Change()

    If Not mflgLoading Then _
        TextChange txtLoteEconomico, mobjArticulo, "LoteEconomico"

End Sub

Private Sub txtLoteEconomico_LostFocus()

    txtLoteEconomico = TextLostFocus(txtLoteEconomico, mobjArticulo, "LoteEconomico")

End Sub

Private Sub txtPrecioCoste_Change()

    'If Not mflgLoading And Not mflgAutorized Then
    If Not mflgLoading Then _
        TextChange txtPrecioCoste, mobjArticulo, "PrecioCoste"

End Sub

Private Sub txtPrecioCoste_LostFocus()

    txtPrecioCoste = TextLostFocus(txtPrecioCoste, mobjArticulo, "PrecioCoste")

End Sub

Private Sub txtPrecioCompra_Change()

    'If Not mflgLoading And Not mflgAutorized Then
    If Not mflgLoading Then _
        TextChange txtPrecioCompra, mobjArticulo, "PrecioCompra"

End Sub

Private Sub txtPrecioCompra_LostFocus()

    txtPrecioCompra.Text = TextLostFocus(txtPrecioCompra, mobjArticulo, "PrecioCompra")

End Sub

Private Sub txtPrecioVenta_Change()

    If Not mflgLoading And Not mflgAutorized Then _
        TextChange txtPrecioVenta, mobjArticulo, "PrecioVenta"

End Sub

Private Sub txtPrecioVenta_LostFocus()

    txtPrecioVenta = TextLostFocus(txtPrecioVenta, mobjArticulo, "PrecioVenta")

End Sub

Private Sub txtPrecioVentaPublico_Change()

    If Not mflgLoading And Not mflgAutorized Then _
        TextChange txtPrecioVentaPublico, mobjArticulo, "PrecioVentaPublico"

End Sub

Private Sub txtPrecioVentaPublico_LostFocus()

    txtPrecioVentaPublico = TextLostFocus(txtPrecioVentaPublico, mobjArticulo, "PrecioVentaPublico")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

Private Sub cboPrenda_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintPrendaSelStart = cboPrenda.SelStart
End Sub

Private Sub cboPrenda_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintPrendaSelStart, cboPrenda
    
End Sub

Private Sub cboModelo_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintModeloSelStart = cboModelo.SelStart
End Sub

Private Sub cboModelo_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintModeloSelStart, cboModelo
    
End Sub

Private Sub cboSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintSerieSelStart = cboSerie.SelStart
End Sub

Private Sub cboSerie_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintSerieSelStart, cboSerie
    
End Sub

Private Sub cboTallaje_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTallajeSelStart = cboTallaje.SelStart
End Sub

Private Sub cboTallaje_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTallajeSelStart, cboTallaje
    
End Sub

Private Sub cboTallaje_Click()
  
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjArticulo.Tallaje = cboTallaje.Text
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub



