VERSION 5.00
Begin VB.Form CrearArticuloEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear artículo"
   ClientHeight    =   3105
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CrearArticuloEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPrecios 
      Caption         =   "Precios"
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtPrecioVentaPublico 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtPrecioCompra 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   700
         Width           =   1815
      End
      Begin VB.TextBox txtPrecioVenta 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1060
         Width           =   1815
      End
      Begin VB.TextBox txtPrecioCoste 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   340
         Width           =   1815
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Precio coste"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Precio compra"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Precio venta"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "PVP"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   270
      End
   End
   Begin VB.ComboBox cboTallaje 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "cboTallaje"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tallaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "CrearArticuloEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintTallajeSelStart As Integer

Private mflgSetFocusPrecioVenta As Boolean
Private mflgSetFocusPrecioCompra As Boolean

Private WithEvents mobjArticulo As Articulo
Attribute mobjArticulo.VB_VarHelpID = -1

Public Sub Component(ArticuloObject As Articulo)

    Set mobjArticulo = ArticuloObject

End Sub

Public Sub SetFocusPrecioVenta()
    
    mflgSetFocusPrecioVenta = True
    mflgSetFocusPrecioCompra = False
    
End Sub

Public Sub SetFocusPrecioCompra()
    
    mflgSetFocusPrecioVenta = False
    mflgSetFocusPrecioCompra = True
    
End Sub

Private Sub cmdCancel_Click()

    mobjArticulo.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    
    ' Si se crea un artículo sin precio de venta que se alimente el precio de venta como el PVP
    If mobjArticulo.PrecioCompra <> 0 And _
       mobjArticulo.PrecioVentaPublico <> 0 And _
       mobjArticulo.PrecioVenta = 0 Then
       mobjArticulo.PrecioVenta = mobjArticulo.PrecioVentaPublico
    End If
  
    mobjArticulo.ApplyEdit
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub Form_Activate()
    
    If mflgSetFocusPrecioCompra Then _
        txtPrecioCompra.SetFocus
        
    If mflgSetFocusPrecioVenta Then _
        txtPrecioVenta.SetFocus
    
End Sub

Private Sub Form_Initialize()

    mflgSetFocusPrecioVenta = False
    mflgSetFocusPrecioCompra = False
  
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjArticulo
        EnableOK .IsValid
        
        If .IsNew Then
          Caption = "Articulo [(nuevo)]"
        
        Else
          Caption = "Articulo [" & .NombreCompleto & "]"
        
        End If
        
        txtPrecioCoste = .PrecioCoste
        txtPrecioCompra = .PrecioCompra
        txtPrecioVenta = .PrecioVenta
        txtPrecioVentaPublico = .PrecioVentaPublico
            
        .BeginEdit
        
        mobjArticulo.AsignarTallajePredeterminado
        ' Cargo los datos del combo despues de asignar la temporada porque esta se
        ' carga con los articulos de una temporada
        LoadCombo cboTallaje, .Tallajes
        cboTallaje.Text = .Tallaje
        
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid

End Sub

Private Sub mobjArticulo_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub cboTallaje_Click()

    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    mobjArticulo.Tallaje = cboTallaje.Text
    txtPrecioCoste = mobjArticulo.PrecioCoste
    txtPrecioCompra = mobjArticulo.PrecioCompra
    txtPrecioVenta = mobjArticulo.PrecioVenta
    txtPrecioVentaPublico = mobjArticulo.PrecioVentaPublico
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub
  
Private Sub txtPrecioCoste_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioCoste

End Sub

Private Sub txtPrecioCoste_Change()
    
    If Not mflgLoading Then _
        TextChange txtPrecioCoste, mobjArticulo, "PrecioCoste"

End Sub

Private Sub txtPrecioCoste_LostFocus()

  txtPrecioCoste = TextLostFocus(txtPrecioCoste, mobjArticulo, "PrecioCoste")

  mobjArticulo.PrecioVenta = mobjArticulo.CalcularPrecioVenta
  txtPrecioVenta.Text = mobjArticulo.PrecioVenta

End Sub
  
Private Sub txtPrecioCompra_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioCompra

End Sub

Private Sub txtPrecioCompra_Change()
    
    If Not mflgLoading Then _
        TextChange txtPrecioCompra, mobjArticulo, "PrecioCompra"

End Sub

Private Sub txtPrecioCompra_LostFocus()

  txtPrecioCompra = TextLostFocus(txtPrecioCompra, mobjArticulo, "PrecioCompra")
  
End Sub
  
Private Sub txtPrecioVenta_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioVenta

End Sub

Private Sub txtPrecioVenta_Change()
    
    If Not mflgLoading Then
        TextChange txtPrecioVenta, mobjArticulo, "PrecioVenta"
    End If

End Sub

Private Sub txtPrecioVenta_LostFocus()

  txtPrecioVenta = TextLostFocus(txtPrecioVenta, mobjArticulo, "PrecioVenta")
  
  mobjArticulo.PrecioVentaPublico = mobjArticulo.CalcularPrecioVentaPublico
  txtPrecioVentaPublico.Text = mobjArticulo.PrecioVentaPublico

End Sub
  
Private Sub txtPrecioVentaPublico_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioVentaPublico

End Sub

Private Sub txtPrecioVentaPublico_Change()

  If Not mflgLoading Then
    TextChange txtPrecioVentaPublico, mobjArticulo, "PrecioVentaPublico"
  End If

End Sub

Private Sub txtPrecioVentaPublico_LostFocus()

  txtPrecioVentaPublico = TextLostFocus(txtPrecioVentaPublico, mobjArticulo, "PrecioVentaPublico")

End Sub
  
' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   IsList = False
End Function

Private Sub cboTallaje_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTallajeSelStart = cboTallaje.SelStart
End Sub

Private Sub cboTallaje_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTallajeSelStart, cboTallaje
    
End Sub
