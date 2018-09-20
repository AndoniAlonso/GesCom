VERSION 5.00
Begin VB.Form NecesidadesMaterialEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Necesidades de Material"
   ClientHeight    =   4410
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NecesidadesMaterialEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Series"
      Height          =   1335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboSerie 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Text            =   "cboSerie"
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton optSerieTodas 
         Caption         =   "Toda&s"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optSerieUna 
         Caption         =   "&Una serie"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Situación de los pedidos"
      Height          =   1215
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   3015
      Begin VB.OptionButton optSituacionServidos 
         Caption         =   "Ser&vidos"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optSituacionTodos 
         Caption         =   "To&dos los pedidos"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optSituacionPendientes 
         Caption         =   "P&endientes de servir"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de material"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optTipoOtros 
         Caption         =   "&Otros materiales"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optTipoTela 
         Caption         =   "&Telas"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   5400
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   4200
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   3015
      Begin VB.TextBox txtPedidoFinal 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPedidoInicial 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optPedidosRango 
         Caption         =   "Pedidos e&ntre el"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optPedidosTodos 
         Caption         =   "Todos los &pedidos"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "y"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "NecesidadesMaterialEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjNecesidadesMaterial As NecesidadesMaterial
Attribute mobjNecesidadesMaterial.VB_VarHelpID = -1

Public Sub Component(NecesidadesMaterialObject As NecesidadesMaterial)

    Set mobjNecesidadesMaterial = NecesidadesMaterialObject

End Sub

Private Sub cboSerie_Click()
  
    On Error GoTo ErrorManager

    mobjNecesidadesMaterial.Serie = cboSerie.Text
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()
Dim objPrintNecesidadesMaterial As PrintNecesidadesMaterial
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
        
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjNecesidadesMaterial.Load

    Set objPrintNecesidadesMaterial = New PrintNecesidadesMaterial
    objPrintNecesidadesMaterial.PrinterNumber = frmPrintOptions.PrinterNumber
    objPrintNecesidadesMaterial.Copies = frmPrintOptions.Copies
    objPrintNecesidadesMaterial.Component mobjNecesidadesMaterial
    objPrintNecesidadesMaterial.PrintObject
    Set objPrintNecesidadesMaterial = Nothing

    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me

    mflgLoading = True

    With mobjNecesidadesMaterial
        EnableOK .IsValid

        Caption = "Fichas de Pedidos"

        ' Aquí se vuelcan los campos del objeto al interfaz
        txtPedidoInicial = .pedidoinicial
        txtPedidoFinal = .pedidofinal
        txtPedidoInicial.Enabled = False
        txtPedidoFinal.Enabled = False

        optTipoTela.Value = True
        optPedidosTodos.Value = True
        optSituacionPendientes.Value = True
        
        optSerieTodas.Value = True
        cboSerie.Enabled = False
        
        LoadCombo cboSerie, .Series
        cboSerie.Text = .Serie

    End With

    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid

End Sub

Private Sub mobjNecesidadesMaterial_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub optSerieTodas_Click()
    
    cboSerie.Enabled = False
    mobjNecesidadesMaterial.Serie = "(Seleccionar uno)"
    
End Sub

Private Sub optSerieUna_Click()

    cboSerie.Enabled = True

End Sub

Private Sub txtPedidoInicial_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoInicial, mobjNecesidadesMaterial, "PedidoInicial"

End Sub

Private Sub txtPedidoInicial_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoInicial

End Sub

Private Sub txtPedidoInicial_LostFocus()

    txtPedidoInicial = TextLostFocus(txtPedidoInicial, mobjNecesidadesMaterial, "PedidoInicial")
    If txtPedidoFinal = 0 Then txtPedidoFinal = txtPedidoInicial

End Sub

Private Sub txtPedidoFinal_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoFinal, mobjNecesidadesMaterial, "PedidoFinal"

End Sub

Private Sub txtPedidoFinal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoFinal

End Sub

Private Sub txtPedidoFinal_LostFocus()

    txtPedidoFinal = TextLostFocus(txtPedidoFinal, mobjNecesidadesMaterial, "PedidoFinal")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean

    IsList = False

End Function

' Opciones de si el material es tela u otros
Private Sub optTipoTela_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.TipoTela

End Sub

Private Sub optTipoOtros_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.TipoOtros

End Sub

' Opciones de delimitacion del pedido.
Private Sub optPedidosRango_Click()

    If mflgLoading Then Exit Sub
    txtPedidoInicial.Enabled = True
    txtPedidoFinal.Enabled = True
    txtPedidoInicial.SetFocus

End Sub

Private Sub optPedidosTodos_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.pedidoinicial = 0
    mobjNecesidadesMaterial.pedidofinal = 0
    txtPedidoInicial.Enabled = False
    txtPedidoFinal.Enabled = False

End Sub

' Opciones de si la situacion es todos, servidos o pendientes.
Private Sub optSituacionTodos_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.SituacionTodos

End Sub

Private Sub optSituacionPendientes_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.SituacionPendientes

End Sub

Private Sub optSituacionServidos_Click()

    If mflgLoading Then Exit Sub
    mobjNecesidadesMaterial.SituacionServidos

End Sub

