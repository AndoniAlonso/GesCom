VERSION 5.00
Begin VB.Form FichasPedidoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FichaPedidos"
   ClientHeight    =   3480
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FichasPedidoEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Colores"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optColorUno 
         Caption         =   "&Un color"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optColorTodos 
         Caption         =   "&Todos los colores"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades"
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txtPedidoFinal 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPedidoInicial 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optPedidosRango 
         Caption         =   "&Pedidos entre el"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optPedidosTodos 
         Caption         =   "T&odos los pedidos"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "y"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "FichasPedidoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjFichasPedido As FichasPedido
Attribute mobjFichasPedido.VB_VarHelpID = -1

Public Sub Component(fichaspedidoObject As FichasPedido)

    Set mobjFichasPedido = fichaspedidoObject

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
Dim objPrintFichasPedido As PrintFichasPedido
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
  
    mobjFichasPedido.Load
    
    Set objPrintFichasPedido = New PrintFichasPedido
    objPrintFichasPedido.PrinterNumber = frmPrintOptions.PrinterNumber
    objPrintFichasPedido.Copies = frmPrintOptions.Copies
    objPrintFichasPedido.Component mobjFichasPedido
    objPrintFichasPedido.PrintObject
    Set objPrintFichasPedido = Nothing
    
    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    
    With mobjFichasPedido
        EnableOK .IsValid
    
        Caption = "Fichas de Pedidos"

        ' Aquí se vuelcan los campos del objeto al interfaz
        txtColor = .Color
        txtPedidoInicial = .pedidoinicial
        txtPedidoFinal = .pedidofinal
        txtColor.Enabled = False
        txtPedidoInicial.Enabled = False
        txtPedidoFinal.Enabled = False
        
        optColorTodos.Value = True
        optPedidosTodos.Value = True
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid

End Sub

Private Sub mobjfichaspedido_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtColor_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtColor
        
End Sub

Private Sub txtColor_Change()

    If Not mflgLoading Then _
        TextChange txtColor, mobjFichasPedido, "Color"

End Sub

Private Sub txtColor_LostFocus()

    txtColor = TextLostFocus(txtColor, mobjFichasPedido, "Color")

End Sub

Private Sub txtPedidoInicial_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoInicial, mobjFichasPedido, "PedidoInicial"

End Sub

Private Sub txtPedidoInicial_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoInicial
        
End Sub

Private Sub txtPedidoInicial_LostFocus()

    txtPedidoInicial = TextLostFocus(txtPedidoInicial, mobjFichasPedido, "PedidoInicial")
    If txtPedidoFinal = 0 Then txtPedidoFinal = txtPedidoInicial

End Sub

Private Sub txtPedidoFinal_Change()

    If Not mflgLoading Then _
        TextChange txtPedidoFinal, mobjFichasPedido, "PedidoFinal"

End Sub

Private Sub txtPedidoFinal_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPedidoFinal
        
End Sub

Private Sub txtPedidoFinal_LostFocus()

    txtPedidoFinal = TextLostFocus(txtPedidoFinal, mobjFichasPedido, "PedidoFinal")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

Private Sub optColorTodos_Click()
    
    If mflgLoading Then Exit Sub
    mobjFichasPedido.Color = vbNullString
    txtColor.Enabled = False
        
End Sub

Private Sub optColorUno_Click()

    If mflgLoading Then Exit Sub
    txtColor.Enabled = True
    txtColor.SetFocus

End Sub

Private Sub optPedidosRango_Click()

    If mflgLoading Then Exit Sub
    txtPedidoInicial.Enabled = True
    txtPedidoFinal.Enabled = True
    txtPedidoInicial.SetFocus

End Sub

Private Sub optPedidosTodos_Click()

    If mflgLoading Then Exit Sub
    mobjFichasPedido.pedidoinicial = 0
    mobjFichasPedido.pedidofinal = 0
    txtPedidoInicial.Enabled = False
    txtPedidoFinal.Enabled = False

End Sub
