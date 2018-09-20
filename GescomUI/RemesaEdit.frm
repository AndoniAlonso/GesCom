VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RemesaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remesas de efectos en gestión de cobro"
   ClientHeight    =   7050
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RemesaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "Cobros &pendientes..."
      Height          =   375
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la remesa"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtSituacionComercial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   340
         Width           =   1695
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Text            =   "cboBanco"
         Top             =   340
         Width           =   3735
      End
      Begin VB.TextBox txtNumeroEfectos 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1420
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFechaDomiciliacion 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58261505
         CurrentDate     =   36938
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha domiciliación"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situación comercial"
         Height          =   195
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Banco de domiciliación"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número de efectos"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Importe remesado"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cobros de la remesa"
      Height          =   4575
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   9855
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7560
         TabIndex        =   14
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   13
         Top             =   4080
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwCobrosPagos 
         Height          =   3615
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6376
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
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "RemesaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintBancoSelStart As Integer

Private WithEvents mobjRemesa As Remesa
Attribute mobjRemesa.VB_VarHelpID = -1

Public Sub Component(RemesaObject As Remesa)

    Set mobjRemesa = RemesaObject

End Sub

Private Sub cboBanco_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    mobjRemesa.Banco = cboBanco.Text
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjRemesa.SetDatosRemesa
    mobjRemesa.ApplyEdit
    mobjRemesa.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    mobjRemesa.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjRemesa.SetDatosRemesa
    mobjRemesa.ApplyEdit
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFechaDomiciliacion_Change()
    
    mobjRemesa.FechaDomiciliacion = dtpFechaDomiciliacion.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjRemesa
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Remesa [(nuevo)]"

        Else
            Caption = "Remesa [" & .Banco & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        dtpFechaDomiciliacion.Value = .FechaDomiciliacion
        txtSituacionComercial = .SituacionComercial
        txtNumeroEfectos = .NumeroEfectos
        txtImporte = FormatoMoneda(.Importe, GescomMain.objParametro.Moneda)
        
        LoadCombo cboBanco, .Bancos
        cboBanco.Text = .Banco
        
        .BeginEdit GescomMain.objParametro.Moneda
        
    End With
    
    lvwCobrosPagos.SmallIcons = GescomMain.mglIconosPequeños
    
'    lvwCobrosPagos.ColumnHeaders.Add , , "CobroID", ColumnSize(6)
    lvwCobrosPagos.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwCobrosPagos.ColumnHeaders.Add , , "Nº Giro", ColumnSize(4), vbRightJustify
    lvwCobrosPagos.ColumnHeaders.Add , , "Cliente", ColumnSize(15)
    lvwCobrosPagos.ColumnHeaders.Add , , "Importe", ColumnSize(10), vbRightJustify
    lvwCobrosPagos.ColumnHeaders.Add , , "Vencimiento", ColumnSize(10)
    LoadCobrosPagos
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwCobrosPagos_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjRemesa_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub


' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

Private Sub cmdEdit_Click()

    Dim frmCobroPago As CobroPagoEdit
  
    On Error GoTo ErrorManager
    Set frmCobroPago = New CobroPagoEdit
    frmCobroPago.Component _
        mobjRemesa.CobrosPagos(Val(lvwCobrosPagos.SelectedItem.Key))
    frmCobroPago.Show vbModal
    'frmCobroPago.Show
    LoadCobrosPagos
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    mobjRemesa.CobrosPagos.RemoveRemesa Val(lvwCobrosPagos.SelectedItem.Key)
    
    LoadCobrosPagos
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadCobrosPagos()

    Dim objCobroPago As CobroPago
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwCobrosPagos.ListItems.Clear
    For lngIndex = 1 To mobjRemesa.CobrosPagos.Count
        Set itmList = lvwCobrosPagos.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objCobroPago = mobjRemesa.CobrosPagos(lngIndex)

        With itmList
'            If objCobroPago.IsNew Then
'                .Text = "(new)"
'
'            Else
'                .Text = objCobroPago.CobroPagoID
'
'            End If

            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("CobroPago").Key
            
            If objCobroPago.IsDeletedremesa Then
'                .Text = .Text & " (b)"
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            End If
            
            .SubItems(1) = objCobroPago.NumeroGiro
            .SubItems(2) = Trim(objCobroPago.Persona)
            .SubItems(3) = FormatoMoneda(objCobroPago.Importe, GescomMain.objParametro.Moneda)
            .SubItems(4) = objCobroPago.Vencimiento
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub cmdPedidos_Click()
   
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
  
    On Error GoTo ErrorManager
  
    Set frmPicker = New PickerList
  
    frmPicker.LoadData "vCobrosPendientes", 0, 0, 0
    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        ' Primero hay que comprobar que no está ya seleccionado anteriormente
        If Not DocumentoSeleccionado(objPickerItemDisplay.DocumentoID) Then _
            RemesarCobro (objPickerItemDisplay.DocumentoID)
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
      
    LoadCobrosPagos
    
    txtNumeroEfectos.Text = mobjRemesa.CobrosPagos.Count
    txtImporte.Text = mobjRemesa.CobrosPagos.Importe
  
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub RemesarCobro(CobroPagoID As Long)

    Dim objCobroAnterior As CobroPago
    
    Set objCobroAnterior = New CobroPago
    
    objCobroAnterior.Load CobroPagoID, GescomMain.objParametro.Moneda
    
    mobjRemesa.CobrosPagos.AddCobroPago objCobroAnterior

    Set objCobroAnterior = Nothing

End Sub

Private Function DocumentoSeleccionado(DocumentoID As Long) As Boolean
Dim objCobroPago As CobroPago
' Se trata de buscar si existe alguna referencia de ese documento en alguna linea de
' albaranes y es nueva (no se ha actualizado).
    For Each objCobroPago In mobjRemesa.CobrosPagos
        If objCobroPago.CobroPagoID = DocumentoID Then
           DocumentoSeleccionado = True
           Exit Function
        End If
    Next
    
    DocumentoSeleccionado = False
    
    Set objCobroPago = Nothing
End Function

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintBancoSelStart = cboBanco.SelStart
End Sub

Private Sub cboBanco_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintBancoSelStart, cboBanco
    
End Sub

