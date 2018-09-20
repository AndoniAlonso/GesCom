VERSION 5.00
Begin VB.Form ParametroEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros"
   ClientHeight    =   4290
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ParametroEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetos Generales de la Aplicación"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtSufijo 
         Height          =   285
         Left            =   4440
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtMoneda 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtServidorContawin 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox txtDireccion 
         Height          =   1005
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   2265
         Width           =   3255
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cboTemporada 
         Height          =   315
         ItemData        =   "ParametroEdit.frx":030A
         Left            =   4440
         List            =   "ParametroEdit.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   705
         Width           =   2055
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   315
         Width           =   2055
      End
      Begin VB.TextBox txtAlfanumero 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1420
         Width           =   1575
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1060
         Width           =   1575
      End
      Begin VB.TextBox txtPropietario 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   700
         Width           =   1575
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   340
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sufijo facturas:"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   1455
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servidor Contawin"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Empresa Inicial"
         Height          =   195
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temporada Inicial"
         Height          =   195
         Left            =   3000
         TabIndex        =   7
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Alfanúmero"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Propietario"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   405
      End
   End
End
Attribute VB_Name = "ParametroEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjParametro As Parametro
Attribute mobjParametro.VB_VarHelpID = -1

Public Sub Component(ParametroObject As Parametro)

    Set mobjParametro = ParametroObject

End Sub

Private Sub btnDireccion_Click()
    
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjParametro.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjParametro.Direccion.DireccionText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjParametro.ApplyEdit
    mobjParametro.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjParametro.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjParametro.ApplyEdit
    Unload Me
    Screen.MousePointer = vbDefault
    GescomMain.GescomTitulo
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorManager

    DisableX Me
    
    mflgLoading = True
    With mobjParametro
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Parámetro [(nuevo)]"
    
        Else
            Caption = "Parámetro [" & .Propietario & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtParametroID = .ParametroID
        txtAlfanumero = .Alfanumero
        txtClave = .Clave
        txtPropietario = .Propietario
        txtUsuario = .Usuario
        'txtEmpresaID = .EmpresaID
        'txtTemporadaID = .TemporadaID
        txtMoneda = .Moneda
        txtDireccion = .Direccion.DireccionText
        txtServidorContawin = .ServidorContawin
        txtSufijo.Text = .Sufijo
      
        LoadCombo cboEmpresa, .Empresas
        cboEmpresa.Text = .Empresa
    
        LoadCombo cboTemporada, .Temporadas
        cboTemporada.Text = .Temporada
    
        .BeginEdit
    
    End With
  
    mflgLoading = False
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboEmpresa_Click()

    If mflgLoading Then Exit Sub
    mobjParametro.Empresa = cboEmpresa.Text

End Sub

Private Sub cboTemporada_Click()

    If mflgLoading Then Exit Sub
    mobjParametro.Temporada = cboTemporada.Text

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjParametro_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtAlfanumero_Change()

    If Not mflgLoading Then _
        TextChange txtAlfanumero, mobjParametro, "Alfanumero"

End Sub

Private Sub txtAlfanumero_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtAlfanumero
        
End Sub

Private Sub txtAlfanumero_LostFocus()

    txtAlfanumero = TextLostFocus(txtAlfanumero, mobjParametro, "Alfanumero")

End Sub

Private Sub txtSufijo_Change()

    If Not mflgLoading Then _
        TextChange txtSufijo, mobjParametro, "Sufijo"

End Sub

Private Sub txtSufijo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtSufijo
        
End Sub

Private Sub txtSufijo_LostFocus()

    txtSufijo = TextLostFocus(txtSufijo, mobjParametro, "Sufijo")

End Sub

Private Sub txtServidorContawin_Change()

    If Not mflgLoading Then _
        TextChange txtServidorContawin, mobjParametro, "ServidorContawin"

End Sub

Private Sub txtServidorContawin_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtServidorContawin
        
End Sub

Private Sub txtServidorContawin_LostFocus()

    txtServidorContawin = TextLostFocus(txtServidorContawin, mobjParametro, "ServidorContawin")

End Sub

Private Sub txtClave_Change()

    If Not mflgLoading Then _
        TextChange txtClave, mobjParametro, "Clave"

End Sub

Private Sub txtClave_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtClave
        
End Sub

Private Sub txtClave_LostFocus()

    txtClave = TextLostFocus(txtClave, mobjParametro, "Clave")

End Sub

Private Sub txtDireccion_DblClick()
    
    Call btnDireccion_Click
    
End Sub

Private Sub txtPropietario_Change()
    
    If Not mflgLoading Then _
        TextChange txtPropietario, mobjParametro, "Propietario"

End Sub

Private Sub txtPropietario_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPropietario
        
End Sub

Private Sub txtPropietario_LostFocus()

    txtPropietario = TextLostFocus(txtPropietario, mobjParametro, "Propietario")

End Sub

Private Sub txtUsuario_Change()

    If Not mflgLoading Then _
        TextChange txtUsuario, mobjParametro, "Usuario"

End Sub

Private Sub txtUsuario_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtUsuario
        
End Sub

Private Sub txtUsuario_LostFocus()

    txtUsuario = TextLostFocus(txtUsuario, mobjParametro, "Usuario")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
