VERSION 5.00
Begin VB.Form TransportistaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transportistas"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TransportistaEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Transportista"
      Height          =   3400
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtZona 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1780
         Width           =   1575
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1420
         Width           =   4215
      End
      Begin VB.TextBox txtDNINIF 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1060
         Width           =   2175
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   700
         Width           =   4215
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   340
         Width           =   4215
      End
      Begin VB.CommandButton btnDireccion 
         Caption         =   "Di&rección"
         Height          =   495
         Left            =   160
         TabIndex        =   11
         Top             =   2155
         Width           =   975
      End
      Begin VB.TextBox txtDireccion 
         Height          =   1005
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2140
         Width           =   4215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contacto"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DNI/NIF"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Titular"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "TransportistaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjTransportista As Transportista
Attribute mobjTransportista.VB_VarHelpID = -1

Public Sub Component(TransportistaObject As Transportista)

    Set mobjTransportista = TransportistaObject

End Sub

Private Sub btnDireccion_Click()
    
    Dim frmDireccion As DireccionEdit
  
    Set frmDireccion = New DireccionEdit
    frmDireccion.Component mobjTransportista.Direccion
    frmDireccion.Show vbModal
    txtDireccion.Text = mobjTransportista.Direccion.DireccionText
  
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjTransportista.ApplyEdit
    mobjTransportista.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjTransportista.IsDirty And Not mobjTransportista.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjTransportista.CancelEdit
            Unload Me
        End If
    Else
        mobjTransportista.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjTransportista.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorManager

    DisableX Me
    
    mflgLoading = True
    With mobjTransportista
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Transportista [(nuevo)]"

        Else
            Caption = "Transportista [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        'txtTransportistaID = .TransportistaID
        txtNombre = .Nombre
        txtTitular = .Titular
        txtContacto = .Contacto
        txtDNINIF = .DNINIF
        txtDireccion = .Direccion.DireccionText
        txtZona = .Zona
        
        .BeginEdit
    
    End With
  
    mflgLoading = False
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjTransportista_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtContacto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtContacto
        
End Sub

Private Sub txtDireccion_DblClick()

    Call btnDireccion_Click
    
End Sub

Private Sub txtDNINIF_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDNINIF
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjTransportista, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjTransportista, "Nombre")

End Sub

Private Sub txtTitular_Change()

    If Not mflgLoading Then _
        TextChange txtTitular, mobjTransportista, "Titular"

End Sub

Private Sub txtTitular_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTitular
        
End Sub

Private Sub txtTitular_LostFocus()

    txtTitular = TextLostFocus(txtTitular, mobjTransportista, "Titular")

End Sub

Private Sub txtContacto_Change()

    If Not mflgLoading Then _
        TextChange txtContacto, mobjTransportista, "Contacto"

End Sub

Private Sub txtContacto_LostFocus()

    txtContacto = TextLostFocus(txtContacto, mobjTransportista, "Contacto")

End Sub

Private Sub txtDNINIF_Change()

    If Not mflgLoading Then _
        TextChange txtDNINIF, mobjTransportista, "DNINIF"

End Sub

Private Sub txtDNINIF_LostFocus()

    txtDNINIF = TextLostFocus(txtDNINIF, mobjTransportista, "DNINIF")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub txtZona_Change()

    If Not mflgLoading Then _
        TextChange txtZona, mobjTransportista, "Zona"
        
End Sub

Private Sub txtZona_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtZona
        
End Sub

Private Sub txtZona_LostFocus()

    txtZona = TextLostFocus(txtZona, mobjTransportista, "Zona")
    
End Sub
