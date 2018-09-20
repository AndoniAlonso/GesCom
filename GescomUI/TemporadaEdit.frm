VERSION 5.00
Begin VB.Form TemporadaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Temporadas"
   ClientHeight    =   2055
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TemporadaEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Temporada"
      ClipControls    =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   700
         Width           =   4455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   340
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "TemporadaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjTemporada As Temporada
Attribute mobjTemporada.VB_VarHelpID = -1

Public Sub Component(TemporadaObject As Temporada)

    Set mobjTemporada = TemporadaObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjTemporada.ApplyEdit
    mobjTemporada.BeginEdit
    Exit Sub
  
ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    Dim Respuesta As VbMsgBoxResult
    
    If mobjTemporada.IsDirty And Not mobjTemporada.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjTemporada.CancelEdit
            Unload Me
        End If
    Else
        mobjTemporada.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjTemporada.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
       
    With mobjTemporada
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Temporada [(nueva)]"

        Else
            Caption = "Temporada [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtCodigo = .Codigo
        .BeginEdit
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjTemporada_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtCodigo_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtCodigo

End Sub

Private Sub txtNombre_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNombre

End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjTemporada, "Nombre"

End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjTemporada, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjTemporada, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjTemporada, "Codigo")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
