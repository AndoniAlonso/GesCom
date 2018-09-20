VERSION 5.00
Begin VB.Form ContextEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresa y Temporada de trabajo"
   ClientHeight    =   2055
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ContextEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Empresa y Temporada de trabajo"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cboTemporadaActual 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   700
         Width           =   2415
      End
      Begin VB.ComboBox cboEmpresaActual 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   340
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Empresa Actual"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temporada Actual"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1305
      End
   End
End
Attribute VB_Name = "ContextEdit"
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
            Caption = "Contexto [(nuevo)]"
    
        Else
            Caption = "Contexto [" & Trim(.EmpresaActual) & "--" & Trim(.TemporadaActual) & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        LoadCombo cboEmpresaActual, .Empresas
        cboEmpresaActual.Text = .EmpresaActual
    
        LoadCombo cboTemporadaActual, .Temporadas
        cboTemporadaActual.Text = .TemporadaActual
    
        .BeginEdit
    
    End With
  
    mflgLoading = False
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboEmpresaActual_Click()

    If mflgLoading Then Exit Sub
    mobjParametro.EmpresaActual = cboEmpresaActual.Text

End Sub

Private Sub cboTemporadaActual_Click()

    If mflgLoading Then Exit Sub
    mobjParametro.TemporadaActual = cboTemporadaActual.Text

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjParametro_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
