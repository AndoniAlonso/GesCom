VERSION 5.00
Begin VB.Form CapturaCodigo 
   Caption         =   "Captura de código de barras"
   ClientHeight    =   4110
   ClientLeft      =   5700
   ClientTop       =   5880
   ClientWidth     =   5100
   Icon            =   "CapturaCodigo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstData 
      Enabled         =   0   'False
      Height          =   2205
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4692
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Finalizar captura"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtBarCode 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblDatosCapturados 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Datos capturados:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Captura de código de barras:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "CapturaCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event CodigoSeleccionado(strCodigo As String)
Event FinCaptura()
Private nDatosCapturados As Integer

Private Sub cmdCancel_Click()
Dim i As Integer
Dim j As Integer

    On Error GoTo ErrorManager
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    
    j = lstData.ListCount - 1
    For i = 0 To j
        ' Lanzar el evento
        RaiseEvent CodigoSeleccionado(lstData.List(0))
        lstData.RemoveItem 0
    Next
    Screen.MousePointer = vbDefault
    
    RaiseEvent FinCaptura
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager
    nDatosCapturados = nDatosCapturados + 1
    lblDatosCapturados.Caption = CStr(nDatosCapturados)
    If txtBarCode.Text = vbNullString Then Exit Sub

    lstData.AddItem txtBarCode.Text

    txtBarCode.Text = vbNullString
    txtBarCode.SetFocus
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()
'    txtBarCode.SetFocus
    nDatosCapturados = 0
    DisableX Me
    CentrarForma Me
    lstData.Clear
End Sub

