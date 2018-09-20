VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ctlEntityProxy 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ScaleHeight     =   345
   ScaleWidth      =   3915
   Begin VB.ComboBox cmbEntity 
      Height          =   315
      Left            =   1600
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1600
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   19660801
      CurrentDate     =   38030
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1600
   End
End
Attribute VB_Name = "ctlEntityProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mepStyle As Long
Private mstrEntityName As String
Private mstrCampoClave As String
Private mstrCampoDescripcion As String
Private mstrClausulaWhere As String
Private mstrProyecto As String
Private mstrServidorPersist As String
Private mstrValorDefecto As String
Private mbolIsAlphanumericList As Boolean

Private mobjTextList As ProxyList
Private mintEntitySelStart As Integer
Private mbolInitialized As Boolean
Private mlngCompare As Long

'Anchura y caption de la etiqueta
Private mLabelWidth As Integer
Private mLabelCaption As String

Public Sub Initialize(Style As Long, EntityName As String, _
                       CampoClave As String, CampoDescripcion As String, strClausulaWhere As String, Proyecto As String, ServidorPersist As String, ValorDefecto As String, _
                       lngCompare As Long, Optional IsAlphanumericList As Boolean = False)
                       
    Select Case Style
    Case epStyleComboBox, epStyleTextBox
        cmbEntity.Visible = True
        cmbEntity.Enabled = True
        dtpFecha.Visible = False
        dtpFecha.Enabled = False
    Case epStyleDatePicker
        cmbEntity.Visible = False
        cmbEntity.Enabled = False
        dtpFecha.Visible = True
        dtpFecha.Enabled = True
    Case Else
        Err.Raise vbObjectError + 1001, "ctlEntityProxy.SetStyle", " No es un tipo de estilo válido:" & "'" & CStr(Style) & "'"
    End Select
    
    mepStyle = Style
    mstrEntityName = EntityName
    mstrCampoClave = CampoClave
    mstrCampoDescripcion = CampoDescripcion
    mstrProyecto = Proyecto
    mstrServidorPersist = ServidorPersist
    mstrClausulaWhere = strClausulaWhere
    mstrValorDefecto = ValorDefecto
    mbolInitialized = True
    mlngCompare = lngCompare
    mbolIsAlphanumericList = IsAlphanumericList
    ' Inicialmente un valor de 1600
    Me.LabelWidth = lblEtiqueta.Width
    Me.LabelCaption = lblEtiqueta.Caption
End Sub

Public Property Get ClausulaWhere() As String
Dim strResultado As String

    If Not mbolInitialized Then Err.Raise vbObjectError + 1001, "ctlEntityProxy.ClausulaWhere", "No se ha establecido el estilo del control"
    
    Select Case mepStyle
    Case epStyleComboBox
        If IsNumeric(mobjTextList.Key(cmbEntity.Text)) Then
            If mobjTextList.Key(cmbEntity.Text) = 0 Then
                strResultado = vbNullString
            Else
                strResultado = mobjTextList.CampoClave & "=" & mobjTextList.Key(cmbEntity.Text)
            End If
        Else
            strResultado = mobjTextList.CampoClave & "= '" & mobjTextList.Key(cmbEntity.Text) & "'"
        End If
    Case epStyleDatePicker
        If IsNull(dtpFecha.Value) Then
            strResultado = vbNullString
        Else
            strResultado = mstrCampoClave & TextCompare & " '" & dtpFecha.Value & "' "
        End If
    End Select
    
    ClausulaWhere = strResultado

End Property

Public Property Get ClausulaWhereTXT() As String
Dim strResultado As String

    If Not mbolInitialized Then Err.Raise vbObjectError + 1001, "ctlEntityProxy.ClausulaWhere", "No se ha establecido el estilo del control"
    
    Select Case mepStyle
    Case epStyleComboBox
        If IsNumeric(mobjTextList.Key(cmbEntity.Text)) Then
            If mobjTextList.Key(cmbEntity.Text) = 0 Then
                strResultado = vbNullString
            Else
                strResultado = mobjTextList.CampoDescripcion & "=" & " '" & cmbEntity.Text & "'"
            End If
        Else
            strResultado = mobjTextList.CampoDescripcion & "=" & " '" & cmbEntity.Text & "'"
        End If
    Case epStyleDatePicker
        strResultado = mstrCampoClave & TextCompare & " '" & dtpFecha.Value & "' "
    End Select
    
    ClausulaWhereTXT = strResultado

End Property

Public Property Get SelectedValue() As String
Dim strResultado As String

    If Not mbolInitialized Then Err.Raise vbObjectError + 1001, "ctlEntityProxy.Value", "No se ha establecido el estilo del control"
    
    Select Case mepStyle
    Case epStyleComboBox
        If mobjTextList.Key(cmbEntity.Text) = 0 Then
            strResultado = vbNullString
        Else
            strResultado = cmbEntity.Text
        End If
    Case epStyleDatePicker
        strResultado = dtpFecha.Value
    End Select
    
    SelectedValue = strResultado

End Property

Public Property Get SelectedKey() As Long
Dim lngResultado As Long

    If Not mbolInitialized Then Err.Raise vbObjectError + 1001, "ctlEntityProxy.Value", "No se ha establecido el estilo del control"
    
    Select Case mepStyle
    Case epStyleComboBox
        lngResultado = mobjTextList.Key(cmbEntity.Text)
    Case epStyleDatePicker
        lngResultado = dtpFecha.Value
    End Select
    
    SelectedKey = lngResultado

End Property

Public Sub LoadControl(strLabel As String)
        
    If Not mbolInitialized Then Err.Raise vbObjectError + 1001, "ctlEntityProxy.ClausulaWhere", "No se ha establecido el estilo del control"
    
    Select Case mepStyle
    Case epStyleComboBox
        mobjTextList.Load mstrEntityName, mstrCampoClave, mstrCampoDescripcion, mstrClausulaWhere, mstrProyecto, mstrServidorPersist, mbolIsAlphanumericList
        LoadCombo cmbEntity, mobjTextList
        If mstrValorDefecto <> vbNullString Then cmbEntity.Text = mstrValorDefecto
    Case epStyleDatePicker
        If mstrValorDefecto <> vbNullString Then
            dtpFecha.Value = mstrValorDefecto
        Else
            dtpFecha.Value = Null
        End If
    End Select
        
    If strLabel <> vbNullString Then lblEtiqueta = strLabel

End Sub

Private Sub UserControl_Initialize()

    Set mobjTextList = New ProxyList
    mbolInitialized = False
    mstrClausulaWhere = ""
End Sub

Private Sub UserControl_Terminate()

    Set mobjTextList = Nothing
    
End Sub

Private Sub cmbEntity_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintEntitySelStart = cmbEntity.SelStart
End Sub

Private Sub cmbEntity_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintEntitySelStart, cmbEntity
    
End Sub

Private Function TextCompare() As String
    Select Case mlngCompare
    Case epCompareEqual
        TextCompare = " ="
    Case epCompareGreater
        TextCompare = " >"
    Case epCompareLess
        TextCompare = " <"
    Case epCompareGreaterEqual
        TextCompare = " >="
    Case epCompareLessEqual
        TextCompare = " <="
    End Select

End Function

Public Property Get LabelWidth() As Integer

    LabelWidth = mLabelWidth

End Property

Public Property Let LabelWidth(ByVal iLabelWidth As Integer)

    mLabelWidth = iLabelWidth

    Call UserControl.PropertyChanged("LabelWidth")
    
    lblEtiqueta.Width = mLabelWidth
    cmbEntity.Left = mLabelWidth
    dtpFecha.Left = mLabelWidth
    UserControl.Width = lblEtiqueta.Width + cmbEntity.Width

End Property

Public Property Get LabelCaption() As String

    LabelCaption = mLabelCaption

End Property

Public Property Let LabelCaption(ByVal strLabelCaption As String)

    mLabelCaption = strLabelCaption

    Call UserControl.PropertyChanged("LabelCaption")
    
    lblEtiqueta.Caption = mLabelCaption

End Property
