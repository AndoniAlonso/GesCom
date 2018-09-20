Attribute VB_Name = "EPMain"
Option Explicit

'Public Enum epStyle
Public Const epStyleComboBox = 1
Public Const epStyleTextBox = 2
Public Const epStyleDatePicker = 3
'End Enum
'Public Enum epCompare
Public Const epCompareEqual = 0    ' =
Public Const epCompareGreater = 1    ' >
Public Const epCompareLess = 2    ' <
Public Const epCompareGreaterEqual = 3  ' >=
Public Const epCompareLessEqual = 4  ' <=
'End Enum


Public Type TextListProps
  Key As String * 30
  Item As String * 255
End Type

Public Type TextListData
  Buffer As String * 285
End Type

Public Sub LoadCombo(Combo As ComboBox, List As ProxyList)
    Dim vntItem As Variant
  
    With Combo
        .Clear
        For Each vntItem In List
            .AddItem vntItem
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub


Public Sub SmartComboKeyPress(KeyAscii As Integer, ByRef mintSelStart As Integer, ByRef cboCombo As ComboBox)
    Dim lCnt       As Long 'Generic long counter
    Dim lMax       As Long
    Dim sComboItem As String
    Dim sComboText As String 'Text currently in combobox
    Dim sText      As String 'Text after keypressed

    With cboCombo
        lMax = .ListCount - 1
        sComboText = .Text
        sText = Left(sComboText, mintSelStart) & Chr(KeyAscii)
        
        KeyAscii = 0 'Reset key pressed
        
        For lCnt = 0 To lMax
            sComboItem = .List(lCnt)
            
            If UCase(sText) = UCase(Left(sComboItem, _
                                         Len(sText))) Then
                .ListIndex = lCnt
                .Text = sComboItem
                .SelStart = Len(sText)
                .SelLength = Len(sComboItem) - (Len(sText))
                
                Exit For
            End If
        Next 'lCnt
    End With
End Sub



