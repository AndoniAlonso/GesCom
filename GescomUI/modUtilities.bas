Attribute VB_Name = "modUtilities"
Option Explicit

Public mstrUsuario As String
Public mstrUserLanguage As String

Public Const cnPARAMARRAYTYPE = 8204
Public Const cnPARAMSTRINGSEP As String = "|"
Public Const cnFORMATSTRINGSTART As String = "["
Public Const cnFORMATSTRINGEND As String = "]"

Public Function Nz(vntEvaluar As Variant, Optional vntValor As Variant = Empty) As Variant
    If IsEmptyValue(vntEvaluar) Then
        Nz = vntValor
    Else
        Nz = vntEvaluar
    End If
End Function

Public Function IsEmptyValue(ByVal varValue As Variant) As Boolean
    IsEmptyValue = True
    If Not IsMissing(varValue) Then
        If Not IsNull(varValue) Then
            IsEmptyValue = IsEmpty(varValue) Or (Len(Trim(varValue)) = 0)
        End If
    End If
End Function


Public Sub GenerateError(ByVal lngNumber As Long, _
                         ByVal strDescription As String, _
                         ByVal strSource As String, _
                         ByVal strTitle As String, _
                         ParamArray arrParameters())
    Dim lngCount As Long
    Dim arrParamArray() As Variant
    
    On Error Resume Next
    
    lngCount = -1
    lngCount = UBound(arrParameters)
    If lngCount = 0 Then
        If Not IsNull(arrParameters) Then
            If VarType(arrParameters(0)) = cnPARAMARRAYTYPE Then
                arrParamArray = arrParameters(0)
            Else
                arrParamArray = arrParameters
            End If
        End If
    Else
        arrParamArray = arrParameters
    End If
    strDescription = ParseFormatString(strDescription, arrParamArray)
    ''''If Len(strTitle) = 0 Then strTitle = mobjExpertisApp.Title
    If Len(strTitle) = 0 Then strTitle = "Gescom"
    If lngNumber >= UserMessageBaseCode And lngNumber <= UserMessageLimitCode Then
        MsgBox strDescription, vbCritical, strTitle
    Else
        'MSG="Ha ocurrido el error X(X) en el módulo X"
        ''''MsgBox ParseFormatString(TraslateWord(55371), lngNumber, strDescription, strSource), vbCritical, strTitle
        MsgBox ParseFormatString(strDescription, lngNumber, strDescription, strSource), vbCritical, strTitle
    End If
End Sub



'----------------------------------------------------------------------------
'   Esta función sustituye los comodines(|[]) de parámetros de la cadena por sus
'   correspondientes valores de la lista de parámetros.
'----------------------------------------------------------------------------
Public Function ParseFormatString(strSourceStr As String, ParamArray arrParameters()) As String
    Dim lngStarPos As Long
    Dim lngFindPos As Long
    Dim lngFormatStartPos As Long
    Dim lngFormatEndPos As Long
    Dim strFormatStr As String
    Dim strParamStr As String
    Dim i As Long
    Dim lngCount As Long
    Dim arrParamArray() As Variant
    
    On Error Resume Next
    
    lngStarPos = 1
    i = -1
    lngCount = -1
    lngCount = UBound(arrParameters)
    If lngCount = 0 Then
        If Not IsNull(arrParameters) Then
            If VarType(arrParameters(0)) = cnPARAMARRAYTYPE Then
                arrParamArray = arrParameters(0)
            Else
                arrParamArray = arrParameters
            End If
        End If
    Else
        arrParamArray = arrParameters
    End If
    lngCount = -1
    lngCount = UBound(arrParamArray)
    Do While True
        lngFindPos = InStr(lngStarPos, strSourceStr, cnPARAMSTRINGSEP, vbTextCompare)
        If lngFindPos = 0 Then Exit Do
        i = i + 1
        'Extraer el formato del parámetro
        lngFormatStartPos = lngFindPos + 1
        lngFormatEndPos = 0
        strFormatStr = vbNullString
        strParamStr = vbNullString
        If Mid$(strSourceStr, lngFormatStartPos, 1) = cnFORMATSTRINGSTART Then
            lngFormatStartPos = lngFormatStartPos + 1
            lngFormatEndPos = InStr(lngFormatStartPos, strSourceStr, cnFORMATSTRINGEND, vbTextCompare)
            If lngFormatEndPos > 0 Then
                strFormatStr = Mid$(strSourceStr, lngFormatStartPos, lngFormatEndPos - lngFormatStartPos)
            End If
        End If
        If i <= lngCount Then
            If Len(strFormatStr) > 0 Then
                If IsEmptyValue(arrParamArray(i)) Then
                    strParamStr = vbNullString
                Else
                    strParamStr = Format(arrParamArray(i), strFormatStr)
                    If Err.Number <> 0 Then
                        strParamStr = CStr(arrParamArray(i))
                        Err.Clear
                    End If
                End If
            Else
                strParamStr = CStr(arrParamArray(i))
            End If
        End If
        ParseFormatString = ParseFormatString & Mid$(strSourceStr, lngStarPos, lngFindPos - lngStarPos) & strParamStr
        If lngFormatEndPos = 0 Then
            lngStarPos = lngFindPos + 1
        Else
            lngStarPos = lngFormatEndPos + 1
        End If
    Loop
    ParseFormatString = ParseFormatString & Mid$(strSourceStr, lngStarPos)
End Function

