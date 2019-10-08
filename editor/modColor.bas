Attribute VB_Name = "modColor"
Option Explicit

Public Type colorTriplet
    red As Long
    green As Long
    blue As Long
End Type

Public Function getColorTriplet(ByRef dlgColor As Object) As colorTriplet
    Dim clr As OLE_COLOR, col As colorTriplet
    
    dlgColor.CancelError = True
    On Error Resume Next

    dlgColor.Flags = cdlCCFullOpen

    dlgColor.ShowColor
    If Err.Number = cdlCancel Then
        Exit Function
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting color." & vbCrLf & Err.Description
        Exit Function
    End If

    On Error GoTo 0
    
    clr = dlgColor.Color
    
    col.red = (clr And &HFF&)
    col.green = (clr And &HFF00&) \ &H100&
    col.blue = (clr And &HFF0000) \ &H10000
    
    getColorTriplet = col
End Function

Public Function getVBColor(ByRef col As colorTriplet) As Long
    getVBColor = VBColor(col.red, col.green, col.blue)
End Function

Private Function VBColor(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    VBColor = r + 256 * (g + 256 * b)
End Function

Public Function isNumber(ByVal strng As String) As Boolean
    isNumber = (strng >= "0" And strng <= "9")
End Function

Public Function colorTripletToCSS(ByRef col As colorTriplet) As String
    colorTripletToCSS = "rgb(" & col.red & "," & col.green & "," & _
            col.blue & ")"
End Function

Public Function cssRGBToVBColor(ByVal cssString As String) As Long
    Dim col As colorTriplet, i As Integer, returnValue As Long, _
            splitted() As String
    
    If Trim$(cssString) <> "" Then
        cssString = Replace$(cssString, "rgb", "", , , vbTextCompare)
        cssString = Replace$(cssString, "(", "", , , vbTextCompare)
        cssString = Replace$(cssString, ")", "", , , vbTextCompare)
        splitted = Split(cssString, ",")
        
        For i = LBound(splitted) To UBound(splitted)
            splitted(i) = Trim$(splitted(i))
            splitted(i) = percentageToByteString(splitted(i))
        Next
            
        col.red = CLng(splitted(0))
        col.green = CLng(splitted(1))
        col.blue = CLng(splitted(2))
        
        returnValue = VBColor(col.red, col.green, col.blue)
    Else
        returnValue = 0
    End If
    
    cssRGBToVBColor = returnValue
End Function

Public Function percentageToByteString(ByVal strng As String) As String
    Dim percentageVal As Single, returnValue As String
    
    If Right$(strng, 1) = "%" Then
        strng = Left$(strng, Len(strng) - 1)
        strng = Replace$(strng, " ", "")
        percentageVal = Val(strng)
        If percentageVal < 0 Then
            percentageVal = 0
        ElseIf percentageVal > 100 Then
            percentageVal = 100
        End If
        returnValue = Str$((percentageVal * 255) \ 100)
    Else
        returnValue = strng
    End If
    
    percentageToByteString = returnValue
End Function
