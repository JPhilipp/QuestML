Attribute VB_Name = "modMisc"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Option Explicit

Public Function returnIf(ByVal state As Boolean, ByVal ifTrue As Variant, ByVal ifFalse As Variant) As Variant
    Dim returnValue As Variant
    If state Then
        returnValue = ifTrue
    Else
        returnValue = ifFalse
    End If
    returnIf = returnValue
End Function

Public Function getFileText(ByVal strFile As String, Optional ByRef success As Boolean) As Variant
    Dim nextFree As Integer
    nextFree = FreeFile
    
    On Error GoTo fileDoesntSeemToExist
    Open strFile For Input As #nextFree
        getFileText = Input(LOF(nextFree), #nextFree)
    Close #nextFree

    success = True
    Exit Function

fileDoesntSeemToExist:
    MsgBox "An error occurred while trying to open file " & _
           strFile, vbExclamation, "Open"
    success = False
End Function

Public Sub setFileText(ByVal pathOf As String, ByVal textOf As String, Optional ByRef success As Boolean)
    Dim fFile As Integer
    fFile = FreeFile
    
    On Error GoTo sftCouldntWrite
    
    Open pathOf For Output As #fFile
        Print #fFile, textOf
    Close fFile
    
    success = True
    Exit Sub
        
sftCouldntWrite:
    success = False
End Sub

Public Function convertToURI(ByVal winPath As String)
    Dim uri As String
    
    uri = "file://"
    uri = uri & Replace$(winPath, "\", "/")
    uri = Replace(uri, " ", "%20")
    
    convertToURI = uri
End Function

Public Function getTextBetween(ByVal text, ByVal startsWith As String, ByVal endsWith As String)
    Dim textBetween As String, startI As Long, endI As Long
    textBetween = ""
    
    startI = InStr(1, text, startsWith, vbBinaryCompare)
    If startI >= 1 Then
        startI = startI + Len(startsWith)
        endI = InStr(startI, text, endsWith, vbBinaryCompare)
        If endI >= 1 Then
            textBetween = Mid$(text, startI, endI - startI)
        End If
    End If
    
    getTextBetween = textBetween
End Function

Public Function getPathOf(ByVal strFile As String) As String
    If InStr(strFile, "\") > 0 Then
        getPathOf = Left$(strFile, InStrRev(strFile, "\"))
    Else
        getPathOf = strFile
    End If
End Function

Public Function fileExists(ByVal strPath As String) As Boolean
    Dim fFile As Integer
    
    fFile = FreeFile
    
    On Error GoTo FEfileNotFound
    Open strPath For Input As #fFile
    Close #fFile
    
    fileExists = True
    Exit Function
    
FEfileNotFound:
    fileExists = False
End Function

Function repeatedReplace(ByVal text As String, ByVal toFind As String, ByVal toReplace As String) As String
    Dim oldText

    Do
        oldText = text
        text = Replace$(text, toFind, toReplace)
    Loop Until text = oldText

    repeatedReplace = text
End Function

Public Function convertToName(ByVal text As String) As String
    Dim i As Integer
    Dim letter As String
    Dim newText As String
    
    text = replaceUmlaute(text)

    For i = 1 To Len(text)
        letter = Mid$(text, i, 1)
        
        If Not isLetterOrNumber(letter) Then
            letter = "_"
        End If
        newText = newText & letter
    Next
    
    If Not isLetter(Left$(newText, 1)) Then
        newText = "n" & newText
    End If
    
    convertToName = newText
End Function

Public Function isLetterOrNumber(ByVal letter As String) As Boolean
    Dim isLetter As Boolean, isNumber As Boolean
    letter = LCase$(letter)
    isLetter = letter >= "a" And letter <= "z"
    isNumber = letter >= "0" And letter <= "9"
    isLetterOrNumber = isLetter Or isNumber
End Function

Public Function isLetter(ByVal letter As String) As Boolean
    letter = LCase$(letter)
    isLetter = letter >= "a" And letter <= "z"
End Function

Public Function replaceUmlaute(ByVal text As String) As String
    text = Replace$(text, "ä", "ae", , , vbBinaryCompare)
    text = Replace$(text, "ö", "oe", , , vbBinaryCompare)
    text = Replace$(text, "ü", "ue", , , vbBinaryCompare)
    
    text = Replace$(text, "Ä", "Ae", , , vbBinaryCompare)
    text = Replace$(text, "Ö", "Oe", , , vbBinaryCompare)
    text = Replace$(text, "Ü", "Ue", , , vbBinaryCompare)
    
    text = Replace$(text, "ß", "ss", , , vbTextCompare)
    
    replaceUmlaute = text
End Function

Public Function nameOfFile(ByVal strFile As String, Optional ByVal cutExtension As Boolean = False) As String
    Dim rValue As String, lastDot As Integer
    rValue = Mid$(strFile, InStrRev(strFile, "\") + 1)
    If cutExtension Then
        lastDot = InStrRev(rValue, ".")
        If lastDot >= 1 Then
            rValue = Left$(rValue, lastDot - 1)
        End If
    End If
    nameOfFile = rValue
End Function

Public Function extension(ByVal strFile As String) As String
    extension = LCase$(Mid$(strFile, InStrRev(strFile, ".") + 1))
End Function

Function goPathUp(ByVal curPath As String) As String
    If Right$(curPath, 1) = "\" Then
        curPath = Left$(curPath, Len(curPath) - Len("\"))
    End If
    goPathUp = Left$(curPath, InStrRev(curPath, "\") - 1)
End Function

Public Sub selectAllText(ByRef objText As Object)
    With objText
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Public Sub openInBrowser(ByRef parHwnd As Long, ByVal URL As String)
    ShellExecute parHwnd, "open", URL, vbNullString, vbNullString, SW_SHOW
End Sub

Public Function trimChar(ByVal text As String, ByVal charToTrim As String) As String
    Dim oldText As String
   
    Do
        oldText = text
        If Left$(text, 1) = charToTrim Then
            text = Right$(text, Len(text) - 1)
        End If
        If Right$(text, 1) = charToTrim Then
            text = Left$(text, Len(text) - 1)
        End If
    Loop Until oldText = text
    
    trimChar = text
End Function

Sub selectLineOfText(ByVal objText As RichTextBox, ByVal line As Integer, Optional ByVal firstCharPos As Integer = 0)
    Dim i As Integer, newLine As Integer, text As String, _
        nextReturn As Integer, lineStart As Integer
    text = objText.text
    For i = 1 To Len(text)
        If Mid$(text, i, Len(vbNewLine)) = vbNewLine Then
            newLine = newLine + 1
            
            If newLine = line - 1 Then
                lineStart = i + Len(vbNewLine)
                objText.SelStart = lineStart - 1
                nextReturn = InStr( _
                        lineStart, text, vbNewLine, vbBinaryCompare)
                If firstCharPos >= 1 Then
                    objText.SelLength = (lineStart + firstCharPos) - lineStart - 1
                ElseIf nextReturn >= 1 Then
                    objText.SelLength = nextReturn - lineStart - 1
                Else
                    objText.SelLength = Len(text) - lineStart - 1
                End If
                
                objText.SetFocus
                
                Exit For
            End If
        End If
    Next
End Sub

Public Function getIndexOfListItem(ByRef lst As ListBox, ByVal nameOf As String) As Integer
    Dim i As Integer, indexOfListItem As Integer
    
    indexOfListItem = -1
    For i = 0 To lst.ListCount - 1
        If lst.List(i) = nameOf Then
            indexOfListItem = i
            Exit For
        End If
    Next
    
    getIndexOfListItem = indexOfListItem
End Function

Public Function getAttributeValueOfText(ByVal text As String, curPos As Long) As String
    Dim attributeValue As String, _
            leftOf As Integer, rightOf As Integer

    attributeValue = ""
    If curPos >= 1 Then
        leftOf = InStrRev(text, """", curPos, vbBinaryCompare)
        rightOf = InStr(curPos + 1, text, """", vbBinaryCompare)
        If leftOf >= 1 And rightOf > leftOf Then
            attributeValue = Mid$(text, leftOf + 1, rightOf - leftOf - 1)
        End If
    End If
    getAttributeValueOfText = attributeValue
End Function

Public Function leftShow(ByVal text As String, ByVal stringLen As Long) As String
    ' same as left$, but adds " ..." if the string had to be cut
    If Len(text) > stringLen Then
        text = Left$(text, stringLen) & " ..."
    End If
    leftShow = text
End Function

Public Sub sortArrayABC(ByRef arr() As Variant)
    Dim i As Long
    Dim tmp As String
    Dim foundNothing As Boolean
    Dim lastSwapped As Long
    Dim startIndex As Integer
    Dim arrMax As Long
    
    arrMax = UBound(arr)
  
    Do While Not foundNothing
        foundNothing = True
        If lastSwapped <= 1 Then
            startIndex = LBound(arr)
        End If
        
        For i = startIndex To arrMax - 1
            If arr(i) > arr(i + 1) Then
                tmp = arr(i + 1)
                arr(i + 1) = arr(i)
                arr(i) = tmp
                foundNothing = False
                lastSwapped = i
                Exit For
            End If
        Next
    Loop
End Sub

Public Sub downsizeArray(ByRef arr() As Variant)
    Dim i As Long
    
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) = "" Then
            ReDim Preserve arr(LBound(arr) To i - 1)
            Exit For
        End If
    Next
End Sub

Public Function rTrimNewline(ByVal text As String) As String
    Dim oldText As String
    Do
        oldText = text
        If Right$(text, Len(vbNewLine)) = vbNewLine Then
            text = Left$(text, Len(text) - Len(vbNewLine))
        End If
    Loop Until oldText = text
    rTrimNewline = text
End Function

Public Function lTrimNewline(ByVal text As String) As String
    Dim oldText As String
    Do
        oldText = text
        If Left$(text, Len(vbNewLine)) = vbNewLine Then
            text = Right$(text, Len(text) - Len(vbNewLine))
        End If
    Loop Until oldText = text
    lTrimNewline = text
End Function

Public Function trimNewline(ByVal text As String) As String
    trimNewline = lTrimNewline(rTrimNewline(text))
End Function

Public Function addUniqueToArray(ByRef arr() As String, ByVal arrValue As String)
    Dim i As Long
    Dim foundIndex As Boolean
    Dim firstFreeIndex As Long
    Dim lowerArrValue As String
    
    firstFreeIndex = -1
    foundIndex = False
    lowerArrValue = LCase$(arrValue)
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = "" Then
            firstFreeIndex = i
            Exit For
        ElseIf LCase$(arr(i)) = arrValue Then
            foundIndex = True
            Exit For
        End If
    Next
    
    If Not foundIndex Then
        If firstFreeIndex <> -1 Then
            arr(firstFreeIndex) = arrValue
        End If
    End If
    
End Function
