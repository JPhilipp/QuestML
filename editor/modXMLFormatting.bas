Attribute VB_Name = "modXMLFormatting"
Option Explicit

Public Function reformatQuestXMLString(ByVal text As String, Optional ByVal globalScope As Boolean = False)
    text = Replace$(text, "<if ", vbNewLine & "<if ")
    text = Replace$(text, "</if><else>", "</if>" & vbNewLine & "<else>")
    text = Replace$(text, "><text", ">" & vbNewLine & "<text")
    text = Replace$(text, "<!-- ", vbNewLine & "<!-- ")
    text = Replace$(text, " -->", " -->" & vbNewLine)
    text = Replace$(text, "><choice ", ">" & vbNewLine & "<choice ")
    text = Replace$(text, "><image ", ">" & vbNewLine & "<image ")
    text = Replace$(text, "><music ", ">" & vbNewLine & "<music ")
    text = Replace$(text, "><randomize ", ">" & vbNewLine & "<randomize ")
    text = Replace$(text, "><state ", ">" & vbNewLine & "<state ")
    text = Replace$(text, "><number ", ">" & vbNewLine & "<number ")
    text = Replace$(text, "><string ", ">" & vbNewLine & "<string ")
        
    If globalScope Then
        text = clearString(text, vbTab)
        
        text = Replace$(text, "<station ", vbNewLine & "<station ")
        text = Replace$(text, "</station>", "</station>" & vbNewLine)
        
        text = Replace$(text, "<about>", vbNewLine & "<about>")
        text = Replace$(text, "</about>", "</about>" & vbNewLine)
        text = Replace$(text, "<style>", vbNewLine & "<style>")
        text = Replace$(text, "</style>", "</style>" & vbNewLine)
        text = Replace$(text, "</quest>", vbNewLine & vbNewLine & "</quest>")
        
        text = precedeTab(text, "title")
        text = precedeTab(text, "author")
        text = precedeTab(text, "homepage")
        text = precedeTab(text, "cover")
        text = precedeTab(text, "email")
    End If
    
    reformatQuestXMLString = text
End Function

Public Sub reformatQuestXMLfile(ByVal xmlLocation As String)
    Dim text As String, success As Boolean
    text = getFileText(xmlLocation, success)
    If success Then
        text = reformatQuestXMLString(text, True)
        setFileText xmlLocation, text
    End If
End Sub

Private Function clearString(ByVal text As String, ByVal strng As String) As String
    clearString = Replace$(text, strng, "")
End Function

Private Function addReturn(ByVal text As String, ByVal strng As String, Optional ByVal isElement As Boolean = True) As String
    If isElement Then strng = wrapAsTag(strng)
    addReturn = Replace$(text, strng, strng & vbNewLine)
End Function

Private Function addDoubleReturn(ByVal text As String, ByVal strng As String, Optional ByVal isElement As Boolean = True) As String
    Dim i As Integer
    If isElement Then strng = wrapAsTag(strng)
    For i = 1 To 2
        text = Replace$(text, strng, strng & vbNewLine)
    Next
    addDoubleReturn = text
End Function

Private Function precedeReturn(ByVal text As String, ByVal strng As String, Optional ByVal isElement As Boolean = True) As String
    If isElement Then strng = wrapAsTag(strng)
    precedeReturn = Replace$(text, strng, vbNewLine & strng)
End Function

Private Function precedeTab(ByVal text As String, ByVal strng As String, Optional ByVal isElement As Boolean = True, Optional ByVal textTab As Boolean = True) As String
    Dim tabStrng As String
    If isElement Then strng = wrapAsTag(strng)
    If textTab Then
        tabStrng = String$(4, " ")
    Else
        tabStrng = vbTab
    End If
    precedeTab = Replace$(text, strng, tabStrng & strng)
End Function

Private Function wrapAsTag(ByVal strng As String) As String
    wrapAsTag = "<" & strng & ">"
End Function
