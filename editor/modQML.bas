Attribute VB_Name = "modQML"
Option Explicit

Public Function toMediaURI(ByVal filePath As String) As String
    Dim text As String
    
    text = filePath
    text = getRelativePath(text)
    text = Replace$(text, "\", "/")
    
    toMediaURI = text
End Function

Public Function getRelativePath(ByVal absolutePath As String)
    Dim relativePath As String
    Dim qmlTopPath As String
    
    relativePath = Mid$(absolutePath, Len(frmMain.qmlTopPath) + 2)
    
    getRelativePath = relativePath
End Function

Sub setTextBoxToWebImagePath(ByRef textBox As Object)
    Dim filePath As String
    
    filePath = frmMain.getMediaFilePath("All image files|*.gif; *.jpg; *.png|" & _
            "*.gif|*.gif|" & _
            "*.jpg|*.jpg|" & _
            "*.png|*.png")
    If filePath <> "" Then
        textBox.text = filePath
    End If
End Sub


