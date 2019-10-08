Attribute VB_Name = "modXML"
Option Explicit

Public Function getXml(ByVal xmlPath As String) As MSXML2.DOMDocument
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim isValid As Boolean
    
    xmlDoc.async = False
    xmlDoc.Load xmlPath
    isValid = (xmlDoc.parseError.errorCode = 0)
    If Not isValid Then
        MsgBox getXMLError(xmlDoc)
    End If
    Set getXml = xmlDoc
End Function

Public Function getXMLError(ByRef xmlDoc As MSXML2.DOMDocument) As String
    Dim strError As String
    
    strError = "Invalid XML file!" & vbNewLine & _
               "File: " & xmlDoc.parseError.URL & vbNewLine & _
               "Line: " & xmlDoc.parseError.line & vbNewLine & _
               "Character: " & xmlDoc.parseError.linepos & vbNewLine & _
               "Source Text: " & xmlDoc.parseError.srcText & vbNewLine & _
               "Description: " & xmlDoc.parseError.reason
    getXMLError = strError
End Function

Public Sub addElmWithAttribAndText(ByRef objDocument As IXMLDOMDocument, ByRef objNode As IXMLDOMElement, ByVal elementName As String, ByVal attributeName As String, ByVal attributeValue As String, ByVal text As String)
    Dim element As IXMLDOMElement
    Dim elementText As IXMLDOMText
    Dim elementAttribute As IXMLDOMAttribute
    
    Set element = objDocument.createElement(elementName)
    Set elementAttribute = objDocument.createAttribute(attributeName)
    Set elementText = objDocument.createTextNode(text)
    
    elementAttribute.Value = attributeValue
    element.appendChild elementText
    element.setAttributeNode elementAttribute
    
    objNode.appendChild element
End Sub

Public Sub addElmWithText(ByRef objDocument As IXMLDOMDocument, ByRef objNode As IXMLDOMElement, ByVal elementName As String, ByVal text As String)
    Dim element As IXMLDOMElement
    Dim elementText As IXMLDOMText
    
    Set element = objDocument.createElement(elementName)
    Set elementText = objDocument.createTextNode(text)
    
    element.appendChild elementText
    
    objNode.appendChild element
End Sub

Public Sub removeTopChildrenOf(ByRef objNode As IXMLDOMElement, ByVal childName As String)
    Dim child As IXMLDOMNode

    For Each child In objNode.childNodes
        If child.nodeName = childName Then
            objNode.removeChild child
        End If
    Next
End Sub

Public Sub removeElementIfExists(ByRef objNode As IXMLDOMElement, ByVal childName As String)
    Dim child As IXMLDOMElement
    For Each child In objNode.childNodes
        If child.nodeName = childName Then
            objNode.removeChild child
        End If
    Next
End Sub

Public Function getChildElementText(ByRef objElement As IXMLDOMElement, ByVal strng As String) As String
    Dim elementText As String
    Dim child As IXMLDOMElement
    elementText = ""
    
    For Each child In objElement.childNodes
        If child.nodeName = strng Then
            If child.childNodes.length >= 1 Then
                elementText = child.firstChild.text
                Exit For
            End If
        End If
    Next
    
    getChildElementText = elementText
End Function

Public Function getChildXML(ByRef xmlElement As IXMLDOMElement, ByVal name As String, Optional ByVal ifEmpty As String = "") As String
    Dim xmlNode As IXMLDOMNode
    Dim xmlText As String
    
    Set xmlNode = xmlElement.selectSingleNode(name)
    If Not (xmlNode Is Nothing) Then
        xmlText = xmlNode.xml
    Else
        xmlText = ifEmpty
    End If
    
    getChildXML = xmlText
End Function

Public Function convertXMLtoText(ByVal xmlText As String) As String
    Dim text As String
    text = xmlText
    text = Replace$(text, "&", "&amp;")
    text = Replace$(text, "<", "&lt;")
    text = Replace$(text, ">", "&gt;")
    convertXMLtoText = text
End Function

Public Function getInnerXml(ByVal objXml As IXMLDOMElement) As String
    Dim child As IXMLDOMNode
    Dim text As String
    
    text = ""
    For Each child In objXml.childNodes
        text = text & child.xml
    Next
    
    getInnerXml = text
End Function
