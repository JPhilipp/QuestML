Attribute VB_Name = "modCheckForTabEditor"
Option Explicit

Public Function getStationOkForTabEditor(ByRef objStation As IXMLDOMElement) As Boolean
    Const doAlert = False
    Dim stationOkForTabEditor As Boolean
    Dim stationNotOkReason As String
    
    If Not containsOnlyTextAndChoices(objStation) Then
        stationOkForTabEditor = False
        stationNotOkReason = _
                "A basic station must contain only text and choices."
    ElseIf hasStateAttributes(objStation) Then
        stationOkForTabEditor = False
        stationNotOkReason = _
                "Text and choices may contain no states."
    ElseIf objStation.getElementsByTagName(elementText).length > 1 Then
        stationOkForTabEditor = False
        stationNotOkReason = _
            "Only a single text element allowed in basic stations."
    ElseIf includesTextFormatting(objStation) Then
        stationOkForTabEditor = False
        stationNotOkReason = _
            "No text formatting allowed."
    ElseIf Not lessEqualChoicesThanBoxes(objStation) Then
        stationOkForTabEditor = False
        stationNotOkReason = _
                "A basic station may not contain more choices then" & _
                "choice-textboxes."
    Else
        stationOkForTabEditor = True
    End If
    
    If doAlert Then
        If Not stationOkForTabEditor Then
            MsgBox "Switching from editor mode to source view." & vbNewLine & _
                "In editor mode, only basic stations can be created." & vbNewLine & _
                    stationNotOkReason, vbInformation
        End If
    End If
    
    getStationOkForTabEditor = stationOkForTabEditor
End Function

Function includesTextFormatting(ByRef objStation As IXMLDOMElement) As Boolean
    Const formattingElements = "emphasis strong display poem span link"
    Dim names() As String
    Dim i As Long
    Dim formattingFound As Boolean
    Dim element As IXMLDOMElement
    
    names = Split(formattingElements, " ")
    For i = LBound(names) To UBound(names)
        Set element = objStation.selectSingleNode(".//" & names(i))
        formattingFound = Not (element Is Nothing)
        If formattingFound Then Exit For
    Next
    
    includesTextFormatting = formattingFound
End Function

Private Function hasStateAttributes(ByRef objStation As IXMLDOMElement) As Boolean
    Dim foundState As Boolean
    Dim defaultRelation As Boolean
    Dim stationAttribute As Boolean
    Dim child As IXMLDOMNode
    Dim attributeChild As IXMLDOMAttribute

    For Each child In objStation.childNodes
        If child.nodeName = elementText Or _
                child.nodeName = elementPath Then
            
            For Each attributeChild In child.Attributes
                If Not ((attributeChild.nodeName = "relation" And _
                        attributeChild.text = "and") Or _
                        (attributeChild.nodeName = "is" And _
                        attributeChild.text = "true") Or _
                        attributeChild.nodeName = "station") Then
                    foundState = True
                    Exit For
                End If
            Next
        
        End If
    Next

    hasStateAttributes = foundState
End Function

Private Function containsOnlyTextAndChoices(ByRef objStation As IXMLDOMElement) As Boolean
    Dim sthElseThanTextOrChoiceFound As Boolean, _
            child As IXMLDOMNode

    For Each child In objStation.childNodes
        If child.nodeName <> elementText And _
                child.nodeName <> elementPath And _
                child.nodeName <> "#text" Then
            sthElseThanTextOrChoiceFound = True
            Exit For
        End If
    Next

    containsOnlyTextAndChoices = _
        Not sthElseThanTextOrChoiceFound
End Function

Private Function lessEqualChoicesThanBoxes(ByRef objStation As IXMLDOMElement) As Boolean
    Dim numberOfChoices As Integer, numberOfBoxes As Integer
    numberOfChoices = _
        objStation.getElementsByTagName(elementPath).length
    numberOfBoxes = frmMain.cboStation.UBound + 1
   
    lessEqualChoicesThanBoxes = numberOfChoices <= numberOfBoxes
End Function

