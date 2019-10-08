Attribute VB_Name = "modXhtmlForPrint"
Option Explicit

Private Const m_arrow = "-&gt;"
Private m_questClone As MSXML2.DOMDocument
Private m_startId As String

Public Function getXhtmlForPrint()
    Dim stations As MSXML2.IXMLDOMNodeList
    Dim station As MSXML2.IXMLDOMElement
    Dim thisElement As MSXML2.IXMLDOMElement
    Dim xhtml As String
    
    xhtml = ""
    Set m_questClone = frmMain.objQuest.cloneNode(True)
    moveStartStationUp
    numberizeStationIds
    
    xhtml = xhtml & getHeader
    Set stations = m_questClone.selectNodes("//station")
    For Each station In stations
        xhtml = xhtml & getTextForStation(station)
    Next

    xhtml = adaptFormatting(xhtml)

    getXhtmlForPrint = xhtml
End Function

Private Function adaptFormatting(ByVal xhtml As String)
    xhtml = Replace$(xhtml, "<emphasis>", "<em>")
    xhtml = Replace$(xhtml, "</emphasis>", "</em>")
    xhtml = Replace$(xhtml, "<display>", "<div class=""display"">")
    xhtml = Replace$(xhtml, "</display>", "</div>")
    xhtml = Replace$(xhtml, "<poem>", "<pre class=""poem"">")
    xhtml = Replace$(xhtml, "</poem>", "</pre>")
    
    adaptFormatting = xhtml
End Function

Private Function getHeader()
    Dim xhtml As String
    Dim thisElement As MSXML2.IXMLDOMElement
    
    xhtml = ""
    
    xhtml = xhtml & "<h1>"
    Set thisElement = m_questClone.selectSingleNode("//title")
    If Not (thisElement Is Nothing) Then
        xhtml = xhtml & thisElement.text
    End If
    Set thisElement = m_questClone.selectSingleNode("//author")
    If Not (thisElement Is Nothing) Then
        xhtml = xhtml & "<br /><span class=""subTitle"">" & _
                thisElement.text & "</span>"
    End If
    xhtml = xhtml & "</h1>" & vbNewLine & vbNewLine
    
    xhtml = xhtml & "<div class=""introduction"">"
    xhtml = xhtml & "Start at " & _
            "<strong>" & m_arrow & " " & m_startId & "</strong>"
    xhtml = xhtml & "</div>" & vbNewLine & vbNewLine

    getHeader = xhtml
End Function

Private Function getTextForStation(ByRef station As MSXML2.IXMLDOMElement) As String
    Dim ifElements As MSXML2.IXMLDOMNodeList
    Dim ifElement As MSXML2.IXMLDOMElement
    Dim elseElement As MSXML2.IXMLDOMElement
    Dim text As String
    Dim thisId As String
    Dim iIf As Long
    
    text = ""
    
    thisId = station.getAttribute("id")
    text = text & "<div class=""station"">" & vbNewLine
    text = text & "<h2>" & thisId & "</h2>" & vbNewLine
    text = text & "<div class=""text"">" & vbNewLine
    
    Set ifElements = station.selectNodes("if")
    If ifElements.length > 0 Then
        iIf = 0
        For Each ifElement In ifElements
            iIf = iIf + 1
            If Not IsNull(ifElement.getAttribute("check")) Then
                text = text & "<em>"
                If iIf = 1 Then
                    text = text & "If "
                Else
                    text = text & "Else, if "
                End If
                text = text & ifElement.getAttribute("check") & ":</em>"
            End If
            text = text & "<div class=""branch"">"
            text = text & getTextForElement(ifElement, thisId)
            text = text & "</div>" & vbNewLine
        Next
        Set elseElement = station.selectSingleNode("else")
        If Not (elseElement Is Nothing) Then
            text = text & "<em>Else</em>:"
            text = text & "<div class=""branch"">"
            text = text & getTextForElement(elseElement, thisId)
            text = text & "</div>" & vbNewLine
        End If
    Else
        text = text & getTextForElement(station, thisId)
    End If
    
    text = Replace$(text, "[", "<span class=""inline"">")
    text = Replace$(text, "]", "</span>")
    
    text = text & "</div>" & vbNewLine
    text = text & "</div>" & vbNewLine & vbNewLine

    getTextForStation = text
End Function

Private Function getTextForElement(ByRef thisElement As MSXML2.IXMLDOMElement, ByVal stationId As String) As String
    Dim subNode As MSXML2.IXMLDOMElement
    Dim text As String
    Dim sStates As String
    
    text = ""
    sStates = ""
    
    For Each subNode In thisElement.selectNodes("randomize")
        text = text & "<div class=""randomize"">" & _
                "Randomize [" & subNode.getAttribute("number") & "] " & _
                subNode.getAttribute("value")
        text = text & "</div>" & vbNewLine
    Next
    
    For Each subNode In thisElement.selectNodes("image")
        text = text & "<div class=""image"">" & _
                "<img src=""../../" & subNode.getAttribute("source") & _
                """ alt="""" /></div>" & vbNewLine
    Next
    
    For Each subNode In thisElement.selectNodes("text")
        If Not IsNull(subNode.getAttribute("check")) Then
           text = text & "<em>If " & subNode.getAttribute("check") & _
                "</em>: "
        End If
        text = text & subNode.xml & vbNewLine
    Next
    
    For Each subNode In thisElement.selectNodes("state | number | string")
        sStates = sStates & capitalize(subNode.getAttribute("process")) & ", "
        sStates = sStates & "set " & subNode.nodeName & " "
        sStates = sStates & "[" & subNode.getAttribute("name") & "] to "
        sStates = sStates & subNode.getAttribute("value") & "<br />" & vbNewLine
    Next
    If sStates <> "" Then
        text = text & "<div class=""states"">" & sStates & "</div>" & vbNewLine
    End If
    
    text = text & "<ul>" & vbNewLine
    For Each subNode In thisElement.selectNodes("choice | input")
        text = text & "<li>"
        If Not IsNull(subNode.getAttribute("check")) Then
            text = text & "<em>If " & subNode.getAttribute("check") & _
                    "</em>: "
        End If
        If subNode.nodeName = "input" Then
            text = text & "<strong>Input [" & _
                    subNode.getAttribute("name") & "]:</strong> "
        End If
        text = text & subNode.text
        text = text & "<strong> " & m_arrow & " " & _
                getChoiceString(subNode, stationId) & "</strong>"
        text = text & "</li>" & vbNewLine
    Next
    text = text & "</ul>" & vbNewLine
    
    text = Replace$(text, "<text", "<div")
    text = Replace$(text, "</text>", "</div>")
    
    getTextForElement = text
End Function

Private Function capitalize(ByVal text As String) As String
    capitalize = UCase$(Left$(text, 1)) & Mid$(text, 2)
End Function

Private Function getChoiceString(ByRef subNode As MSXML2.IXMLDOMElement, ByVal stationId As String)
    Const seperator = ", "
    Dim text As String
    Dim thisStation As String
    Dim stations As MSXML2.IXMLDOMNodeList
    Dim station As MSXML2.IXMLDOMElement
    Dim sStations As String
    Dim refererId As String
    Dim xPath As String
    Dim iBack As Long
    
    text = ""
    thisStation = subNode.getAttribute("station")
    If thisStation = "back" Then
        sStations = ""
        iBack = 0
        xPath = "//station//choice[@station = '" & stationId & "']"
        Set stations = m_questClone.selectNodes(xPath)
        For Each station In stations
            refererId = "(" & station.getAttribute("station") & ")"
            If Not InStr(sStations, refererId) >= 1 Then
                iBack = iBack + 1
                sStations = sStations & refererId & seperator
            End If
        Next
        If sStations <> "" Then
            sStations = Replace$(sStations, "(", "")
            sStations = Replace$(sStations, ")", "")
            sStations = Left$(sStations, Len(sStations) - Len(seperator))
            If iBack > 1 Then
                text = text & " back to last station (" & sStations & ")"
            Else
                text = text & " " & sStations
            End If
        End If
    Else
        text = thisStation
    End If
    
    getChoiceString = text
End Function

Private Function numberizeStationIds()
    Dim stations As MSXML2.IXMLDOMNodeList
    Dim station As MSXML2.IXMLDOMElement
    Dim choices As MSXML2.IXMLDOMNodeList
    Dim choice As MSXML2.IXMLDOMElement
    Dim thisId As String
    Dim newId As String
    Dim i As Long
    Dim xPath As String
    
    i = 0
    xPath = "//station"
    Set stations = m_questClone.selectNodes(xPath)
    For Each station In stations
        i = i + 1
        thisId = station.getAttribute("id")
        newId = i
        If thisId = "start" Then
            m_startId = newId
        End If
        station.setAttribute "id", newId
        xPath = "//choice[@station = '" & thisId & "']"
        Set choices = m_questClone.selectNodes(xPath)
        For Each choice In choices
            choice.setAttribute "station", newId
        Next
    Next
End Function

Private Sub moveStartStationUp()
    Dim startStation As MSXML2.IXMLDOMElement
    Dim firstStation As MSXML2.IXMLDOMElement
    Dim stations As MSXML2.IXMLDOMNodeList
    Dim xPath As String
    
    xPath = "//station"
    Set stations = m_questClone.selectNodes(xPath)
    If stations.length >= 1 Then
        xPath = "//station[@id = 'start']"
        Set startStation = m_questClone.selectSingleNode(xPath)
        xPath = "//station"
        Set firstStation = m_questClone.selectSingleNode(xPath)
        If Not ((startStation Is Nothing) Or _
                (firstStation Is Nothing)) Then
            m_questClone.documentElement.insertBefore _
                    startStation.cloneNode(True), firstStation
            m_questClone.documentElement.removeChild startStation
        End If
    End If
End Sub

