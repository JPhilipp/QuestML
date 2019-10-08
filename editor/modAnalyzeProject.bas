Attribute VB_Name = "modAnalyzeProject"
Option Explicit

Public Function prepareAnalysisPage() As String
    Dim template As String
    Dim newFile As String
    Dim success As Boolean
    Dim graphFileName As String
    
    template = getFileText(frmMain.qmlTopPath & "\tool\qmledit" & _
            "\graph.tpl", success)
    
    If success Then
        newFile = getAnalysisFileString(template)
        graphFileName = frmMain.qmlTopPath & _
                "\help\graph_temp.htm"
        setFileText graphFileName, newFile
    End If
    
    prepareAnalysisPage = graphFileName
End Function

Private Function getAnalysisFileString(ByVal template As String)
    Const startNode = "<param name=""center"" value=""[startNode]"">"
    Dim fileString As String, startParam As String
    
    fileString = template
    fileString = Replace$(fileString, "[description]", _
            getAnalyzeString)
    If startNodeExists Then
        startParam = "<param name=""center"" value=""start"" />"
    Else
        startParam = ""
    End If
    fileString = Replace$(fileString, "[startParam]", _
            startParam)
    fileString = Replace$(fileString, "[nodes]", _
            getNodesString)
    
    getAnalysisFileString = fileString
End Function

Private Function startNodeExists() As Boolean
    startNodeExists = frmMain.objQuest.documentElement. _
            selectNodes("./station[@id = ""start""]").length = 1
End Function

Private Function getNodesString() As String
    Const seperator = ","
    Dim nodesString As String
    Dim child As IXMLDOMElement
    Dim choices As IXMLDOMNodeList
    Dim choice As IXMLDOMElement
    Dim nodeString As String
    Dim pathTo As String
    Dim leadsToOtherChapter As Boolean

    For Each child In frmMain.objQuest.documentElement.childNodes
        
        If child.nodeName = elementStation Then
            Set choices = child.selectNodes(".//" & elementPath)
            For Each choice In choices
                nodeString = child.getAttribute("id")
                pathTo = choice.getAttribute("station")
                leadsToOtherChapter = InStr(pathTo, ":") >= 1
                If Not leadsToOtherChapter And pathTo <> "back" Then
                    nodeString = nodeString & "-" & pathTo
                    nodesString = nodesString & nodeString & seperator
                End If
            Next
        
        End If
        
    Next
    
    If Len(nodesString) > Len(seperator) Then
        nodesString = Left$(nodesString, _
                Len(nodesString) - Len(seperator))
    End If
        
    getNodesString = nodesString
End Function

Private Function getAnalyzeString() As String
    Dim text As String
    Dim stations As Long
    
    stations = getNumberOfElements(elementStation)
    text = "The project consists of " & stations & " stations " & _
            "with each having an average of " & getAverageChoices(stations) & " choices. " & _
            "A station has an average of " & getAverageIfs(stations) & " if-statements. " & _
            "Numbers and states are set " & getNumberOfStateSets & " times. " & _
            getEndingStationsString & _
            getChoicesToNothing
    
    getAnalyzeString = text
End Function

Private Function getEndingStationsString()
    Dim text As String, child As IXMLDOMElement, _
            stationsWithoutChoice As Long, _
            stationSamples As String, stationSample As String
    
    For Each child In frmMain.objQuest.documentElement.childNodes
        If child.nodeName = elementStation Then
            If child.selectNodes(".//choice").length = 0 Then
                
                stationsWithoutChoice = stationsWithoutChoice + 1
                
                If stationsWithoutChoice < 4 Then
                    stationSamples = stationSamples & _
                            child.getAttribute("id") & ", "
                End If
                
            End If
        End If
    Next
    
    If Len(stationSamples) > Len(", ") Then
        stationSamples = Left$( _
                stationSamples, Len(stationSamples) - Len(", "))
    End If
    
    If stationsWithoutChoice = 0 Then
        text = "All stations can be continued."
    Else
        If stationsWithoutChoice > 1 Then
            text = stationsWithoutChoice & " stations can't be " & _
                    "continued. Names of these stations are " & _
                    stationSamples & "..."
        ElseIf stationsWithoutChoice = 1 Then
            text = "1 station, """ & stationSamples & """, can't " & _
                    "be continued."
        End If
    End If
    
    getEndingStationsString = text
End Function

Private Function getChoicesToNothing() As String
    Const seperator = ", ", stationsToList = 4
    Dim choices As IXMLDOMNodeList, choice As IXMLDOMElement, _
            text As String, stationId As String, listedI As Integer
    
    Set choices = frmMain.objQuest.documentElement. _
            selectNodes("//" & elementPath)
    
    For Each choice In choices
        stationId = choice.getAttribute("station")
        If stationId <> "back" And InStr(stationId, ":") < 1 Then
            If Not frmMain.stationExists(stationId) Then
                listedI = listedI + 1
                If listedI <= stationsToList Then
                    text = text & stationId & ", "
                Else
                    Exit For
                End If
            End If
        End If
    Next
    
    If Len(text) > Len(seperator) Then
        text = Left$(text, Len(text) - Len(seperator))
        If listedI = 1 Then
            text = " One choice leads to the unwritten station " & _
                    text & "... "
        Else
            text = " Some choices lead to unwritten stations like " & _
                    text & "... "
        End If
    Else
        text = " No choices lead to unwritten stations. "
    End If
    
    getChoicesToNothing = text
End Function

Private Function getNumberOfElements(ByVal elementName As String) As Long
    getNumberOfElements = frmMain.objQuest.documentElement. _
        selectNodes("//" & elementName).length
End Function

Private Function getAverageChoices(ByVal stations As Long) As Single
    getAverageChoices = getAverageOf(stations, _
            getNumberOfElements(elementPath))
End Function

Private Function getAverageIfs(ByVal stations As Long) As Single
    getAverageIfs = getAverageOf(stations, _
            getNumberOfElements(elementIf))
End Function

Private Function getAverageOf(ByVal whole As Long, ByVal part As Long) As Single
    Const singlePattern = "#.#"
    Dim returnValue As Single
    If part > 0 Then
        returnValue = Format$(part / whole, singlePattern)
    End If
    getAverageOf = returnValue
End Function

Private Function getNumberOfStateSets() As Long
    getNumberOfStateSets = getNumberOfElements("set") & _
            getNumberOfElements("name")
End Function
