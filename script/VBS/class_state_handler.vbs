option explicit

class classStateHandler

    private m_xmlStates

    public sub class_initialize
        resetStates
    end sub

    public sub setFromStatesString(byVal statesString)
        dim sPair
        dim iPair
        dim sNameValue
        dim sName
        dim sValue
        dim arrStateName(30)
        dim arrNumberName(30)
        dim arrStringName(30)
        dim arrStateValue(30)
        dim arrNumberValue(30)
        dim arrStringValue(30)
        dim sTypeIndexSubtype
        dim sType
        dim sIndex
        dim sSubtype
        dim i
        dim vValue

        if statesString <> "" then

            for i = lBound(arrStateName) to uBound(arrStateName)
                arrStateName(i) = ""
                arrNumberName(i) = ""
                arrStringName(i) = ""
            next

            sPair = split(statesString, "&")
            for iPair = lBound(sPair) to uBound(sPair)

                sNameValue = split( sPair(iPair), "=" )
                sName = sNameValue( lBound(sNameValue) )
                sValue = sNameValue( uBound(sNameValue) )

                sTypeIndexSubtype = split(sName, "_")
                sType = sTypeIndexSubtype( lBound(sTypeIndexSubtype) )
                sIndex = sTypeIndexSubtype( lBound(sTypeIndexSubtype) + 1 )
                sSubtype = sTypeIndexSubtype( lBound(sTypeIndexSubtype) + 2 )

                select case sType
                    case "state"
                        i = cLng(sIndex)
                        if i >= lBound(arrStateName) and i <= uBound(arrStateName) then
                            if sSubtype = "name" then
                                arrStateName(i) = sValue
                            elseif sSubtype = "value" then
                                arrStateValue(i) = sValue
                            end if
                        end if

                    case "number"
                        i = cLng(sIndex)
                        if i >= lBound(arrNumberName) and i <= uBound(arrNumberName) then
                            if sSubtype = "name" then
                                arrNumberName(i) = sValue
                            elseif sSubtype = "value" then
                                arrNumberValue(i) = sValue
                            end if
                        end if

                    case "string"
                        i = cLng(sIndex)
                        if i >= lBound(arrStringName) and i <= uBound(arrStringName) then
                            if sSubtype = "name" then
                                arrStringName(i) = sValue
                            elseif sSubtype = "value" then
                                arrStringValue(i) = sValue
                            end if
                        end if

                end select

            next

            for i = lBound(arrStateName) to uBound(arrStateName)
                if arrStateName(i) <> "" then
                    vValue = cBool(arrStateValue(i) = "true")
                    setState arrStateName(i), vValue
                end if
                if arrNumberName(i) <> "" then
                    vValue = arrNumberValue(i)
                    if vValue = "" then
                        vValue = 0
                    end if
                    vValue = cLng(vValue)
                    setNumber arrNumberName(i), vValue
                end if
                if arrStringName(i) <> "" then
                    vValue = cStr( arrStringValue(i) )
                    setString arrStringName(i), vValue
                end if
            next

        end if
    end sub

    public function getSessionDataAsString
        dim sXml

        sXml = m_xmlStates.xml
        sXml = replace(sXml, "/>", "/>" & vbNewline & "        ")

        getSessionDataAsString = sXml
    end function

    public sub setSessionDataFromXml(byRef xmlSession)
        dim xPath
        dim oStates
        dim oState

        resetStates

        xPath = "//states/*"
        set oStates = xmlSession.selectNodes(xPath)
        for each oState in oStates
            select case oState.nodeName
                case "state"
                    processSetNode oState
                case "number"
                    processNumberNode oState
                case "string"
                    processStringNode oState
            end select
        next
    end sub

    public sub resetStates
        set m_xmlStates = getXmlString("<?xml version=""1.0""?><states></states>")
    end sub

    public sub handlePreStates(byRef station)
        dim oStates
        dim oState
        dim xPath
        dim ifElement

        xPath = "if"
        set ifElement = station.selectSingleNode(xPath)
        if not (ifElement is nothing) then
            xPath = "state | number| string"
            set oStates = station.selectNodes(xPath)
            for each oState in oStates
                select case oState.nodeName
                    case "state"
                        processSetNode oState
                    case "number"
                        processNumberNode oState
                    case "string"
                        processStringNode oState
                end select
            next
        end if
    end sub

    public function getNodeState(byRef stateNode)
        dim thisCheck
        dim checkValue

        thisCheck = true
        checkValue = stateNode.getAttribute("check")
        if not isNull(checkValue) then
            checkValue = replace(checkValue, "equal", "=")
            checkValue = replace(checkValue, "greater", ">")
            checkValue = replace(checkValue, "lower", "<")
            checkValue = replace(checkValue, "= >", "> =")
            checkValue = replace(checkValue, "= <", "< =")
            checkValue = replace(checkValue, "> =", ">=")
            checkValue = replace(checkValue, "< =", "<=")
            checkValue = replace(checkValue, "=>", ">=")
            checkValue = replace(checkValue, "=<", "<=")
            checkValue = replace(checkValue, "'", """")
            checkValue = replaceAllValuesQuote(checkValue)
            if checkValue <> "" then
                thisCheck = eval(checkValue)
                thisCheck = cBool(thisCheck)
            end if
        end if

        getNodeState = thisCheck
    end function

    public sub setStates(byRef child)
        select case child.nodeName
            case "state"
                processSetNode child
            case "number"
                processNumberNode child
            case "string"
                processStringNode child
        end select
    end sub

    public sub setState(byVal thisName, byVal thisValue)
        dim thisElement

        thisValue = returnIf(thisValue, "true", "false")

        set thisElement = setValue("state", thisName, thisValue)

        if thisValue = "false" then
            thisElement.parentNode.removeChild thisElement
        end if
    end sub

    public sub setString(byVal thisName, byVal thisValue)
        dim thisElement

        thisValue = replaceAllValues(thisValue)

        set thisElement = setValue("string", thisName, thisValue)
    end sub

    public sub setNumber(byVal thisName, byVal thisValue)
        dim thisElement

        thisValue = replaceAllValues(thisValue)
        if thisValue <> "" then
            thisValue = eval(thisValue)
        end if

        set thisElement = setValueNumber(thisName, thisValue)
    end sub

    public sub setNumberWithMinMax(byVal thisName, byVal thisValue, byVal min, byVal max)
        dim thisElement

        thisValue = replaceAllValues(thisValue)
        if thisValue <> "" then
            thisValue = eval(thisValue)
        end if

        set thisElement = setValueNumberWithMinMax(thisName, thisValue, min, max)
    end sub

    public function getState(byVal thisName)
        dim thisValue

        thisValue = getValue("state", thisName)
        thisValue = cBool(thisValue = "true")

        getState = thisValue
    end function

    public function getNumber(byVal thisName)
        dim thisValue

        thisValue = getValue("number", thisName)
        if thisValue = "" then
            thisValue = 0
        end if

        getNumber = thisValue
    end function

    public function getString(byVal thisName)
        dim thisValue

        thisValue = getValue("string", thisName)

        getString = thisValue
    end function

    public function replaceAllValues(byVal text)
        const quoteString = false
        replaceAllValues = cStr( replaceAllValuesOption(text, quoteString) )
    end function

    private function replaceAllValuesQuote(byVal text)
        const quoteString = true
        replaceAllValuesQuote = cStr( replaceAllValuesOption(text, quoteString) )
    end function

    private function replaceAllValuesOption(byVal text, byVal quoteString)
        text = replaceValuesOf("string", text, false, quoteString)
        text = replaceValuesOf("number", text, false, quoteString)
        text = replaceValuesOf("state", text, false, quoteString)

        text = replaceValuesOf("number", text, true, quoteString)
        text = replaceFunctionValues(text)

        replaceAllValuesOption = cStr(text)
    end function

    public sub addVisits(byVal stationId)
        dim allName
        dim thisName

        allName = "qmlVisits(*)"
        thisName = "qmlVisits(" & stationId & ")"
        setNumber allName, getNumber(allName) + 1
        setNumber thisName, getNumber(thisName) + 1
    end sub

    public function getStatesInformation(byVal stationId)
        dim xhtml
        dim xmlTemplate
        dim stateList
        dim numberList
        dim stringList
        dim xPath
        dim stateElements
        dim stateElement
        dim thisValue
        dim thisName
        dim internalState
        dim i
        dim sStart
        dim sEnd
        dim min
        dim max

        stateList = ""
        numberList = ""
        stringList = ""

        for i = 1 to 2
            xPath = "//state|//number|//string"
            set stateElements = m_xmlStates.selectNodes(xPath)
            for each stateElement in stateElements
                thisName = stateElement.getAttribute("name")
                thisValue = stateElement.getAttribute("value")
                internalState = inStr(thisName, "qml") = 1

                if i = 1 then
                    sStart = "<li>"
                    sEnd = "</li>"
                else ' if i = 2 then
                    sStart = "<li><em>"
                    sEnd = "</em></li>"
                end if

                if ( i = 1 and (not internalState) ) or (i = 2 and internalState) then
                    select case stateElement.nodeName
                        case "state"
                            stateList = stateList & sStart & thisName & " = " & thisValue & sEnd
                        case "number"
                            numberList = numberList & sStart & thisName & " = " & thisValue & sEnd
                            min = stateElement.getAttribute("min")
                            max = stateElement.getAttribute("max")
                            if ( not isNull(min) ) or ( not isNull(max) ) then
                                numberList = numberList & " (min=" & min & ", max=" & max & ")"
                            end if
                        case "string"
                            stringList = stringList & sStart & thisName & " = """ & thisValue & """" & sEnd
                    end select
                end if
            next
        next

        if stateList <> "" then
            stateList = "<ul>" & stateList & "</ul>"
        end if
        if numberList <> "" then
            numberList = "<ul>" & numberList & "</ul>"
        end if
        if stringList <> "" then
            stringList = "<ul>" & stringList & "</ul>"
        end if

        set xmlTemplate = getXml("script/states_node.xml")
        xhtml = xmlTemplate.documentElement.xml
        ' xhtml = replace( xhtml, "[xml]", replace( xmlToText(m_xmlStates.xml) , "&gt;", "&gt;<br />") )
        xhtml = replace(xhtml, "[stationId]", """" & stationId & """")
        xhtml = replace(xhtml, "[stateList]", stateList)
        xhtml = replace(xhtml, "[numberList]", numberList)
        xhtml = replace(xhtml, "[stringList]", stringList)

        getStatesInformation = xhtml
    end function

    ' private __________________________________________________________

    private function getValue(byVal thisNodeName, byVal thisName)
        dim thisElement
        dim xPath
        dim thisValue
        dim min
        dim max

        thisValue = ""
        xPath = "//" & thisNodeName & "[@name = '" & thisName & "']"
        set thisElement = m_xmlStates.selectSingleNode(xPath)
        if not (thisElement is nothing) then
            thisValue = thisElement.getAttribute("value")

            if isNull(thisValue) then
                thisValue = ""

            elseif thisNodeName = "number" then
                min = thisElement.getAttribute("min")
                max = thisElement.getAttribute("max")
                if not isNull(min) then
                    if cLng(thisValue) < cLng(min) then
                        thisValue = min
                    end if
                end if
                if not isNull(max) then
                    if cLng(thisValue) > cLng(max) then
                        thisValue = max
                    end if
                end if

            end if
        end if

        getValue = thisValue
    end function

    private function setValue(byVal thisNodeName, byVal thisName, byVal thisValue)
        dim thisElement
        dim xPath

        xPath = "//" & thisNodeName & "[@name = '" & thisName & "']"
        set thisElement = m_xmlStates.selectSingleNode(xPath)
        if thisElement is nothing then
            set thisElement = m_xmlStates.createElement(thisNodeName)
            set thisElement = m_xmlStates.documentElement.appendChild(thisElement)
        end if

        thisElement.setAttribute "name", thisName
        thisElement.setAttribute "value", thisValue

        set setValue = thisElement
    end function

    private function setValueNumber(byVal thisName, byVal thisValue)
        dim thisElement
        dim xPath
        dim thisNodeName
        dim min
        dim max

        thisNodeName = "number"
        xPath = "//" & thisNodeName & "[@name = '" & thisName & "']"
        set thisElement = m_xmlStates.selectSingleNode(xPath)
        if thisElement is nothing then
            set thisElement = m_xmlStates.createElement(thisNodeName)
            set thisElement = m_xmlStates.documentElement.appendChild(thisElement)
        end if

        min = thisElement.getAttribute("min")
        max = thisElement.getAttribute("max")

        if not isNull(min) then
            thisElement.setAttribute "min", min
            if cLng(thisValue) < cLng(min) then
                thisValue = min
            end if
        end if
        if not isNull(max) then
            thisElement.setAttribute "max", max
            if cLng(thisValue) > cLng(max) then
                thisValue = max
            end if
        end if

        thisElement.setAttribute "name", thisName
        thisElement.setAttribute "value", thisValue

        set setValueNumber = thisElement
    end function

    private function setValueNumberWithMinMax(byVal thisName, byVal thisValue, byVal min, byVal max)
        dim thisElement
        dim xPath
        dim thisNodeName

        thisNodeName = "number"
        xPath = "//" & thisNodeName & "[@name = '" & thisName & "']"
        set thisElement = m_xmlStates.selectSingleNode(xPath)
        if thisElement is nothing then
            set thisElement = m_xmlStates.createElement(thisNodeName)
            set thisElement = m_xmlStates.documentElement.appendChild(thisElement)
        end if

        if not isNull(min) then
            thisElement.setAttribute "min", min
            if cLng(thisValue) < cLng(min) then
                thisValue = min
            end if
        end if
        if not isNull(max) then
            thisElement.setAttribute "max", max
            if cLng(thisValue) > cLng(max) then
                thisValue = max
            end if
        end if
        thisElement.setAttribute "name", thisName
        thisElement.setAttribute "value", thisValue

        set setValueNumberWithMinMax = thisElement
    end function

    private function getIsTrue(byVal state, byVal relation)
        dim isTrue
    
        isTrue = false
        if state then
            isTrue = true
        elseif relation = "and" then
            isTrue = false
        end if
    
        getIsTrue = cBool(isTrue)
    end function
    
    private sub processSetNode(byRef setNode)
        dim state
        dim stateNew
    
        state = lCase( setNode.getAttribute("name") )
        stateNew = cBool( "true" = setNode.getAttribute("value") )
        setState state, stateNew
    end sub

    private sub processNumberNode(byRef numberNode)
        dim numberName
        dim numberValue
        dim min
        dim max

        numberName = numberNode.getAttribute("name")
        numberValue = numberNode.getAttribute("value")
        min = numberNode.getAttribute("min")
        max = numberNode.getAttribute("max")

        setNumberWithMinMax numberName, numberValue, min, max
    end sub
    
    private sub processStringNode(byRef stringNode)
        dim stringName
        dim stringValue
    
        stringName = stringNode.getAttribute("name")
        stringValue = stringNode.getAttribute("value")
        stringValue = replaceValuesOf("string", stringValue, false, false)

        setString stringName, stringValue
    end sub
    
    private sub setVisitsFromString(byVal strng)
        ' format "visits(start)=1" or "visits(*)=1"
        dim splitted
    
        splitted = split(strng, "=")
        setNumber splitted(0), splitted(1)
    end sub

    private function replaceValuesOf(byVal sValueType, byVal text, byVal forceDefault, byVal quoteString)
        const startString = "["
        dim startsAt
        dim endsAt
        dim lengthOf
        dim valueName
        dim splitted
        dim sValue
        dim doUse

        startsAt = 0
        do
            if isNull(text) then
                text = ""
            end if

            startsAt = instr(startsAt + 1, text, startString)
            if startsAt >= 1 then
                lengthOf = instr( mid( text, startsAt + len(startString) ), "]" )
                if lengthOf >= 1 then
                    valueName = mid(text, startsAt + len(startString), lengthOf - 1)

                    doUse = false
                    select case sValueType
                        case "state"
                            sValue = getState(valueName)
                            sValue = returnIf( cBool(sValue), "true", "false" )
                            doUse = cBool(sValue) or forceDefault
                        case "number"
                            sValue = getNumber(valueName)
                            doUse = sValue <> 0 or forceDefault
                        case "string"
                            sValue = getString(valueName)
                            if sValue <> "" then
                                if quoteString then
                                    sValue = """" & sValue & """"
                                else
                                    sValue = "" & sValue & ""
                                end if
                            end if
                            doUse = sValue <> "" or forceDefault
                    end select

                    if doUse then
                        text = left(text, startsAt - 1) & sValue & _
                                mid( text, startsAt + len(startString) + lengthOf )
                    end if
                end if
            end if
        loop until not startsAt >= 1
    
        replaceValuesOf = text
    end function

    private function replaceFunctionValues(byVal text)
        dim oldText

        do
            oldText = text
            text = doReplaceFunctionValues(text)
            if oldText = text then
                exit do
            end if
        loop
   
        replaceFunctionValues = text
    end function
    
    private function doReplaceFunctionValues(byVal text)
        const startString = "{"
        dim startsAt
        dim endsAt
        dim lengthOf
        dim newText
        dim functionString
        dim functionOk
        dim returnValue
        dim oInlineFunction

        newText = text
        startsAt = instr(newText, startString)
        if startsAt >= 1 then

            lengthOf = instr( mid( newText, startsAt + len(startString) ), "}" )
            if lengthOf >= 1 then
                functionString = mid(newText, startsAt + len(startString), lengthOf - 1)

                set oInlineFunction = new classInlineFunction
                oInlineFunction.setInlineString functionString
                oInlineFunction.setXmlStates m_xmlStates
                oInlineFunction.process
                returnValue = oInlineFunction.getXhtml
                newText = left(newText, startsAt - 1) & returnValue & _
                        mid( newText, startsAt + len(startString) + lengthOf )
            end if
        end if

        doReplaceFunctionValues = newText
    end function
        
end class