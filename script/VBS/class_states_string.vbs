class classStatesString

    public function getStatesFromChoice(byVal choice, byRef oStateHandler)
        dim xPath
        dim stateElements
        dim stateElement
        dim thisName
        dim thisNodeName
        dim thisValue
        dim iState
        dim iNumber
        dim iString
        dim statesString

        iState = 0
        iNumber = 0
        iString = 0
        statesString = ""

        xPath = ".//state | .//number | .//string"
        set stateElements = choice.selectNodes(xPath)
        for each stateElement in stateElements

            thisNodeName = stateElement.nodeName
            thisName = stateElement.getAttribute("name")
            thisValue = stateElement.getAttribute("value")

            if isNull(thisValue) then
                select case thisNodeName
                    case "state"
                        thisValue = true
                    case "number"
                        thisValue = 0
                    case "string"
                        thisValue = ""
                end select
            end if

            thisName = oStateHandler.replaceAllValues(thisName)
            thisValue = oStateHandler.replaceAllValues(thisValue)

            select case thisNodeName
                case "state"
                    thisValue = returnIf(thisValue, "true", "false")
                case "number"
                    thisValue = cStr(thisValue)
                    thisValue = eval(thisValue)
                case "string"
                    thisValue = cStr(thisValue)
            end select

            if statesString <> "" then
                statesString = statesString & "&amp;"
            end if

            select case thisNodeName
                case "state"
                    iState = iState + 1
                    statesString = statesString & "state_" & iState & "_name=" & thisName & "&amp;"
                    statesString = statesString & "state_" & iState & "_value=" & thisValue
                case "number"
                    inumber = inumber + 1
                    statesString = statesString & "number_" & iNumber & "_name=" & thisName & "&amp;"
                    statesString = statesString & "number_" & iNumber & "_value=" & thisValue
                case "string"
                    iString = iString + 1
                    statesString = statesString & "string_" & iString & "_name=" & thisName & "&amp;"
                    statesString = statesString & "string_" & iString & "_value=" & thisValue
            end select

        next

        getStatesFromChoice = statesString
    end function

    public function getStatesString
        dim sValue
        dim sStates
        dim sNumbers
        dim sStrings
    
        sValue = ""
        sStates = getStatesOfType("state")
        sNumbers = getStatesOfType("number")
        sStrings = getStatesOfType("string")
    
        if sStates <> "" then
            sStates = sStates & "&"
        end if
        if sNumbers <> "" then
            sNumbers = sNumbers & "&"
        end if
        if sStrings <> "" then
            sStrings = sStrings & "&"
        end if
    
        sValue = sStates & sNumbers & sStrings
        if sValue <> "" then
            sValue = left( sValue, len(sValue) - len("&") )
        end if
    
        getStatesString = sValue
    end function
    
    ' private ________________________________

    private function getStatesOfType(byVal sType)
        const statesMax = 30
        dim sValue
        dim i
        dim thisPair
    
        sValue = ""
        for i = 1 to statesMax
            thisPair = getStatesStringByPrefix( sType & "_" & cStr(i) & "_" )
            if thisPair <> "" then
                if sValue <> "" then
                    sValue = sValue & "&"
                end if
                sValue = sValue & thisPair
            end if
        next
    
        getStatesOfType = sValue
    end function
    
    private function getStatesStringByPrefix(byVal sPrefix)
        dim thisName
        dim thisValue
        dim sValue
    
        sValue = ""
        thisName = request.queryString(sPrefix & "name")
        thisValue = request.queryString(sPrefix & "value")
        if thisName <> "" then
            sValue = sPrefix & "name=" & thisName & "&" & _
                    sPrefix & "value=" & thisValue
        end if
    
        getStatesStringByPrefix = sValue
    end function

end class