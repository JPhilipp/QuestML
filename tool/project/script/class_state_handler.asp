<%

class classStateHandler

    private m_arrState()
    private m_arrNumber()
    private m_arrNumberName()
    private m_arrNumberMin()
    private m_arrNumberMax()
    private m_arrNumberMinSet()
    private m_arrNumberMaxSet()
    private m_arrString()
    private m_arrStringName()

    ' Session _____________________________________________    

    public function getSessionDataAsString
        dim sXml
        dim i

        sXml = ""
        sXml = sXml & "<states>" & vbNewline
        for i = lBound(m_arrState) to uBound(m_arrState)
            if m_arrState(i) <> "" then
                sXml = sXml & "<state name=""" & xmlToText( m_arrState(i) ) & """ " & _
                       " value=""true"" />" & vbNewline
            end if
        next
        for i = lBound(m_arrNumber) to uBound(m_arrNumber)
            if m_arrNumber(i) <> 0 or m_arrNumberMinSet(i) or m_arrNumberMaxSet(i) then
                sXml = sXml & "<number name=""" & m_arrNumberName(i) & """ " & _
                        "value=""" & m_arrNumber(i)
                if m_arrNumberMinSet(i) then
                    sXml = sXml & " min=""" & m_arrNumberMin & """"
                end if
                if m_arrNumberMaxSet(i) then
                    sXml = sXml & " min=""" & m_arrNumberMax & """"
                end if
                sXml = sXml & """ />" & vbNewline
            end if
        next
        for i = lBound(m_arrStringName) to uBound(m_arrStringName)
            if m_arrStringName(i) <> "" then
                sXml = sXml & "<string name=""" & m_arrStringName(i) & """ " & _
                        "value=""" & xmlToText( m_arrString(i) ) & """/>" & vbNewline
            end if
        next

        sXml = sXml & "</states>" & vbNewline

        getSessionDataAsString = sXml
    end function

    public sub setSessionDataFromXml(byRef xmlSession)
        dim xPath
        dim oStates
        dim oState

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

    public sub redimArrays
        redim m_arrState(10000)
        redim m_arrNumber(5000)
        redim m_arrNumberName(5000)
        redim m_arrNumberMin(5000)
        redim m_arrNumberMax(5000)
        redim m_arrNumberMinSet(5000)
        redim m_arrNumberMaxSet(5000)
        redim m_arrString(5000)
        redim m_arrStringName(5000)
    end sub

    ' _____________________________________________

    public sub resetStates
        dim i
    
        for i = lbound(m_arrState) to ubound(m_arrState)
            m_arrState(i) = ""
        next
        for i = lbound(m_arrNumber) to ubound(m_arrNumber)
            m_arrNumberName(i) = ""
            m_arrNumber(i) = 0
            setArrNumberMin i, g_numberDefaultMin
            setArrNumberMax i, g_numberDefaultMax
        next
        for i = lbound(m_arrString) to ubound(m_arrString)
            m_arrStringName(i) = ""
            m_arrString(i) = ""
        next
    end sub

    public sub handleCheckStates(byRef station)
        dim checkStatesAgain
    
        do
            checkStatesAgain = checkStates(station)
        loop until not checkStatesAgain
    end sub

    public sub handleRandomize(byRef station)
        dim nodes
        dim node
    
        set nodes = station.selectNodes("randomize")
        for each node in nodes
            randomizeNumber node.getAttribute("number"), node.getAttribute("value")
        next
    end sub

    public sub setStatesBoth(byRef element)
        setStates element, "before"
        setStates element, "after"
    end sub

    public function getNodeState(byRef stateNode)
        dim thisCheck
        dim checkValue

        thisCheck = true
        checkValue = stateNode.getAttribute("check")
        if not isNull(checkValue) then
            checkValue = replace(checkValue, "below", "<")
            checkValue = replace(checkValue, "above", ">")
            checkValue = replace(checkValue, "equal", "=")
            checkValue = replace(checkValue, "= >", "> =")
            checkValue = replace(checkValue, "= <", "< =")
            checkValue = replace(checkValue, "=>", ">=")
            checkValue = replace(checkValue, "=<", "<=")
            checkValue = replaceAllValues(checkValue)
            if checkValue <> "" then
                thisCheck = eval(checkValue)
                thisCheck = cBool(thisCheck)
            end if
        end if
        
        getNodeState = thisCheck
    end function

    public sub setStates(byRef oStation, byVal process)
        dim child
    
        for each child in oStation.childNodes
            if child.nodeName = "state" then
                if child.getAttribute("process") = process then
                    processSetNode child
                end if
            elseif child.nodeName = "number" then
                if child.getAttribute("process") = process then
                    processNumberNode child
                end if
            elseif child.nodeName = "string" then
                if child.getAttribute("process") = process then
                    processStringNode child
                end if
            end if
        next
    end sub

    public sub setState(byRef parStateName, byRef newState)
        dim index
        dim stateName
        dim isNotState
    
        stateName = parStateName
        isNotState = false

        if lcase( left( stateName, len(g_notState) ) ) = g_notState then
            stateName = mid(stateName, len(g_notState) + 1)
            isNotState = true
        end if
    
        index = getIndex(m_arrState, stateName)
        if index = g_noIndexFound then
            index = getFreeIndex(m_arrState)
        end if
    
        if isNotState then newState = not newState
    
        if newState then
            m_arrState(index) = lcase(stateName)
        else
            m_arrState(index) = ""
        end if
    end sub

    public sub setString(byVal stringName, byVal stringValue)
        dim oldValue
        dim midValue
        dim newValue
    
        oldValue = 0
        midValue = 0
        newValue = 0
    
        oldValue = getStringOfName(stringName)

        stringValue = replaceAllValues(stringValue)
        midValue = stringValue
    
        if left(midValue, 1) = "+" then
            midValue = cStr(replace(midValue, "+", ""))
            newValue = cStr(oldValue) + cStr(midValue)
        else
            midValue = replace(midValue, "=", "")
            newValue = cStr(midValue)
        end if
    
        setStringOfName stringName, newValue
    end sub

    public function getStringOfName(byVal parStringName)
        dim i
        dim stringValue
        dim foundName
        dim stringName
    
        stringValue = ""
        stringName = lcase(parStringName)
    
        for i = lBound(m_arrStringName) to uBound(m_arrStringName)
            if lCase( m_arrStringName(i) ) = stringName then
                stringValue = m_arrString(i)
                exit for
            end if
        next

        getStringOfName = cStr(stringValue)
    end function

    public function replaceAllValues(byVal text)
        const defaultToZero = true

        text = replaceStringValues(text)
        text = replaceNumberValues(text, not defaultToZero)
        text = replaceStateValues(text)        

        text = replaceNumberValues(text, defaultToZero)

        text = replaceFunctionValues(text)

        replaceAllValues = cStr(text)
    end function

    public sub addVisits(byVal stationId)
        const endString = ")"
        const allSum = "*"
        const addOne = "+1"

        setNumberOfName g_visitsStartString & stationId & endString, addOne
        setNumberOfName g_visitsStartString & allSum & endString, addOne
    end sub

    public function getStatesInformation(byVal stationId)
        dim i
        dim xhtml
        dim xmlTemplate
        dim setList
        dim numberList
        dim stringList

        setList = ""
        numberList = ""
        stringList = ""

        for i = lbound(m_arrState) to ubound(m_arrState)
            if not m_arrState(i) = "" then
                setList = setList & "<li>" & m_arrState(i) & "</li>" & vbNewline
            end if
        next
        for i = lbound(m_arrNumber) to ubound(m_arrNumber)
            if m_arrNumberName(i) <> "" and _
                    lcase( left( m_arrNumberName(i), len(g_visitsStartString) ) ) <> g_visitsStartString then
    
                numberList = numberList & "<li>" & m_arrNumberName(i) & " is " & m_arrNumber(i)
                if getm_arrNumberMin(i) <> g_numberDefaultMin and getm_arrNumberMax(i) <> g_numberDefaultMax then
                    numberList = numberList & " (" & getm_arrNumberMin(i) & " to " & getm_arrNumberMax(i) &  ")"
                end if
                numberList = numberList & "</li>" & vbNewline
            end if
        next
        for i = lbound(m_arrString) to ubound(m_arrString)
            if m_arrStringName(i) <> "" and m_arrString(i) <> "" then
                stringList = stringList & "<li>" & m_arrStringName(i) & " is """ & m_arrString(i) & """</li>"
            end if
        next

        if setList <> "" then
            setList = "<ul>" & setList & "</ul>"
        end if
        if numberList <> "" then
            numberList = "<ul>" & numberList & "</ul>"
        end if
        if stringList <> "" then
            stringList = "<ul>" & stringList & "</ul>"
        end if

        set xmlTemplate = getXml("script/states_node.xml")
        xhtml = xmlTemplate.documentElement.xml
        xhtml = replace(xhtml, "[stationId]", stationId)
        xhtml = replace(xhtml, "[setList]", setList)
        xhtml = replace(xhtml, "[numberList]", numberList)
        xhtml = replace(xhtml, "[stringList]", stringList)

        getStatesInformation = xhtml
    end function

    public function getm_arrNumberMin(byVal i)
        if m_arrNumberMinSet(i) then
            getm_arrNumberMin = cLng( m_arrNumberMin(i) )
        else
            getm_arrNumberMin = cLng(g_numberDefaultMin)
        end if
    end function

    public function getm_arrNumberMax(byVal i)
        if m_arrNumberMaxSet(i) then
            getm_arrNumberMax = cLng( m_arrNumberMax(i) )
        else
            getm_arrNumberMax = cLng(g_numberDefaultMax)
        end if
    end function

    public sub setNumberOfName(byVal parNumberName, byVal numberValue)
        dim i
        dim numberName
    
        numberName = replaceAllValues(parNumberName)
        i = getNumberIndex(numberName)
    
        m_arrNumberName(i) = numberName
        numberValue = replaceAllValues(numberValue)
        numberValue = eval(numberValue)
        m_arrNumber(i) = numberValue
        pushNumberInMinMaxByIndex i
        clearNumberNameIfDefaultValues i
    end sub
    
    ' private __________________________________________________________

    private function checkStates(byRef station)
        dim checkStatesAgain
        dim child
        dim Ok
        dim isTrue
        dim inputCheck
        dim input
    
        input = ""
        checkStatesAgain = false
    
        handleRandomize station
    
        for each child in station.childNodes
            if child.nodeName = "input" then
                handleInput child
    
            elseif child.nodeName = "if" then
                isTrue = getNodeState(child)
    
                if isTrue then
                    if child.lastChild.nodeName = "choose" then
                        processChoose child, station, checkStatesAgain
                    else
                        set station = child
                    end if
                    exit for
                end if
            
            elseif child.nodeName = "else" then
                if child.lastChild.nodeName = "choose" then
                    processChoose child, station, checkStatesAgain
                else
                    set station = child
                end if
            end if
        next
    
        checkStates = cBool(checkStatesAgain)
    end function
  
    private sub processChoose(byRef refChild, byRef refStation, byRef refCheckStatesAgain)
        setStatesBoth refChild
        set refStation = getStation( getLink(refChild.lastChild) )
        addVisits refStation.getAttribute("id")
        refCheckStatesAgain = true
    end sub
    
    private sub handleInput(byRef child)
        dim input
    
        if g_isServerVersion then
            sendError "Input is not supported in Server-QML"
        else
            input = prompt( replaceStringValues( child.getAttribute("text") ), "" )
            if isNull(input) then
                input = ""
            end if
        end if
    
        input = convertInput(child, input)
    
        setString child.getAttribute("name"), input
        setString "qmlInput", input
    end sub
    
    private function getUserInput
        getUserInput = getStringOfName("qmlInput")
    end function
    
    private function convertInput(byRef inputNode, byVal parStringInput)
        dim stringInput
        dim max
    
        stringInput = parStringInput
    
        if inputNode.getAttribute("name") <> "" then
            if inputNode.getAttribute("max") <> "" then
                stringInput = left( stringInput, cLng( inputNode.getAttribute("max") ) )
            end if
        end if
    
        convertInput = stringInput
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
    
        numberName = numberNode.getAttribute("name")
        numberValue = numberNode.getAttribute("value")
    
        processNumberMinMax numberName, numberNode
        setNumberOfName numberName, numberValue
    end sub
    
    private sub processStringNode(byRef stringNode)
        dim stringName
        dim stringValue
    
        stringName = stringNode.getAttribute("name")
        stringValue = stringNode.getAttribute("value")
        stringValue = replaceStringValues(stringValue)
    
        setString stringName, stringValue
    end sub
    
    private sub processNumberMinMax(byVal numberName, byRef numberNode)
        dim i
    
        i = getNumberIndex(numberName)
    
        if numberNode.getAttribute("min") <> "" then
            setArrNumberMin i, cLng( replaceAllValues( numberNode.getAttribute("min") ) )
        end if
    
        if numberNode.getAttribute("max") <> "" then
            setArrNumberMax i, cLng( replaceAllValues( numberNode.getAttribute("max") ) )
        end if
    
        correctNumberMinMax i
    end sub
    
    private sub correctNumberMinMax(byRef i)
        const integerLimitMin = -32768
        const integerLimitMax = 32767
        dim tempMin
    
        if getm_arrNumberMin(i) > getm_arrNumberMax(i) then
            tempMin = getm_arrNumberMin(i)
            setArrNumberMin i, getm_arrNumberMax(i)
            setArrNumberMax i, tempMin
        end if
        if getm_arrNumberMin(i) < integerLimitMin then
            setArrNumberMin i, cLng(integerLimitMin)
        end if
        if getm_arrNumberMax(i) > integerLimitMax then
            setArrNumberMax i, cLng(integerLimitMax)
        end if
    end sub
    
    private sub setNumberMinMax(byVal numberName, byVal min, byVal max)
        dim i
    
        i = getNumberIndex(numberName)
        setArrNumberMin i, min
        setArrNumberMax i, max
    
        correctNumberMinMax i
    end sub
    
    private function getState(byVal parStateName)
        dim index
        dim state
        dim isNotState
        dim stateName
    
        stateName = parStateName
        state = false
        isNotState = false
    
        if lcase( left( stateName, len(g_notState) ) ) = g_notState then
            stateName = mid( stateName, len(g_notState) + 1 )
            isNotState = true
        end if
        stateName = lcase(stateName)
    
        index = getIndex(m_arrState, lcase(stateName))
    
        state = cBool( not (index = g_noIndexFound) )
        if isNotState then
            state = not state
        end if
    
        getState = cBool(state)
    end function
    
    private function getIndex(byRef array, byVal content)
        dim index
        dim i
    
        index = g_noIndexFound
        for i = lbound(array) to ubound(array)
            if array(i) = content then    
                index = i
                exit for
            end if
        next
    
        getIndex = index
    end function
    
    private function getFreeIndex(byRef array)
        dim index
        dim i
    
        for i = lbound(array) to ubound(array)
            if array(i) = "" then    
                index = i
                exit for
            end if
        next
    
        getFreeIndex = index
    end function
    
    private function getNumberIndex(byVal parNumberName)
        dim numberIndex
        dim i
        dim foundName
        dim numberName
    
        foundName = false
        numberIndex = 0
    
        numberName = lcase(parNumberName)
    
        for i = lbound(m_arrNumberName) to ubound(m_arrNumberName)
            if lcase(m_arrNumberName(i)) = numberName then
                numberIndex = i
                foundName = true
                exit for
            end if
        next
    
        if not foundName then
            for i = lbound(m_arrNumberName) to ubound(m_arrNumberName)
                if m_arrNumberName(i) = "" then
                    m_arrNumberName(i) = numberName
                    m_arrNumber(i) = 0
                    setArrNumberMin i, g_numberDefaultMin
                    setArrNumberMax i, g_numberDefaultMax
                    numberIndex = i
                    exit for
                end if
            next
        end if
    
        getNumberIndex = cLng(numberIndex)
    end function
    
    private function getStringIndex(byVal parStringName)
        dim stringIndex
        dim i
        dim foundName
        dim stringName
    
        foundName = false
        stringIndex = 0
    
        stringName = lcase(parStringName)
    
        for i = lbound(m_arrStringName) to ubound(m_arrStringName)
            if lcase(m_arrStringName(i)) = stringName then
                stringIndex = i
                foundName = true
                exit for
            end if
        next
    
        if not foundName then
            for i = lbound(m_arrStringName) to ubound(m_arrStringName)
                if m_arrStringName(i) = "" then
                    m_arrStringName(i) = stringName
                    m_arrString(i) = 0
                    stringIndex = i
                    exit for
                end if
            next
        end if
    
        getStringIndex = cStr(stringIndex)
    end function
    
    private sub setVisitsFromString(byVal strng)
        ' format "visits(start)=1" or "visits(*)=1"
        dim splitted
    
        splitted = split(strng, "=")
        setNumberOfName splitted(0), splitted(1)
    end sub
    
    private sub setStringOfName(byVal stringName, byVal stringValue)
        dim i
    
        i = getStringIndex(stringName)
    
        m_arrStringName(i) = stringName
        m_arrString(i) = cStr(stringValue)
    end sub
    
    private sub clearNumberNameIfDefaultValues(byVal i)
        if m_arrNumber(i) = 0 and _
                getm_arrNumberMin(i) = g_numberDefaultMin and _
                getm_arrNumberMax(i) = g_numberDefaultMax then
    
            m_arrNumberName(i) = ""
        end if
    end sub
    
    private sub pushNumberInMinMax(byVal numberName)
        dim i
    
        i = getNumberIndex(numberName)
    
        pushNumberInMinMaxByIndex i
    end sub
    
    private sub pushNumberInMinMaxByIndex(byVal i)
        dim min
        dim max
        
        min = getm_arrNumberMin(i)
        max = getm_arrNumberMax(i)
    
        m_arrNumber(i) = numberIntoMinMax(m_arrNumber(i), min, max)
    end sub
    
    private function getNumberOfName(byVal parNumberName)
        dim i
        dim numberValue
        dim foundName
        dim numberName
    
        numberValue = 0
        numberName = lcase(parNumberName)
        numberName = replaceAllValues(numberName)
    
        for i = lbound(m_arrNumberName) to ubound(m_arrNumberName)
            if lcase(m_arrNumberName(i)) = numberName then
                numberValue = m_arrNumber(i)
                exit for
            end if
        next
    
        getNumberOfName = numberValue
    end function
    
    private function replaceNumberValues(byVal text, byVal defaultToZero)
        dim oldText
    
        do
            oldText = text
            text = doReplaceNumberValues(text, defaultToZero)
            if oldText = text then
                exit do
            end if
        loop
    
        replaceNumberValues = text
    end function
    
    private function replaceStringValues(byVal text)
        dim oldText
    
        do
            oldText = text
            text = doReplaceStringValues(text)
            if oldText = text then
                exit do
            end if
        loop
    
        replaceStringValues = text
    end function
    
    private function replaceStateValues(byVal text)
        dim oldText

        do
            oldText = text
            text = doReplaceStateValues(text)
            if oldText = text then
                exit do
            end if
        loop
   
        replaceStateValues = text
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
    
    private function doReplaceNumberValues(byVal text, byVal defaultToZero)
        const startString = "["
        dim startsAt
        dim endsAt
        dim lengthOf
        dim newText
        dim numberName
        dim numberOfName
        
        newText = text
        startsAt = instr(newText, startString)
        numberOfName = ""
    
        if startsAt >= 1 then
            lengthOf = instr( mid( newText, startsAt + len(startString) ), "]" )
    
            if lengthOf >= 1 then
                numberName = mid( newText, startsAt + len(startString), lengthOf - 1 )
    
                numberOfName = getNumberOfName(numberName)

                if cLng(numberOfName) <> 0 or defaultToZero then
                    newText = left(newText, startsAt - 1) & _
                            numberOfName & _
                            mid(newText, startsAt + len(startString) + lengthOf)
                end if
    
            end if
    
        end if
    
        doReplaceNumberValues = newText
    end function
    
    private function verboseNumber(byVal oldNumber)
        dim newNumber
        
        newNumber = oldNumber
        select case cLng(newNumber)
            case 11, 12, 13
                newNumber = newNumber & "th"
            case else
    
                select case right(newNumber, 1)
                    case 1
                        newNumber = newNumber & "st"
                    case 2
                        newNumber = newNumber & "nd"
                    case 3
                        newNumber = newNumber & "rd"
                    case else
                        newNumber = newNumber & "th"
                end select
    
        end select
    
        verboseNumber = cStr(newNumber)
    end function
    
    private function formatNumber(byVal oldNumber, byVal numberFormat)
        dim newNumber
    
        newNumber = oldNumber
        if numberFormat <> "" then
            if len(newNumber) < len(numberFormat) then
                newNumber = newNumber & mid( numberFormat, len(newNumber) + 1 )
            end if
        end if
    
        formatNumber = cStr(newNumber)
    end function
    
    private function doReplaceStringValues(byVal text)
        const startString = "["
        dim startsAt
        dim endsAt
        dim lengthOf
        dim newText
        dim stringName
        dim splitted
        dim sValue

        newText = text
        startsAt = instr(newText, startString)

        if startsAt >= 1 then
            lengthOf = instr( mid( newText, startsAt + len(startString) ), "]" )
            if lengthOf >= 1 then
                stringName = mid(newText, startsAt + len(startString), lengthOf - 1)
                sValue = getStringOfName(stringName)
                if sValue <> "" then
                    newText = left(newText, startsAt - 1) & sValue & _
                            mid( newText, startsAt + len(startString) + lengthOf )
                end if
            end if
        end if
    
        doReplaceStringValues = newText
    end function

    private function doReplaceStateValues(byVal text)
        const startString = "["
        dim startsAt
        dim endsAt
        dim lengthOf
        dim stateName
        dim boolValue

        startsAt = instr(text, startString)

        if startsAt >= 1 then
            lengthOf = instr( mid( text, startsAt + len(startString) ), "]" )
            if lengthOf >= 1 then
                stateName = mid(text, startsAt + len(startString), lengthOf - 1)
                boolValue = getState(stateName)
                if boolValue then
                    text = left(text, startsAt - 1) & _
                            returnIf( boolValue, "true", "false") & _
                            mid( text, startsAt + len(startString) + lengthOf )
                end if
            end if
        end if
    
        doReplaceStateValues = text
    end function

    private function doReplaceFunctionValues(byVal text)
        const startString = "[function "
        dim startsAt
        dim endsAt
        dim lengthOf
        dim newText
        dim functionString
        dim functionOk
        dim functionName
        dim returnValue
        dim splitted

        newText = text
        startsAt = instr(newText, startString)

        if startsAt >= 1 then
            lengthOf = instr( mid( newText, startsAt + len(startString) ), "]" )
            if lengthOf >= 1 then
                functionString = mid(newText, startsAt + len(startString), lengthOf - 1)
                functionName = left( functionString, inStr(functionString, "(") - 1 )

                select case lCase(functionName)
                    case "randomnumber"
                        functionString = mid( functionString, inStr(functionString, "(") + 1 )
                        functionString = left( functionString, len(functionString) -1 )
                        splitted = split(functionString, ",")
                        returnValue = randomNumber( trim( splitted(0) ), trim( splitted(1) ) )
                    case "containsword"
                        functionString = mid( functionString, inStr(functionString, "(") + 1 )
                        functionString = left( functionString, len(functionString) -1 )
                        splitted = split(functionString, ",")
                        returnValue = containsWord( trim( splitted(0) ), trim( splitted(1) ) )
                    case else
                        returnValue = ""
                end select

                newText = left(newText, startsAt - 1) & _
                        returnValue & _
                        mid( newText, startsAt + len(startString) + lengthOf )
            end if
        end if
    
        doReplaceFunctionValues = newText
    end function
        
    private function convertString(byVal oldString, byVal conversion)
        dim newString

        newString = oldString
    
        select case conversion
            case "lower"
                newString = lcase(newString)
            case "upper"
                newString = ucase(newString)
            case "propercase"
                newString = properCase(newString)
            case "trim"
                newString = trim(newString)
        end select
    
        convertString = newString
    end function
    
    private function checkNumber(byRef numberString)
        dim splitted
        dim isOk
        dim numberName
        dim operator
        dim numberCheckValue
        dim numberRealValue
    
        numberCheckValue = 0
        numberRealValue = 0
    
        isOk = false
        numberString = trimDoubleSpaces(numberString)
        numberString = replaceAllValues(numberString)
    
        splitted = Split(numberString, " ")
    
        if ubound(splitted) = 2 then
            numberName = splitted(0)
            operator = splitted(1)
            numberCheckValue = splitted(2)
        else
            numberName = splitted(0)
            operator = splitted(1) & " " & splitted(2)
            numberCheckValue = splitted(3)
        end if
    
        numberRealValue = getNumberOfName(numberName)
    
        select case lcase(operator)
            case "is"
                isOk = ( cLng(numberRealValue) = cLng(numberCheckValue) )
            case "is below"
                isOk = ( cLng(numberRealValue) < cLng(numberCheckValue) )
            case "is equalbelow"
                isOk = ( cLng(numberRealValue) <= cLng(numberCheckValue) )
            case "is above"
                isOk = ( cLng(numberRealValue) > cLng(numberCheckValue) )
            case "is equalabove"
                isOk = ( cLng(numberRealValue) >= cLng(numberCheckValue) )
            case "is not"
                isOk = ( cLng(numberRealValue) <> cLng(numberCheckValue) )
            case else
                sendError "Erronous operator """ & numberString & """"
        end select
    
        checkNumber = isOk
    end function
    
    private sub randomizeNumber(byVal name, byRef numberValue)
        dim oldValue
        dim newValue
    
        oldValue = getNumberOfName(name)
        numberValue = trimDoubleSpaces( trim(numberValue) )
        numberValue = replaceAllValues(numberValue)
    
        if instr(numberValue, " to ") then
            newValue = getRandomNumberOfString(numberValue)
        else
            numberValue = replace(numberValue, " ", "")
            newValue = cLng(oldValue) + getRandomNumber( 0, cLng(numberValue) )
        end if
    
        setNumberOfName name, newValue
    end sub
    
    private function getRandomNumberOfString(byVal numberString)
        ' handles strings like: "1 to 6"
        dim randomNumber
        dim splitted

        splitted = split(numberString)
        randomNumber = getRandomNumber(splitted(0), splitted(2))
        getRandomNumberOfString = cLng(randomNumber)
    end function
    
    private function getRandomNumber(byVal parMin, byVal parMax)
        dim min
        dim max
    
        randomize
    
        if parMin < parMax then
            min = cLng(parMin)
            max = cLng(parMax)
        else
            min = cLng(parMax)
            max = cLng(parMin)
        end if
    
        getRandomNumber = cLng( (rnd * (max - min)) + min )
    end function
    
    private sub setArrNumberMin(byVal i, byVal value)
        if cLng(value) = g_numberDefaultMin then
            if m_arrNumberMinSet(i) <> false then
                m_arrNumberMinSet(i) = false
            end if
            if m_arrNumberMin(i) <> 0 then
                m_arrNumberMin(i) = 0
            end if
        else
            m_arrNumberMinSet(i) = true
            m_arrNumberMin(i) = cLng(value)
        end if
    end sub
    
    private sub setArrNumberMax(byVal i, byVal value)
        if cLng(value) = g_numberDefaultMax then
            if m_arrNumberMaxSet(i) <> false then
                m_arrNumberMaxSet(i) = false
            end if
            if m_arrNumberMax(i) <> 0 then
                m_arrNumberMax(i) = 0
            end if
        else
            m_arrNumberMaxSet(i) = true
            m_arrNumberMax(i) = cLng(value)
        end if
    end sub

    ' ****** user functions ******

    private function randomNumber(byVal min, byVal max)
        dim lNumber

        randomize
        lNumber = cLng( rnd * (max - min + 1) ) + min
        randomNumber = lNumber
    end function

    private function containsWord(byVal sentence, byVal parCheckWord)
        dim splitted
        dim checkWord
        dim i
        dim foundWord
    
        checkWord = lcase(parCheckWord)
        splitted = splitWords( lcase(sentence) )
    
        foundWord = false
        for i = lbound(splitted) to ubound(splitted)
            if splitted(i) = parCheckWord then
                foundWord = true
                exit for
            end if
        next
    
        containsWord = cBool(foundWord)
    end function

end class

%>