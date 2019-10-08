option explicit

sub handleCheckStates(station)
    dim checkStatesAgain
    do
        checkStatesAgain = checkStates(station)
    loop until not checkStatesAgain
end sub

function checkStates(station)
    dim checkStatesAgain
    dim child
    dim chanceOk
    dim isTrue
    dim inputCheck
    dim input

    input = ""
    checkStatesAgain = false

    handleRandomize station

    for each child in station.childNodes
        if child.nodeName = elementInput then
            handleInput child

        elseif child.nodeName = elementIf then
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

sub handleRandomize(station)
    dim nodes
    dim node
    set nodes = station.selectNodes(elementRandomize)
    for each node in nodes
        randomizeNumber node.getAttribute("number"), node.getAttribute("value")
    next
end sub

sub processChoose(refChild, refStation, refCheckStatesAgain)
    setStatesBoth refChild
    set refStation = getStation( getLink(refChild.lastChild) )
    addVisits refStation.getAttribute("id")
    refCheckStatesAgain = true
end sub

sub setStatesBoth(element)
    setStates element, "before"
    setStates element, "after"
end sub

sub handleInput(child)
    dim input

    if serverVersion then
        sendError language("Input is currently not supported in Server-QML", _
                "Input ist in der QML-Server Version zur Zeit nicht enthalten.")
    else
        input = prompt(replaceStringValues(child.getAttribute("text")), "")
        if isNull(input) then input = ""
    end if

    input = convertInput(child, input)

    setString child.getAttribute("name"), input
    setString "qmlInput", input
end sub

function getUserInput
    getUserInput = getStringOfName("qmlInput")
end function

function convertInput(inputNode, parStringInput)
    dim stringInput
    dim max
    stringInput = parStringInput

    if inputNode.getAttribute("name") <> "" then
        if inputNode.getAttribute("max") <> "" then
            stringInput = left(stringInput, cInt(inputNode.getAttribute("max")))
        end if
    end if

    convertInput = stringInput
end function

function getNodeState(stateNode)
    dim relation
    dim isTrue
    dim state
    dim stateCheck
    dim numberIndex
    dim numberString
    dim numberOk
    dim stringOk
    dim inputCheck
    dim chanceOk
    dim input
    input = getUserInput

    relation = stateNode.getAttribute("relation")
    
    isTrue = true ' cBool(relation = relationAnd)

    for numberIndex = 1 to maximumStateAttributes

        if numberIndex = 1 then
            numberString = ""
        else
            numberString = numberIndex
        end if
    
        if input <> "" then
            if stateNode.getAttribute("input" & numberString) <> "" then
                inputCheck = lcase(stateNode.getAttribute("input" & numberString))
                isTrue = getIsTrue(input = inputCheck, relation)
                if (relation = relationAnd and not isTrue) or _
                       (relation = relationOr and isTrue) then
                    exit for
                end if
            end if
        end if
    
        if stateNode.getAttribute("state" & numberString) <> "" then
            state = stateNode.getAttribute("state" & numberString)
            stateCheck = "true" = stateNode.getAttribute("is")
            isTrue = getIsTrue( getState(state) = stateCheck, relation )

            if (relation = relationAnd and not isTrue) or _
               (relation = relationOr and isTrue) then
                exit for
            end if
        end if
    
        if stateNode.getAttribute("number" & numberString) <> "" then
            numberOk = checkNumber(stateNode.getAttribute("number" & numberString))
            isTrue = getIsTrue(numberOk, relation)

            if (relation = relationAnd and not isTrue) or _
                   (relation = relationOr and isTrue) then
                exit for
            end if
        end if

        if stateNode.getAttribute("string" & numberString) <> "" then
            stringOk = checkString(stateNode.getAttribute("string" & numberString))
            isTrue = getIsTrue(stringOk, relation)

            if (relation = relationAnd and not isTrue) or _
                   (relation = relationOr and isTrue) then
                exit for
            end if
        end if
    
    next
    
    if stateNode.getAttribute("chance") <> "" then
    
        chanceOk = getChance(stateNode.getAttribute("chance"))
        if chanceOk then
            isTrue = true
        elseif relation = relationAnd then
            isTrue = false
        end if
    end if
    
    getNodeState = cBool(isTrue)
end function

function getIsTrue(state, relation)
    dim isTrue
    isTrue = false
    if state then
        isTrue = true
    elseif relation = relationAnd then
        isTrue = false
    end if

    getIsTrue = cBool(isTrue)
end function

sub setStates(oStation, process)
    dim child

    for each child in oStation.childNodes
        if child.nodeName = "set" then
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

sub processSetNode(setNode)
    dim state
    dim stateNew

    state = lcase( setNode.getAttribute("state") )
    stateNew = cBool("true" = setNode.getAttribute("is"))
    setState state, stateNew
end sub

sub processNumberNode(numberNode)
    dim numberName
    dim numberValue

    numberName = numberNode.getAttribute("name")
    numberValue = numberNode.getAttribute("value")
    numberValue = replaceNumberValues(numberValue)

    processNumberMinMax numberName, numberNode
    setNumber numberName, numberValue
end sub

sub processStringNode(stringNode)
    dim stringName
    dim stringValue

    stringName = stringNode.getAttribute("name")
    stringValue = stringNode.getAttribute("value")
    stringValue = replaceStringValues(stringValue)

    setString stringName, stringValue
end sub

sub processNumberMinMax(numberName, numberNode)
    dim i
    i = getNumberIndex(numberName)

    if numberNode.getAttribute("min") <> "" then
        setArrNumberMin i, cLng(replaceNumberValues(numberNode.getAttribute("min")))
    end if

    if numberNode.getAttribute("max") <> "" then
        setArrNumberMax i, cLng(replaceNumberValues(numberNode.getAttribute("max")))
    end if

    correctNumberMinMax i
end sub

sub correctNumberMinMax(i)
    const integerLimitMin = -32768, integerLimitMax = 32767
    dim tempMin
    if getArrNumberMin(i) > getArrNumberMax(i) then
        tempMin = getArrNumberMin(i)
        setArrNumberMin i, getArrNumberMax(i)
        setArrNumberMax i, tempMin
    end if
    if getArrNumberMin(i) < integerLimitMin then
        setArrNumberMin i, cInt(integerLimitMin)
    end if
    if getArrNumberMax(i) > integerLimitMax then
        setArrNumberMax i, cInt(integerLimitMax)
    end if
end sub

sub setNumberMinMax(numberName, min, max)
    dim i
    i = getNumberIndex(numberName)
    setArrNumberMin i, min
    setArrNumberMax i, max

    correctNumberMinMax i
end sub

sub setState(parStateName, newState)
    dim index
    dim stateName
    dim isNotState

    stateName = parStateName
    isNotState = false

    if lcase(left(stateName, len(notState))) = notState then
        stateName = mid(stateName, len(notState) + 1)
        isNotState = true
    end if

    index = getIndex(arrState, stateName)
    if index = noIndexFound then
        index = getFreeIndex(arrState)
    end if

    if isNotState then newState = not newState

    if newState then
        arrState(index) = lcase(stateName)
    else
        arrState(index) = ""
    end if
end sub

function getState(parStateName)
    dim index
    dim state
    dim isNotState
    dim stateName

    stateName = parStateName
    state = false
    isNotState = false

    if lcase(left(stateName, len(notState))) = notState then
        stateName = mid(stateName, len(notState)+1)
        isNotState = true
    end if
    stateName = lcase(stateName)

    index = getIndex(arrState, lcase(stateName))

    state = cBool(not (index = noIndexFound))
    if isNotState then state = not state

    getState = cBool(state)
end function

function getIndex(array, content)
    dim index, i
    index = noIndexFound
    for i = lbound(array) to ubound(array)
        if array(i) = content then    
            index = i
            exit for
        end if
    next

    getIndex = index
end function

function getFreeIndex(array)
    dim index, i
    for i = lbound(array) to ubound(array)
        if array(i) = "" then    
            index = i
            exit for
        end if
    next

    getFreeIndex = index
end function

sub resetStates
    dim i
    for i = lbound(arrState) to ubound(arrState)
        arrState(i) = ""
    next
    for i = lbound(arrNumber) to ubound(arrNumber)
        arrNumberName(i) = ""
        arrNumber(i) = 0
        setArrNumberMin i, numberDefaultMin
        setArrNumberMax i, numberDefaultMax
    next
    for i = lbound(arrString) to ubound(arrString)
        arrStringName(i) = ""
        arrString(i) = ""
    next
    setQmlStartVariables
    setQmlVariables
end sub

sub setNumber(numberName, numberValue)
    dim oldValue
    dim midValue
    dim newValue

    oldValue = 0
    midValue = 0
    newValue = 0

    oldValue = getNumberOfName(numberName)

    if instr(numberValue, "[") >= 1 then
        numberValue = replaceNumberValues(numberValue)
    end if
    midValue = numberValue

    select case left(midValue, 1)
        case "+"
            midValue = cInt(replace(midValue, "+", ""))
            newValue = cInt(oldValue) + cInt(midValue)
        case "-"
            midValue = cInt(replace(midValue, "-", ""))
            newValue = cInt(oldValue) - cInt(midValue)
        case "*"
            midValue = cInt(replace(midValue, "*", ""))
            newValue = cInt(oldValue) * cInt(midValue)
        case "/"
            midValue = cInt(replace(midValue, "/", ""))
            if cInt(midValue) <> 0 then
                newValue = cInt(oldValue) \ cInt(midValue)
            else
                newValue = cInt(oldValue)
            end if
        case "^"
            midValue = cInt(replace(midValue, "^", ""))
            newValue = cInt(oldValue) ^ cInt(midValue)
        case else
            midValue = replace(midValue, "=", "") '
            newValue = cInt(midValue)
    end select

    setNumberOfName numberName, newValue
end sub

sub setString(stringName, stringValue)
    dim oldValue
    dim midValue
    dim newValue

    oldValue = 0
    midValue = 0
    newValue = 0

    oldValue = getStringOfName(stringName)

    if instr(stringValue, "[") >= 1 then
        stringValue = replaceStringValues(stringValue)
        stringValue = replaceNumberValues(stringValue)
    end if
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

function getNumberIndex(parNumberName)
    dim numberIndex
    dim i
    dim foundName
    dim numberName

    foundName = false
    numberIndex = 0

    numberName = lcase(parNumberName)

    for i = lbound(arrNumberName) to ubound(arrNumberName)
        if lcase(arrNumberName(i)) = numberName then
            numberIndex = i
            foundName = true
            exit for
        end if
    next

    if not foundName then
        for i = lbound(arrNumberName) to ubound(arrNumberName)
            if arrNumberName(i) = "" then
                arrNumberName(i) = numberName
                arrNumber(i) = 0
                setArrNumberMin i, numberDefaultMin
                setArrNumberMax i, numberDefaultMax
                numberIndex = i
                exit for
            end if
        next
    end if

    getNumberIndex = cInt(numberIndex)
end function

function getStringIndex(parStringName)
    dim stringIndex
    dim i
    dim foundName
    dim stringName

    foundName = false
    stringIndex = 0

    stringName = lcase(parStringName)

    for i = lbound(arrStringName) to ubound(arrStringName)
        if lcase(arrStringName(i)) = stringName then
            stringIndex = i
            foundName = true
            exit for
        end if
    next

    if not foundName then
        for i = lbound(arrStringName) to ubound(arrStringName)
            if arrStringName(i) = "" then
                arrStringName(i) = stringName
                arrString(i) = 0
                stringIndex = i
                exit for
            end if
        next
    end if

    getStringIndex = cStr(stringIndex)
end function

sub setNumberFromCookie(parStrng)
    ' of format "x=10" or "x=10(-100 100)"
    ' and format "visits(start)=1" or "visits(*)=1"
    dim numberName
    dim numberValue
    dim min
    dim max
    dim minMaxStrng
    dim strng
    dim bracket
    dim splitted

    strng = parStrng

    if lcase(left( strng, len(visitsStartString) )) = visitsStartString then
        setVisitsFromString strng
    else
        min = numberDefaultMin
        max = numberDefaultMax
        numberName = ""
        numberValue = 0
    
        bracket = instr(strng, "(")
        if bracket >= 1 then
            minMaxStrng = mid(strng, bracket + 1)
            minMaxStrng = left( minMaxStrng, len(minMaxStrng) - len(")") )
            splitted = split(minMaxStrng)
            min = cInt(splitted(0))
            max = cInt(splitted(1))
            strng = left(strng, bracket-1)
        else
            min = numberDefaultMin
            max = numberDefaultMax
        end if
    
        numberName = left(strng, instr(strng, "=") -1)
        numberValue = cInt(mid(strng, instr(strng, "=") +1))
    
        setNumberMinMax numberName, min, max
        setNumberOfName numberName, numberValue
    end if
end sub

sub setVisitsFromString(strng)
    ' format "visits(start)=1" or "visits(*)=1"
    dim splitted
    splitted = split(strng, "=")
    setNumber splitted(0), splitted(1)
end sub

sub setStringFromCookie(parStrng)
    dim splitted
    splitted = split(parStrng, "=")
    setString splitted(0), splitted(1)
end sub

sub setNumberOfName(parNumberName, numberValue)
    dim i
    dim numberName

    numberName = replaceAllValues(parNumberName)

    i = getNumberIndex(numberName)

    arrNumberName(i) = numberName
    arrNumber(i) = cInt(numberValue)
    pushNumberInMinMaxByIndex i
    clearNumberNameIfDefaultValues i
end sub

sub setStringOfName(stringName, stringValue)
    dim i
    i = getStringIndex(stringName)

    arrStringName(i) = stringName
    arrString(i) = cStr(stringValue)
end sub

sub clearNumberNameIfDefaultValues(i)
    if arrNumber(i) = 0 and _
            getArrNumberMin(i) = numberDefaultMin and _
            getArrNumberMax(i) = numberDefaultMax then

        arrNumberName(i) = ""
    end if
end sub

sub pushNumberInMinMax(numberName)
    dim i

    i = cInt( getNumberIndex(numberName) )

    pushNumberInMinMaxByIndex i
end sub

sub pushNumberInMinMaxByIndex(i)
    dim min
    dim max
    
    min = getArrNumberMin(i)
    max = getArrNumberMax(i)

    arrNumber(i) = numberIntoMinMax(arrNumber(i), min, max)
end sub

function getNumberOfName(parNumberName)
    dim i
    dim numberValue
    dim foundName
    dim numberName

    numberValue = 0
    numberName = lcase(parNumberName)
    numberName = replaceAllValues(numberName)

    for i = lbound(arrNumberName) to ubound(arrNumberName)
        if lcase(arrNumberName(i)) = numberName then
            numberValue = arrNumber(i)
            exit for
        end if
    next

    getNumberOfName = cInt(numberValue)
end function

function getStringOfName(parStringName)
    dim i
    dim stringValue
    dim foundName
    dim stringName

    stringValue = ""
    stringName = lcase(parStringName)

    for i = lbound(arrStringName) to ubound(arrStringName)
        if lcase(arrStringName(i)) = stringName then
            stringValue = arrString(i)
            exit for
        end if
    next

    getStringOfName = cStr(stringValue)
end function

function replaceAllValues(parText)
    dim text

    text = parText
    text = replaceStringValues(text)
    text = replaceNumberValues(text)
    text = replaceStateValues(text)

    replaceAllValues = cStr(text)
end function

function replaceNumberValues(text)
    dim oldText
    dim newText

    newText = text

    if instr(newText, "[") >= 1 then
        do
            oldText = newText
            newText = doReplaceNumberValues(newText)
            if oldText = newText then exit do
        loop
    end if

    replaceNumberValues = newText
end function

function replaceStringValues(text)
    dim oldText
    dim newText

    newText = text

    if instr(newText, "[") >= 1 then
        do
            oldText = newText
            newText = doReplaceStringValues(newText)
            if oldText = newText then exit do
        loop
    end if

    replaceStringValues = newText
end function

function replaceStateValues(text)
    dim oldText
    dim newText
    newText = text

    if instr(newText, "[") >= 1 then
        do
            oldText = newText
            newText = doReplaceStateValues(newText)
            if oldText = newText then exit do
        loop
    end if
    newText = newText

    replaceStateValues = newText
end function

function doReplaceNumberValues(text)
    const startString = "[number "
    const formatString = "format "
    const verboseString = "verbose"
    dim startsAt
    dim endsAt
    dim lengthOf
    dim newText
    dim numberName
    dim numberFormat
    dim numberOfName
    dim numberVerbose
    
    newText = text
    startsAt = instr(newText, "[number ")
    numberFormat = ""
    numberOfName = ""
    numberVerbose = false

    if startsAt >= 1 then
        lengthOf = instr( mid(newText, startsAt + len(startString)), "]" )

        if lengthOf >= 1 then
            numberName = mid(newText, startsAt + len(startString), lengthOf - 1)
            if instr(numberName, formatString) >= 1 then
                numberFormat = mid( numberName, instr(numberName, formatString) + len(formatString) )
                numberName = left( numberName, instr(numberName, formatString) - 2 )
            elseif instr(numberName, verboseString) >= 1 then
                numberVerbose = true
                numberName = left( numberName, instr(numberName, " ") - 1)
            end if

            numberOfName = getNumberOfName(numberName)
            numberOfName = formatNumber(numberOfName, numberFormat)
            if numberVerbose then
                numberOfName = verboseNumber(numberOfName)
            end if
            newText = left(newText, startsAt - 1) & _
                    numberOfName & _
                    mid(newText, startsAt + len(startString) + lengthOf)

        end if

    end if

    doReplaceNumberValues = newText
end function

function verboseNumber(oldNumber)
    dim newNumber
    
    newNumber = oldNumber
    select case cInt(newNumber)
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

function formatNumber(oldNumber, numberFormat)
    dim newNumber

    newNumber = oldNumber
    if numberFormat <> "" then
        if len(newNumber) < len(numberFormat) then
            newNumber = newNumber & mid( numberFormat, len(newNumber) + 1 )
        end if
    end if

    formatNumber = cStr(newNumber)
end function

function doReplaceStringValues(text)
    const startString = "[string "
    dim startsAt
    dim endsAt
    dim lengthOf
    dim newText
    dim stringName
    dim conversion
    dim splitted

    newText = text
    startsAt = instr(newText, "[string ")

    if startsAt >= 1 then
        lengthOf = instr( mid(newText, startsAt + len(startString)), "]" )
        if lengthOf >= 1 then
            stringName = mid(newText, startsAt + len(startString), lengthOf - 1)

            if instr(stringName, " ") >= 1 then
                splitted = split(stringName, " ")
                conversion = lcase(splitted(0))
                stringName = splitted(1)
            end if

            newText = left(newText, startsAt - 1) & _
                    convertString(getStringOfName(stringName), conversion) & _
                    mid(newText, startsAt + len(startString) + lengthOf)

        end if
    end if

    doReplaceStringValues = newText
end function

function doReplaceStateValues(text)
    const startString = "[state "
    const seperator = ", "
    dim startsAt
    dim endsAt
    dim lengthOf
    dim newText
    dim stateStart
    dim i
    dim state
    dim lenStateStart
    dim newString

    newText = text
    startsAt = instr(newText, "[state ")

    newString = ""
    if startsAt >= 1 then

        lengthOf = instr( mid(newText, startsAt + len(startString)), "]" )

        if lengthOf >= 1 then
            stateStart = mid(newText, startsAt + len(startString), lengthOf - 1)
            lenStateStart = len(stateStart)

            for i = lbound(arrState) to ubound(arrState)
                state = lcase(arrState(i))
                if state <> "" then

                    if lcase( left(state, lenStateStart) ) = stateStart then
                        newString = newString & _
                                mid(arrState(i), lenStateStart + 1) & seperator
                    end if

                end if
            next

            if newString <> "" then
                newString = left( newString, len(newString) - len(seperator) )
            end if

            newText = left(newText, startsAt - 1) & _
                    newString & _
                    mid(newText, startsAt + lenStateStart + lengthOf + 3)

        end if

    end if

    doReplaceStateValues = newText
end function

function convertString(oldString, conversion)
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

function checkNumber(numberString)
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

    numberString = replaceNumberValues(numberString)

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
            isOk = (cInt(numberRealValue) = cInt(numberCheckValue))
        case "is below"
            isOk = (cInt(numberRealValue) < cInt(numberCheckValue))
        case "is equalbelow"
            isOk = (cInt(numberRealValue) <= cInt(numberCheckValue))
        case "is above"
            isOk = (cInt(numberRealValue) > cInt(numberCheckValue))
        case "is equalabove"
            isOk = (cInt(numberRealValue) >= cInt(numberCheckValue))
        case "is not"
            isOk = (cInt(numberRealValue) <> cInt(numberCheckValue))
        case else
            sendError "Erronous operator """ & numberString & """"
    end select

    checkNumber = isOk
end function

function checkString(stringString)
    dim splitted
    dim stringName
    dim operator
    dim stringCheckValue
    dim stringRealValue
    dim lastSpace

    stringCheckValue = 0
    stringRealValue = 0

    stringString = rtrim(stringString)
    stringString = replace(stringString, "[", "'[")
    stringString = replace(stringString, "]", "]'")

    stringString = replaceStringValues(stringString)
    stringString = replaceNumberValues(stringString)

    if instr(stringString, "'") >= 1 then
        stringCheckValue = mid(stringString, instr(stringString, "'") + 1)
        stringCheckValue = left(stringCheckValue, len(stringCheckValue) - 1)
        stringString = left(stringString, instr(stringString, "'") - 1)
        stringString = rtrim(stringString)
    else
        lastSpace = instrrev(stringString, " ")
        stringCheckValue = mid(stringString, lastSpace + 1)
        stringString = left(stringString, lastSpace - 1)
    end if

    stringString = trimDoubleSpaces(stringString)
  
    splitted = Split(stringString, " ")

    if ubound(splitted) = 1 then
        stringName = splitted(0)
        operator = splitted(1)
    else
        stringName = splitted(0)
        operator = splitted(1) & " " & splitted(2)
    end if

    stringRealValue = getStringOfName(stringName)

    checkString = getStringIsOk(operator, stringRealValue, stringCheckValue)
end function

function getStringIsOk(operator, stringRealValue, stringCheckValue)
    dim isOk
    isOk = false

    select case lcase(operator)
        case "is"
            isOk = (cStr(stringRealValue) = cStr(stringCheckValue))
        case "islike"
            isOk = ( lcase(cStr(stringRealValue)) = lcase(cStr(stringCheckValue)) )
        case "islike not", "not islike", "isunlike"
            isOk = ( lcase(cStr(stringRealValue)) <> lcase(cStr(stringCheckValue)) )
        case "is not"
            isOk = ( cStr(stringRealValue) <> cStr(stringCheckValue) )
        case "contains"
            isOk = cBool( instr(cStr(stringRealValue), cStr(stringCheckValue)) >= 1 )
        case "containslike"
            isOk = cBool( instr(lcase(cStr(stringRealValue)), lcase(cStr(stringCheckValue))) >= 1 )
        case "contains not", "not contains"
            isOk = not cBool( instr(cStr(stringRealValue), cStr(stringCheckValue)) >= 1 )
        case "containslike not", "not containslike"
            isOk = not cBool( instr(lcase(cStr(stringRealValue)), lcase(cStr(stringCheckValue))) >= 1 )
        case "containsword"
            isOk = stringContainsWord( cStr(stringRealValue), cStr(stringCheckValue) )
        case "containsword not", "not containsWord"
            isOk = not stringContainsWord( cStr(stringRealValue), cStr(stringCheckValue) )
        case "length"
            isOk = cBool( len(cStr(stringRealValue)) = cInt(stringCheckValue) )
        case "length not", "not length"
            isOk = cBool( len(cStr(stringRealValue)) = cInt(stringCheckValue) )
        case "length above", "above", "above length"
            isOk = cBool( len(cStr(stringRealValue)) > cInt(stringCheckValue) )
        case "length below", "below", "below length"
            isOk = cBool( len(cStr(stringRealValue)) < cInt(stringCheckValue) )
        case else
            sendError language("Erronous operator in string check", _
                    "Fehlerhafter Operator bei String-Test")
    end select

    getStringIsOk = cBool(isOk)
end function

function stringContainsWord(sentence, parCheckWord)
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

    stringContainsWord = cBool(foundWord)
end function

sub randomizeNumber(name, numberValue)
    dim oldValue, newValue

    oldValue = getNumberOfName(name)
    numberValue = trimDoubleSpaces(ltrim(rtrim(numberValue)))
    numberValue = replaceNumberValues(numberValue)

    if instr(numberValue, " to ") then
        newValue = getRandomNumberOfString(numberValue)
    else
        numberValue = replace(numberValue, " ", "")
        newValue = cInt(oldValue) + getRandomNumber(0, cInt(numberValue))
    end if

    setNumberOfName name, newValue
end sub

function getRandomNumberOfString(numberString)
    ' handles strings like: "1 to 6"
    dim randomNumber, splitted
    splitted = split(numberString)
    randomNumber = getRandomNumber(splitted(0), splitted(2))
    getRandomNumberOfString = cInt(randomNumber)
end function

function getRandomNumber(parMin, parMax)
    dim min
    dim max

    if parMin < parMax then
        min = cInt(parMin)
        max = cInt(parMax)
    else
        min = cInt(parMax)
        max = cInt(parMin)
    end if
        
    getRandomNumber = cInt( (rnd * (max - min)) + min )
end function

sub toggleDisplayStates
    gAlwaysDisplayInfo = not gAlwaysDisplayInfo
    if gAlwaysDisplayInfo then
        displayStates gLastStation
    else
        hideStates
    end if
end sub

sub displayStates(stationId)
    if gDebug then
        if not gDebugInfoIsDisplayed then
            objPage.all.stateDisplay.style.display = "block"
            doDisplayStates stationId
            gDebugInfoIsDisplayed = true
        end if
    end if
end sub

sub hideStates
    if gDebug then
        if gDebugInfoIsDisplayed then
            objPage.all.stateDisplay.style.display = "none"
            gDebugInfoIsDisplayed = false
        end if
    end if
end sub

sub addVisits(stationId)
    const endString = ")"
    const allSum = "*"
    const addOne = "+1"
    setNumber visitsStartString & stationId & endString, addOne
    setNumber visitsStartString & allSum & endString, addOne
end sub

sub doDisplayStates(stationId)
    dim i, strng

    strng = strng & "<h1 class=""stationDisplay"" " & _
            "onMouseDown=""grabEl(this.offsetParent)"">Station """ & stationId & """</h1>"
    strng = strng & "<div class=""statesDisplay"">"
    strng = strng & "<h2>States</h2>" & _
            "<p class=""listDescript"">Listed are the true states:</p>"
    strng = strng & "<ul>"

    for i = lbound(arrState) to ubound(arrState)
        if not arrState(i) = "" then
            strng = strng & "<li>" & arrState(i) & "</li>"
        end if
    next

    strng = strng & "</ul>"
    strng = strng & "</div>"

    strng = strng & "<div class=""numbersDisplay"">"
    strng = strng & "<h2>Numbers</h2>" & _
            "<p class=""listDescript"">Listed are the numbers which are not 0 or have other than default limits:</p>"
    strng = strng & "</ul>"
    for i = lbound(arrNumber) to ubound(arrNumber)
        if arrNumberName(i) <> "" and left(arrNumberName(i), len("qml")) <> ("qml") and _
                lcase(left(arrNumberName(i), len(visitsStartString))) <> visitsStartString then

            strng = strng & "<li>" & arrNumberName(i) & " is " & arrNumber(i)
            if getArrNumberMin(i) <> numberDefaultMin and getArrNumberMax(i) <> numberDefaultMax then
                strng = strng & " (" & getArrNumberMin(i) & " to " & getArrNumberMax(i) &  ")"
            end if
            strng = strng & "</li>"
        end if
    next
    strng = strng & "</ul>"
    strng = strng & "</div>"

    strng = strng & "<div class=""stringsDisplay"">"
    strng = strng & "<h2>Strings</h2>" & _
            "<p class=""listDescript"">Listed are the strings which are not empty:</p>"
    strng = strng & "</ul>"
    for i = lbound(arrString) to ubound(arrString)
        if arrStringName(i) <> "" and left(arrStringName(i), len("qml")) <> ("qml") then
            strng = strng & "<li>" & arrStringName(i) & " is " & arrString(i) & "</li>"
        end if
    next
    strng = strng & "</ul>"
    strng = strng & "</div>"

    strng = strng & "<div class=""closeDisplay"" onclick=""toggleDisplayStates()"">Close</div>"

    objPage.all.stateDisplay.innerHTML = strng
end sub

sub setArrNumberMin(i, value)
    if cInt(value) = numberDefaultMin then
        if arrNumberMinSet(i) <> false then
            arrNumberMinSet(i) = false
        end if
        if arrNumberMin(i) <> 0 then
            arrNumberMin(i) = 0
        end if
    else
        arrNumberMinSet(i) = true
        arrNumberMin(i) = cInt(value)
    end if
end sub

function getArrNumberMin(i)
    if arrNumberMinSet(i) then
        getArrNumberMin = cInt(arrNumberMin(i))
    else
        getArrNumberMin = cInt(numberDefaultMin)
    end if
end function

sub setArrNumberMax(i, value)
    if cInt(value) = numberDefaultMax then
        if arrNumberMaxSet(i) <> false then
            arrNumberMaxSet(i) = false
        end if
        if arrNumberMax(i) <> 0 then
            arrNumberMax(i) = 0
        end if
    else
        arrNumberMaxSet(i) = true
        arrNumberMax(i) = cInt(value)
    end if
end sub

function getArrNumberMax(i)
    if arrNumberMaxSet(i) then
        getArrNumberMax = cInt(arrNumberMax(i))
    else
        getArrNumberMax = cInt(numberDefaultMax)
    end if
end function

sub setQmlStartVariables
    setString "qmlSecondsStart", timer
    setString "qmlVersion", qmlVersionNumber
    if serverVersion then
        setString "qmlServer", "true"
    else
        setString "qmlServer", "false"
    end if
end sub

sub setQmlVariables
    dim seconds

    setString "qmlLastStation", gLastStation

    seconds = cInt(timer - cLng(getStringOfName("qmlSecondsStart")))
    if seconds > 30000 then seconds = 0
    setNumber "qmlSeconds", seconds
    setNumber "qmlMinutes", cInt(seconds / 60)

    setString "qmlTime", time
    setString "qmlDay", verboseWeekday(date)
end sub