class classInlineFunction

    private m_inlineString
    private m_xhtml
    private m_xmlStates

    public sub setInlineString(byVal inlineString)
        m_inlineString = inlineString
    end sub

    public sub setXmlStates(byRef xmlStates)
        set m_xmlStates = xmlStates.cloneNode(true)
    end sub

    public sub process
        dim firstSpace
        dim parameters
        dim functionName
        dim sEvaluate

        m_xhtml = ""

        m_inlineString = replace(m_inlineString, "'", """")
        firstSpace = inStr(m_inlineString, " ")
        if firstSpace > 0 then
            functionName = left(m_inlineString, firstSpace - 1)
            parameters = mid(m_inlineString, firstSpace + 1)
        else
            functionName = m_inlineString
            parameters = ""
        end if

        sEvaluate = "qml_" & functionName & "(" & parameters & ")"
        ' response.write sEvaluate
        m_xhtml = eVal(sEvaluate)
    end sub

    public function getXhtml()
        getXhtml = m_xhtml
    end function

    ' private ________________

    private function qml_random(byVal min, byVal max)
        dim lNumber

        randomize
        lNumber = cLng( rnd * (max - min) ) + min

        qml_random = lNumber
    end function

    private function qml_states(byVal sValue)
        const seperator = ", "
        dim xhtml
        dim subNode
        dim subNodes
        dim xPath
        dim thisName
        dim thisEnd

        xhtml = ""
        xPath = "//state"
        set subNodes = m_xmlStates.selectNodes(xPath)
        for each subNode in subNodes
            thisName = subNode.getAttribute("name")
            if len(thisName) >= len(sValue) then
                if left( thisName, len(sValue) ) = sValue then
                    thisEnd = mid( thisName, len(sValue) + 1 )
                    xhtml = xhtml & thisEnd & seperator
                end if
            end if
        next
        if xhtml <> "" then
            xhtml = left( xhtml, len(xhtml) - len(seperator) )
        end if

        qml_states = xhtml
    end function

    private function qml_contains(byVal stringAll, byVal stringToCheck)
        dim isTrue

        isTrue = inStr(stringAll, stringToCheck) >= 1
        qml_contains = returnIf(isTrue, "true", "false")
    end function

    private function qml_containsWord(byVal sentence, byVal checkWord)
        dim splitted
        dim i
        dim foundWord

        checkWord = lcase(checkWord)
        splitted = splitWords( lcase(sentence) )

        foundWord = false
        for i = lbound(splitted) to ubound(splitted)
            if splitted(i) = checkWord then
                foundWord = true
                exit for
            end if
        next
    
        qml_containsWord = returnIf(foundWord, "true", "false")
    end function

    private function qml_verbose(byVal thisNumber)
        select case cLng(thisNumber)
            case 11, 12, 13
                thisNumber = thisNumber & "th"
            case else

                select case right(thisNumber, 1)
                    case 1
                        thisNumber = thisNumber & "st"
                    case 2
                        thisNumber = thisNumber & "nd"
                    case 3
                        thisNumber = thisNumber & "rd"
                    case else
                        thisNumber = thisNumber & "th"
                end select
    
        end select
    
        qml_verbose = cStr(thisNumber)
    end function

    private function qml_lower(byVal text)
        qml_lower = lCase(text)
    end function

    private function qml_upper(byVal text)
        qml_upper = uCase(text)
    end function

    private function qml_repeatString(byVal text, byVal n)
        dim i
        dim newText

        newText = ""
        for i = 1 to n
            newText = newText & text
        next

        qml_repeatString = newText
    end function

end class