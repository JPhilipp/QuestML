option explicit

function getNodeType(byRef nodeTypeNumber)
    dim nodeTypeText

    select case nodeTypeNumber
        case 1: nodeTypeText = "element"
        case 2: nodeTypeText = "attribute"
        case 3: nodeTypeText = "text"
        case 4: nodeTypeText = "cdataSection"
        case 5: nodeTypeText = "entityReference"
        case 6: nodeTypeText = "entity"
        case 7: nodeTypeText = "processingInstructions"
        case 8: nodeTypeText = "comment"
        case 9: nodeTypeText = "document"
        case 10: nodeTypeText = "documentType"
        case 11: nodeTypeText = "documentFragment"
        case 12: nodeTypeText = "notation"
    end select
    
    getNodeType = nodeTypeText
end function

sub showErrorOf(byRef obj)
    dim strError

    strError = "Invalid XML document!" & vbNewline & _
               "File: " & obj.parseError.url & vbNewline & _
               "Line: " & obj.parseError.line & vbNewline & _
               " --- Character: " & obj.parseError.linepos & vbNewline & _
               "Source Text: " & obj.parseError.srcText & vbNewline & _
               "Description: " & obj.parseError.reason
    sendError strError
end sub

function returnIf(byVal state, byVal ifTrue, byVal ifFalse)
    dim returnValue

    if state then
        returnValue = ifTrue
    else
        returnValue = ifFalse
    end if
    returnIf = returnValue
end function

function trimDoubleSpaces(byVal strng)
    dim oldString
    dim newString

    newString = strng
    do
        oldString = newString
        newString = replace(newString, "  ", " ")
    loop until oldString = newString

    trimDoubleSpaces = newString
end function

function repeatedReplace(byVal parText, byVal toFind, byVal toReplace)
    dim text
    dim oldText

    text = parText

    do
        oldText = text
        text = replace(text, toFind, toReplace)
    loop until text = oldText

    repeatedReplace = text
end function

function numberIntoMinMax(byVal oldNumber, byVal min, byVal max)
    dim newNumber

    newNumber = oldNumber
    if newNumber < min then
        newNumber = min
    elseif newNumber > max then
        newNumber = max
    end if

    numberIntoMinMax = newNumber
end function

function properCase(byVal text)
    dim newText
    dim splitted
    dim i
    dim thisWord
    dim singleWord

    splitted = split(text, " ")
    newText = ""

    for i = lbound(splitted) to ubound(splitted)
        thisWord = splitted(i)
        if len(thisWord) >= 2 then
            singleWord = ucase( left(thisWord, 1) ) & mid(thisWord, 2)
        else
            singleWord = thisWord
        end if
        newText = newText & singleWord & " "
    next

    properCase = rtrim(newText)
end function

function splitWords(byVal inputText)
    const chars = ".!?,;:""'()[]{}"
    dim strReplacedText
    dim i

    strReplacedText = inputText

    For i = 1 To Len(chars)
        strReplacedText = Trim(Replace(strReplacedText, _
            Mid(chars, i, 1), " "))
    Next

    Do While InStr(strReplacedText, "  ")
        strReplacedText = Replace(strReplacedText, "  ", " ")
    Loop

    splitWords = split(strReplacedText, " ")
end function

function getXmlString(byVal xmlString)
    dim xmlDoc
    dim isValid

    if g_isServerVersion then
        set xmlDoc = server.createObject("Microsoft.XMLDOM")
    else
        set xmlDoc = createObject("Microsoft.XMLDOM")
    end if

    xmlDoc.async = false
    xmlDoc.loadXML xmlString
    isValid = cBool(xmlDoc.parseError.errorCode = 0)
    if not isValid then
        showErrorOf xmlDoc
    end if

    set getXmlString = xmlDoc
end function

Function getInnerXml(byRef objXml)
    Dim child
    Dim text
    
    text = ""
    For Each child In objXml.childNodes
        text = text & child.xml
    Next

    getInnerXml = text
End Function

function getWochentag(byRef datum)
    getWochentag = getWochentagOfIndex(weekday(datum))
end function

function compareStrings(byVal oldStringCheck, byVal stringOriginal)
    ' return true if first parameter is
    ' "hello world", "*lo world", "hello wo*", or "*lo wo*"
    ' and second is "hello world"

    const wildcard = "*"
    dim check
    dim wildcardLeft
    dim wildcardRight
    dim stringCheck
    dim areSame
    stringCheck = oldStringCheck

    wildcardLeft = cBool(left(stringCheck, len(wildcard)) = wildcard)
    wildcardRight = cBool(right(stringCheck, len(wildcard)) = wildcard)

    if stringCheck = wildcard then
        areSame = true
    elseif wildcardLeft or wildcardRight then
        stringCheck = replace(stringCheck, wildcard, "")
        set check = new RegExp
        check.ignoreCase = true

        if wildcardLeft and wildcardRight then
            check.pattern = "\B" & stringCheck
        elseif wildcardLeft then
            check.pattern = stringCheck & "$"
        elseif wildcardRight then
            check.pattern = "^" & stringCheck
        end if
        areSame = check.test(stringOriginal)

    else
        areSame = lcase(stringCheck) = lcase(stringOriginal)
    end if

    compareStrings = cBool(areSame)
end function

function getXml(byVal xmlPath)
    dim xmlDoc
    dim isValid

    if g_isServerVersion then
        set xmlDoc = server.createObject("Microsoft.XMLDOM")
    else
        set xmlDoc = CreateObject("Microsoft.XMLDOM")
    end if
    xmlDoc.async = false

    if g_isServerVersion then
        xmlDoc.load server.mapPath(xmlPath)
    else
        xmlDoc.load xmlPath
    end if

    isValid = cBool(xmlDoc.parseError.errorCode = 0)
    if not isValid then
        showErrorOf xmlDoc
    end if

    set getXml = xmlDoc
end function

function toProperCase(byVal text)
    dim newText

    newText = cStr(text)
    newText = ucase( left(newText, 1) ) & lcase( mid(newText, 2) )

    toProperCase = cStr(newText)
end function

function xmlToText(byVal text)
    text = replace(text, "&", "&amp;")
    text = replace(text, """", "&quot;")
    text = replace(text, "<", "&lt;")
    text = replace(text, ">", "&gt;")

    xmlToText = text
end function

function textToXml(byVal text)
    text = replace(text, "&gt;", ">")
    text = replace(text, "&lt;", "<")
    text = replace(text, "&quot;", """")
    text = replace(text, "&amp;", "&")

    textToXml = text
end function


private sub sendMessage(byVal message)
    if g_isServerVersion then
        response.write "<p>" & message & "</p>"
    else
        msgBox message
    end if
end sub

private sub sendError(byVal message)
    if g_isServerVersion then
        response.write "<p class=""error"">" & message & "</p>"
    else
        msgBox message
    end if
end sub

function getTaggedValue(byVal thisName, byVal thisValue)
    getTaggedValue = getTaggedAttributedValue(thisName, "", "", thisValue)
end function

function getTaggedAttributedValue(byVal thisName, attributeName, attributeValue, byVal thisValue)
    dim sXml

    sXml = ""
    if not isEmpty(thisValue) then
        if varType(thisValue) = vbBoolean then
            thisValue = returnIf(thisValue, "true", "false")
        end if
        thisValue = textToXml(thisValue)
        sXml = sXml & "<" & thisName
        if attributeName <> "" then
            sXml = sXml & " " & attributeName & "=""" & attributeValue & """"
        end if
        sXml = sXml & ">" & thisValue & _
                "</" & thisName & ">" & vbNewline
    end if

    getTaggedAttributedValue = sXml
end function

function verboseBoolean(byVal state)
    verboseBoolean = cStr( returnIf(state, "true", "false") )
end function

function getIsoDateCompact(byRef ofDate)
    dim isoDate

    isoDate = getIsoDate(ofDate)
    isoDate = replace(isoDate, "-", "")
    isoDate = replace(isoDate, ":", "")
    isoDate = replace(isoDate, " ", "")

    getIsoDateCompact = isoDate
end function

function getIsoDate(byRef ofDate)
    dim isoDate

    isoDate = ""
    isoDate = isoDate & year(ofDate) & "-" & getPad( month(ofDate) ) & "-" & _
            getPad( day(ofDate) ) & " "
    isoDate = isoDate & getPad( hour(ofDate) ) & ":" & getPad( minute(ofDate) ) & _
            ":" & getPad( second(ofDate) )

    getIsoDate = isoDate
end function

function getPad(byVal num)
    if num < 10 then
        num = "0" & num
    end if

    getPad = num
end function