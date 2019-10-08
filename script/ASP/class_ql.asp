<%

' QL-Converter class v0.5

class Ql
    dim m_source ' as string
    dim m_spacing

    public sub setSource(byVal source)
        m_source = source
    end sub

    public function getQml
        const sComment = "//"
        dim qml ' as string
        dim allLine ' ()
        dim lines(10000) ' ()
        dim i ' as long
        dim indent(10000) ' as integer
        dim thisIndent ' as integer
        dim keyword ' as string
        dim inStation ' as boolean
        dim sValue ' as string
        dim sName ' as string
        dim sState ' as string
        dim splitted ' ()
        dim oldIndent ' as long
        dim oldThisIndent ' as long
        dim i2 ' as long
        dim sText ' as string
        dim sHead ' as string
        dim maxLines ' as long
        dim lineStart ' as string
        dim thisLine ' as string
        dim lenComment ' as integer
        dim iLevel ' as integer
        dim inIf ' as boolean
        dim inElse ' as boolean
        dim keyword2 ' as string
        dim sStates ' as string
        dim ifI ' as long
        dim sClass ' as string
        dim sCheck ' as string

        m_spacing = 4
        qml = ""

        m_source = replace(m_source, space(m_spacing), " ")
        m_source = replace(m_source, vbNewline & vbNewline, vbNewline)
        m_source = replace(m_source, " <--", " --> back")
        m_source = replace(m_source, "&", "&amp;")

        allLine = split(m_source, vbNewline)
        lenComment = len(sComment)
        i2 = 0
        maxLines = 0
        for i = 0 to uBound(allLine)
            thisLine = trim( allLine(i) )
            if len(thisLine) > lenComment then
                if mid(thisLine, 1, lenComment) = sComment then

                elseif mid(thisLine, 1, 1) = """" then
                    splitted = split(thisLine, """,")
                    qml = qml & "<about>" & vbNewline
                    qml = qml & "    <title>" & mid( splitted(0), 2, len( splitted(0) ) - 1 ) & _
                            "</title>" & vbNewline
                    qml = qml & "    <author>" & trim( splitted(1) ) & "</author>" & vbNewline
                    qml = qml & "</about>" & vbNewline & vbNewline
                else
                    lines(i2) = allLine(i)
                    maxLines = i2
                    i2 = i2 + 1
                end if
            end if
        next

        for i = 0 to maxLines
            indent(i) = getIndent( lines(i) )
            lines(i) = lTrim( lines(i) )
        next

        inIf = false
        inStation = false
        iLevel = 0
        thisIndent = 0
        for i = 0 to maxLines
            oldThisIndent = thisIndent
            thisIndent = indent(i)

            if oldThisIndent > thisIndent then
                if inIf then
                    iLevel = iLevel - 1
                    qml = qml & space(iLevel * m_spacing) & _
                            "</if>" & vbNewline
                    inIf = false
                end if
                if inElse then
                    iLevel = iLevel - 1
                    qml = qml & space(iLevel * m_spacing) & _
                            "</else>" & vbNewline
                    inElse = false
                end if
            end if

            if thisIndent = 0 then
                if inStation then
                    qml = qml & "</station>" & vbNewline & vbNewline
                    qml = qml & "<station id=""" & lines(i) & """>" & vbNewline
                else
                    qml = qml & "<station id=""" & lines(i) & """>" & vbNewline
                    inStation = true
                    iLevel = iLevel + 1
                end if

            else

                keyword = getFirstWord( lines(i) )
                select case keyword
                    case "?"
                        inIf = true
                        sValue = mid( lines(i), len(keyword) + 2 )
                        sValue = replaceCheckValue(sValue)
                        qml = qml & space(iLevel * m_spacing) & _
                                "<if check=""" & sValue & """>" & vbNewline
                        iLevel = iLevel + 1

                    case "..."
                        inElse = true
                        qml = qml & space(iLevel * m_spacing) & _
                                "<else>" & vbNewline
                        iLevel = iLevel + 1

                    case "_", "%", "$"
                        qml = qml & getState( iLevel, keyword, lines(i) )

                    case "---"
                        oldIndent = thisIndent
                        sText = ""
                        sValue = mid( lines(i), len(keyword) + 2 )
                        sClass = ""
                        sCheck = ""

                        ifI = inStr(sValue, "? ")
                        if ifI > 0 then
                            sClass = mid( sValue, 1, ifI - 1 )
                            sClass = rTrim(sClass)

                            sCheck = mid(sValue, ifI + 1)
                            sCheck = trim( replaceCheckValue(sCheck) )
                            if sCheck <> "" then
                                sCheck = " check=""" & sCheck & """ "
                            end if
                        else
                            sClass = sValue
                        end if
                        if sClass <> "" then
                            sClass = " class=""" & sClass & """ "
                        end if

                        iLevel = iLevel + 1
                        for i2 = i + 1 to maxLines
                            if indent(i2) > oldIndent then
                                sText = sText & space(iLevel * m_spacing) & trim( lines(i2) ) & _
                                        vbNewline
                            else
                                i = i2 - 1
                                exit for
                            end if
                        next
                        iLevel = iLevel - 1

                        sValue = removeDoubleSpaces(sCheck & sClass)
                        qml = qml & space(iLevel * m_spacing) & "<text" & rTrim(sValue) & _
                                ">" & vbNewline & replaceTextValue(sText) & space(iLevel * m_spacing) & "</text>" & vbNewline

                    case "-//"
                        sValue = mid( lines(i), len(keyword) + 2)
                        sValue = replaceTextValue(sValue)
                        qml = qml & "    <comment>" & sValue & "</comment>" & vbNewline

                    case "+++"
                        sValue = ""
                        oldIndent = thisIndent
                        iLevel = iLevel + 1
                        for i2 = i + 1 to maxLines
                            if indent(i2) > oldIndent then
                                sValue = sValue & space(iLevel * m_spacing) & _
                                        "<in station=""" & lines(i2) & """ />" & vbNewline
                            else
                                i = i2 - 1
                                exit for
                            end if
                        next
                        iLevel = iLevel - 1

                        qml = qml & space(iLevel * m_spacing) & "<include>" & vbNewline & _
                                sValue & space(iLevel * m_spacing) & "</include>" & vbNewline

                    case "("
                        sValue = mid( lines(i), len(keyword) + 2)
                        sValue = trim(sValue)
                        sValue = mid( sValue, 1, len(sValue) - 2 )
                        select case getExtension(sValue)
                            case "gif", "jpg", "jpeg", "png"
                                keyword = "image"
                                sValue = "media/" & sValue
                            case "mid", "wav", "wave", "mp3", "ram"
                                keyword = "music"
                                sValue = "media/" & sValue
                            case else
                                keyword = "embed"
                        end select
                        if keyword <> "" then
                            sValue = replaceTextValue(sValue)
                            qml = qml & space(iLevel * m_spacing) & _
                                    "<" & keyword & " source=""" & sValue & """/>" & vbNewline
                        end if

                    case "-->"
                        oldIndent = thisIndent
                        sText = ""
                        sState = ""
                        splitted = split( lines(i), " ? ")
                        if uBound(splitted) > 0 then
                            sValue = splitted(0)
                            sState = splitted(1)
                            sValue = mid( sValue, len(keyword) + 2 )
                            sState = replaceCheckValue(sState)
                            sState = " check=""" & sState & """"
                        else
                            sValue = mid( lines(i), len(keyword) + 2 )
                        end if

                        qml = qml & space(iLevel * m_spacing) & _
                            "<choice station=""" & sValue & """" & sState &">" & vbNewline

                        sStates = ""
                        iLevel = iLevel + 1
                        for i2 = i + 1 to maxLines
                            if indent(i2) > oldIndent then
                                keyword2 = getFirstWord( lines(i2) )
                                select case keyword2
                                    case "_", "%", "$"
                                        sStates = sStates & _
                                                getState( iLevel, keyword2, lines(i2) )
                                    case else
                                        sText = sText & space(iLevel * m_spacing) & _
                                                trim( lines(i2) ) & vbNewline
                                end select
                            else
                                i = i2 - 1
                                exit for
                            end if
                        next

                        qml = qml & replaceTextValue(sText) & sStates

                        iLevel = iLevel - 1
                        qml = qml & space(iLevel * m_spacing) & "</choice>" & vbNewline

                end select

            end if
        next

        qml = qml & "</station>" & vbNewline

        qml = replace(qml, vbNewline & "<station id="""">" & vbNewline & "</station>" & vbNewline, "")

        sHead = ""
        sHead = sHead & "<?xml version=""1.0"" encoding=""iso-8859-1"" ?>" & vbNewline
        sHead = sHead & "<!DOCTYPE quest SYSTEM ""../script/quest.dtd"">" & vbNewline

        qml = sHead & "<quest>" & vbNewline & vbNewline & qml & vbNewline  & "</quest>"

        getQml = qml
    end function

    private function getState(byVal iLevel, byVal keyword, lineI)
        dim sNameValue ' as string
        dim sValue ' as string
        dim sName ' as string
        dim firstI ' as long
        dim sState ' as string

        sNameValue = mid( lineI, len(keyword) + 2 )
        sNameValue = lTrim(sNameValue)
        firstI = inStr(sNameValue, "=")

        if firstI > 0 then
            sName = mid(sNameValue, 1, firstI - 1)
            sName = trim(sName)
            sValue = mid(sNameValue, firstI + 1)
            sValue = trim(sValue)
        elseif keyword = "_" then
            sName = sNameValue
            sValue = "true"
        end if

        sValue = replaceCheckValue(sValue)
        select case keyword
            case "_": keyword = "state"
            case "%": keyword = "number"
            case "$": keyword = "string"
        end select

        sState = space(iLevel * m_spacing) & _
                "<" & keyword & " name=""" & sName & """ " & _
                "value=""" & sValue & """/>" & vbNewline

        getState = sState
    end function

    private function replaceCheckValue(byVal sValue)
        sValue = replace(sValue, "!", " not ")
        sValue = replace(sValue, ">", " greater ")
        sValue = replace(sValue, "<", " lower ")
        sValue = removeDoubleSpaces(sValue)
        replaceCheckValue = sValue
    end function

    private function removeDoubleSpaces(byVal sValue)
        dim oldValue

        oldValue = ""
        while oldValue <> sValue
            oldValue = sValue
            sValue = replace(sValue, "  ", " ")
        wend

        removeDoubleSpaces = sValue
    end function

    private function replaceTextValue(byVal sValue)
        sValue = replace(sValue, "<", "&lt;")
        sValue = replace(sValue, ">", "&gt;")
        sValue = replace(sValue, " **", " <strong>")
        sValue = replace(sValue, "** ", "</strong> ")
        sValue = replace(sValue, "**" & vbNewline, "</strong>" & vbNewline)
        sValue = replace(sValue, " *", " <emphasis>")
        sValue = replace(sValue, "* ", "</emphasis> ")
        sValue = replace(sValue, "*" & vbNewline, "</emphasis>" & vbNewline)

        sValue = replace(sValue, "||", "<break/>")
        replaceTextValue = sValue
    end function

    private function getFirstWord(byVal line)
        dim words
        dim firstWord

        firstWord = ""
        words = split(line, " ")
        firstWord = words(0)

        getFirstWord = firstWord
    end function

    private function getIndent(byVal line)
        dim i ' as long
        dim iIndention ' as long

        iIndention = 0
        for i = 1 to len(line)
            if mid(line, i, 1) = " " then
                iIndention = iIndention + 1
            else
                exit for
            end if
        next

        getIndent = iIndention
    end function

    private function getExtension(byVal fileName)
        dim lastI
        dim extension

        extension = ""
        lastI = inStrRev(fileName, ".")
        if lastI > 0 then
            extension = mid(fileName, lastI + 1)
            extension = lCase(extension)
        end if

        getExtension = extension
    end function

end class

%>