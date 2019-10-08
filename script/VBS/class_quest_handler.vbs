option explicit

class classQuestHandler

    private m_objPage
    private m_objQuest
    private m_oStateHandler
    private m_questName
    private m_contentType
    private m_debug
    private m_stationId
    private m_sessionId
    private m_statesString

    ' persistent via save/load:
    private m_lastStation
    private m_beforeLastStation
    private m_firstQuestName
    private m_defaultImage
    private m_defaultMusic
    private m_musicLoop
    private m_linkInlineStyle
    private m_language
    private m_gameOver

    public sub setStatesString(byVal sValue)
        m_statesString = sValue
    end sub

    public sub setSessionId(byVal sessionId)
        m_sessionId = sessionId
    end sub

    public sub setContentType(byVal contentType)
        m_contentType = contentType
    end sub

    public sub setQuestName(byVal questName)
        m_questName = questName
    end sub

    public sub setStationId(byVal stationId)
        m_stationId = stationId
    end sub

    public sub init
        dim pageTitle

        set m_oStateHandler = new classStateHandler

        randomize

        setQmlStartVariables
        setQmlVariables

        m_linkInlineStyle = ""

        m_firstQuestName = m_questName
        m_defaultImage = g_none
        m_defaultMusic = g_none
        m_musicLoop = "0"

        if m_contentType = "" then
            m_contentType = "text/html"
        end if

        loadXmlQuestFile mapPathIf(m_questName & ".xml")
        if m_objQuest.parseError.errorCode = 0 then
            m_debug = getDebug
            setObjPage

            setStyle
            m_language = getLanguage

            pageTitle = getPageTitle 
            setDocTitle pageTitle

            m_oStateHandler.setString "qmlTitle", pageTitle
        else
            showErrorOf m_objQuest
        end if

        if m_sessionId = "" then
            m_sessionId = getNewSessionId
        else
            loadQuest
        end if
    end sub

    public sub doHandleStation
        dim displayGotten
        dim inputString
        dim station

        inputString = ""
    
        set station = getStation(m_stationId)
        if (station is nothing) then
            exit sub
        end if

        handleTopChoose station, m_stationId

        m_oStateHandler.setString "qmlStation", m_stationId

        handleStationSettings station

        handleCheckStates station

        displayGotten = getDisplay(station, false)
        handleInclude m_stationId, displayGotten

        output displayGotten

        handleStatesInformation
   
        m_beforeLastStation = m_lastStation
        m_lastStation = m_stationId
        m_oStateHandler.setString "qmlLastStation", m_lastStation

        m_oStateHandler.addVisits m_stationId

        saveQuest

        if g_isServerVersion then
            handleServerOutput
        end if
    end sub

    public function getObjPage
        set getObjPage = m_objPage
    end function

    ' private __________________________________________________________

    private sub handleCheckStates(byRef station)
        dim child
        dim checkStatesAgain
        dim chooseElement

        do
            m_oStateHandler.handlePreStates station
            checkStatesAgain = false

            for each child in station.childNodes
                if child.nodeName = "if" then
                    if m_oStateHandler.getNodeState(child) then
                        set chooseElement = child.selectSingleNode("choose")
                        if not (chooseElement is nothing) then
                            processChoose child, station, chooseElement
                            checkStatesAgain = true
                        else
                            set station = child
                        end if
                        exit for
                    end if
                
                elseif child.nodeName = "else" then
                    set chooseElement = child.selectSingleNode("choose")
                    if not (chooseElement is nothing) then
                        processChoose child, station, chooseElement
                        checkStatesAgain = true
                    else
                        set station = child
                    end if
                end if
            next
        loop until not checkStatesAgain
    end sub

    private sub processChoose(byRef ifElseElement, byRef station, byRef chooseElement)
        dim sStation
        dim child

        for each child in ifElseElement.childNodes
            m_oStateHandler.setStates child
        next
        sStation = getLink(chooseElement)
        set station = getStation(sStation)
        m_oStateHandler.addVisits station.getAttribute("id")
    end sub
   
    private function handleStatesInformation
        dim statesInformation

        if m_debug then
            statesInformation = m_oStateHandler.getStatesInformation(m_stationId)
            if g_isServerVersion then
                serverOutputToId "stateDisplay", statesInformation
            else
                m_objPage.all.stateDisplay.innerHTML = statesInformation
            end if
        end if
    end function

    private sub handleServerOutput
        dim oServerResponse

        set oServerResponse = new classServerResponse
        oServerResponse.setContentType m_contentType
        oServerResponse.setSessionId m_sessionId
        oServerResponse.setQuestName m_questName
        oServerResponse.setObjPage m_objPage
        oServerResponse.process
    end sub

    private sub setObjPage
        if g_isServerVersion then
            setObjPageServer
        else
            set m_objPage = document
        end if
    end sub

    private sub setObjPageServer
        dim xHtml
        dim stateDisplay
        dim bodyNode
        dim xPath

        set m_objPage = createObject("Microsoft.XMLDOM")
        set xHtml = getXml("script\page.xml")
        if m_debug then
            set stateDisplay = xHtml.createElement("div")
            stateDisplay.setAttribute "id", "stateDisplay"
            xPath = "//body[@id = 'bodyNode']"
            set bodyNode = xHtml.selectSingleNode(xPath)
            bodyNode.appendChild stateDisplay
        end if

        m_objPage.load xHtml
    end sub
    
    private function mapPathIf(byVal filePath)
        dim newFilePath
    
        if g_isServerVersion then
            newFilePath = server.mapPath(filePath)
        else
            newFilePath = filePath
        end if
    
        mapPathIf = newFilePath
    end function

    private sub setDocTitle(byVal text)
        dim objTitle
    
        if g_isServerVersion then
            set objTitle = m_objPage.documentElement.selectSingleNode("//title")
            objTitle.text = text
        else
            m_objPage.title = text
        end if
    end sub
    
    private sub handleTopChoose(byRef station, byVal stationId)
        dim choose
        dim sStation
        dim child

        set choose = station.selectSingleNode("choose")
        if not (choose is nothing) then
            m_oStateHandler.addVisits stationId
            m_oStateHandler.handlePreStates station
            for each child in station.childNodes
                m_oStateHandler.setStates child
            next
            sStation = choose.getAttribute("station")
            stationId = getLink(choose)
            set station = getStation(stationId)
        end if
    end sub
    
    private sub handleInclude(byVal stationId, byRef oldDisplay)
        dim includeIn
        dim inNode
        dim includeNode
        dim doInclude
        dim includeState
        dim station
        dim newDisplay
    
        set includeIn = m_objQuest.documentElement.selectNodes("//in")
    
        for each inNode in includeIn
    
            if compareStrings(inNode.getAttribute("station"), stationId) then

                if m_oStateHandler.getNodeState(inNode) then
                    set includeNode = inNode.parentNode
                    if m_oStateHandler.getNodeState(includeNode) then
                        set station = includeNode.parentNode
                        handleCheckStates station
                        newDisplay = getDisplay(station, true)
    
                        if includeNode.getAttribute("process") = "after" then
                            oldDisplay = combineDisplay(oldDisplay, newDisplay)
                        elseif includeNode.getAttribute("process") = "before" then
                            oldDisplay = combineDisplay(newDisplay, oldDisplay)
                        else ' if includeNode.getAttribute("process") = "exclusive" then
                            oldDisplay = newDisplay
                        end if
        
                    end if
                end if
    
            end if
    
        next
    end sub
    
    private function combineDisplay(byRef station1, byRef station2)
        dim station
        dim parags
        dim parag
        dim lastParag
        dim listEntry
        dim listEntries
        dim list1
    
        set station = getXmlString("<top></top>")
    
        if instr(station1, "<top>") < 1 then
            set station1 = getXmlString("<top>" & station1 & "</top>")
        else
            set station1 = getXmlString(station1)
        end if
        if instr(station2, "<top>") < 1 then
            set station2 = getXmlString("<top>" & station2 & "</top>")
        else
           set station2 = getXmlString(station2)
        end if
    
        set parags = station2.documentElement.selectNodes("p")
        for each parag in parags
            set lastParag = station1.documentElement.selectSingleNode("./p[end()]")
            if lastParag is nothing then
                set lastParag = station1.documentElement.appendChild( m_objQuest.createElement("p") )
            end if
            if not (lastParag.nextSibling is nothing) then
                station1.documentElement.insertBefore parag.cloneNode(true), lastParag.nextSibling
            else
                station1.documentElement.appendChild parag.cloneNode(true)
            end if
        next
    
        set listEntries = station2.documentElement.selectNodes("//li")
        set list1 = station1.documentElement.selectSingleNode("ul")
        if list1 is nothing then
            set list1 = station1.documentElement.appendChild( station1.createElement("ul") )
        end if
        for each listEntry in listEntries
            list1.appendChild listEntry.cloneNode(true)
        next
        
        set station = station1    
    
        combineDisplay = getInnerXml(station)
    end function
    
    private sub handleStationSettings(byRef station)
        if station.getAttribute("states") = "reset" then
            m_oStateHandler.resetStates
            setQmlStartVariables
            setQmlVariables
        end if
    end sub
    
    private sub addStyle(byVal selector, byVal property, byVal value)
        dim selectedNode
        dim oldStyle
        dim newStyle
        dim xslPattern

        if selector = "body" then
            xslPattern = "//" & selector
        else
            xslPattern = "//div[@id =""" & selector & """]"
        end if
    
        set selectedNode = m_objPage.documentElement.selectSingleNode(xslPattern)
    
        if not (selectedNode is nothing) then
            oldStyle = selectedNode.getAttribute("style")
            newStyle = " " & property & ":" & value & ";"
            selectedNode.setAttribute "style", oldStyle & newStyle
        end if
    end sub

    private function loadXmlQuestFile(byVal source)
        set m_objQuest = CreateObject("Microsoft.XMLDOM")
        m_objQuest.validateOnParse = true
        m_objQuest.async = false
        m_objQuest.load(source)
    end function
    
    private function getDebug
        getDebug = "true" = m_objQuest.documentElement.getAttribute("debug")
    end function
    
    private function getLanguage
        getLanguage = m_objQuest.documentElement.getAttribute("language")
    end function
    
    private function getPageTitle
        dim title
    
        title = m_objQuest.selectSingleNode("//title").text
    
        getPageTitle = title
    end function

    private function constructHref(byVal stationLink, byVal statesString)
        dim href
        dim station
        dim splitted
        dim questName

        if inStr(stationLink, ":") >= 1 then
            splitted = split(stationLink, ":")
            questName = splitted(0)
            station = splitted(1)
        else
            questName = m_questName
            station = stationLink
        end if
        
        if g_isServerVersion then
            station = replace(station, " ", "%20")
            href = g_aspFileName & "?quest=" & questName & "&amp;" & _
                    "station=" & station & "&amp;" & _
                    "t=" & getIsoDateCompact(now) & "&amp;" & _
                    "session=" & m_sessionId & "&amp;" & _
                    "content=" & m_contentType
            if statesString <> "" then
                href = href & "&amp;" & statesString
            end if
        else
            href = "javascript:handleStation('" & questName & "', '" & station & "', " & _
                    "'" & m_sessionId & "', '" & m_contentType & "', '" & statesString & "')"
        end if

        constructHref = href
    end function
   
    private function getDisplay(byRef stationNode, byRef toInclude)
        dim child
        dim text
        dim path
        dim image
        dim imageMap
        dim imageMapString
        dim musicSource
        dim supressMusic
        dim listType
        dim includesImagemap
        dim imageSource
    
        imageSource = g_none
        musicSource = g_none
    
        includesImagemap = not (stationNode.selectSingleNode("choice[@area]") is nothing)
    
        for each child In stationNode.childNodes
            select case child.nodeName
                case "text"
                    displayText child, text, "source", "text", imageMapString, includesImagemap, imageSource
                case "image"
                    displayImage child, text, "source", "text", imageMapString, includesImagemap, imageSource, false
                case "music"
                    displayMusic child, text, musicSource, "source", supressMusic
                case "choice"
                    displayPath child, text, imageMap, path, "source", "text", imageMapString, includesImagemap, imageSource
                case "input"
                    displayInput child, text
                case "table"
                    text = text & getTable(child)
                case "embed"
                    text = text & getEmbed(child)
                case "state", "number", "string"
                    m_oStateHandler.setStates child
            end select
        next
    
        checkIfGameOver path, toInclude, stationNode
        handleMusic musicSource, supressMusic
        if includesImagemap then
            text = text & "<map id=""imapa"" name=""imapa"">" & imageMap & "</map>"
        end if
    
        text = cleanUpText(text)
        getDisplay = image & vbNewline & text & vbNewline & path
    end function

    private function getEmbed(byRef child)
        dim xhtml
        dim sSource
        dim thisStyle
        dim thisClass
        dim xmlEmbed

        xhtml = ""
        if m_oStateHandler.getNodeState(child) then
            sSource = child.getAttribute("source")
            if isNull( child.getAttribute("class") ) then
                thisClass = "qmlEmbed"
            else
                thisClass = element.getAttribute("class")
            end if
   
            thisStyle = getClassStyle(thisClass)

            if child.getAttribute("merge") = "false" then
                if thisStyle <> "" then
                    thisStyle = " style=""" & thisStyle & """ "
                end if
                xhtml = "<iframe " & thisStyle & "src=""" & sSource & """></iframe>"
            else
                set xmlEmbed = getXml(sSource)
                if thisStyle <> "" then
                    xmlEmbed.documentElement.setAttribute "style", thisStyle
                end if
                xhtml = xmlEmbed.documentElement.xml
            end if
        end if

        getEmbed = xhtml
    end function
    
    private function getTable(byRef parTable)
        dim xhtml
        dim table
        dim elements
        dim element
    
        set table = parTable.cloneNode(true)
        set elements = table.selectNodes(".//*")
        insertStyle table
        for each element in elements
            insertStyle element
        next
    
        xhtml = "<br /><br />" & table.xml & "<br /><br />"
    
        getTable = xhtml
    end function
    
    private sub insertStyle(byRef element)
        dim thisClass
        dim thisStyle
     
        if isNull( element.getAttribute("class") ) then
            thisClass = "qml" + toPropercase(element.nodeName)
        else
            thisClass = element.getAttribute("class")
        end if
    
        thisStyle = getClassStyle(thisClass)
    
        if thisStyle <> "" then
            element.setAttribute "style", thisStyle
        end if
        element.removeAttribute "class"
    end sub
   
    private sub handleMusic(byVal musicSource, byVal supressMusic)
        if musicSource <> g_none or m_defaultMusic <> g_none then
            if supressMusic then
                backgroundMusic.src = ""
            else
                if musicSource = g_none then
                    musicSource = m_defaultMusic
                end if

                if not backgroundMusic.loop = m_musicLoop then
                    backgroundMusic.loop = m_musicLoop
                end if
                backgroundMusic.src = musicSource
    
            end if
        end if
    end sub
    
    private sub checkIfGameOver(byRef path, byRef toInclude, byRef stationNode)
        if path <> "" then
            path = "<ul id=""choices"">" & path & "</ul>"
        elseif not toInclude then
            if ( stationNode.selectSingleNode(".//choice") is nothing ) then
                m_gameOver = true
            end if
        end if
    end sub
    
    private sub displayInput(byRef child, byRef text)
        dim station
        dim stringName

        station = m_oStateHandler.replaceAllValues( child.getAttribute("station") )
        'station = replace(station, " ", "%20")
        stringName = child.getAttribute("name")
        if isNull(stringName) then
            stringName = "qmlInput"
        else
            stringName = m_oStateHandler.replaceAllValues(stringName)
        end if

        if g_isServerVersion then
            text = text & vbNewline
            text = text & "<form method=""get"" action=""" & g_aspFileName & """>" & vbNewline
            text = text & "<input type=""hidden"" name=""quest"" value=""" & m_questName & """ />" & vbNewline
            text = text & "<input type=""hidden"" name=""session"" value=""" & m_sessionId & """ />" & vbNewline
            text = text & "<input type=""hidden"" name=""content"" value=""" & m_contentType & """ />" & vbNewline
            text = text & "<input type=""hidden"" name=""station"" value=""" & station & """ />" & vbNewline
            text = text & "<input type=""hidden"" name=""string_1_name"" value=""" & stringName & """ />" & vbNewline
            text = text & "<input type=""hidden"" name=""t"" value=""" & getIsoDateCompact(now) & """ />" & vbNewline
            text = text & "<input type=""text"" name=""string_1_value"" />" & vbNewline
            text = text & "<input type=""submit"" value=""" & child.text & """ />" & vbNewline
            text = text & "</form>" & vbNewline

        else
            text = text & "<form>"
            text = text & "<input type=""text"" onkeydown=""call trapKey()"" name=""string_1_value"" />" & vbNewline
            text = text & "<input type=""button"" value=""" & child.text & """ onclick=""" 
            text = text & "handleStation '" & m_questName & "', '" & station & "', " & _
                    "'" & m_sessionId & "', '" & m_contentType & "', " & _
                    "'string_1_name=" & stringName & "&amp;string_1_value=' + string_1_value.value"
            text = text & """ />" & vbNewline
            text = text & "</form>" & vbNewline

        end if

    end sub

    private sub displayPath(byRef child, byRef text, byRef imageMap, byRef path, byRef sSource, byRef sText, byRef imageMapString, byRef includesImagemap, byRef imageSource)
        dim pathText
        dim linkStyle   
        dim classStyle
        dim statesString
        dim oStatesString

        set oStatesString = new classStatesString
        statesString = oStatesString.getStatesFromChoice(child, m_oStateHandler)

        if m_oStateHandler.getNodeState(child) then
            if child.getAttribute("area") <> "" then
                imageMap = imageMap & getImageMapString( _
                        child.getAttribute("area"), _
                        getLink(child), _
                        child.text)
            else
                linkStyle = m_linkInlineStyle
                classStyle = getClassStyle("qmlLink")
                if classStyle <> "" then
                    linkStyle = replace(linkStyle, ";""", ";" & classStyle & """")
                end if
                pathText = "<a " & linkStyle & " " & _
                        "href=""" & constructHref( getLink(child), statesString) & """>" & _
                        getText(child, sSource, sText, imageMapString, includesImagemap, imageSource) & "</a>"
                path = path & wrapListWithClass(child, pathText, "qmlChoice")
            end if
        end if
    end sub
    
    private sub displayText(byRef child, byRef text, byRef sSource, byRef sText, byRef imageMapString, byRef includesImagemap, byRef imageSource)
        if m_oStateHandler.getNodeState(child) then
            text = text & wrapWithParagraphClass(child, getText(child, sSource, sText, imageMapString, includesImagemap, imageSource) )
        end if
    end sub

    private sub displayMusic(byRef child, byRef text, byRef musicSource, byRef sSource, byRef supressMusic)
        if m_oStateHandler.getNodeState(child) then
            musicSource = child.getAttribute(sSource)
            musicSource = m_oStateHandler.replaceAllValues(musicSource)
            m_musicLoop = returnIf(child.getAttribute("loop") = "true", "-1", "0")
            if child.getAttribute("default") = "true" then
                m_defaultMusic = musicSource
            end if

            supressMusic = (musicSource = g_none)
        end if
    end sub
    
    private sub displayImage(byRef child, byRef text, byRef sSource, byRef sText, byRef imageMapString, byRef includesImagemap, byRef imageSource, byRef isInline)
        dim imageClass
        dim supressImage
        dim thisImage
        dim altText
    
        if m_oStateHandler.getNodeState(child) then
            imageSource = child.getAttribute("source")
            imageSource = m_oStateHandler.replaceAllValues(imageSource)
            supressImage = (imageSource = g_none)
    
            if not supressImage then
                altText = child.getAttribute("text")
                altText = m_oStateHandler.replaceAllValues(altText)
                imageMapString = returnIf(includesImagemap, " usemap=""#imapa""", "")
    
                thisImage = "<img src=""" & imageSource & """" & _
                        " alt=""" & altText & """ " & imageMapString & " />"
    
                if isNull( child.getAttribute("class") ) then
                    imageClass = "qmlImage"
                else
                    imageClass = child.getAttribute("class")
                end if
                if not isInline then
                    thisImage = wrapWithElementClass(thisImage, "p", imageClass, "")
                end if
                text = text & thisImage
    
                if child.getAttribute("default") = "true" then
                    m_defaultImage = imageSource
                end if
            end if
        end if
    end sub
    
    private function wrapWithElementClass(byVal content, byVal nodeName, byVal className, byRef realClass)
        dim thisStyle
        dim elementWithClass
    
        thisStyle = getClassStyle(className)

        if thisStyle <> "" then
            thisStyle = " style=""" & thisStyle & """"
        end if
        if realClass <> "" then
            realClass = " class=""" & realClass & """"
        end if
    
        elementWithClass = "<" & nodeName & thisStyle & realClass & ">" & _
                content & "</" & nodeName & ">"
    
        wrapWithElementClass = elementWithClass
    end function
    
    private function wrapListWithClass(byRef listNode, byVal text, byVal defaultClass)
        dim listWithClass
        dim className
        dim classStyle
    
        className = listNode.getAttribute("class")
        if isNull( className ) then className = defaultClass
        classStyle = getClassStyle(className)
    
        if classStyle <> "" then
            if not instr(classStyle, "list-style-type") >= 1 then
                classStyle = "list-style-type: none;" & classStyle
            end if
            listWithClass = "<li style=""" & classStyle & """>" & text & "</li>"
        else
            listWithClass = "<li><p>" & text & "</p></li>"
        end if
    
        wrapListWithClass = listWithClass
    end function
    
    private function wrapWithParagraphClass(byRef thisNode, byVal text)
        dim paragraphWithClass
        dim classNode
        dim className
    
        className = thisNode.getAttribute("class")
    
        if className <> "" then
            paragraphWithClass = "<p style=""display: inline; " & getClassStyle(className) & """>" & text & "</p>"
        else
            paragraphWithClass = "<p style=""display: inline"">" & text & "</p>"
        end if
        
        wrapWithParagraphClass = paragraphWithClass
    end function
    
    private function getClassStyle(byRef parClassName)
        dim className
        dim classStyle
        dim classNode
        dim inherits
        dim parentClass
        dim parentClassStyle
        dim i
    
        className = m_oStateHandler.replaceAllValues(parClassName)
    
        classStyle = ""
        parentClassStyle = ""
        set classNode = m_objQuest.documentElement. _
                selectSingleNode("//class[@name = """ & className & """]")
    
        if not (classNode is nothing) then
            classStyle = classNode.getAttribute("style")
    
            inherits = classNode.getAttribute("inherits")
            if inherits <> "" then
                inherits = trim(trimDoubleSpaces(inherits))
                if instr(inherits, " ") >= 1 then
                    parentClass = split(inherits, " ")
    
                    for i = lbound(parentClass) to ubound(parentClass)
                        parentClassStyle = parentClassStyle & ";" & getClassStyle( parentClass(i) )
                    next
                else
                    parentClassStyle = getClassStyle(inherits)
                end if
    
                classStyle = ";" & parentClassStyle & ";" & classStyle & ";"
            end if
    
            classStyle = m_oStateHandler.replaceAllValues(classStyle)
            classStyle = removeSemicolonPairs(classStyle)
            classStyle = replace(classStyle, """", "'")
        end if
    
        getClassStyle = classStyle
    end function
    
    private function removeSemicolonPairs(byVal oldText)
        dim text

        text = oldText
        text = repeatedReplace(text, "  ", " ")
        text = repeatedReplace(text, " ;", ";")
        text = repeatedReplace(text, "; ", ";")
        text = repeatedReplace(text, ";;", ";")

        removeSemicolonPairs = text
    end function
    
    private function getImageMapString(byVal area, byVal link, byVal text)
        dim imageMapString
        dim map

        imageMapString = "<area shape=""poly"" coords=""[area]"" " & _
                " href=""" & constructHref("[link]", "") & """ alt=""[text]"" title=""[text]"" />"
    
        map = imageMapString
        map = replace(map, "[area]", area)
        map = replace(map, "[link]", link)
        map = replace(map, "[text]", text)

        getImageMapString = map
    end function
    
    private function cleanUpText(byVal parText)
        dim text
        dim oldText
    
        text = parText
    
        do
            oldText = text
            text = replace(text, "<p></p>", "")
            text = replace(text, "<p><br /></p>", "")
        loop until oldText = text
    
        cleanUpText = text
    end function
    
    private function getText(byRef node, byRef sSource, byRef sText, byRef imageMapString, byRef includesImagemap, byRef imageSource)
        dim child
        dim text
        dim convertedText
        dim choice
        dim choiceClass
    
        for each child In node.childNodes
            if getNodeType(child.nodeType) = "element" then
                select case child.nodeName
                    case "choice"
                        text = text & getInlineChoice(child)
                    case "break"
                        text = text & "<br />"
                        if child.getAttribute("type") = "strong" then
                            text = text & "<br />"
                        end if
                    case "emphasis"
                        text = text & wrapWithElementClass(child.firstChild.text, "em", "qmlEmphasis", "")
                    case "strong"
                        text = text & wrapWithElementClass(child.firstChild.text, "strong", "qmlStrong", "")
                    case "poem"
                        text = text & "</p><pre class=""poem"">" & child.firstChild.text & "</pre><p>"
                    case "display"
                        text = text & wrapWithElementClass(child.firstChild.text, "span", "qmlDisplay", "display")
                    case "link"
                        text = text & "<a href=""" & child.getAttribute("to") & """ " & _
                               "target=""_" & child.getAttribute("target") & """ class=""hyperlink"">" & _
                                 child.firstChild.text & "</a>"
                    case "image"
                        displayImage child, text, sSource, sText, imageMapString, includesImagemap, imageSource, true
                end select
            else
                convertedText = child.data
                convertedText = m_oStateHandler.replaceAllValues(convertedText)
    
                text = text & convertedText
            end if
        next
    
        getText = text
    end function
    
    private function getInlineChoice(byRef node)
        dim choice
        dim thisClass
        dim thisStyle
    
        if m_oStateHandler.getNodeState(node) then
            if isNull( node.getAttribute("class") ) then
                thisClass = "qmlInlineChoice"
            else
                thisClass = node.getAttribute("class")
            end if
            thisStyle = getClassStyle(thisClass)
            if thisStyle <> "" then
                thisStyle = "style=""" & thisStyle & """ "
            end if
            choice = "<a " & thisStyle & " " & _
                    "href=""" & constructHref( getLink(node), "" ) & """>" & _
                    node.text & "</a>"
        end if
    
        getInlineChoice = choice
    end function
    
    private function getLink(byRef choice)
        dim leadsTo

        leadsTo = choice.getAttribute("station")
        leadsTo = m_oStateHandler.replaceAllValues(leadsTo)
        if leadsTo = "back" then
            leadsTo = m_lastStation
        end if

        getLink = leadsTo
    end function
    
    private function getStation(byVal id)
        dim xPath

        xPath = "//station[@id = '" & id & "']"
        set getStation = m_objQuest.selectSingleNode(xPath)
    end function
    
    private sub outputStatus(byVal display)
        if g_isServerVersion then
            serverOutputToId "statusNode", display
        else
            m_objPage.all.statusNode.innerHTML = display
        end if
    end sub
    
    private sub output(byVal display)
        if g_isServerVersion then
            serverOutputToId "displayNode", display
        else
            m_objPage.all.displayNode.innerHTML = display
        end if
    end sub
    
    private sub serverOutputToId(byVal id, byVal display)
        dim displayNode
        dim content
        dim xPath
    
        set content = createObject("Microsoft.XMLDOM")
        xPath = "//div[@id = '" & id & "']"
        set displayNode = m_objPage.documentElement.selectSingleNode(xPath)
        content.loadXML "<div>" & display & "</div>"
        if content.parseError.errorCode <> 0 then
            showErrorOf content
        else
            if displayNode.childNodes.length > 0 then
                displayNode.removeChild displayNode.childNodes(0)
            end if
            displayNode.appendChild content.documentElement
        end if
    end sub
    
    private sub setStyle
        if g_isServerVersion then
            setStyleServer
        else
            setStyleClient
        end if
    end sub
    
    private sub setStyleClient
        dim child
        dim styleChild
        dim marginHasBeenSet
        dim linksDecoration
        dim linksColor
        dim doPositionContent
        dim doPositionStatus
    
        linksDecoration = ""
        linksColor = ""
        marginHasBeenSet = false
        doPositionContent = false
        doPositionStatus = false
    
        for each child in m_objQuest.documentElement.childNodes
            if child.nodeName = "style" then
                for each styleChild in child.childNodes
                    select case styleChild.nodeName
    
                        case "background"
                            if styleChild.getAttribute("color") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.backgroundColor = styleChild.getAttribute("color")
                            end if
                            if styleChild.getAttribute("image") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.backgroundImage = _
                                    convertToUrl(styleChild.getAttribute("image"))
                            end if
                            m_objPage.all.bodyNode.style.backgroundRepeat = _
                                styleChild.getAttribute("repeat")
    
                        case "font"
                            if styleChild.getAttribute("color") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.color = styleChild.getAttribute("color")
                                linksColor = "color: " & styleChild.getAttribute("color") & ";"
                            end if
                            if styleChild.getAttribute("family") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.fontFamily = styleChild.getAttribute("family")
                            end if
                            if styleChild.getAttribute("size") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.fontSize = styleChild.getAttribute("size")
                            end if
                            if styleChild.getAttribute("weight") <> g_defaultValue then
                                m_objPage.all.bodyNode.style.fontWeight = styleChild.getAttribute("weight")
                            end if
                            if not styleChild.getAttribute("links") = "underlined" then
                                linksDecoration = "text-decoration: none;"
                            end if
    
                        case "content"
                            if styleChild.getAttribute("width") <> g_defaultValue then
                                m_objPage.all.displayNode.style.width = styleChild.getAttribute("width")
                            end if
                            if styleChild.getAttribute("left") <> g_defaultValue then
                                m_objPage.all.displayNode.style.left = styleChild.getAttribute("left")
                                doPositionContent = true
                            end if
                            if styleChild.getAttribute("top") <> g_defaultValue then
                                m_objPage.all.displayNode.style.top = styleChild.getAttribute("top")
                                doPositionContent = true
                            end if
    
                    end select
                next
                exit for
            end if
        next

        if doPositionContent then
            m_objPage.all.displayNode.style.position = "absolute"
        end if
        if doPositionStatus then
            m_objPage.all.statusNode.style.position = "absolute"
        end if
    
        if linksDecoration = "" and linksColor = "" then
            m_linkInlineStyle = ""
        else
            m_linkInlineStyle = " style=""" & linksDecoration & linksColor & """ "
        end if
    end sub

    private sub setStyleServer
        dim child
        dim styleChild
        dim marginHasBeenSet
        dim linksDecoration
        dim linksColor
        dim doPositionContent
        dim doPositionStatus
        dim pageBodyNode
        dim pageStatusNode
        dim pageDisplayNode
        dim bodyNodeStyle
        dim statusNodeStyle
        dim displayNodeStyle
    
        linksDecoration = ""
        linksColor = ""
        marginHasBeenSet = false
        doPositionContent = false
        doPositionStatus = false
    
        set pageBodyNode = m_objPage.documentElement.selectSingleNode("//body")
        set pageDisplayNode = m_objPage.documentElement.selectSingleNode("//div[@id =""displayNode""]")
        set pageStatusNode = m_objPage.documentElement.selectSingleNode("//div[@id =""statusNode""]")
    
        for each child in m_objQuest.documentElement.childNodes
            if child.nodeName = "style" then
                for each styleChild in child.childNodes
                    select case styleChild.nodeName
    
                        case "background"
                            if styleChild.getAttribute("color") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "background-color: " & styleChild.getAttribute("color") & ";"
                            end if
                            if styleChild.getAttribute("image") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "background-image: " & convertToUrl(styleChild.getAttribute("image")) & ";"
                            end if
                            bodyNodeStyle = bodyNodeStyle & _
                                    "background-repeat: " & styleChild.getAttribute("repeat") & ";"
    
                        case "font"
                            if styleChild.getAttribute("color") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "color: " & styleChild.getAttribute("color") & ";"
                                linksColor = "color: " & styleChild.getAttribute("color") & ";"
                            end if
                            if styleChild.getAttribute("family") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "font-family: " & styleChild.getAttribute("family") & ";"
                            end if
                            if styleChild.getAttribute("size") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "font-size: " & styleChild.getAttribute("size") & ";"
                            end if
                            if styleChild.getAttribute("weight") <> g_defaultValue then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "font-weight: " & styleChild.getAttribute("weight") & ";"
                            end if
                            if not styleChild.getAttribute("links") = "underlined" then
                                bodyNodeStyle = bodyNodeStyle & _
                                        "text-decoration: none;"
                                    linksDecoration = "text-decoration: none;"
                            end if
    
                        case "content"
                            if styleChild.getAttribute("width") <> g_defaultValue then
                                displayNodeStyle = displayNodeStyle & _
                                        "width: " & styleChild.getAttribute("width") & ";"
                            end if
                            if styleChild.getAttribute("left") <> g_defaultValue then
                                displayNodeStyle = displayNodeStyle & _
                                        "left: " & styleChild.getAttribute("left") & ";"
                                doPositionContent = true
                            end if
                            if styleChild.getAttribute("top") <> g_defaultValue then
                                displayNodeStyle = displayNodeStyle & _
                                        "margin-top: " & styleChild.getAttribute("top") & ";"
                                doPositionContent = true
                            end if
    
                    end select
                next
                exit for
            end if
        next

        if doPositionContent then
            displayNodeStyle = displayNodeStyle & _
                    "position: absolute;"
        end if
        if doPositionStatus then
            statusNodeStyle = statusNodeStyle & _
                    "position: absolute;"
        end if
    
        if linksDecoration = "" and linksColor = "" then
            m_linkInlineStyle = ""
        else
            m_linkInlineStyle = " style=""" & linksDecoration & linksColor & """ "
        end if
    
        pageBodyNode.setAttribute "style", bodyNodeStyle
        pageDisplayNode.setAttribute "style", displayNodeStyle
        pageStatusNode.setAttribute "style", statusNodeStyle
    end sub
    
    private function convertToUrl(byVal filePath)
        dim newString
    
        newString = filePath
        if instr(newString, "url") < 1 then
            newString = "url('" & newString & "')"
        end if
        convertToUrl = newString
    end function

    private function stationExists(byVal id)
        dim stationNode
    
        set stationNode = m_objQuest.documentElement.selectSingleNode("//station[@id = """ & id & """]")
        stationExists = not (stationNode is nothing)
    end function
    
    private function language(byVal textEnglish, byVal textGerman)
        if m_language = "german" then
            language = textGerman
        else
            language = textEnglish
        end if
    end function

    private sub saveQuest
        const intervalMinute = "n"
        dim sessionData
        dim dateTimeOut

        sessionData = getSessionDataAsString

        if g_isServerVersion then
            dateTimeOut = dateAdd(intervalMinute, 1, now)
            dateTimeOut = getIsoDate(dateTimeOut)
            session("data") = sessionData
            ' setFileText "tool/session/" & m_sessionId & ".xml", sessionData
        else
            g_clientSessionData = sessionData
        end if
    end sub

    private sub loadQuest
        dim sessionData
        dim xmlSessionData

        if g_isServerVersion then
            ' removeTimedOutSessions
            sessionData = session("data")
            ' getFileText("tool/session/" & m_sessionId & ".xml")
        else
            sessionData = g_clientSessionData
        end if

        set xmlSessionData = getXmlString(sessionData)
        setSessionDataFromXml xmlSessionData

        m_oStateHandler.setFromStatesString m_statesString
    end sub

    private sub removeTimedOutSessions
        ' todo
    end sub
    
    ' ________________________________

    private sub setSessionDataFromXml(byRef xmlSession)
        dim questElements
        dim questElement
        dim thisValue
        dim xPath

        xPath = "//quest/*"
        set questElements = xmlSession.selectNodes(xPath)
        for each questElement in questElements
            select case questElement.nodeName
                case "beforeLastStation": m_beforeLastStation = questElement.text
                case "defaultImage": m_defaultImage = questElement.text
                case "defaultMusic": m_defaultMusic = questElement.text
                case "musicLoop": m_musicLoop = questElement.text
                case "firstQuestName": m_firstQuestName = questElement.text
                case "linkInlineStyle": m_linkInlineStyle = questElement.text
                case "language": m_language = questElement.text
                case "gameOver": m_gameOver = questElement.text
                case "lastStation": m_lastStation = questElement.text
            end select
        next

        m_oStateHandler.setSessionDataFromXml xmlSession
    end sub

    private function getSessionDataAsString
        dim sXml

        sXml = ""
        sXml = "<?xml version=""1.0""?>" & vbNewline
        sXml = sXml & "<qmlSession>" & vbNewline
        sXml = sXml & "<quest>" & vbNewline
        sXml = sXml & getTaggedValue("beforeLastStation", m_beforeLastStation)
        sXml = sXml & getTaggedValue("defaultImage", m_defaultImage)
        sXml = sXml & getTaggedValue("defaultMusic", m_defaultMusic)
        sXml = sXml & getTaggedValue("musicLoop", m_musicLoop)
        sXml = sXml & getTaggedValue("firstQuestName", m_firstQuestName)
        sXml = sXml & getTaggedValue("linkInlineStyle", m_linkInlineStyle)
        sXml = sXml & getTaggedValue("language", m_language)
        sXml = sXml & getTaggedValue("gameOver", m_gameOver)
        sXml = sXml & getTaggedValue("lastStation", m_lastStation)
        sXml = sXml & "</quest>" & vbNewline
        sXml = sXml & m_oStateHandler.getSessionDataAsString
        sXml = sXml & "</qmlSession>" & vbNewline

        sXml = getXmlString(sXml).xml
        
        getSessionDataAsString = sXml
    end function

    ' ________________________________

    private function verboseWeekday(byVal ofDate)
        dim strDay
        
        select case weekday(ofDate)
            case 1
                strDay = language("sunday", "Sonntag")
            case 2
                strDay = language("monday", "Montag")
            case 3
                strDay = language("tuesday", "Dienstag")
            case 4
                strDay = language("wednesday", "Mittwoch")
            case 5
                strDay = language("thursday", "Donnerstag")
            case 6
                strDay = language("friday", "Freitag")
            case 7
                strDay = language("saturday", "Samstag")
        end select
    
        verboseWeekday = strDay
    end function

    private sub setQmlStartVariables
        '' m_oStateHandler.setString "qmlSecondsStart", timer
        m_oStateHandler.setString "qmlVersion", g_qmlVersionNumber
        if g_isServerVersion then
            m_oStateHandler.setString "qmlServer", "true"
        else
            m_oStateHandler.setString "qmlServer", "false"
        end if
    end sub
    
    private sub setQmlVariables
        '' dim seconds
    
        m_oStateHandler.setString "qmlLastStation", m_lastStation
    
        '' seconds = cLng( timer - cLng( m_oStateHandler.getStringOfName("qmlSecondsStart") ) )
        '' if cLng(seconds) > 50000 then
        ''     seconds = 0
        '' end if
        '' m_oStateHandler.setNumber "qmlSeconds", seconds
        '' m_oStateHandler.setNumber "qmlMinutes", cLng(seconds / 60)
    
        '' m_oStateHandler.setString "qmlTime", time
        '' m_oStateHandler.setString "qmlDay", verboseWeekday(date)
    end sub

    private function getNewSessionId
        dim sessionId
        dim i
        dim compactQmlVersion
        dim period
    
        compactQmlVersion = g_qmlVersionNumber
        period = inStr(compactQmlVersion, ".")
        if period > 1 then
            compactQmlVersion = left(compactQmlVersion, period - 1)
        end if

        sessionId = ""
        sessionId = sessionId & "QML" & compactQmlVersion & "-"
        sessionId = sessionId & getIsoDateCompact(now) & "-"
        for i = 1 to 8
            sessionId = sessionId & cInt( rnd * 9 )
        next

        getNewSessionId = sessionId
    end function

end class