option explicit

const serverVersion = false
const aspFileName = "qml.asp"
const qmlVersionNumber = "1.4"

const elementStation = "station"
const elementBreak = "break"
const elementImage = "image"
const elementText = "text"
const elementPath = "choice"
const elementIf = "if"
const elementInput = "input"
const elementRandomize = "randomize"
const elementEmphase = "emphasis"
const elementStrong = "strong"
const elementMusic = "music"
const elementTable = "table"
const elementComponent = "component"

const attributeSource = "source"
const attributeTextAlternative = "text"

const defaultValue = "default"
const noIndexFound = -1
const cNone = "none"

const relationAnd = "and"
const relationOr = "or"

const maximumStateAttributes = 9
const notState = "not "

const numberDefaultMin = -10000
const numberDefaultMax =  10000

const visitsStartString = "visits("

dim objPage

dim objQuest
dim gStation
dim gLastStation
dim gBeforeLastStation
dim gDefaultImage
dim gDefaultMusic
dim gMusicLoop
dim gQuestName
dim gFirstQuestName
dim gDidGoBeyondStart
dim gSavingAllowed
dim gLinkInlineStyle
dim gIsEncoded
dim gDebug
dim gDebugInfoIsDisplayed
dim gAlwaysDisplayInfo
dim gLanguage
dim gGameOver

dim arrState()
dim arrNumber()
dim arrNumberName()
dim arrNumberMin()
dim arrNumberMax()
dim arrNumberMinSet()
dim arrNumberMaxSet()

dim arrString()
dim arrStringName()

redimArrays

sub start(param)
    gLinkInlineStyle = ""
    gSavingAllowed = true
    resetStates
    prepareStart param, true
end sub

sub prepareStart(param, doShowCover)
    dim stationId
    dim name
    dim colon

    setObjPage

    colon = instr(param, ":")
    if colon >= 1 then
        name = left(param, colon - 1)
        stationId = mid(param, colon + 1)
    else
        name = param
        stationId = "start"
    end if

    if stationId <> "start" then
        doShowCover = false
    end if

    setQmlStartVariables
    setQmlVariables
    startQuest name, stationId, doShowCover
end sub

sub setObjPage
    if serverVersion then
        set objPage = createObject("Microsoft.XMLDOM")
        objPage.loadXML "<html><head><title>QML</title><meta name=""expires"" content=""0"" /><link rel=""stylesheet"" href=""style/quest.css"" type=""text/css"" media=""screen"" /><link rel=""stylesheet"" href=""style/components.css"" type=""text/css"" media=""screen"" /></head><body id=""bodyNode""> <div id=""statusNode""> </div> <div id=""displayNode""></div><div id=""stateDisplay""> </div></body></html>"
        if objPage.parseError.errorCode <> 0 then
            showErrorOf objPage
        end if
    else
        set objPage = document
    end if
end sub

sub startQuest(name, stationId, doShowCover)
    dim quest
    gQuestName = name

    if doShowCover then
        gFirstQuestName = name
    end if

    gDidGoBeyondStart = false
    randomize timer
    gDefaultImage = cNone
    gDefaultMusic = cNone
    gSavingAllowed = true
    gMusicLoop = "0"

    loadQuest mapPathIf(gQuestName & ".xml")

    if objQuest.parseError.errorCode = 0 then
        setStyle
        gIsEncoded = getIsEncoded
        gDebug = getDebug
        gLanguage = getLanguage

        setDocTitle getPageTitle

        if doShowCover then
            doShowCover = cBool(objQuest.selectSingleNode("//about").getAttribute("show") = "true")
        end if

        if doShowCover then
            showCover
        else
            handleStation stationId
        end if
    else
        showError
    end if
end sub

sub setDocTitle(text)
    dim objTitle
    if serverVersion then
        set objTitle = objPage.documentElement.selectSingleNode("//title")
        objTitle.text = text
    else
        objPage.title = text
    end if
end sub

sub handleStation(stationId)
    dim displayGotten
    dim inputString
    dim station
    inputString = ""

    if instr(stationId, ":") >= 1 then
        handleChapterJump stationId

    else
        handleIfFirstStation

        set station = getStation(stationId)
        if (station is nothing) then exit sub
        handleTopChoose station, stationId

        setString "qmlStation", stationId


        handleStationSettings station

        handleCheckStates station

        setStates station, "before"
        displayGotten = getDisplay(station, false, false)
        handleInclude stationId, displayGotten

        setStates station, "after"

        output displayGotten

        gDebugInfoIsDisplayed = false
        if gAlwaysDisplayInfo then displayStates stationId

        gBeforeLastStation = gLastStation
        gLastStation = stationId
        setQmlVariables

        addVisits stationId

        showStatus
        setString "qmlInput", ""
    end if
end sub

sub handleTopChoose(station, stationId)
    dim choose
    set choose = station.selectSingleNode("choose")
    if not (choose is nothing) then
        addVisits stationId
        stationId = choose.getAttribute("station")
        set station = getStation(stationId)
        setStatesBoth choose.parentNode
    end if
end sub

sub handleInclude(stationId, oldDisplay)
    dim includeIn
    dim inNode
    dim includeNode
    dim doInclude
    dim includeState
    dim station
    dim newDisplay

    set includeIn = objQuest.documentElement.selectNodes("//in")

    for each inNode in includeIn

        if compareStrings(inNode.getAttribute("station"), stationId) then

            if getNodeState(inNode) then
                set includeNode = inNode.parentNode
                if getNodeState(includeNode) then
                    set station = includeNode.parentNode
                    handleCheckStates station
                    setStates station, "before"
                    newDisplay = getDisplay(station, false, true)
                    setStates station, "after"

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

function combineDisplay(station1, station2)
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
            set lastParag = station1.documentElement.appendChild( objQuest.createElement("p") )
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

sub handleStationSettings(station)
    handleSavingSettings station

    if station.getAttribute("states") = "reset" then
        resetStates
    end if
end sub

sub handleSavingSettings(station)
    if station.getAttribute("saving") = "on" then
        gSavingAllowed = true
    elseif station.getAttribute("saving") = "off" then
        gSavingAllowed = false
    end if
end sub

sub handleChapterJump(stationId)
    gLastStation = ""
    gBeforeLastStation = ""
    resetStates
    prepareStart stationId, false
end sub

sub handleIfFirstStation
    if not gDidGoBeyondStart then
        if stationExists("information") then
            if serverVersion then
                addStyle "statusNode", "display", "block"
            else
                objPage.all.statusNode.style.display = "block"
            end if
        end if
        setString "qmlTitle", getPageTitle
        gDidGoBeyondStart = true
    end if
end sub

sub addStyle(selector, property, value)
    dim selectedNode, oldStyle, newStyle, xslPattern
    if selector = "body" then
        xslPattern = "//" & selector
    else
        xslPattern = "//div[@id =""" & selector & """]"
    end if

    set selectedNode = objPage.documentElement.selectSingleNode(xslPattern)

    if not (selectedNode is nothing) then
        oldStyle = selectedNode.getAttribute("style")
        newStyle = " " & property & ":" & value & ";"
        selectedNode.setAttribute "style", oldStyle & newStyle
    end if
end sub

function loadQuest(source)
    set objQuest = CreateObject("Microsoft.XMLDOM")
    objQuest.validateOnParse = true
    objQuest.async = false
    objQuest.load(source)
end function

function getIsEncoded
    getIsEncoded = "true" = objQuest.documentElement.getAttribute("encoded")
end function

function getDebug
    getDebug = "true" = objQuest.documentElement.getAttribute("debug")
end function

function getLanguage
    getLanguage = objQuest.documentElement.getAttribute("language")
end function

function getPageTitle
    dim title

    title = objQuest.selectSingleNode("//title").text

    getPageTitle = encodeIf(title)
end function

sub showStatus
    const statusId = "information"
    dim oStatusNode, statusText
    set oStatusNode = objQuest.documentElement. _
            selectSingleNode("station[@id = """ & statusId & """]")
    if not (oStatusNode is nothing) then
        statusText = getDisplay(oStatusNode, true, true)
        outputStatus statusText
    end if
end sub

sub showCover
    dim child
    dim aboutNode
    dim text
    dim author
    dim title
    dim email
    dim homepage
    dim contact
    dim footnote
    dim cover
    dim intro
    dim startLink

    set aboutNode = objQuest.selectSingleNode("//about")

    for each child in aboutNode.childNodes
        select case child.nodeName
            case "title"
                if child.getAttribute("show") = "true" then
                    title = "<h1>" & encodeIf(child.text) & "</h1>"
                end if
            case "author"
                if child.getAttribute("show") = "true" then
                    author = "<p class=""author"">" & language("By", "Von") & " " & encodeIf(child.text) & "</p>"
                end if

            case "cover"
                cover = "<p class=""coverImage""><img src=""" & _
                        child.getAttribute("source") & """ alt="""" /></p>"
            case "intro"
                intro = "<p class=""intro"">" & encodeIf(child.text) & "</p>"
            case "email"
                email = "Email: <a " & gLinkInlineStyle & " href=""mailto:" & _
                        encodeIf(child.text) & """>" & encodeIf(child.text) & "</a>"
            case "homepage"
                homepage = "Homepage: <a " & gLinkInlineStyle & " href=""" & _
                           encodeIf(child.text) & " "">" & encodeIf(child.text) & "</a>"
        end select
    next

    startLink = "<ul id=""choices""><li><a " & gLinkInlineStyle & _
                "href=""" & constructHref("start") & """>Start</a></li></ul>"
    contact = "<p class=""contact"">" & email & vbNewline & "<br />" & _
              homepage & "</p>"
    footnote = getFootnote

    text = title & author & cover & intro & startLink & vbNewline & contact
    text = left(text, 1000) & vbNewline & footnote

    text = "<div class=""cover"">" & text & "</div>"
    output text
end sub

function constructHref(station)
    dim href
    if serverVersion then
        href = aspFileName & "?station=" & station & getTimeParameter
    else
        href = "javascript:handleStation('" & station & "')"
    end if
    constructHref = href
end function

function getTimeParameter
    dim parameter
    parameter = qmlVersionNumber
    parameter = parameter & now
    parameter = replace(parameter, " ", "")
    parameter = replace(parameter, ".", "")
    parameter = replace(parameter, ":", "")
    getTimeParameter = "&amp;t=" & parameter
end function

function getFootnote
    dim footnote
    footnote = "<p class=""footnote"">"

    if gLanguage = "german" then
        footnote = footnote & _
               "Diese Quest ist in QML (Quest Markup Language) geschrieben. " & _
               "QML ist Freeware Copyright &copy; 2000, 2001 by <a " & gLinkInlineStyle & _
               "href=""mailto:lenssen@hitnet.rwth-aachen.de?subject=QML"">" & _
               "Philipp Lenssen</a>. Besuchen Sie <a " & gLinkInlineStyle & _
               "href=""http://www.outer-court.com/goodies/qml.htm"">The " & _
               "Outer Court</a> für mehr Informationen und aktuelle Versionen."
    else
        footnote = footnote & _
               "This quest is written in QML (Quest Markup Language). " & _
               "QML is Freeware Copyright © 2000, 2001 by <a " & gLinkInlineStyle & _
               "href=""mailto:lenssen@hitnet.rwth-aachen.de?subject=QML"">" & _
               "Philipp Lenssen</a>. Visit <a " & gLinkInlineStyle & _
               "href=""http://www.outer-court.com/goodies/qml.htm"">The " & _
               "Outer Court</a> for more information and the latest updates."

    end if

    footnote = footnote & "</p>"

    getFootnote = footnote
end function

function getChance(chanceString)
    const chanceLimitNormal = 10
    const chanceLimitPercentage = 100
    dim chanceLimit
    dim chanceNumber
    dim erronousChance

    erronousChance = false
    chanceString = replace(chanceString, " ", "")

    if instr(chanceString, "%") >= 1 then
        chanceLimit = chanceLimitPercentage
        chanceNumber = cInt(mid(chanceString, 1, len(chanceString) - 1))
        if chanceNumber < 0 or chanceNumber > chanceLimitPercentage then
            erronousChance = true
        end if
    else
        chanceLimit = chanceLimitNormal
        chanceNumber = cInt(chanceString)
        if chanceNumber < 0 or chanceNumber > chanceLimitNormal then
            erronousChance = true
        end if
    end if

    if erronousChance then
        sendError "Erronous chance: " & chanceString & chr(13) & chr(10) & _
               "Must be 0 to 10 or 0% to 100%"
    end if

    getChance = (rnd * chanceLimit < chanceNumber)
end function

function getDisplay(stationNode, isStatus, toInclude)
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

    imageSource = cNone
    musicSource = cNone

    includesImagemap = not (stationNode.selectSingleNode("choice[@area]") is nothing)

    for each child In stationNode.childNodes
        select case child.nodeName
            case elementText
                displayText child, text, isStatus, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource
            case elementImage
                displayImage child, text, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource, false
            case elementMusic
                displayMusic child, text, musicSource, attributeSource, supressMusic, isStatus
            case elementPath
                displayPath child, text, isStatus, imageMap, path, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource
            case elementTable
                text = text & getTable(child)
            case elementComponent
                text = text & getComponent(child)
        end select
    next

    if not isStatus then
        checkIfGameOver path, toInclude, stationNode
        handleMusic musicSource, supressMusic
        if includesImagemap then
            text = text & "<map id=""imapa"" name=""imapa"">" & imageMap & "</map>"
        end if
    end if

    text = cleanUpText(text)
    getDisplay = image & vbNewline & text & vbNewline & path
end function

function getTable(parTable)
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

sub insertStyle(element)
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

function getComponent(component)
    const prefix = "component"
    dim xhtml
    dim oXhtml
    dim valuesOf
    dim nameOf
    dim returns
    dim isValid

    nameOf = component.getAttribute("name")
    if left( nameOf, len(prefix) ) = prefix then
        nameOf = mid( nameOf, len(prefix) + 1 )
    end if
    nameOf = ucase( left(nameOf, 1) ) & mid(nameOf, 2)
    nameOf = "component" & nameOf

    valuesOf = component.getAttribute("values")
    valuesOf = replaceAllValues(valuesOf)

    returns = component.getAttribute("returns")
    returns = lcase(returns)

    xhtml = ""
    if returns = "xhtml" then
        set oXhtml = getComponentJS(nameOf, valuesOf)
        isValid = cBool(oXhtml.parseError.errorCode = 0)
        if isValid then
            xhtml = oXhtml.xml
        else
            showErrorOf oXhtml
        end if
    else ' if returns = "void" then
        handleComponentJS nameOf, valuesOf
    end if

    getComponent = xhtml
end function

sub handleMusic(musicSource, supressMusic)
    if musicSource <> cNone or gDefaultMusic <> cNone then
        if supressMusic then
            backgroundMusic.src = ""
        else
            if musicSource = cNone then
                musicSource = gDefaultMusic
            end if

            if not backgroundMusic.loop = gMusicLoop then
                backgroundMusic.loop = gMusicLoop
            end if
            backgroundMusic.src = musicSource

        end if
    end if
end sub

sub checkIfGameOver(path, toInclude, stationNode)
    if path <> "" then
        path = "<ul id=""choices"">" & path & "</ul>"
    elseif not toInclude then
        if ( stationNode.selectSingleNode(".//choice") is nothing ) then
            gGameOver = true
        end if
    end if
end sub

sub displayPath(child, text, isStatus, imageMap, path, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource)
    dim pathText
    dim linkStyle   
    dim classStyle

    if not isStatus then
        if getNodeState(child) then
            if child.getAttribute("area") <> "" then
                imageMap = imageMap & getImageMapString( _
                        child.getAttribute("area"), _
                        getLink(child), _
                        encodeIf(child.text))
            else
                linkStyle = gLinkInlineStyle
                classStyle = getClassStyle("qmlLink")
                if classStyle <> "" then
                    linkStyle = replace(linkStyle, ";""", ";" & classStyle & """")
                end if
                pathText = "<a " & linkStyle & " " & _
                        "href=""" & constructHref( getLink(child) ) & """>" & _
                        getText(child, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource) & "</a>"
                path = path & wrapListWithClass(child, pathText, "qmlChoice")
            end if
        end if
    end if
end sub

sub displayText(child, text, isStatus, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource)
    if getNodeState(child) then
        text = text & wrapWithParagraphClass(child, getText(child, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource), _
                isStatus)
    end if
end sub

sub displayMusic(child, text, musicSource, attributeSource, supressMusic, isStatus)
    if getNodeState(child) then
        if not isStatus then
            musicSource = child.getAttribute(attributeSource)
            musicSource = replaceAllValues(musicSource)
            gMusicLoop = returnIf(child.getAttribute("loop") = "true", "-1", "0")
            if child.getAttribute("default") = "true" then
                gDefaultMusic = musicSource
            end if

            supressMusic = (musicSource = cNone)
        end if
    end if
end sub

sub displayImage(child, text, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource, isInline)
    dim imageClass
    dim supressImage
    dim thisImage
    dim altText

    if getNodeState(child) then
        imageSource = child.getAttribute(attributeSource)
        imageSource = replaceAllValues(imageSource)
        supressImage = (imageSource = cNone)

        if not supressImage then
            altText = child.getAttribute(attributeTextAlternative)
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
                gDefaultImage = imageSource
            end if
        end if
    end if
end sub

function wrapWithElementClass(content, nodeName, className, realClass)
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

function wrapListWithClass(listNode, text, defaultClass)
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

function wrapWithParagraphClass(thisNode, text, isStatus)
    dim paragraphWithClass
    dim classNode
    dim className

    className = thisNode.getAttribute("class")

    if isStatus then
       paragraphWithClass = "<p>" & text & "</p>"
    elseif className <> "" then
        paragraphWithClass = "<p style=""display: inline; " & getClassStyle(className) & """>" & text & "</p>"
    else
        paragraphWithClass = "<p style=""display: inline"">" & text & "</p>"
    end if
    
    wrapWithParagraphClass = paragraphWithClass
end function

function getClassStyle(parClassName)
    dim className
    dim classStyle
    dim classNode
    dim inherits
    dim parentClass
    dim parentClassStyle
    dim i

    className = replaceAllValues(parClassName)

    classStyle = ""
    parentClassStyle = ""
    set classNode = objQuest.documentElement. _
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

        classStyle = replaceAllValues(classStyle)
        classStyle = removeSemicolonPairs(classStyle)
        classStyle = replace(classStyle, """", "'")
    end if

    getClassStyle = classStyle
end function

function removeSemicolonPairs(oldText)
    dim text
    text = oldText
    text = repeatedReplace(text, "  ", " ")
    text = repeatedReplace(text, " ;", ";")
    text = repeatedReplace(text, "; ", ";")
    text = repeatedReplace(text, ";;", ";")
    removeSemicolonPairs = text
end function

function getImageMapString(area, link, text)
    dim imageMapString
    dim map
    imageMapString = "<area shape=""poly"" coords=""[area]"" " & _
            " href=""" & constructHref("[link]") & """ alt=""[text]"" title=""[text]"" />"

    map = imageMapString
    map = replace(map, "[area]", area)
    map = replace(map, "[link]", link)
    map = replace(map, "[text]", text)
    getImageMapString = map
end function

function cleanUpText(parText)
    dim text, oldText
    text = parText

    do
        oldText = text
        text = replace(text, "<p></p>", "")
        text = replace(text, "<p><br /></p>", "")
    loop until oldText = text

    cleanUpText = text
end function

function getText(node, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource)
    dim child
    dim text
    dim convertedText
    dim choice
    dim choiceClass

    for each child In node.childNodes
        if getNodeType(child.nodeType) = "element" then
            select case child.nodeName
                case elementPath
                    text = text & getInlineChoice(child)
                case elementBreak
                    text = text & "<br />"
                    if child.getAttribute("type") = "strong" then
                        text = text & "<br />"
                    end if
                case elementEmphase
                    text = text & wrapWithElementClass(encodeIf(child.firstChild.text), "em", "qmlEmphasis", "")
                case elementStrong
                    text = text & wrapWithElementClass(encodeIf(child.firstChild.text), "strong", "qmlStrong", "")
                case "poem"
                    text = text & "</p><pre class=""poem"">" & encodeIf(child.firstChild.text) & "</pre><p>"
                case "display"
                    text = text & wrapWithElementClass(encodeIf(child.firstChild.text), "span", "qmlDisplay", "display")
                case "link"
                    text = text & "<a href=""" & child.getAttribute("to") & """ " & _
                           "target=""_" & child.getAttribute("target") & """ class=""hyperlink"">" & _
                             encodeIf(child.firstChild.text) & "</a>"
                case elementImage
                    displayImage child, text, attributeSource, attributeTextAlternative, imageMapString, includesImagemap, imageSource, true
            end select
        else
            convertedText = encodeIf(child.data)
            convertedText = replaceNumberValues(convertedText)
            convertedText = replaceStringValues(convertedText)
            convertedText = replaceStateValues(convertedText)

            text = text & convertedText
        end if
    next

    getText = text
end function

function getInlineChoice(node)
    dim choice
    dim thisClass
    dim thisStyle

    if getNodeState(node) then
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
                "href=""" & constructHref( getLink(node) ) & """>" & _
                encodeIf(node.text) & "</a>"
    end if

    getInlineChoice = choice
end function

function encodeIf(text)
    dim newText, i, letter
    newText = ""

    if gIsEncoded then
        for i = 1 to len(text)
            letter = mid(text, i, 1)
            if letter >= "a" and letter <= "z" then
                if letter = "a" then
                    letter = "z"
                else
                    letter = chr(asc(letter) - 1)
                end if
            end if
            newText = newText & letter
        next
    else
        newText = text
    end if

    encodeIf = newText
end function

function getLink(path)
    dim leadsTo
    leadsTo = replaceStringValues( path.getAttribute("station") )
    leadsTo = replaceNumberValues(leadsTo)
    if leadsTo = "back" then leadsTo = gLastStation

    getLink = leadsTo
end function

function getStation(id)
    set getStation = objQuest.documentElement.selectSingleNode("//station[@id = """ & id & """]")
end function

sub outputStatus(display)
    if serverVersion then
        serverOutputToId "statusNode", display
    else
        objPage.all.statusNode.innerHTML = display
    end if
end sub

sub output(display)
    if serverVersion then
        serverOutputToId "displayNode", display
    else
        objPage.all.displayNode.innerHTML = display
    end if
end sub

sub serverOutputToId(id, display)
    dim displayNode, content
    set content = createObject("Microsoft.XMLDOM")
    set displayNode = objPage.documentElement.selectSingleNode("//div[@id =""" & id & """]")
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

sub setStyle
    if serverVersion then
        setStyleServer
    else
        setStyleClient
    end if
end sub

sub setStyleClient
    dim child, styleChild, marginHasBeenSet, _
        linksDecoration, linksColor, _
        doPositionContent, doPositionStatus
    linksDecoration = ""
    linksColor = ""
    marginHasBeenSet = false
    doPositionContent = false
    doPositionStatus = false

    for each child in objQuest.documentElement.childNodes
        if child.nodeName = "style" then
            for each styleChild in child.childNodes
                select case styleChild.nodeName

                    case "background"
                        if styleChild.getAttribute("color") <> defaultValue then
                            objPage.all.bodyNode.style.backgroundColor = styleChild.getAttribute("color")
                        end if
                        if styleChild.getAttribute("image") <> defaultValue then
                            objPage.all.bodyNode.style.backgroundImage = _
                                convertToUrl(styleChild.getAttribute("image"))
                        end if
                        objPage.all.bodyNode.style.backgroundRepeat = _
                            styleChild.getAttribute("repeat")

                    case "font"
                        if styleChild.getAttribute("color") <> defaultValue then
                            objPage.all.bodyNode.style.color = styleChild.getAttribute("color")
                            linksColor = "color: " & styleChild.getAttribute("color") & ";"
                        end if
                        if styleChild.getAttribute("family") <> defaultValue then
                            objPage.all.bodyNode.style.fontFamily = styleChild.getAttribute("family")
                        end if
                        if styleChild.getAttribute("size") <> defaultValue then
                            objPage.all.bodyNode.style.fontSize = styleChild.getAttribute("size")
                        end if
                        if styleChild.getAttribute("weight") <> defaultValue then
                            objPage.all.bodyNode.style.fontWeight = styleChild.getAttribute("weight")
                        end if
                        if not styleChild.getAttribute("links") = "underlined" then
                            linksDecoration = "text-decoration: none;"
                        end if

                    case "content"
                        if styleChild.getAttribute("width") <> defaultValue then
                            objPage.all.displayNode.style.width = styleChild.getAttribute("width")
                        end if
                        if styleChild.getAttribute("left") <> defaultValue then
                            objPage.all.displayNode.style.left = styleChild.getAttribute("left")
                            doPositionContent = true
                        end if
                        if styleChild.getAttribute("top") <> defaultValue then
                            objPage.all.displayNode.style.top = styleChild.getAttribute("top")
                            doPositionContent = true
                        end if

                    case "information"
                        if styleChild.getAttribute("left") <> defaultValue then
                            objPage.all.statusNode.style.left = styleChild.getAttribute("left")
                            objPage.all.statusNode.style.position = "absolute"
                            doPositionStatus = true
                        end if
                        if styleChild.getAttribute("top") <> defaultValue then
                            objPage.all.statusNode.style.top = styleChild.getAttribute("top")
                            objPage.all.statusNode.style.position = "absolute"
                            doPositionStatus = true
                        end if
                        if styleChild.getAttribute("width") <> defaultValue then
                            objPage.all.statusNode.style.width = styleChild.getAttribute("width")
                        end if
                        if styleChild.getAttribute("height") <> defaultValue then
                            objPage.all.statusNode.style.height = styleChild.getAttribute("height")
                        end if
                        if styleChild.getAttribute("backgroundColor") <> defaultValue then
                            objPage.all.statusNode.style.backgroundColor = styleChild.getAttribute("backgroundColor")
                        end if
                        if styleChild.getAttribute("color") <> defaultValue then
                            objPage.all.statusNode.style.color = styleChild.getAttribute("color")
                        end if
                        if styleChild.getAttribute("fontSize") <> defaultValue then
                            objPage.all.statusNode.style.fontSize = styleChild.getAttribute("fontSize")
                        end if
                        if styleChild.getAttribute("padding") <> defaultValue then
                            objPage.all.statusNode.style.padding = styleChild.getAttribute("padding")
                        end if
                        if styleChild.getAttribute("textAlign") <> defaultValue then
                            objPage.all.statusNode.style.textAlign = styleChild.getAttribute("textAlign")
                        end if

                end select
            next
            exit for
        end if
    next

    if doPositionContent then
        objPage.all.displayNode.style.position = "absolute"
    end if
    if doPositionStatus then
        objPage.all.statusNode.style.position = "absolute"
    end if

    if linksDecoration = "" and linksColor = "" then
        gLinkInlineStyle = ""
    else
        gLinkInlineStyle = " style=""" & linksDecoration & linksColor & """ "
    end if
end sub

function convertToUrl(filePath)
    dim newString
    newString = filePath
    if instr(newString, "url") < 1 then
        newString = "url('" & newString & "')"
    end if
    convertToUrl = newString
end function

sub saveGame
    dim savingAllowedHere
    dim station
    set station = getStation(gLastStation)

    if gDidGoBeyondStart then
        if station.getAttribute("saving") = "allowed" then
            savingAllowedHere = true
        elseif station.getAttribute("saving") = "forbidden" then
            savingAllowedHere = false
        else
            savingAllowedHere = gSavingAllowed
        end if
    end if

    if savingAllowedHere then
        if confirm(language( _ 
                   "This game will be saved. Older saved games of this" & chr(13) & chr(10) & _
                   "adventure will be overwritten.", _
                   "Das Spiel wird gespeichert. Ältere Speicherstände dieses" & chr(13) & chr(10) & _
                   "Abenteuers werden dabei überschrieben." _
                    )) then
            SetCookie "[QML]" + gFirstQuestName, getFileSaveText(station)
            sendMessage language("Game was saved", "Spiel wurde gespeichert")
        end if
    else
        sendError language("Saving is not allowed here.", "Speichern ist hier nicht erlaubt")
    end if

end sub

sub loadGame
    if gDidGoBeyondStart then
        if confirm(language( _
                   "The saved game for this adventure will be loaded." & chr(13) & chr(10) & _
                   "Your current adventure will be lost.", _
                   "Das gespeicherte Spiel für dieses Abenteuer wird geladen." & chr(13) & chr(10) & _
                   "Das jetzige Abenteuer geht dabei verloren." _
                   )) then
            setVariablesFromString GetCookie("[QML]" + gFirstQuestName)
            sendMessage language("Game was loaded", "Spiel wurde geladen")
        end if
    else
        sendMessage language("You can only load from within the game.", "Es kann nur innerhalb des Spiels geladen werden.")
    end if
end sub

sub setVariablesFromString(textString)
    const modeStates = 1
    const modeNumbers = 2
    const modeStrings = 3
    dim splitted
    dim i
    dim stationId
    dim variableMode

    variableMode = modeStates

    splitted = Split(textString, "|")

    gLastStation = splitted(1)
    gDefaultImage = splitted(2)
    gDefaultMusic = splitted(3)
    gMusicLoop = splitted(4)
    gSavingAllowed = splitted(5)

    resetStates
    for i = 6 to ubound(splitted)

        select case splitted(i)
            case "[switchToNumbers]"
                variableMode = modeNumbers
            case "[switchToStrings]"
                variableMode = modeStrings

            case else

                select case variableMode
                    case modeStates
                        setState "" & splitted(i), true
                    case modeNumbers
                        setNumberFromCookie splitted(i)
                    case modeStrings
                        setStringFromCookie splitted(i)
                end select

        end select
    next

    prepareStart splitted(0), false
end sub

function getFileSaveText(station)
    dim textString
    dim i
    dim textValue

    textString = ""
    textString = textString & gQuestName & ":" & _
                              station.getAttribute("id") & "|"
    textString = textString & gBeforeLastStation & "|"
    textString = textString & gDefaultImage & "|"
    textString = textString & gDefaultMusic & "|"
    textString = textString & gMusicLoop & "|"
    textString = textString & gSavingAllowed & "|"

    for i = lbound(arrState) to ubound(arrState)
        if arrState(i) <> "" then
            textString = textString & "|" & arrState(i)
        end if
    next

    textString = textString & "|[switchToNumbers]"
    for i = lbound(arrNumber) to ubound(arrNumber)
        if arrNumberName(i) <> "" then
            textString = textString & "|" & _
                         arrNumberName(i) & "=" & arrNumber(i)
            if getArrNumberMin(i) <> numberDefaultMin or getArrNumberMax(i) <> numberDefaultMax then
                textString = textString & "(" & getArrNumberMin(i) & " " & getArrNumberMax(i) & ")"
            end if
        end if
    next

    textString = textString & "|[switchToStrings]"
    for i = lbound(arrString) to ubound(arrString)
        if arrString(i) <> "" then
            textValue = arrString(i)
            textValue = replace(textValue, "|", "/")
            textValue = replace(textValue, "=", "-")
            textString = textString & "|" & arrStringName(i) & "=" & textValue
        end if
    next

    getFileSaveText = textString
end function

function stationExists(id)
    dim stationNode
    set stationNode = objQuest.documentElement.selectSingleNode("//station[@id = """ & id & """]")
    stationExists = not (stationNode is nothing)
end function

function language(textEnglish, textGerman)
    if gLanguage = "german" then
        language = textGerman
    else
        language = textEnglish
    end if
end function

sub redimArrays
    redim arrState(10000)
    redim arrNumber(5000)
    redim arrNumberName(5000)
    redim arrNumberMin(5000)
    redim arrNumberMax(5000)
    redim arrNumberMinSet(5000)
    redim arrNumberMaxSet(5000)
    redim arrString(5000)
    redim arrStringName(5000)
end sub
