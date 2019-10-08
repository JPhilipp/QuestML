<%

class classServerResponse

    private m_objPage
    private m_contentType
    private m_sessionId
    private m_questName

    public sub setObjPage(byRef objPage)
        set m_objPage = objPage
    end sub

    public sub setSessionId(byVal sessionId)
        m_sessionId = sessionId
    end sub

    public sub setQuestName(byVal questName)
        m_questName = questName
    end sub

    public sub setContentType(byVal contentType)
        m_contentType = contentType
    end sub

    sub process
        dim responseString

        select case m_contentType
            case "text/html", "text/xhtml"
                responseString = getPageAsXhtml
            case "text/xml"
                responseString = getPageAsXml
            case "text/plain"
                responseString = getPageAsText
            case else
                responseString = "Error: Request for unsupported " & _
                        " content-type """ & m_contentType & """. " & _
                        "Supported are text/html, text/xhtml, text/xml, text/plain."
                m_contentType = "text/plain"
                response.status = "Bad request 400"
        end select

        response.contentType = m_contentType
        response.write responseString
    end sub

    private function getPageAsXhtml
        dim pageString

        pageString = m_objPage.documentElement.xml
        pageString = replace(pageString, "<br/>", "<br />")
        pageString = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" " & _
                """DTD/xhtml1-strict.dtd"">" & pageString

        getPageAsXhtml = pageString
    end function

    private function getPageAsXml
        dim sXml
        dim xPath
        dim text
        dim subElement
        dim subElements
        dim sStation

        sXml = ""

        sXml = sXml & "<?xml version=""1.0""?>"
        sXml = sXml & "<qmlMessage " & _
                "quest=""" & m_questName & """ session=""" & m_sessionId & """>" & vbNewline

        xPath = "//div/div/p"
        set subElements = m_objPage.selectNodes(xPath)
        sXml = sXml & "<text>"
        for each subElement in subElements
            sXml = sXml & subElement.text & vbNewline
        next
        sXml = sXml & "</text>" & vbNewline

        xPath = "//li/p/a"
        set subElements = m_objPage.selectNodes(xPath)
        for each subElement in subElements
            sStation = subElement.getAttribute("href")
            sStation = mid( sStation, inStr(sStation, "&station=") + len("&station=") )
            sStation = left( sStation, inStr(sStation, "&") - 1 )
            sXml = sXml & "<choice station=""" & sStation & """>"
            sXml = sXml & subElement.text
            sXml = sXml & "</choice>" & vbNewline
        next

        sXml = sXml & "</qmlMessage>"

        sXml = getXmlString(sXml).xml

        getPageAsXml = sXml
    end function

    private function getPageAsText
        getPageAsText = m_objPage.documentElement.text
    end function

end class

%>