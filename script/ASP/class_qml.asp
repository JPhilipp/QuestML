<%

' QML-Converter class v0.1

class Qml
    dim m_source ' as msxml2.domdocument40
    dim m_spacing

    public sub setSource(byRef source)
        set m_source = source
    end sub

    public function getQl
        const sComment = "//"
        dim ql ' as string
        dim station ' as element
        dim stations ' as elementList
        dim indent ' as integer

        m_spacing = 4
        ql = ""

        indent = 0
        set stations = m_source.selectNodes("//station")
        for each station in stations
            ql = ql & station.getAttribute("id") & vbNewline
            ql = ql & getMainContent(station, indent) & vbNewline
        next

        getQl = ql
    end function

    private function getMainContent(byRef topNode, byVal indent) ' as string
        dim ql ' as string
        dim nNode ' as element
        dim nNodes ' as elementList
        dim sNodename ' as string
        dim sValue ' as string
        dim sCheck ' as string

        ql = ""
        indent = indent + 1
        set nNodes = topNode.selectNodes("*")
        for each nNode in nNodes
            sNodename = nNode.nodeName
            select case sNodename
                case "text"
                    sValue = nNode.getAttribute("class")
                    if isNull(sValue) then
                        sValue = ""
                    else
                        sValue = " " & sValue
                    end if
                    sCheck = nNode.getAttribute("check")
                    if not isNull(sCheck) then
                        sValue = sValue & " ? " & sCheck
                    end if

                    ql = ql & space(indent * m_spacing) & "---"
                    ql = ql & sValue
                    ql = ql & vbNewline
                    ql = ql & getText(nNode, indent) & vbNewline
                case "choice"
                    ql = ql & space(indent * m_spacing)
                    sValue = nNode.getAttribute("station")
                    if sValue = "back" then
                        ql = ql & "<--"
                    else
                        ql = ql & "--> " & sValue
                    end if

                    sValue = nNode.getAttribute("check")
                    if not isNull(sValue) then
                        ql = ql & " ? " & sValue
                    end if
                    ql = ql & vbNewline

                    ql = ql & getText(nNode, indent) & vbNewline
                case "choose"
                    ql = ql & space(indent * m_spacing) & "->>" & vbNewline
                    ql = ql & getText(nNode, indent) & vbNewline
                case "if"
                    ql = ql & space(indent * m_spacing) & "?" & vbNewline
                    ql = ql & getMainContent(nNode, indent) & vbNewline
                case "else"
                    ql = ql & space(indent * m_spacing) & "..." & vbNewline
                    ql = ql & getMainContent(nNode, indent)
                case "state", "number", "string"
                    ql = ql & space(indent * m_spacing) & getStateSymbol(sNodename) & " "
                    ql = ql & nNode.getAttribute("name") & " = "
                    ql = ql & nNode.getAttribute("value") & vbNewline
            end select
        next

        getMainContent = ql
    end function

    private function getStateSymbol(byVal sNodename)
        dim sState

        sState = ""
        select case sNodename
            case "state": sState = "_"
            case "number": sState = "%"
            case "string": sState = "$"
        end select

        getStateSymbol = sState
    end function

    private function getText(byRef topNode, byVal indent) ' as string
        dim ql ' as string
        dim i ' as integer
        dim nNode ' as node

        ql = ""

        indent = indent + 1
        for each nNode in topNode.childNodes
            select case nNode.nodeName
                case "emphasis"
                    ql = ql & " *" & getText(nNode, indent) & "* "
                case "strong"
                    ql = ql & " **" & getText(nNode, indent) & "** "
                case "#text"
                    ql = ql & space(indent * m_spacing) & nNode.data
                case "break"
                    ql = ql & "||"
                    if nNode.getAttribute("type") = "strong" then
                        ql = ql & "||"
                    end if
            end select
        next

        for i = 4 to 1 step - 1
            ql = replace( ql, space( (indent + i) * m_spacing ), space(indent * m_spacing) )
        next

        getText = ql
    end function

end class

%>