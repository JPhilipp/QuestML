<%@ language="vbscript" %>
<%
    option explicit
    const g_isServerVersion = true
%>
<!-- #include file="misc.asp" -->
<!-- #include file="class_ql.asp" -->
<!-- #include file="class_qml.asp" -->
<%

main

sub main
    dim questName
    dim output
    dim sText
    dim preview

    questName = request.queryString("quest")
    output = request.queryString("output")
    preview = ( request.queryString("preview") = "true" )
    if questName = "" then
        questName = "simple"
    end if
    if output = "" then
        output = "qml"
    end if

    if output = "qml" then
        sText = getQml(questName, preview)
    else ' if output = "ql" then
        sText = getQl(questName, preview)
    end if

    response.write sText
end sub

function getQml(byVal questName, byVal textPreview)
    dim oQl
    dim qml


    set oQl = new QL
    oQl.setSource( getFileText("../quest/" & questName & ".ql") )
    qml = oQl.getQml

    if textPreview then
        qml = xmlToText(qml)
        qml = "<pre>" & qml & "</pre>"
        qml = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""DTD/xhtml1-strict.dtd"">" & vbNewline & _
                "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">" & vbNewline & _
                "<head>" & vbNewline & _
                "    <title>QL to QML conversion</title>" & vbNewline & _
                "</head>" & vbNewline & _
                "<body>" & vbNewline & _
                qml & vbNewline & _
                "</body>" & vbNewline & _
                "</html>"
    else
        response.contentType = "text/xml"
    end if

    getQml = qml
end function

function getQl(byVal questName, byVal textPreview)
    dim oQml
    dim ql
    dim oXml

    set oQml = new Qml
    
    set oXml = getXml("../quest/" & questName & ".xml")
    oQml.setSource oXml
    ql = oQml.getQl

    if textPreview then
        ql = "<pre>" & ql & "</pre>"
        ql = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""DTD/xhtml1-strict.dtd"">" & vbNewline & _
                "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"" lang=""en"">" & vbNewline & _
                "<head>" & vbNewline & _
                "    <title>QL to QML conversion</title>" & vbNewline & _
                "</head>" & vbNewline & _
                "<body>" & vbNewline & _
                ql & vbNewline & _
                "</body>" & vbNewline & _
                "</html>"
    else
        response.contentType = "text/plain"
    end if

    getQl = ql
end function

%>