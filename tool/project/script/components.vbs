' Use this file for your own VBScript-Functions to be called
' by QML components

' The content of this file serves sample purpose
' and can be deleted without risk, except that
' sample quest "component.htm" will not run anymore

option explicit

function componentTestTable(byVal value1, byVal value2, byVal value3)
    dim xhtml
    dim oXhtml
    dim i

    xhtml = ""
    xhtml = xhtml & "<table class=""specialTable"">"
    xhtml = xhtml & "<tr>" & _
            "<th class=""strongHeader"">Add</th>" & _
            "<th>Value 1</th><th>Value 2</th><th>Value 3</th>" & _
            "</tr>"
    for i = 0 to 10
        xhtml = xhtml & "<tr>" & _
                "<th class=""headerAdd"">+" & i & "</th>" & _
                "<td>" & value1 + i & "</td>" & _
                "<td>" & value2 + i & "</td>" & _
                "<td>" & value3 + i & "</td>" & _
                "</tr>"
    next
    xhtml = xhtml & "</table>"

    set oXhtml = getXmlString(xhtml)

    set componentTestTable = oXhtml
end function
