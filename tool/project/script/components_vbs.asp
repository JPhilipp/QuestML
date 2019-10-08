<script language="vbscript" runat="server">

function componentTestTable(value1, value2, value3)
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

</script>