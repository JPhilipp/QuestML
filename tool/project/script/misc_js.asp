<script language="JScript" runat="Server">

function getComponentJS(nameOf, valuesOf)
{
    var xhtml;
    var functionCall;

    functionCall = "xhtml = " + nameOf + "(" + valuesOf + ");";
    eval(functionCall);

    return xhtml;
}

function handleComponentJS(nameOf, valuesOf)
{
    var functionCall;

    functionCall = "xhtml = " + nameOf + "(" + valuesOf + ");";
    eval(functionCall);
}

</script>