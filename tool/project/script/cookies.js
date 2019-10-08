function getCookieVal(offset)
{
    var endstr = document.cookie.indexOf (";", offset);
    if (endstr == -1)
        endstr= document.cookie.length;
    return unescape( document.cookie.substring(offset, endstr) );
}

function GetCookie(name)
{
    var arg = name + "=";
    var alen = arg.length;
    var clen = document.cookie.length;
    var i = 0;
    while (i < clen)
    {
        var j= i + alen;
        if (document.cookie.substring(i, j) == arg)
            return getCookieVal(j);
        i = document.cookie.indexOf(" ", i) + 1;
        if (i==0)
            break;
    }
    return null;
}

function SetCookie(name, value)
{
    var expDate = "";
    var oneYear = 265 * 24 * 60 * 60 * 1000;
    var expDate = new Date();
    expDate.setTime(expDate.getTime() + oneYear);
    document.cookie = name + "=" + escape(value) + "; expires=" + expDate.toGMTString();
}
