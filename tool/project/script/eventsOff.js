function turnOffEvents()
{
    document.oncontextmenu = function () { return false; }
    document.onselectstart = function () { return false; }
}
