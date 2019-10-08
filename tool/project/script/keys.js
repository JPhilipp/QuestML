document.onkeypress = doKey;

function doKey(e)
{
    var number = 0, whichASC = 0, whichKey = "";
    whichASC = event.keyCode;
    whichKey = String.fromCharCode(whichASC).toLowerCase();
    if (whichKey == "s")
        saveGame();
    else if (whichKey == "l")
        loadGame();
    else if (whichKey == " ")
        toggleDisplayStates();
}
