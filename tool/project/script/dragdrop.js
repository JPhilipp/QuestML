var currentX, currentY, whichEl, curPos;
currentX = currentY = null;
grabbedElm = null;

function grabEl(elm)
{
    grabbedElm = elm;
    event.cancelBubble = true;
    grabbedElm.style.pixelLeft = grabbedElm.offsetLeft;
    grabbedElm.style.pixelTop = grabbedElm.offsetTop;
    currentX = (event.clientX + document.body.scrollLeft);
    currentY = (event.clientY + document.body.scrollTop); 
    document.onmousemove = moveEl;
    document.onmouseup = dropEl;
    event.returnValue = false;
}

function moveEl()
{
    newX = (event.clientX + document.body.scrollLeft);
    newY = (event.clientY + document.body.scrollTop);
    distanceX = (newX - currentX);
    distanceY = (newY - currentY);
    currentX = newX;
    currentY = newY;
    if (grabbedElm.style.width != "100%")
    {
        grabbedElm.style.pixelLeft += distanceX;
        grabbedElm.style.pixelTop += distanceY;
    }
    event.returnValue = false;
}

function dropEl()
{
   document.onmousemove = document.onmouseup = null;
}


