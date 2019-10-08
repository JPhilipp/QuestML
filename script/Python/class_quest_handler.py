# Python QML-Interpreter 0.942

import xml.dom.minidom
from xml.dom.minidom import *
import cgi
import string
import re
import random

class QuestHandler:

    states = {}

    def getOutput(self):
        xhtml = ""

        self.storeStates()
        random.seed()
        pathXml = "./qml/quest/" + self.states['quest'] + ".xml"
        qmlDoc = xml.dom.minidom.parse(pathXml)

        content = self.getContent(qmlDoc)
        header = self.getHeader(qmlDoc)
        footer = self.getFooter()

        xhtml = header + content + footer
        xhtml = xhtml.encode("latin-1")
        return xhtml

    def storeStates(self):
        form = cgi.FieldStorage()
        self.states['quest'] = 'test'
        self.states['station'] = 'start'
        if form:
            for key in form.keys():
                if not key == 'submitForm':
                    s = self.decodeField(form[key])
                    thisKey = key
                    if string.find(thisKey, 'qmlInput') == 0:
                        s = "'" + s + "'"
                        if thisKey != 'qmlInput':
                            thisKey = string.replace(thisKey, 'qmlInput', '')
                    self.states[thisKey] = s

    def decodeField(self, field):
       if isinstance( field, type([]) ):
           return map(self.decodeField, field)
       elif hasattr(field, "file") and field.file:
           return (field.filename, field.file)
       else:
           return field.value

    def getContent(self, qmlDoc):
        xhtml = ''
        stationNode = self.getElementById(qmlDoc.documentElement, self.states['station'])
        xhtml += '<div class="content">\n'

        mainContent = self.getMainContent(stationNode, qmlDoc)
        mainContent = self.getIncludeContent(stationNode, qmlDoc, mainContent)

        xhtml += mainContent
        xhtml += '</div>\n'
        xhtml = string.replace(xhtml, '<br></br>', '<br />')
        return xhtml

    def getIncludeContent(self, stationNode, qmlDoc, mainContent):
        targetId = stationNode.getAttribute('id')
        for stationNode in qmlDoc.documentElement.getElementsByTagName('station'):
            mainContent = self.getIncludeContentOf(stationNode, targetId, mainContent, qmlDoc)

        return mainContent

    def getIncludeContentOf(self, stationNode, targetNeeded, mainContent, qmlDoc):
        for includeNode in stationNode.getElementsByTagName('include'):
            for inNode in includeNode.getElementsByTagName('in'):
                target = inNode.getAttribute('station')
                isTarget = (target == targetNeeded or target == '*' )

                if not isTarget:
                    regex = string.replace(target, '*', '.*')
                    pattern = re.compile(regex)
                    isTarget = re.search(pattern, targetNeeded)

                if isTarget:
                    includeContent = self.getMainContent(stationNode, qmlDoc)
                    process = inNode.getAttribute('process')
                    if (not process) or process == '':
                        process = 'before'
                    if process == 'before':
                        mainContent = includeContent + mainContent
                    elif process == 'after':
                        mainContent += includeContent
                    else: # if process == 'exclusive':
                        mainContent = includeContent

        return mainContent

    def getMainContent(self, topNode, qmlDoc):
        xhtml = ''

        nodes = topNode.childNodes
        for node in nodes:

            sClass = ''

            thisName = node.nodeName
            if thisName == 'if':
                if self.checkElement(node):
                    xhtml += self.getMainContent(node, qmlDoc)
                    break
            elif thisName == 'else':
                xhtml += self.getMainContent(node, qmlDoc)
            if thisName == 'text':
                classValue = node.getAttribute('class')
                if classValue:
                    sClass += ' class="' + classValue + '"'
                if self.checkElement(node):
                    xhtml += '<span' + sClass + '>' + self.getNewInnerXml(node) + '</span>\n'
            elif thisName == 'image':
                classValue = node.getAttribute('class')
                if classValue:
                    sClass += ' class="' + classValue + '"'
                else:
                    sClass = ' class="qmlImage"'
                if self.checkElement(node):
                    imageSource = node.getAttribute('source')
                    altValue = node.getAttribute('text')
                    xhtml += '<div' + sClass + '><img src="/qml/' + imageSource + '" alt="' + altValue + '" /></div>\n'
            elif thisName == 'state' or thisName == 'number' or thisName == 'string':
                stateName = node.getAttribute('name')
                stateValue = node.getAttribute('value')
                if thisName == 'state' and stateValue == '':
                    stateValue = '1'
                if thisName == 'string':
                    stateValue = "'" + stateValue + "'"
                stateValue = self.getEval(stateValue)
                if thisName == 'string':
                    stateValue = "'" + str(stateValue) + "'"
                self.states[stateName] = str(stateValue)
            elif thisName == 'embed':
                xhtml += '<iframe src="' + node.getAttribute('source') + '"></iframe>'
            elif thisName == 'choose':
                self.states['station'] = node.getAttribute("station")
                stationNode = self.getElementById(qmlDoc.documentElement, self.states['station'])
                xhtml += self.getMainContent(stationNode, qmlDoc)
            elif thisName == 'choice' or thisName == 'input':
                if self.checkElement(node):
                    xhtml += self.getChoice(node, thisName)

        return xhtml

    def getChoice(self, node, thisName):
        xhtml = ''
        sClass = ''
        classValue = node.getAttribute('class')
        if classValue:
            sClass += ' class="' + classValue + '"'
        innerXml = self.getNewInnerXml(node)
        innerXml = string.replace(innerXml, '"', '&quot;')
        xhtml += '<form action="/cgi-bin/index.py" method="post"' + sClass +'><div>\n'
        xhtml += self.getHiddenStates()
        xhtml += '<input type="hidden" name="station" value="' + node.getAttribute("station") + '" />\n'
        if thisName == 'input':
            xhtml += self.getInput(node)
        xhtml += '<input type="submit" name="submitForm" value="' + innerXml + '" class="submit" />\n'
        xhtml += '</div></form>\n'
        return xhtml

    def getInput(self, node):
        xhtml = ''
        inputName = 'qmlInput' + node.getAttribute('name')
        xhtml += '<input type="text" name="' + inputName + '" /> '
        return xhtml

    def checkElement(self, element):
        evalString = element.getAttribute('check')
        returnValue = self.getEval(evalString)
        return returnValue

    def getEval(self, evalString):
        returnValue = 1
        if evalString != '':
            evalString = self.replaceStates(evalString, 0)
            returnValue = self.safeEval(evalString)

        return returnValue

    def safeEval(self, s):
        s = string.replace(s, "__", "")
        s = string.replace(s, "file", "")
        s = string.replace(s, "eval", "")
        # return rexec.RExec.r_eval(rexec.RExec(), s)
        return eval(s)

    def replaceStates(self, s, cutQuotation):
        for key in self.states.keys():
            keyValue = self.states[key]
            if cutQuotation:
                quotationLeft = string.find(keyValue, "'") == 0
                quotationRight = string.rfind(keyValue, "'") == len(keyValue) - 1
                if quotationLeft and quotationRight:
                    keyValue = keyValue[1:-1]
            s = string.replace(s, '[' + key + ']', keyValue)

        s = string.replace(s, 'true', '1')
        s = string.replace(s, 'false', '0')

        pattern = re.compile(r'\[.*?\]')
        s = pattern.sub('0', s)

        s = string.replace(s, ' lower ', ' < ')
        s = string.replace(s, ' greater ', ' > ')
        s = string.replace(s, ' = ', ' == ')

        s = string.replace(s, '\'{', '{')
        s = string.replace(s, '}\'', '}')

        s = string.replace(s, '{states ', 'self.qml_states(')
        s = string.replace(s, '{random ', 'self.qml_random(')
        s = string.replace(s, '{lower ', 'self.qml_lower(')
        s = string.replace(s, '{upper ', 'self.qml_upper(')
        s = string.replace(s, '{contains ', 'self.qml_contains(')
        s = string.replace(s, '}', ')')

        return s

    def getHiddenStates(self):
        xhtml = ''
        for key in self.states.keys():
            if key != 'lastStation':
                thisName = key
                if key == 'station':
                    thisName = 'lastStation'
                xhtml += '<input type="hidden" name="' + thisName + '" '
                xhtml += 'value="' + str(self.states[key]) + '" />\n'
        return xhtml

    def getNewInnerXml(self, topNode):
        s = ''

        for subNode in topNode.childNodes:
            if subNode.nodeType == Node.TEXT_NODE:
                s += self.replaceStates(subNode.data, 1)
            elif subNode.nodeType == Node.ELEMENT_NODE:
                newName = self.getNewElementName(subNode)
                classValue = subNode.getAttribute('class')
                if classValue and classValue != '':
                    classValue = ' class="' + classValue + '"'
                innerXml = self.getNewInnerXml(subNode)
                s += '<' + newName + classValue +'>' + innerXml
                s += '</' + newName + '>'

        return s

    def getNewElementName(self, node):
        oldName = node.nodeName

        if oldName == 'emphasis':
            newName = 'em'
        elif oldName == 'poem':
            newName = 'pre'
        elif oldName == 'image':
            newName = 'img'
        elif oldName == 'break':
            newName = 'br'
        else:
            newName = oldName

        return newName

    # general XML

    def getElementById(self, topNode, targetId, idName = 'id'):
        foundNode = 0
        nodes = topNode.childNodes
        for node in nodes:
            if node.nodeType == Node.ELEMENT_NODE:
                thisId = node.getAttribute(idName)
                if thisId == targetId:
                    foundNode = node
                    break

        return foundNode

    # more
    
    def getHeader(self, qmlDoc):
        xhtml = ''
        xhtml += """\
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "DTD/xhtml1-strict.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
            <title>QML</title>
        """
        xhtml += self.getHeaderStyle(qmlDoc)

        xhtml += """\
        </head>
        <body>
        """
        return xhtml

    def getHeaderStyle(self, qmlDoc):
        xhtml = ''
        xhtml += '\n<style media="screen" type="text/css"><!--\n'
        xhtml += '.submit\n'
        xhtml += '{\n'
        xhtml += 'background-color: transparent;\n'
        xhtml += 'color: inherit;\n'
        xhtml += 'border: 0;\n'
        xhtml += 'text-decoration: underline;\n'
        xhtml += 'cursor: pointer;\n'
        xhtml += 'margin: 0;\n'
        xhtml += 'padding: 0;\n'
        xhtml += 'text-align: left;\n'
        xhtml += 'font-family: inherit;\n'
        xhtml += '}\n'
        xhtml += self.getQmlStyle(qmlDoc)
        xhtml += '--></style>\n'

        xhtml = string.replace(xhtml, '\n', '\n        ')
        xhtml += '\n'
        return xhtml

    def getQmlStyle(self, qmlDoc):
        css = ''

        backgroundColor = ''
        backgroundImage = ''
        backgroundRepeat = ''
        fontColor = ''
        fontFamily = ''
        fontSize = ''
        contentLeft = ''
        contentTop = ''
        contentWidth = ''
        contentPosition = ''
        classStyle = ''

        topNodes = qmlDoc.documentElement.getElementsByTagName('style')
        if topNodes:
            for node in topNodes[0].childNodes:
                thisName = node.nodeName
                if thisName == 'background':
                    backgroundColor = node.getAttribute('color')
                    backgroundImage = node.getAttribute('image')
                    backgroundRepeat = node.getAttribute('repeat')
                elif thisName == 'font':
                    fontColor = node.getAttribute('color')
                    fontFamily = node.getAttribute('family')
                    fontSize = node.getAttribute('size')
                elif thisName == 'content':
                    contentLeft = node.getAttribute('left')
                    contentTop = node.getAttribute('top')
                    contentWidth = node.getAttribute('width')
                elif thisName == 'class':
                    classStyle += self.getClassStyle(node, topNodes[0])

        if backgroundColor != '':
            backgroundColor = 'background-color: ' + backgroundColor + ';\n'
        if backgroundImage != '':
            backgroundImage = 'background-image: url(/qml/' + backgroundImage + ');\n'
        if backgroundRepeat != '':
            backgroundRepeat = 'background-repeat: ' + backgroundRepeat + ';\n'
        if fontColor != '':
            fontColor = 'color: ' + fontColor + ';\n'
        if fontFamily != '':
            fontFamily = 'font-family: ' + fontFamily + ';\n'
        if fontSize != '':
            fontSize = 'font-size: ' + fontSize + ';\n'
        if contentLeft != '':
            contentLeft = 'left: ' + contentLeft + ';\n'
        if contentTop != '':
            contentTop = 'top: ' + contentTop + ';\n'
        if contentLeft != '' or contentTop != '':
            contentPosition = 'position: absolute;\n'
        if contentWidth != '':
            contentWidth = 'width: ' + contentWidth + ';\n'

        css += 'body\n{\n'
        css += backgroundColor
        css += backgroundImage
        css += backgroundRepeat
        css += fontColor
        css += fontFamily
        css += fontSize
        css += '}\n'

        css += '.content\n{\n'
        css += contentPosition
        css += contentLeft
        css += contentTop
        css += contentWidth
        css += '}\n'

        css += classStyle

        return css

    def getClassStyle(self, node, topNode):
        css = ''

        sName = node.getAttribute('name')
        css += '.' + sName + '\n{\n'
        css += self.getClassStyleContent(topNode, sName)
        css += '}\n'

        return css

    def getClassStyleContent(self, topNode, wantedName):
        css = ''

        for node in topNode.childNodes:
            if node.nodeName == 'class':
                name = node.getAttribute('name')
                if name == wantedName:
                    sInherits = node.getAttribute('inherits')
                    if sInherits and sInherits != '':
                        css += self.getClassStyleContent(topNode, sInherits)
                    thisCss = node.getAttribute('style')
                    thisCss = self.replaceStates(thisCss, 1)
                    thisCss = string.replace(thisCss, "'", '')
                    thisCss = string.replace(thisCss, 'url(', 'url(/qml/')
                    css += thisCss + '\n'

        return css

    def getFooter(self):
        xhtml = """\

        </body>
        </html>
        """
        return xhtml

    # QML Inline functions

    def qml_states(self, wanted):
        s = ''
        seperator = ', '
        for key in self.states.keys():
            if self.states[key]:
                if string.find(key, wanted) == 0:
                    s += key[len(wanted) + 1:] + seperator
        if len(s) > 0:
            s = s[:-len(seperator)]
        return s

    def qml_random(self, min, max):
        return random.randrange(min, max)

    def qml_lower(self, s):
        return string.lower(s)

    def qml_upper(self, s):
        return string.upper(s)

    def qml_contains(self, sAll, sSub):
        return ( string.find(sAll, sSub) != -1 )
