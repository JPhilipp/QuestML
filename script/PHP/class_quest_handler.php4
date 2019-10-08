<?

/* PHP-QuestHandler Version 1.0 */

class classQuestHandler
{
    var $stationId = "";
    var $questName = "";
    var $lastStation = "";
    var $debugMode = "";
    var $states = array();
    var $debugInfo = "";
    var $internalStateStart = "state_";
    var $internalSpace = "___";
    var $qmlVersion = "QML PHP 1.0";
    var $g_hasDied = false;

    function setByQuery()
    {
        global $HTTP_COOKIE_VARS;
        $station = '';
        $quest = '';
        $station = '';

        if ( isset($_POST['station']) )
        {
            $station = ( isset( $_POST['station'] ) ) ? $_POST['station'] : 'start';
            $quest = ( isset( $_POST['quest'] ) ) ? $_POST['quest'] : 'test';
            $lastStation = ( isset($_POST['lastStation']) ) ? $_POST['lastStation'] : '';
        }
        else
        {
            $station = ( isset( $_GET['station'] ) ) ? $_GET['station'] : 'start';
            $quest = ( isset( $_GET['quest'] ) ) ? $_GET['quest'] : 'test';
            $lastStation = ( isset($_GET['lastStation']) ) ? $_GET['lastStation'] : '';
        }

        $this->questName = $quest;
        $this->stationId = $station;
        $this->lastStation = $lastStation;

        if ($station == 'start') { setcookie($this->questName . '_died', '0'); }
        $this->hasDied = $HTTP_COOKIE_VARS[$this->questName . '_died'] == '1';
    }

    function process()
    {
        $xhtml = "";
        $outputStatus = false;

        if ($outputStatus) echo "ready<br />"; flush();

        $qml = $this->getFileText("quest/" . $this->questName . ".xml");
        $xPath = new XPath();
        $xPath->importFromString($qml);

        if ($outputStatus) echo "(1/5) did XPath class<br />"; flush();

        $result = $xPath->match("quest");
        $this->debugMode = ($xPath->getAttributes($result[0], "debug") == "true");

        $this->setStates();
        $this->setInternalStates($xPath);
        if ($outputStatus) echo "(2/5) did states<br />"; flush();

        $xhtmlContent = $this->getContent($xPath);
        $xhtmlContent = "<div id=\"content\">\n\n$xhtmlContent\n\n</div>";
        if ($outputStatus) echo "(3/5) did content<br />"; flush();

        $xhtmlHeader = $this->getHeader($xPath);
        if ($outputStatus) echo "(4/5) did header<br />"; flush();

        $xhtmlFooter = $this->getFooter();

        $xhtml = $xhtmlHeader . $xhtmlContent . $xhtmlFooter;
        if ($outputStatus) echo "(5/5) outputting now<br />"; flush();

        echo $xhtml;
    }

    function setStates()
    {
        $params = ( isset( $_POST ) ) ? $_POST : $_GET;
        foreach ($params as $key => $value)
        {
            $thisKey = str_replace($this->internalSpace, " ", $key);
            $lenStart = strlen($this->internalStateStart);
            $stringStart = substr($thisKey, 0, $lenStart);
            if ($stringStart == $this->internalStateStart)
            {
                $stringEnd = substr($thisKey, $lenStart);

                $thisValue = $value;
                if ($stringEnd == "qmlInput")
                {
                    $thisValue = str_replace("'", "", $thisValue);
                    $thisValue = "'" . $thisValue . "'";
                }
                $thisValue = str_replace("\\", "", $thisValue);
                $this->states[$stringEnd] = $thisValue;
            }
        }
    }

    function setInternalStates($xPath)
    {
        $this->states["qmlStation"] = $this->stationId;
        $this->states["qmlLastStation"] = $this->lastStation;
        $this->states["qmlVersion"] = $this->qmlVersion;

        $today = getdate(); 
        $this->states["qmlTime"] = $today["hours"] . ":" .
                $today["minutes"] . ":" . $today["seconds"];
        $this->states["qmlDay"] = $today["weekday"];
        $this->states["qmlServer"] = "true";
    }

    function getContent($xPath)
    {
        $oldThisStation = "";
        $xhtml = '';

        if ($this->stationId != 'start' && $this->hasDied) {
            $xhtml .= '<p>This quest has already ended.</p>';

            $result = $xPath->match("quest/about/homepage");
            $homepage = $xPath->getData($result[0]);
            if ($homepage != null) {
                $xhtml .= '<p><a href="' . $this->toAttribute($homepage) .'" class="qmlChoice">Home</a></p>';
            }
        }
        else {
            while ($oldThisStation != $this->stationId)
            {
                $oldThisStation = $this->stationId;
                $result = $xPath->match("quest/station[@id = '" . $this->stationId . "']");
                $xmlStation = $xPath->exportAsXml($result[0]);
                $xhtml = $this->getStation($xmlStation);
            }
    
            $xhtml = $this->addStationIncludes($xPath, $xhtml);
    
            if ( strpos($xhtml, '<form') === false ) {
                $twoDays = time() + 60 * 60 * 24 * 2;
                setcookie($this->questName . '_died', '1', $twoDays);
            }
        }

        return $xhtml;
    }

    function addStationIncludes($xPath, $xhtml)
    {
        $result = $xPath->match("quest/station/include");
        $countResult = count($result);
        for ($i = 0; $i < $countResult; $i++)
        {
            if ( $this->getNodeState($xPath, $result[0]) )
            {
                $resultIn = $xPath->match($result[$i] . "/in");
                if ( $this->getNodeState($xPath, $resultIn[0]) )
                {
                    $includeIn = $xPath->getAttributes($resultIn[0], "station");
                    if ( $this->getIncludeInMatches($includeIn, $this->stationId) )
                    {
                        $sXPath = $resultIn[0] . "/../../.";
                        $thisResult = $xPath->match($sXPath);
                        $xmlStation = $xPath->exportAsXml($thisResult[0]);
                        $includeXhtml = $this->getStation($xmlStation);
                        $xhtml = $includeXhtml . "\n" . $xhtml;
                    }
                }
            }
        }

        return $xhtml;
    }

    function getIncludeInMatches($includeIn, $stationId)
    {
        $includeInMatch = "";
        if ($includeIn != "")
        {
            $includeInMatch = ($includeIn == "*" || $includeIn == $stationId);
            if (!$includeInMatch)
            {
                $regex = "/" . $includeIn . "/";
                $regex = str_replace("*", ".*", $regex);
                $includeInMatch = preg_match($regex, $stationId);
            }
        }

        return $includeInMatch;
    }

    function getStation($xmlStation)
    {
        $xhtml = "";

        $xPath = new XPath();
        $xPath->importFromString($xmlStation);

        $context = "station";
        $xhtml .= $this->getStationMainContent($xPath, $context);

        $newContext = $this->handleIfElse($xPath, $context);
        if ($newContext != $context)
            $xhtml .= $this->getStationMainContent($xPath, $newContext);

        $xhtml .= $this->getChoices($xPath, $newContext);

        $xhtml = $this->getReplacedText($xhtml);

        return $xhtml;
    }

    function handleIfElse($xPath, $context)
    {
        $newContext = "";
        $result = $xPath->match("$context/if");
        $countResult = count($result);
        if ($countResult > 0)
        {
            for ($i = 0; $i < $countResult && $newContext == ""; $i++)
            {
                if ( $this->getNodeState($xPath, $result[$i]) )
                {
                    $this->debugInfo .= "ifResult=" . $result[$i] . "<br />";
                    $newContext = $result[$i];
                }
            }
            if ($newContext == "")
                $newContext = "$context/else";
        }
        else
            $newContext = $context;

        return $newContext;
    }

    function getStationMainContent($xPath, $context)
    {
        $xhtml = "";
        $sXPath = "$context/*";
        $result = $xPath->match($sXPath);

        $countResult = count($result);
        for ($i = 0; $i < $countResult; $i++)
        {
            $nodeName = $xPath->nodeName($result[$i]);
            switch ($nodeName)
            {
                case "choose":
                    $this->stationId = $this->getReplacedStates(
                            $xPath->getAttributes($result[$i], "station") );
                    break;
                case "text":
                    if ( $this->getNodeState($xPath, $result[$i]) )
                    {
                        $thisText = $this->getReplacedStates(
                                $xPath->exportAsXml($result[$i]) );
                        $thisText = $this->getInnerXml($thisText);
                        $sClass = $xPath->getAttributes($result[$i], "class");
                        if ($sClass == "")
                            $sClass = "qmlText";
                        $sClass = " class=\"" . $sClass . "\"";
                        $xhtml .= "<span" . $sClass . ">" . $thisText . "</span>\n";
                    }
                    break;

                case "table":
                    if ( $this->getNodeState($xPath, $result[$i]) )
                    {
                        $xhtml .= $this->getReplacedStates(
                                $xPath->exportAsXml($result[$i]) );
                    }
                    break;
                case "comment":
                    $this->debugInfo .= "\n<div class=\"comment\"><em>Comment:</em> " .
                            $this->getReplacedStates( $xPath->getData($result[$i]) ) .
                            "</div>\n";
                    break;
                case "embed":
                    if ( $this->getNodeState($xPath, $result[$i]) )
                    {
                        $source = $this->getReplacedStates(
                                $xPath->getAttributes($result[$i], "source") );
                        $xhtml .= "<iframe class=\"qmlEmbed\" src=\"$source\"></iframe>";
                    }
                    break;
                case "image":
                    if ( $this->getNodeState($xPath, $result[$i]) )
                    {
                        $sSource = $this->getReplacedStates(
                                $xPath->getAttributes($result[$i], "source") );
                        $sAlt = $this->getReplacedStates(
                                $xPath->getAttributes($result[$i], "text") );
                        $sClass = $xPath->getAttributes($result[$i], "class");
                        if ($sClass == "")
                            $sClass = "qmlImage";
                        $xhtml .= "<p><img src=\"" . $sSource . "\"" .
                                " alt=\"" . $sAlt . "\" class=\"" . $sClass . "\" />" .
                                "</p>\n";
                    }
                    break;

                case "state":
                case "number":
                case "string":
                    $stateName = $xPath->getAttributes($result[$i], "name");
                    $stateValue = $xPath->getAttributes($result[$i], "value");

                    if ($nodeName == "state")
                    {
                        if ($stateValue == "")
                            $stateValue = "true";
                    }
                    else if ($nodeName == "string")
                        $stateValue = "'" . $stateValue . "'";

                    $stateValue = $this->getNodeStateByValue($stateValue);

                    if ($nodeName == "state")
                        $stateValue = ($stateValue == 0)
                                ? "false" : "true";

                    $this->states[$stateName] = $stateValue;

                    break;
            }
        }

        return $xhtml;
    }

    function getInnerXml($text)
    {
        $firstCloser = strPos($text, ">");
        $lastOpener = strRPos($text, "<");
        $text = subStr($text, $firstCloser + 1, $lastOpener - $firstCloser - 1);
        return $text;
    }

    function getChoices($xPath, $context)
    {
        global $HTTP_SERVER_VARS;
        global $SERVER_NAME;
        global $PHP_SELF;

        $xhtml = "";

        $result = $xPath->match("$context/*");

        $path = "http" . ($_SERVER["HTTPS"] == "on" ? "s" : "") .
                "://" . $_SERVER["SERVER_NAME"] . strRev( strStr( strRev($_SERVER["PHP_SELF"]) , "/" ) );
        $formStart = "action=\"./\" method=\"post\">\n";

        foreach ($this->states as $key => $value) {
            $formStart .= $this->getHidden($this->internalStateStart . $key, $value);
        }

        $countResult = count($result);
        for ($i = 0; $i < $countResult; $i++)
        {
            $nodeName = $xPath->nodeName($result[$i]);
            switch ($nodeName)
            {
                case "choice":
                case "input":
                    if ( $this->getNodeState($xPath, $result[$i]) )
                    {
                        $guid = $this->getGuid();
                        $formName = 'f' . $this->toName($guid);
                        $submitId = 's' . $this->toName($guid);
                        $xhtml .= "\n\n<form name=\"" . $formName . "\" ";

                        $thisClass = $xPath->getAttributes($result[$i], "class");
                        if ($thisClass == "") $thisClass = "qmlChoice";
                        // $xhtml .= " class=\"$thisClass\" ";

                        $xhtml .= $formStart;

                        if ($nodeName == "input")
                        {
                            $name = $xPath->getAttributes($result[$i], "name");
                            if ($name == "") $name = "state_qmlInput";
                            $xhtml .= "<input type=\"text\" name=\"$name\" />";
                        }

                        $xhtml .= $this->getChoiceStates($xPath, $result[$i]);
        
                        $thisStation = $xPath->getAttributes($result[$i], "station");
                        $thisStation = $this->getReplacedStates($thisStation);
                        if ($thisStation == "back")
                            $thisStation = $this->lastStation;
            
                        $xhtml .= $this->getHidden("quest", $this->questName);
                        $xhtml .= $this->getHidden("station", $thisStation);
                        $xhtml .= $this->getHidden("lastStation", $this->stationId);

                        $xhtml .= "    <input type=\"submit\" id=\"" . $this->toAttribute($submitId) . "\" class=\"submit\" " .
                                "value=\"" . $this->toAttribute( $xPath->getData($result[$i]) ) . "\" />\r";

                        $xhtml .= '    <script type="text/javascript"><!--' . "\n";
                        $xhtml .= '        document.getElementById("' . $submitId . '").style.display = "none";' . "\r";
                        $xhtml .= '        document.write("<a class=\"' . $thisClass . '\" href=\"javascript:document.forms.' . $formName . '.submit()\">' .
                                $this->toAttribute( $xPath->getData($result[$i]) ) . '</a>")' . "\r";
                        $xhtml .= "\n" . '    //--></script>' . "\r";
            
                        $xhtml .= "</form>";
                    }
                    break;
            }
        }

        return $xhtml;
    }

    function getChoiceStates($xPath, $context)
    {
        $xhtml = "";

        $sXPath = "$context/*";
        $result = $xPath->match($sXPath);
        $countResult = count($result);
        for ($i = 0; $i < $countResult; $i++)
        {
            $nodeName = $xPath->nodeName($result[$i]);
            switch ($nodeName)
            {
                case "state":
                case "number":
                case "string":
                    $stateName = $xPath->getAttributes($result[$i], "name");
                    $stateValue = $xPath->getAttributes($result[$i], "value");
            
                    if ($nodeName == "state" && $stateValue == "") $stateValue = "true";
                    if ($nodeName == "string") $stateValue = "'" . $stateValue . "'";

                    $stateValue = $this->getNodeStateByValue($stateValue);
        
                    $xhtml .= $this->getHidden($this->internalStateStart . $stateName, $stateValue);
                    break;
            }
        }

        return $xhtml;
    }

    function getReplacedStates($xmlStation)
    {
        foreach ($this->states as $key => $value)
            $xmlStation = str_replace("[" . $key . "]", $value, $xmlStation);

        $xmlStation = preg_replace("%\[[^\]]*\]%", "0", $xmlStation);

        return $xmlStation;
    }

    function getReplacedText($xhtml)
    {
        $xhtml = str_replace("<display", "<br /><strong", $xhtml);
        $xhtml = str_replace("</display>", "</strong>", $xhtml);
        $xhtml = str_replace("<poem", "<pre", $xhtml);
        $xhtml = str_replace("</poem>", "</pre>", $xhtml);
        $xhtml = str_replace("<link to=", "<a class=\"qmlLink\" href=", $xhtml);
        $xhtml = str_replace("</link>", "</a>", $xhtml);
        $xhtml = str_replace(" target=\"", " target=\"_", $xhtml);
        $xhtml = str_replace("<emphasis", "<em", $xhtml);
        $xhtml = str_replace("</emphasis>", "</em>", $xhtml);
        $xhtml = str_replace("<break type=\"strong\"/>", "<br/><br/>", $xhtml);
        $xhtml = str_replace("<break/>", "<br/>", $xhtml);

        return $xhtml;
    }

    function getNodeState($xPath, $result)
    {
        $checkValue = $xPath->getAttributes($result, "check");
        return $this->getNodeStateByValue($checkValue);
    }

    function getNodeStateByValue($checkValue)
    {
        $state = true;

        $checkValue = $this->getReplacedStates($checkValue);
        $evalString = "";
        if ($checkValue != "")
        {
            $checkValue = str_replace("not ", "!", $checkValue);
            $checkValue = str_replace(" and ", " && ", $checkValue);
            $checkValue = str_replace(" equal", "=", $checkValue);
            $checkValue = str_replace(" greater", ">", $checkValue);
            $checkValue = str_replace(" lower", "<", $checkValue);
            $checkValue = str_replace("= >", "> =", $checkValue);
            $checkValue = str_replace("= <", "< =", $checkValue);
            $checkValue = str_replace("> =", ">=", $checkValue);
            $checkValue = str_replace("< =", "<=", $checkValue);
            $checkValue = str_replace("=>", ">=", $checkValue);
            $checkValue = str_replace("=<", "<=", $checkValue);
            $checkValue = str_replace("'", "\"", $checkValue);

            $checkValue = str_replace("=", "==", $checkValue);
            $checkValue = str_replace("!==", "!=", $checkValue);
            $checkValue = str_replace("<==", "<=", $checkValue);
            $checkValue = str_replace(">==", ">=", $checkValue);

            $checkValue = $this->replaceFunctions($checkValue);
            $checkValue = $this->removeForbiddenInput($checkValue);
            $evalString = "\$state = " . $checkValue . ";";

            eval($evalString);
        }

        return $state;
    }

    function removeForbiddenInput($evalString)
    {
        $forbidden = array(
            "unlink",
            "unset",
            "delete",
            "rmdir",
            "eval",
            "phpinfo",
            "chgrp",
            "chmod",
            "chown",
            "copy",
            "file",
            "file_exists",
            "filegroup",
            "fileowner",
            "flock",
            "fopen",
            "fputs",
            "fread",
            "fscanf",
            "fseek",
            "ftruncate",
            "fwrite",
            "set_file_buffer",
            "link",
            "mkdir",
            "readfile",
            "rename",
            "symlink",
            "tmpfile",
            "system",
            "exec",
            "passthru",
            "popen",
            "touch",
            "{",
            "}",
            ";",
            "\".\""
            );

        $countForbidden = count($forbidden);
        for ($i = 0; $i < $countForbidden; $i++)
        {
            $lowerEval = strToLower($evalString);
            $newEvalString = str_replace($forbidden[$i], "", $lowerEval);
            if ( $newEvalString != $lowerEval)
                $evalString = $newEvalString;
        }

        $newString = "";
        while ($newString != $evalString)
        {
            $newString = $evalString;
            $evalString = str_replace( "( ", "(", $evalString );
            $evalString = str_replace( ") ", ")", $evalString );
        }
        $evalString = str_replace( "()", "", $evalString );

        return $evalString;
    }

    function replaceFunctions($text)
    {
        $text = str_replace( "\"{", "{", $text );
        $text = str_replace( "}\"", "}", $text );
        $text = str_replace( "{random ", "\$this->qml_random(", $text );
        $text = str_replace( "{states ", "\$this->qml_states(", $text );
        $text = str_replace( "{contains ", "\$this->qml_contains(", $text );
        $text = str_replace( "{containsWord ", "\$this->qml_containsWord(", $text );
        $text = str_replace( "{verbose ", "\$this->qml_verbose(", $text );
        $text = str_replace( "{lower ", "\$this->qml_lower(", $text );
        $text = str_replace( "{upper ", "\$this->qml_upper(", $text );
        $text = str_replace( "{repeatString ", "\$this->qml_repeatString(", $text );
        $text = str_replace( "}", ")", $text );

        return $text;
    }

    function getHidden($thisName, $thisValue)
    {
        $formName = str_replace(" ", $this->internalSpace, $thisName);
        return "    <input type=\"hidden\" name=\"$formName\" value=\"$thisValue\" />\r";
    }

    function getHeader($xPath)
    {
        $xhtml = "";

        $result = $xPath->match("quest/about/title");
        $title = $xPath->getData($result[0]);
        $result = $xPath->match("quest/about/author");
        $author = $xPath->getData($result[0]);

        $xhtml .= "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"DTD/xhtml1-strict.dtd\">\r";
        $xhtml .= "<html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"en\" lang=\"en\">\r";
        $xhtml .= "<head>\r";
        $xhtml .= "    <title>$title by $author</title>\n";
        $xhtml .= "    <style><!--";
        $xhtml .= $this->getHeadStyle($xPath);
        $xhtml .= "\n    --></style>\n";
        $xhtml .= "\n    <meta name=\"expires\" content=\"0\" />\n";
        $xhtml .= "</head>\n";
        $xhtml .= "<body>\n\n";
        return $xhtml;
    }

    function getFooter()
    {
        $xhtml = "";

        if ($this->debugMode)
        {
            $xhtml .= "\n\n<div class=\"debugInfo\">";

            foreach ($this->states as $key => $value)
                $xhtml .= "$key=<em>$value</em> | \n";

            if ($this->debugInfo != "")
                $xhtml .= "<p><strong>Info</strong>:<br />\n" . $this->debugInfo . "</p>\n";

            $xhtml .= "</div>";
        }

        $xhtml .= "\r\r</body>\r";
        $xhtml .= "</html>\r";
        return $xhtml;
    }

    function getFileText($filePath)
    {
        $fileText = "";
    
        $fileArray = file($filePath);
        $countFile = count($fileArray);
        for ($i = 0; $i < $countFile; $i++)
            $fileText .= $fileArray[$i];
    
        return $fileText;
    }

    /* Style *****************************/

    function getHeadStyle($xPath)
    {
        $style = "";
        $backgroundColor = "";
        $backgroundImage = "";
        $backgroundRepeat = "";
        $fontColor = "";
        $fontFamily = "";
        $fontSize = "";
        $contentLeft = "";
        $contentTop = "";
        $contentWidth = "";
        $contentPosition = "";

        $style .= "\n.submit {\n background-color: transparent; color: inherit; border: 0;\n" .
                "    text-decoration: underline; cursor: pointer;\n" . 
                "    margin: 0; padding: 0; text-align: left; font-family: inherit;\n}\n";

        $result = $xPath->match("quest/style/background");
        if ($result)
        {
            $backgroundColor = $xPath->getAttributes($result[0], "color");
            $backgroundImage = $xPath->getAttributes($result[0], "image");
            $backgroundRepeat = $xPath->getAttributes($result[0], "repeat");

            if ($backgroundColor != "")
                $backgroundColor = "    background-color: $backgroundColor;\n";
            if ($backgroundImage != "")
                $backgroundImage = "    background-image: url($backgroundImage);\n";
            if ($backgroundRepeat != "")
                $backgroundRepeat = "    background-repeat: $backgroundRepeat;\n";
        }

        $result = $xPath->match("quest/style/font");
        if ($result)
        {
            $fontColor = $xPath->getAttributes($result[0], "color");
            $fontFamily = $xPath->getAttributes($result[0], "family");
            $fontSize = $xPath->getAttributes($result[0], "size");
            if ($fontColor != "")
                $fontColor = "    color: $fontColor;\n";
            if ($fontFamily != "")
            {
                $fontFamily = "    font-family: $fontFamily;\n";
                $fontFamily = str_replace("'", "\"", $fontFamily);
            }
            if ($fontSize != "")
                $fontSize = "    font-size: $fontSize;\n";
        }

        $result = $xPath->match("quest/style/content");
        if ($result)
        {
            $contentLeft = $xPath->getAttributes($result[0], "left");
            $contentTop = $xPath->getAttributes($result[0], "top");
            $contentWidth = $xPath->getAttributes($result[0], "width");

            if ($contentTop == 'default') { $contentTop = ''; }
            if ($contentLeft == 'default') { $contentLeft = ''; }

            else if ($contentLeft != "") {
                $contentLeft = "    left: $contentLeft;\n";
            }

            if ($contentTop != "") {
                $contentTop = "    top: $contentTop;\n";
            }

            if ($contentWidth != "" && $contentWidth != "default") {
                $contentWidth = "    width: $contentWidth;\n";
            }
        }

        $contentPosition = ($contentLeft != "" || $contentTop != "") ?
                "    position: relative;\n" : "";

        $style .= "body {\n" .
                $backgroundColor . $backgroundImage . $backgroundRepeat .
                $fontColor . $fontFamily . $fontSize .
                "}\n";
        $style .= "#content {\n" .
                $contentPosition . $contentLeft . $contentTop . $contentWidth .
                "}\n";

        $style .= $this->getStyleClasses($xPath);

        if ($this->debugMode)
            $style .= ".debugInfo { margin-top: 16px; padding: 8px; background-color: #eee; ".
                    "color: #000; font-family: arial, sans-serif; border: 1px solid black;" .
                    "width: 500px; }";

        $style = str_replace("\n", "\n        ", $style);

        $style = $this->getReplacedStates($style);
        return $style;
    }

    function getStyleClasses($xPath)
    {
        $style = "";

        $result = $xPath->match("quest/style/class");
        $countResult = count($result);
        for ($i = 0; $i < $countResult; $i++)
        {
            $className = $xPath->getAttributes($result[$i], "name");
            $classStyle = $this->getStyleOfClass($xPath, $className);
            $style .= ".$className { $classStyle }\n";
        }

        return $style;
    }

    function getStyleOfClass($xPath, $className)
    {
        $style = "";

        $result = $xPath->match("quest/style/class[@name = '" . $className . "']");
        if ($result)
        {
            $style = $xPath->getAttributes($result[0], "style");
    
            $classInherits = $xPath->getAttributes($result[0], "inherits");
            if ($classInherits != "")
                $style .= ";" . $this->getStyleOfClass($xPath, $classInherits);
    
            $style = str_replace(";;", "; ", $style);
        }
        return $style;
    }

    /* QML internal functions ************/

    function qml_random($min, $max)
    {
        return rand($min, $max);
    }

    function qml_states($startsWith)
    {
        $text = "";

        $lenStartsWith = strLen($startsWith);
        foreach ($this->states as $key => $value)
        {
            $stringStart = substr($key, 0, $lenStartsWith);
            if ($stringStart == $startsWith)
            {
                $stringEnd = substr($key, $lenStartsWith);
                if ($text != "") $text .= ", ";
                $text .= $stringEnd;
            }
        }

        return $text;
    }

    function qml_contains($sString, $sSubstring)
    {
        $returnValue = "false";
        if ($sString != "" && $sSubstring != "")
        {
            $sString = strToLower($sString);
            $sSubstring = strToLower($sSubstring);
            $i = strpos($sString, $sSubstring);
            $returnValue = ($i === false) ? "false" : "true";
        }

        return $returnValue;
    }

    function qml_containsWord($sString, $sSubstring)
    {
        // ?! preg_match('/\b" . $sSubstring . "\b/i', $sString);

        $didFind = false;
        if ($sString != "" && $sSubstring != "")
        {
            $sString = strToLower($sString);
            $sSubstring = strToLower($sSubstring);
        
            $words = split('[/.-\,]():; ', $sString);
            $countWords = count($words);
            for ($i = 0; $i < $countWords && !$didFind; $i++)
            {
                // echo $words[$i] . "-<br />";
                $didFind = ($words[$i] == $sSubstring);
            }
        }
        
        $returnValue = ($didFind) ? "true" : "false";
        return $returnValue;
    }

    function qml_verbose($n)
    {
        $n = "$n";
        switch ($n)
        {
            case "11":
            case "12":
            case "13":
                $n .= "th";
                break;
            default:

                switch( substr($n, -1) )
                {
                    case "1":
                        $n .= "st";
                        break;
                    case "2":
                        $n .= "nd";
                        break;
                    case "3":
                        $n .= "rd";
                        break;
                    case "4":
                        $n .= "th";
                        break;
                }
        }
    
        return $n;
    }

    function qml_lower($text)
    {
        return strToLower($text);
    }

    function qml_upper($text)
    {
        return strToUpper($text);
    }

    function qml_repeatString($text, $max)
    {
        $newText = "";
        if ($max >= 1 && $max <= 10000)
        {
            for ($i = 0; $i < $max; $i++)
                $newText .= $text;
        }
        return $newText;
    }

    function toAttribute($s)
    {
        $s = $this->toXml($s);
        $s = str_replace('"', '&quot;', $s);
        return $s;
    }

    function toXml($s)
    {
        $s = str_replace('&', '&amp;', $s);
        $s = str_replace('<', '&lt;', $s);
        $s = str_replace('>', '&gt;', $s);
        return $s;
    }

    function toName($s)
    {
        $s = strtolower($s);
        $abc = 'abcdefghijklmnopqrstuvwxyz0123456789';
        $sNew = '';
    
        for ($i = 0; $i < strlen($s); $i++)
        {
            $letter = substr($s, $i, 1);
            if ( strpos($abc, $letter) === false )
            {
                $letter = '';
            }
    
            $sNew .= $letter;
        }
    
        return $sNew;
    }
    
    function getGuid()
    {
        mt_srand( (double) microtime() * 10000 );
        $charid = strtoupper(md5(uniqid(rand(), true)));
        $hyphen = chr(45);
        $uuid = substr($charid, 0, 8) . $hyphen
               . substr($charid, 8, 4) . $hyphen
               . substr($charid,12, 4) . $hyphen
               . substr($charid,16, 4) . $hyphen
               . substr($charid,20,12);
    
        return $uuid;
    }
    
}

?>