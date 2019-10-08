<?

// PHP5-enabled XPath wrapper for downwards-compatible but
// fast/ native Php.XPath usage
// By Philipp Lenssen 2006. Php.XPath was started by Nigel Swinson

class XPath
{

    private $dom = null;

    public function importFromString($sXml)
    {
        if ($sXml != '') {
            $this->dom = new DOMDocument();
            $this->dom->loadXML($sXml);
        }
    }
    
    public function match($sXpath, $context = null)
    {
        $retval = null;

        if ( !is_string($sXpath) )
        {
            die($sXpath);
            $sXpath = $d;
        }

        if ($this->dom != null)
        {
            $retval = array();
            $xpath = new DOMXPath($this->dom);

            $nodes = null;
            if ( $context == null ) {
                $nodes = $xpath->query($sXpath);
            }
            else {
                if ( is_array($context) ) { $context = $context[0]; }
                $nodes = $xpath->query($sXpath, $context);
            }

            $i = 0;
            foreach ($nodes as $node) {
                $retval[$i] = $node;
                $i++;
            }
        }

        return $retval;
    }
    
    public function getAttributes($node, $attributeName)
    {
        $retval = null;
        if ($node != null) {
            $retval = $node->getAttribute($attributeName);
        }
        return $retval;
    }
    
    public function getData($node)
    {
        return $node->firstChild->data; // or $node->firstChild->data ?
    }
    
    public function exportAsXml($node)
    {
        return $this->dom->saveXML($node);
    }
    
    public function nodeName($node)
    {
        return $node->nodeName;
    }

}

?>