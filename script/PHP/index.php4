<?

require_once("script/XPath.class.php4");
require_once("script/class_quest_handler.php4");

main();

function main()
{
    $oQuestHandler = new classQuestHandler();
    $oQuestHandler->setByQuery();
    $oQuestHandler->process();
}

?>