<?

require_once('script/XPath.class.php5');
require_once('script/class_quest_handler.php5');

main();

function main()
{
    $oQuestHandler = new classQuestHandler();
    $oQuestHandler->setByQuery();
    $oQuestHandler->process();
}

?>