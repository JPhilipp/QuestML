#!/usr/bin/python

import class_quest_handler

questHandler = class_quest_handler.QuestHandler()
print 'Content-type: text/html\n\n'
print questHandler.getOutput()