<%

option explicit

const g_isServerVersion = true
const g_aspFileName = "default.asp"
const g_qmlVersionNumber = "1.4"
const g_defaultValue = "default"
const g_noIndexFound = -1
const g_none = "none"
const g_notState = "not "
const g_numberDefaultMin = -30000
const g_numberDefaultMax =  30000
const g_visitsStartString = "visits("

dim g_clientSessionData

if g_isServerVersion then
    processRequest
end if

sub processRequest
    dim station

    station = request.queryString("station")
    if station = "" then
        station = "start"
    end if
    handleStation request.queryString("quest"), _
            station, _
            request.queryString("session"), _
            request.queryString("content")
end sub

sub handleStation(byVal questName, byVal stationId, byVal sessionId, byVal contentType)
    dim oQuestHandler

    set oQuestHandler = new classQuestHandler
    oQuestHandler.setQuestName questName
    oQuestHandler.setStationId stationId
    oQuestHandler.setSessionId sessionId
    oQuestHandler.setContentType contentType
    oQuestHandler.init
    oQuestHandler.doHandleStation
end sub

%>