<%

option explicit

const g_isServerVersion = true
const g_aspFileName = "default.asp"
const g_qmlVersionNumber = "2.032"
const g_defaultValue = "default"
const g_noIndexFound = -1
const g_none = "none"
const g_numberDefaultMin = -50000
const g_numberDefaultMax =  50000

dim g_clientSessionData

if g_isServerVersion then
    processRequest
end if

sub processRequest
    dim station
    dim oStatesString
    dim statesString

    set oStatesString = new classStatesString
    statesString = oStatesString.getStatesString

    station = request.queryString("station")
    if station = "" then
        station = "start"
    end if
    handleStation request.queryString("quest"), _
            station, _
            request.queryString("session"), _
            request.queryString("content"), _
            statesString
end sub

sub handleStation(byVal questName, byVal stationId, byVal sessionId, byVal contentType, byVal statesString)
    dim oQuestHandler

    set oQuestHandler = new classQuestHandler
    oQuestHandler.setQuestName questName
    oQuestHandler.setStationId stationId
    oQuestHandler.setSessionId sessionId
    oQuestHandler.setContentType contentType
    oQuestHandler.setStatesString statesString
    oQuestHandler.init
    oQuestHandler.doHandleStation
end sub

%>