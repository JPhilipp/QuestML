Attribute VB_Name = "modConstants"
Option Explicit

Public Const elementStation = "station", _
        elementBreak = "break", _
        elementImage = "image", _
        elementText = "text", _
        elementPath = "choice", _
        elementIf = "if", _
        elementInput = "input", _
        elementRandomize = "randomize", _
        elementEmphase = "emphasis", _
        elementStrong = "strong", _
        elementMusic = "music"
Public Const nameOfStartStation = "start"
Public Const defaultStationStatesAttribute = "persist", _
        defaultStationSavingAttribute = "default"
Public Const qmlPreviewName = "qmlpreview"
Public Const stationBackName = "back"
Public Const graphFilePath = "tool\add-on\graph_layout"

Public Const defaultFontName = "Arial"
Public Const defaultFontSize = "10"
Public Const defaultTab = 0

Public Enum tabType
    tabEditor = 0
    tabsource = 1
    tabPreview = 2
End Enum
