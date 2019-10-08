Attribute VB_Name = "modDeclare"
Option Explicit

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const SB_THUMBPOSITION = 4
Public Const SB_VERT = 1
Public Const SB_HORZ = 0

' to open a folder ------------------------

Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const MAX_PATH = 260

Public Declare Function SHGetPathFromIDList _
   Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' ------------------------
