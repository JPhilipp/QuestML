Attribute VB_Name = "modHiliteSource"
Option Explicit

Private Type tpeDomLocation
    inTag As Boolean
    inAttribute As Boolean
    inValue As Boolean
    inVariable As Boolean
End Type

Private oldSelStart As Long
Private oldsellength As Long
Private oldScrollPositionY As Long
Private oldScrollPositionX As Long
Private updateIsLocked As Boolean

Public Sub setSourceAndHilite(ByRef txt As RichTextBox, ByRef textSource As String)
    lockText txt
    txt = textSource
    hilite txt
    unlockText txt
End Sub

Public Sub hiliteSource(ByRef txt As RichTextBox)
    lockText txt
    hilite txt
    unlockText txt
End Sub

Public Sub lockText(ByRef txt As RichTextBox)
    frmMain.eventsOn = False
    memorizeSelection txt
    lockUpdate txt
End Sub

Public Sub unlockText(ByRef txt As RichTextBox)
    recallSelection txt
    releaseLockUpdate
    frmMain.eventsOn = True
End Sub

Private Sub hilite(ByRef txt As RichTextBox)
    Const elementColor = &HC00000
    Const attributeColor = 7159651
    Const valueColor = &H808000
    Const textColor = 0
    Const errorColor = &HC0&

    Dim i As Long
    Dim letter As String
    Dim color As Long
    Dim dom As tpeDomLocation
    Dim closeNext As Boolean
    Dim closeValueNext As Boolean
    Dim closeVariableNext As Boolean
    Dim lastColorIndex As Long
    Dim lastColor As Long
    Dim lenText As Long
    Dim domError As Boolean
    Dim valueQuote As String
    Dim decoration As Boolean
    Dim lastDecoration As Boolean
    Dim lastDecorationIndex As Long
    Dim text As String

    lastColor = -1
    lastColorIndex = 2
    lastDecorationIndex = 2
    lastDecoration = False
    
    With txt
        text = .text
        lenText = Len(text)
        For i = 1 To lenText
            
            If closeNext Then
                closeNext = False
                dom.inTag = False
                dom.inAttribute = False
            ElseIf closeValueNext Then
                closeValueNext = False
                dom.inValue = False
                valueQuote = ""
            ElseIf closeVariableNext Then
                closeVariableNext = False
                dom.inVariable = False
            End If
            
            letter = Mid$(text, i, 1)
            Select Case letter
                Case "<"
                    If dom.inAttribute Then
                        domError = True
                    ElseIf Not dom.inTag Then
                        dom.inTag = True
                    End If
                Case ">"
                    If dom.inValue Then
                        domError = True
                    Else
                        closeNext = True
                    End If
                Case "["
                    If (Not dom.inTag) Or dom.inValue Then
                        dom.inVariable = True
                    End If
                Case "]"
                    If dom.inVariable Then
                        closeVariableNext = True
                    End If
                Case " "
                    If dom.inTag Then
                        dom.inAttribute = True
                    End If
                Case """", "'"
                    If dom.inTag And Not dom.inAttribute Then
                        domError = True
                    ElseIf dom.inAttribute Then
                        If Not dom.inValue Then
                            dom.inValue = True
                            valueQuote = letter
                            
                        ElseIf valueQuote = letter Then
                            closeValueNext = True
                        End If
                    End If
            End Select
            
            If domError Then
                color = errorColor
                decoration = False
            ElseIf dom.inVariable Then
                decoration = True
            ElseIf dom.inValue Then
                color = valueColor
                decoration = False
            ElseIf dom.inAttribute Then
                color = attributeColor
                decoration = False
            ElseIf dom.inTag Then
                color = elementColor
                decoration = False
            Else
                color = textColor
                decoration = False
            End If
            
            If color <> lastColor Or decoration <> lastDecoration Or i = lenText Then
                txt.SelStart = lastColorIndex - 1
                txt.SelLength = i - lastColorIndex + 1
                
                txt.SelColor = lastColor
                txt.SelItalic = lastDecoration

                lastColor = color
                lastColorIndex = i
                
                lastDecoration = decoration
                lastDecorationIndex = i
            End If
            
        Next
    End With
End Sub

Private Sub lockUpdate(ByRef txt As RichTextBox)
    If Not updateIsLocked Then
        LockWindowUpdate txt.hwnd
        updateIsLocked = True
    End If
End Sub

Private Sub releaseLockUpdate()
    If updateIsLocked Then
        LockWindowUpdate 0&
        updateIsLocked = False
    End If
End Sub

Private Sub memorizeSelection(ByRef txt As RichTextBox)
    memorizeFocus txt
    oldSelStart = txt.SelStart
    oldsellength = txt.SelLength
End Sub

Private Sub memorizeFocus(ByRef txt As RichTextBox)
    oldScrollPositionX = GetScrollPos(txt.hwnd, SB_HORZ)
    oldScrollPositionY = GetScrollPos(txt.hwnd, SB_VERT)
End Sub

Private Sub recallSelection(ByRef txt As RichTextBox)
    txt.SelStart = oldSelStart
    txt.SelLength = oldsellength
    recallFocus txt
End Sub

Private Sub recallFocus(ByRef txt As RichTextBox)
    Call PostMessage(txt.hwnd, WM_VSCROLL, _
        SB_THUMBPOSITION Or (65536 * oldScrollPositionY), 0)
    Call PostMessage(txt.hwnd, WM_HSCROLL, _
        SB_THUMBPOSITION Or (65536 * oldScrollPositionX), 0)
End Sub

