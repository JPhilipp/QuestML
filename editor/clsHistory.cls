VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A history collecting strings for things like
' back/ forward buttons of a browser

Const maxEntries = 30

Private history(1 To maxEntries) As String
Private pointer As Long

Private Sub Class_Initialize()
    resetPointer
End Sub

Public Sub clear()
    resetHistory
    resetPointer
End Sub

Public Sub add(ByVal name As String)
    If pointer + 1 > UBound(history) Then
        shiftArrayLeft history()
    Else
        pointer = pointer + 1
    End If
    history(pointer) = name
End Sub

Public Function back() As Boolean
    Dim oldPointer As Long
    oldPointer = pointer
    movePointer -1
    back = oldPointer <> pointer
End Function

Public Function canGoBack() As Boolean
    canGoBack = indexIsOk(pointer - 1)
End Function

Public Function canGoForward() As Boolean
    canGoForward = indexIsOk(pointer + 1)
End Function

Public Function forward() As Boolean
    Dim oldPointer As Long
    oldPointer = pointer
    movePointer 1
    If getValue = "" Then movePointer -1
    forward = oldPointer <> pointer
End Function

Public Function getValue() As String
    getValue = getValueOfIndex(pointer)
End Function

Private Sub clearAfter(ByVal index As Long)
    Dim i As Long
    If index < UBound(history) Then
        For i = index + 1 To UBound(history)
            history(i) = ""
        Next
    End If
End Sub

Public Sub remove(ByVal name As String)
    Dim i As Long
    For i = LBound(history) To UBound(history)
        If history(i) = name Then
            history(i) = ""
        End If
    Next
    collectGarbage
End Sub

Public Function getDebugString() As String
    Dim i As Long
    Dim debugString As String
    
    For i = LBound(history) To UBound(history)
        If i = pointer Then
            debugString = debugString & "*" & history(i) & "*"
        Else
            debugString = debugString & history(i)
        End If
        If i < UBound(history) Then
            debugString = debugString & " --> "
        End If
    Next
    
    getDebugString = debugString
End Function

'--- private ---

Private Function indexIsOk(ByVal i As Long) As Boolean
    Dim isOk As Boolean
    isOk = False
    If i >= LBound(history) And i <= UBound(history) Then
        isOk = history(i) <> ""
    End If
    indexIsOk = isOk
End Function

Private Function getIndex(ByVal name As String) As Long
    Dim i As Long
    Dim index As Long
    index = -1
    For i = LBound(history) To UBound(history)
        If history(i) = name Then
            index = i
            Exit For
        End If
    Next
    getIndex = index
End Function

Private Sub shiftArrayLeft(ByRef arr() As String)
    shiftArrayLeftStartAt arr(), LBound(arr)
End Sub

Private Sub shiftArrayLeftStartAt(ByRef arr() As String, ByVal startAt As Long)
    Dim i As Long
    For i = startAt To UBound(arr) - 1
        arr(i) = arr(i + 1)
    Next
    arr(UBound(arr)) = ""
End Sub

Private Function getValueOfIndex(ByVal i As Long) As String
    Dim returnValue As String
    returnValue = ""
    If i >= LBound(history) And i <= UBound(history) Then
        returnValue = history(i)
    End If
    getValueOfIndex = returnValue
End Function

Private Function movePointer(ByVal n As Long)
    pointer = pushIntoLimits(pointer + n, LBound(history), UBound(history))
End Function

Private Function pushIntoLimits(ByVal n As Long, ByVal limit1 As Long, ByVal limit2 As Long) As Long
    Dim min As Long
    Dim max As Long
    
    If limit1 > limit2 Then
        min = limit2
        max = limit1
    Else
        min = limit1
        max = limit2
    End If
    
    If n < min Then
        n = min
    ElseIf n > max Then
        n = max
    End If
    
    pushIntoLimits = n
End Function

Private Function resetHistory()
    Dim i As Long
    For i = LBound(history) To UBound(history)
        history(i) = ""
    Next
End Function

Private Function resetPointer()
    pointer = LBound(history) - 1
End Function

Private Function collectGarbage()
    Dim foundEmptyEntry As Boolean
    Do
        foundEmptyEntry = removeFirstEmptyEntry
    Loop Until Not foundEmptyEntry
    readjustPointer
End Function

Sub readjustPointer()
    Dim firstEmpty As Long
    firstEmpty = getIndex("")
    If firstEmpty <= pointer Then
        pointer = firstEmpty - 1
    End If
End Sub

Private Function removeFirstEmptyEntry()
    Dim indexOfEmpty As Long
    Dim foundEmptyEntry As Boolean
    
    foundEmptyEntry = False
    indexOfEmpty = getFirstEmptyInbetween
    
    If indexOfEmpty <> -1 Then
        foundEmptyEntry = True
        shiftArrayLeftStartAt history(), indexOfEmpty
    End If
    
    removeFirstEmptyEntry = foundEmptyEntry
End Function

Private Function getFirstEmptyInbetween() As Long
    Dim i As Long
    Dim firstEmptyInbetween As Long
    
    firstEmptyInbetween = -1
    For i = LBound(history) To UBound(history) - 1
        If history(i) = "" And history(i + 1) <> "" Then
            firstEmptyInbetween = i
            Exit For
        End If
    Next
    
    getFirstEmptyInbetween = firstEmptyInbetween
End Function
