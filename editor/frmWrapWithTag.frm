VERSION 5.00
Begin VB.Form frmWrapWithTag 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Wrap Tag selection"
   ClientHeight    =   1572
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2736
   Icon            =   "frmWrapWithTag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1572
   ScaleWidth      =   2736
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cboTag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1344
      ItemData        =   "frmWrapWithTag.frx":0442
      Left            =   120
      List            =   "frmWrapWithTag.frx":0464
      Sorted          =   -1  'True
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   252
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   972
   End
End
Attribute VB_Name = "frmWrapWithTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lastKeyPressed As Integer

Private Sub cboTag_Change()
    Dim i As Integer, txt As String, origStart As Integer
    txt = cboTag.text
    If txt = "" Or lastKeyPressed = 8 Then Exit Sub
    
    For i = 0 To cboTag.ListCount
        If Left$(cboTag.List(i), Len(txt)) = LCase$(txt) Then
            txt = cboTag.List(i)
            Exit For
        End If
    Next
    origStart = cboTag.SelStart
    If txt <> cboTag.text Then
        cboTag.text = txt
        cboTag.SelStart = origStart
        cboTag.SelLength = Len(txt) - origStart
    End If
End Sub

Private Sub cboTag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOk_Click
    ElseIf KeyCode = vbKeyEscape Then
        Hide
    End If
End Sub

Private Sub cboTag_KeyPress(KeyAscii As Integer)
    lastKeyPressed = KeyAscii
End Sub

Private Sub cmdOk_Click()
    Dim strTag As String, startTag As String, endTag As String, _
        addAttr As String, oldPos As Long, newPos As Long, _
        pasteText As String
    
    strTag = cboTag.text
    
    If Len(strTag) Then
        uniteList LCase$(strTag)
        
        oldPos = frmMain.txtSource.SelStart
        Select Case strTag
            Case "choice"
                addAttr = " station="""""
                newPos = oldPos + Len(strTag) + Len(addAttr)
        End Select
        startTag = "<" & strTag & addAttr & ">"
        endTag = "</" & strTag & ">"
        
        With frmMain.txtSource
            pasteText = Trim$(.SelText)
            If pasteText = "" Then pasteText = Clipboard.GetText
            pasteText = Trim$(pasteText)
            pasteText = LRTrimTab(pasteText)
            .SelText = startTag & pasteText & endTag
            If Len(addAttr) Then .SelStart = newPos
        End With
    End If
    
    Hide
End Sub

Private Sub uniteList(ByVal strTag As String)
    If Not isInList(cboTag, strTag) Then cboTag.AddItem strTag
End Sub

Private Function isInList(ByRef objCombo As ComboBox, ByVal strng As String) As Boolean
    Dim i As Integer
    For i = 0 To objCombo.ListCount - 1
        If objCombo.List(i) = strng Then
            isInList = True
            Exit For
        End If
    Next
End Function

Public Function LRTrimTab(ByVal txt As String) As String
    Dim oldText As String
    Do
        oldText = txt
        If Left$(txt, 1) = Chr(9) Then
            txt = Right$(txt, Len(txt) - 1)
        End If
        If Right$(txt, 1) = Chr(9) Then
            txt = Left$(txt, Len(txt) - 1)
        End If
    Loop Until oldText = txt
    
    LRTrimTab = txt
End Function

