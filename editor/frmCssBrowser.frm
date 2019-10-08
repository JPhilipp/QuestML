VERSION 5.00
Begin VB.Form frmCssBrowser 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CSS2 Properties"
   ClientHeight    =   4164
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   4680
   Icon            =   "frmCssBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4164
   ScaleWidth      =   4680
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ListBox lstGroup 
      Height          =   816
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   2532
   End
   Begin VB.ListBox lstCssVal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2448
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1812
   End
   Begin VB.CommandButton comPaste 
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   372
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   1332
   End
   Begin VB.ListBox lstCssProp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2448
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2532
   End
   Begin VB.Label Label3 
      Caption         =   "Groups"
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2532
   End
   Begin VB.Label Label2 
      Caption         =   "Values"
      Height          =   252
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "Properties"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "frmCssBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Used TopStyle 1.5 as quick-reference

Private arrCss(1 To 500) As String

Private Sub comPaste_Click()
    Clipboard.SetText lstCssProp.List(lstCssProp.ListIndex) & ": " & _
        lstCssVal.List(lstCssVal.ListIndex) & ";"
    Unload Me
End Sub

Private Sub Form_Load()
    buildGroups
End Sub

Private Sub Form_Activate()
    lstGroup.ListIndex = 0 ' -> buildLstCss "all"
    'hiliteDef
End Sub

Private Sub buildGroups()
    With lstGroup
        .AddItem "all"
        .AddItem "aural"
        .AddItem "color"
        .AddItem "background"
        .AddItem "position"
        .AddItem "print"
        .AddItem "size"
        .AddItem "text"
    End With
End Sub

Private Sub buildLstCss(ByVal group As String)
    
    Const cssAnyLength = "ex | em | px | cm | mm | pc | in | pt", _
        cssAnyCol = "rgb(,,) | rgb(%,%,%) | #", _
        cssAnyBorderstyle = "none | dotted | dashed | solid | double | groove | ridge | inset | outset"
    
    lstCssProp.Clear
    lstCssVal.Clear
    
    addCss "", "", "", "", True ' initialization
    
    ' groupStr are needed if the property doesn't contain the groupStr

    addCss group, "azimuth", "inherit | deg | grad | rad | left-side | far-left | left | center-left | center | center-right | far-right | right-side | behind | leftwards | rightwards"
    addCss group, "background-attachment", "scroll | fixed"
    addCss group, "background-color", "transparent | " & cssAnyCol
    addCss group, "background-image", "url() | none"
    addCss group, "background-repeat", "no-repeat | repeat-x | repeat-y | repeat"
    addCss group, "background-position", "top | center | bottom | left | right"
    addCss group, "border-collapse", "inherit | collapse | seperate"
    addCss group, "border-color", cssAnyCol
    addCss group, "border-style", cssAnyBorderstyle
    addCss group, "border-width", "thin | medium | thick | % " & cssAnyLength
    addCss group, "bottom", "inherit | auto"
    addCss group, "color", "inherit | rgb(,,) | rgb(%,%,%) | #"
    addCss group, "caption-side", "inherit | top | bottom | left | right"
    addCss group, "clear", "inherit | none | left | right | both", "position"
    addCss group, "clip", "inherit | auto", "size"
    addCss group, "content", "url() | inherit | open-quote | close-quote | no-open-quote | no-close-quote"
    addCss group, "counter-increment", "inherit | none"
    addCss group, "counter-reset", "inherit | none"
    addCss group, "cue-before", "url() | inherit | none"
    addCss group, "cue-after", "url() | inherit | none"
    addCss group, "cursor", "url() | inherit | auto | crosshair | default | pointer | move | e-resize | ne-resize | nw-resize | n-resize | se-resize | sw-resize | s-resize | w-resize"
    addCss group, "direction", "inherit | ltr | rtl"
    addCss group, "display", "inherit | block | inline | list-item | none"
    addCss group, "elevation", "inherit | deg | grad | rad | below | level | above | higher | lower"
    addCss group, "empty-cells", "inherit | show | hide", "table"
    addCss group, "float", "inherit | left | right | none"
    addCss group, "font-family", "inherit | serif | sans-serif | cursive | fantasy | monospace", "text"
    addCss group, "font-size", "% | " & cssAnyLength, "text"
    addCss group, "font-size-adjust", "inherit | none", "text"
    addCss group, "font-stretch", "inherit | normal | wider | narrower | ultra-condensed | extra-condensed | condensed | semi-condensed | semi-expanded | expanded | extra-expanded | ultra-expanded", "text"
    addCss group, "font-style", "inherit | normal | italic | oblique", "text"
    addCss group, "font-variant", "inherit | normal | small-caps", "text"
    addCss group, "font-weight", "inherit | normal | bolder | bold | lighter | 100 | 200 | 300 | 400 | 500 | 600 | 700 | 800 | 900", "text"
    addCss group, "height", "% | " & cssAnyLength & " | inherit | auto", "size"
    addCss group, "left", "% | " & cssAnyLength & " | inherit | auto", "position"
    addCss group, "letter-spacing", cssAnyLength & " | inherit | normal", "text"
    addCss group, "line-height", "% | " & cssAnyLength & "inherit | normal"
    addCss group, "list-style-image", "url() | inherit | none"
    addCss group, "list-style-position", "inherit | inside | outside"
    addCss group, "list-style-type", "inherit | disc | circle | square | decimal | lower-roman | upper-roman | lower-alpha | upper-alpha | none"
    addCss group, "margin", "% |" & cssAnyLength
    addCss group, "marker-offset", cssAnyLength & " | inherit | auto"
    addCss group, "marks", "inherit | crop | cross | none"
    addCss group, "max-height", "% | " & cssAnyLength & " | inherit | none", "size"
    addCss group, "max-width", "% | " & cssAnyLength & " | inherit | none", "size"
    addCss group, "min-height", "% | " & cssAnyLength & " | inherit | none", "size"
    addCss group, "min-width", "% | " & cssAnyLength & " | inherit | none", "size"
    addCss group, "orphans", "inherit"
    addCss group, "outline-color", cssAnyCol & " | inherit | invert"
    addCss group, "outline-style", "inherit | " & cssAnyBorderstyle
    addCss group, "outline-width", cssAnyLength & " | inherit | thin | medium | thick"
    addCss group, "overflow", "inherit | visible | hidden | scroll | auto"
    addCss group, "padding", "% | " & cssAnyLength & " | inherit"
    addCss group, "page", "inherit | auto", "print"
    addCss group, "page-break-after", "inherit | auto | always | avoid | left | right", "print"
    addCss group, "page-break-before", "inherit | auto | always | avoid | left | right", "print"
    addCss group, "page-break-inside", "inherit | avoid | auto", "print"
    addCss group, "pause-before", "% | ms | s", "aural"
    addCss group, "pause-after", "% | ms | s", "aural"
    addCss group, "pitch", "inherit | hz | khz | x-low | low | medium | high | x-high", "aural"
    addCss group, "pitch-range", "inherit", "aural"
    addCss group, "play-during", "url() | inherit | mix | repeat | auto | none", "aural"
    addCss group, "position", "inherit | static | relative | absolute | fixed"
    addCss group, "quotes", "inherit | none"
    addCss group, "richness", "inherit", "aural"
    addCss group, "right", "% | " & cssAnyLength & " | inherit | auto", "position"
    addCss group, "size", cssAnyLength & " | inherit | auto | portrait | landscape"
    addCss group, "speak", "inherit | normal | none | spell-out", "aural"
    addCss group, "speak-header", "inherit | once | always", "aural"
    addCss group, "speak-numeral", "inherit | digits | continuous", "aural"
    addCss group, "speak-punctuation", "inherit | code | none", "aural"
    addCss group, "speech-rate", "inherit | x-slow | slow | medium | fast | x-fast | faster", "aural"
    addCss group, "stress", "inherit", "aural"
    addCss group, "table-layout", "inherit | auto | fixed", "table"
    addCss group, "text-align", "left | right | center | inherit"
    addCss group, "text-decoration", "inherit | none | underline | overline | line-through | blink"
    addCss group, "text-indent", "% | " & cssAnyLength & " | inherit"
    addCss group, "text-shadow", "inherit"
    addCss group, "text-transform", "inherit | capitalize | uppercase | lowercase | none"
    addCss group, "top", "% | " & cssAnyLength & " | inherit | auto", "position"
    addCss group, "unicode-bidi", "inherit | normal | embed | bidi-override"
    addCss group, "vertical-align", "% | inherit | baseline | sub | super | top | text-top | middle | bottom | text-bottom", "position"
    addCss group, "visibility", "inherit | visible | hidden | collapse"
    addCss group, "voice-family", "male | female | child", "aural"
    addCss group, "volume", "% | inherit | silent | x-soft | soft | medium | loud | x-loud", "aural"
    addCss group, "white-space", "inherit | normal | pre | nowrap", "text"
    addCss group, "widows", "inherit"
    addCss group, "width", "% | " & cssAnyLength & " | inherit | auto", "size"
    addCss group, "word-spacing", cssAnyLength & " | inherit | normal", "text"
    addCss group, "z-index", "inherit | auto", "position"
    'addCss group, "", ""
    
End Sub

Private Sub addCss(ByVal group As String, ByVal cssProp As String, ByVal cssOpts As String, Optional ByVal specGroup As String = "", Optional ByVal initArrI As Boolean = False)
    
    Dim groupOk As Boolean
    Static arrI As Integer
    If initArrI Then
        arrI = 0
        Exit Sub
    End If
    
    If group = "all" Or group = specGroup Then
        groupOk = True
    ElseIf InStr(1, cssProp, group, vbTextCompare) > 0 Then
        groupOk = True
    End If
    
    If groupOk Then
        arrI = arrI + 1
        lstCssProp.AddItem (cssProp)
        arrCss(arrI) = cssOpts & "|"
    End If
            
End Sub

Private Sub lstCssProp_Click()

    Dim strVals As String, _
        lastI As Integer, newI As Integer, curWord As String
    strVals = arrCss(lstCssProp.ListIndex + 1)
    
    lstCssVal.Clear
    Do
        newI = InStr(lastI + 1, strVals, "|", vbTextCompare)
        If newI > 0 Then
            curWord = Mid$(strVals, lastI + 1, newI - lastI)
            curWord = Left$(curWord, Len(curWord) - 1)
            curWord = Trim$(curWord)
            lstCssVal.AddItem curWord
            lastI = newI
        Else
            Exit Do
        End If
    Loop

End Sub

Private Sub lstCssProp_DblClick()
    lstCssVal.Clear
    comPaste_Click
End Sub

Private Sub lstCssVal_DblClick()
    comPaste_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    handleKey KeyCode
End Sub

Private Sub lstCssProp_KeyDown(KeyCode As Integer, Shift As Integer)
    handleKey KeyCode
End Sub

Private Sub lstCssVal_KeyDown(KeyCode As Integer, Shift As Integer)
    handleKey KeyCode
End Sub

Private Sub handleKey(ByVal KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub lstGroup_Click()
    buildLstCss lstGroup.List(lstGroup.ListIndex)
End Sub
