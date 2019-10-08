VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Program options"
   ClientHeight    =   2268
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5028
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2268
   ScaleWidth      =   5028
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CheckBox chkAutoHandleDebugMode 
      Caption         =   "Auto-handle debug mode"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   4692
   End
   Begin VB.CheckBox chkAutoCreateChoice 
      Caption         =   "Automatically create a choice for new stations (source mode)"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      WhatsThisHelpID =   1
      Width           =   4812
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   372
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog dlgCommonOptions 
      Left            =   1680
      Top             =   360
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.TextBox txtFontName 
      BackColor       =   &H8000000A&
      Height          =   288
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton cmdSetFont 
      Caption         =   "Set..."
      Height          =   252
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   852
   End
   Begin VB.ComboBox cboStartMode 
      Height          =   288
      ItemData        =   "frmOptions.frx":0442
      Left            =   2160
      List            =   "frmOptions.frx":044F
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   120
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Font for editable text:"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "Mode to start program in:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDefault_Click()
    If vbOK = MsgBox("All options will be set back to the default values.", _
            vbOKCancel) Then
        resetOptionsToDefault
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    adaptDialogToRegistry
End Sub

Private Sub adaptDialogToRegistry()
    setListStartModeFromRegistry
    setFontNameTextboxFromRegistry
    setAutoCreateChoiceFromRegistry
    setAutoHandleDebugModeFromRegistry
End Sub

Private Sub resetOptionsToDefault()
    SaveSetting "qmlEdit", "options", "defaultTab", defaultTab
    SaveSetting "qmlEdit", "options", "fontName", defaultFontName
    SaveSetting "qmlEdit", "options", "fontSize", defaultFontSize
    SaveSetting "qmlEdit", "options", "autoCreateChoice", True
    SaveSetting "qmlEdit", "options", "autoHandleDebugMode", True
End Sub

Private Sub cmdOk_Click()
    saveFontToRegistry
    SaveSetting "qmlEdit", "options", "defaultTab", _
            cboStartMode.ListIndex
    SaveSetting "qmlEdit", "options", "autoCreateChoice", _
            (chkAutoCreateChoice.Value = 1)
    SaveSetting "qmlEdit", "options", "autoHandleDebugMode", _
            (chkAutoHandleDebugMode.Value = 1)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSetFont_Click()
    With dlgCommonOptions
        .Flags = cdlCFScreenFonts
        .ShowFont
        setFontNameTextboxFromDialog
    End With
End Sub

Private Sub saveFontToRegistry()
    With dlgCommonOptions
        If Trim$(.fontName) <> "" Then
            SaveSetting "qmlEdit", "options", "fontName", .fontName
            SaveSetting "qmlEdit", "options", "fontSize", .fontSize
        End If
    End With
End Sub

Private Sub setListStartModeFromRegistry()
    Dim currentDefaultTab As tabType
    currentDefaultTab = GetSetting("qmlEdit", "options", "defaultTab", defaultTab)
    cboStartMode.ListIndex = currentDefaultTab
End Sub

Private Sub setFontNameTextboxFromRegistry()
    Dim fontName As String, fontSize As Integer
    fontName = GetSetting("qmlEdit", "options", "fontName", defaultFontName)
    fontSize = GetSetting("qmlEdit", "options", "fontSize", defaultFontSize)
    setFontNameTextbox fontName, fontSize
End Sub

Private Sub setFontNameTextboxFromDialog()
    With dlgCommonOptions
        If Trim$(.fontName) <> "" Then
            setFontNameTextbox .fontName, .fontSize
        End If
    End With
End Sub

Private Sub setAutoCreateChoiceFromRegistry()
    With chkAutoCreateChoice
        If GetSetting("qmlEdit", "options", _
                "autoCreateChoice", True) Then
            .Value = 1
        Else
            .Value = 0
        End If
    End With
End Sub

Private Sub setAutoHandleDebugModeFromRegistry()
    With chkAutoHandleDebugMode
        If GetSetting("qmlEdit", "options", _
                "autoHandleDebugMode", True) Then
            .Value = 1
        Else
            .Value = 0
        End If
    End With
End Sub

Private Sub setFontNameTextbox(ByVal fontName As String, ByVal fontSize As Integer)
    txtFontName.text = fontName & ", size " & Str$(fontSize)
End Sub

