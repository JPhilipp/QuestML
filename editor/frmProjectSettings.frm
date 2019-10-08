VERSION 5.00
Begin VB.Form frmProjectSettings 
   Caption         =   "Project settings"
   ClientHeight    =   1560
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   4140
   Icon            =   "frmProjectSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4140
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cboDebug 
      Height          =   288
      ItemData        =   "frmProjectSettings.frx":0442
      Left            =   1560
      List            =   "frmProjectSettings.frx":044C
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   600
      Width           =   2532
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   288
      ItemData        =   "frmProjectSettings.frx":045D
      Left            =   1560
      List            =   "frmProjectSettings.frx":0467
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Debug mode in browser active:"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label lblLang 
      Caption         =   "Language:"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "frmProjectSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    adaptDialogToNode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    adaptNodeToDialog
    Unload Me
End Sub

Private Sub adaptDialogToNode()
    cboLanguage.text = _
            frmMain.objQuest.documentElement.getAttribute("language")
    cboDebug.text = _
            frmMain.objQuest.documentElement.getAttribute("debug")
End Sub

Private Sub adaptNodeToDialog()
    frmMain.objQuest.documentElement.setAttribute _
            "language", cboLanguage.text
    frmMain.objQuest.documentElement.setAttribute _
            "debug", cboDebug.text
End Sub

