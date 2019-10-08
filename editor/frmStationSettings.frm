VERSION 5.00
Begin VB.Form frmStationSettings 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Station settings"
   ClientHeight    =   1056
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4164
   Icon            =   "frmStationSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1056
   ScaleWidth      =   4164
   StartUpPosition =   1  'Fenstermitte
   Begin VB.ComboBox cboStates 
      Height          =   288
      ItemData        =   "frmStationSettings.frx":0442
      Left            =   1560
      List            =   "frmStationSettings.frx":044C
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "States:"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "frmStationSettings"
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
    Dim objStation As IXMLDOMElement
    
    frmMain.setStationFromSelectedId objStation
    cboStates.text = objStation.getAttribute("states")
End Sub

Private Sub adaptNodeToDialog()
    Dim objStation As IXMLDOMElement
    
    frmMain.setStationFromSelectedId objStation
    If cboStates.text = defaultStationStatesAttribute Then
        objStation.removeAttribute "states"
    Else
        objStation.setAttribute "states", cboStates.text
    End If
End Sub
