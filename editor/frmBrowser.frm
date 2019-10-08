VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "QML-Edit Browser"
   ClientHeight    =   6348
   ClientLeft      =   132
   ClientTop       =   624
   ClientWidth     =   9444
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6348
   ScaleWidth      =   9444
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList imgWebHi 
      Left            =   720
      Top             =   5880
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":01F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgWebLo 
      Left            =   120
      Top             =   5880
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":02A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":034A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrWeb 
      Align           =   1  'Oben ausrichten
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9444
      _ExtentX        =   16658
      _ExtentY        =   550
      ButtonWidth     =   529
      ButtonHeight    =   508
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgWebLo"
      HotImageList    =   "imgWebHi"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Back"
            Object.Tag             =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Forward"
            Object.Tag             =   "Forward"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   5412
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6252
      ExtentX         =   11028
      ExtentY         =   9546
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents webBrowser_EventListener As MSHTML.HTMLDocument
Attribute webBrowser_EventListener.VB_VarHelpID = -1

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Function webBrowser_EventListener_oncontextmenu() As Boolean
    webBrowser_EventListener_oncontextmenu = False
End Function

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Set webBrowser_EventListener = pDisp.Document
End Sub

Private Sub Form_Load()
    webBrowser.Navigate frmMain.browserStartURL
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        webBrowser.Move 0, tbrWeb.Height, _
                ScaleWidth, ScaleHeight - tbrWeb.Height
    End If
End Sub

Private Sub mnuFilePrint_Click()
    Dim status As OLECMDF
    MousePointer = vbHourglass
    DoEvents

    On Error Resume Next
    status = webBrowser.QueryStatusWB(OLECMDID_PRINT)

    If Err.Number = 0 Then
        If status And OLECMDF_ENABLED Then
            webBrowser.ExecWB OLECMDID_PRINT, _
                OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "Print is disabled."
        End If
    End If

    MousePointer = vbDefault
End Sub

Private Sub tbrWeb_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case LCase$(Button.Tag)
        Case "back"
            webBrowser.GoBack
        Case "forward"
            webBrowser.GoForward
    End Select
    On Error GoTo 0
End Sub


