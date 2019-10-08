VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "QML-Edit Help"
   ClientHeight    =   6144
   ClientLeft      =   132
   ClientTop       =   624
   ClientWidth     =   6264
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6144
   ScaleWidth      =   6264
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
            Picture         =   "frmHelp.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":04EE
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
            Picture         =   "frmHelp.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelp.frx":0642
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
      Width           =   6264
      _ExtentX        =   11049
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
            Object.Tag             =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
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

Public startURL As String

Private WithEvents webBrowser_EventListener As MSHTML.HTMLDocument
Attribute webBrowser_EventListener.VB_VarHelpID = -1

Private Function webBrowser_EventListener_oncontextmenu() As Boolean
    webBrowser_EventListener_oncontextmenu = False
End Function

Private Sub webBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Set webBrowser_EventListener = pDisp.Document
End Sub

Private Sub Form_Load()
    WebBrowser.Navigate startURL
End Sub

Private Sub Form_Resize()
    WebBrowser.Move 0, tbrWeb.Height, ScaleWidth, ScaleHeight - tbrWeb.Height
End Sub

Private Sub mnuFilePrint_Click()
    Dim status As OLECMDF
    MousePointer = vbHourglass
    DoEvents

    On Error Resume Next
    status = WebBrowser.QueryStatusWB(OLECMDID_PRINT)

    If Err.Number = 0 Then
        If status And OLECMDF_ENABLED Then
            WebBrowser.ExecWB OLECMDID_PRINT, _
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
            WebBrowser.GoBack
        Case "forward"
            WebBrowser.GoForward
    End Select
    On Error GoTo 0
End Sub


