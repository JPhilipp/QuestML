VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProjectAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "About Project"
   ClientHeight    =   5208
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5484
   Icon            =   "frmProjectAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5208
   ScaleWidth      =   5484
   StartUpPosition =   1  'Fenstermitte
   Begin MSComDlg.CommonDialog dlgCommonProject 
      Left            =   240
      Top             =   960
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOptional 
      Caption         =   "Show optional >>>"
      Height          =   372
      Left            =   3720
      TabIndex        =   19
      Top             =   1560
      Width           =   1692
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4320
      TabIndex        =   16
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Caption         =   "Optional"
      Height          =   3084
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5292
      Begin VB.TextBox txtIntro 
         Height          =   612
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   18
         Top             =   1320
         Width           =   3492
      End
      Begin VB.CommandButton cmdInsertToday 
         Caption         =   "Today"
         Height          =   252
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   852
      End
      Begin VB.TextBox txtDate 
         Height          =   288
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   2532
      End
      Begin VB.TextBox txtEmail 
         Height          =   288
         Left            =   1680
         TabIndex        =   11
         Top             =   2640
         Width           =   3492
      End
      Begin VB.TextBox txtHomepage 
         Height          =   288
         Left            =   1680
         TabIndex        =   9
         Top             =   2280
         Width           =   3492
      End
      Begin VB.CommandButton cmdSetCoverImage 
         Caption         =   "Set..."
         Height          =   252
         Left            =   4320
         TabIndex        =   7
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox txtCover 
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   2532
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   5160
         Y1              =   792
         Y2              =   792
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5160
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Introduction text:"
         Height          =   612
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Date of creation:"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Your email:"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Your homepage:"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Cover image:"
         Height          =   372
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1212
      End
   End
   Begin VB.TextBox txtAuthor 
      Height          =   288
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   3612
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   3612
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author of quest:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title of quest:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1572
   End
End
Attribute VB_Name = "frmProjectAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const dialogDefaultHeight = 2340
Private Const dialogExtendedHeight = 5544
Private Const dialogDefaultCaption = "Show optional >>>"
Private Const dialogExtendedCaption = "Hide optional <<<"

Dim dialogIsExtended As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dialogIsExtended = Not CBool(GetSetting("qmlEdit", "options", "aboutIsExtended", False))
    toggleDialogDefaultExtended
    adaptDialogToNode
End Sub

Private Sub cmdOk_Click()
    If checkAllValidity Then
        saveAboutFromDialog
        Unload Me
    End If
End Sub

Private Sub saveAboutFromDialog()
    Dim child As IXMLDOMElement
    Dim objAbout As IXMLDOMElement
    Dim objOldAbout As IXMLDOMElement
    Dim objCover As IXMLDOMElement
    
    Set objAbout = frmMain.objQuest.documentElement.firstChild
    
    For Each child In objAbout.childNodes
        Select Case child.nodeName
            Case "title"
                child.firstChild.text = txtTitle.text
            Case "author"
                child.firstChild.text = txtAuthor.text
        End Select
    Next
    
    removeElementIfExists objAbout, "date"
    removeElementIfExists objAbout, "homepage"
    removeElementIfExists objAbout, "email"
    removeElementIfExists objAbout, "cover"
    removeElementIfExists objAbout, "intro"

    If txtDate.text <> "" Then
        addElmWithText frmMain.objQuest, objAbout, _
                "date", txtDate.text
    End If

    If txtCover.text <> "" Then
        Set objCover = frmMain.objQuest.createElement("cover")
        objCover.setAttribute "source", txtCover.text
        objAbout.appendChild objCover
    End If
    If txtIntro.text <> "" Then
        addElmWithText frmMain.objQuest, objAbout, _
                "intro", txtIntro.text
    End If
    
    If txtHomepage.text <> "" Then
        addElmWithText frmMain.objQuest, objAbout, _
                "homepage", txtHomepage.text
    End If
    If txtEmail.text <> "" Then
        addElmWithText frmMain.objQuest, objAbout, _
                "email", txtEmail.text
    End If
    
    
End Sub

Private Sub cmdSetCoverImage_Click()
    setTextBoxToWebImagePath txtCover
End Sub

Private Sub cmdInsertToday_Click()
    txtDate.text = Date
End Sub

Private Sub cmdOptional_Click()
    toggleDialogDefaultExtended
End Sub

Private Sub toggleDialogDefaultExtended()
    If dialogIsExtended Then
        Height = dialogDefaultHeight
        cmdOptional.Caption = dialogDefaultCaption
        SaveSetting "qmlEdit", "options", "extendedAbout", True
    Else
        Height = dialogExtendedHeight
        cmdOptional.Caption = dialogExtendedCaption
    End If
    dialogIsExtended = Not dialogIsExtended
    SaveSetting "qmlEdit", "options", "aboutIsExtended", dialogIsExtended
End Sub

Private Function checkAllValidity()
    Dim success As Boolean
    If Not checkTitleValidity Then
        txtTitle.SetFocus
    ElseIf Not checkAuthorValidity Then
        txtAuthor.SetFocus
    Else
        checkHomepageValidity
        checkEmailValidity
        success = True
    End If
    
    checkAllValidity = success
End Function

Private Sub txtHomepage_LostFocus()
    checkHomepageValidity
End Sub

Private Sub txtEmail_LostFocus()
    checkEmailValidity
End Sub

Private Sub checkHomepageValidity()
    With txtHomepage
        .text = Trim$(.text)
        If Not .text = "" Then
            If InStr(.text, "://") < 1 Then
                .text = "http://" & .text
            End If
        End If
    End With
End Sub

Private Sub checkEmailValidity()
    With txtEmail
        .text = Trim$(.text)
        If Not .text = "" Then
            If InStr(.text, "@") < 1 Then
                MsgBox "Email should contain an ""@"" character." & vbNewLine & _
                        "If you're using AOL, you would put" & vbNewLine & _
                        """@aol.com"" behind your nickname.", vbExclamation
            End If
        End If
    End With
End Sub

Private Function checkTitleValidity()
    Dim isValid As Boolean
    If Trim$(txtTitle.text) = "" Then
        MsgBox "Please enter a title for the story."
        isValid = False
    Else
        isValid = True
    End If
    checkTitleValidity = isValid
End Function

Private Function checkAuthorValidity()
    Dim isValid As Boolean
    If Trim$(txtAuthor.text) = "" Then
        MsgBox "Please enter your name."
        isValid = False
    Else
        isValid = True
    End If
    checkAuthorValidity = isValid
End Function

Private Sub adaptDialogToNode()
    Dim objAbout As IXMLDOMElement
    Dim objCover As IXMLDOMElement
    Dim title As IXMLDOMElement
    Dim author As IXMLDOMElement
    
    Set objAbout = frmMain.objQuest.selectSingleNode("//about")
    
    Set title = objAbout.selectSingleNode("title")
    Set author = objAbout.selectSingleNode("author")
    
    txtTitle.text = title.text
    txtAuthor.text = author.text
    
    txtDate.text = getChildElementText(objAbout, "date")
    txtHomepage.text = getChildElementText(objAbout, "homepage")
    txtEmail.text = getChildElementText(objAbout, "email")
    
    Set objCover = objAbout.selectSingleNode("cover")
    If Not (objCover Is Nothing) Then
        txtCover.text = objCover.getAttribute("source")
    End If
    
    txtIntro.text = getChildElementText(objAbout, "intro")
End Sub

