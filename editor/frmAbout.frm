VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "About"
   ClientHeight    =   2352
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3348
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2352
   ScaleWidth      =   3348
   StartUpPosition =   1  'Fenstermitte
   WhatsThisHelp   =   -1  'True
   Begin VB.Label Label4 
      Caption         =   "Additional program and language design by Dave Posladek, Mike Marinos and Fred Clarke."
      Height          =   612
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3012
   End
   Begin VB.Image Image1 
      Height          =   576
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblLink 
      Caption         =   "http://www.questml.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1764
   End
   Begin VB.Label Label3 
      Caption         =   "More about the Quest Markup Language:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3096
   End
   Begin VB.Label Label2 
      Caption         =   "Freeware © 2000 - 2002 by Philipp Lenssen"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   3252
   End
   Begin VB.Label lblVersion 
      Caption         =   "1.x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1824
      TabIndex        =   0
      Top             =   444
      Width           =   1008
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion = App.Major & "." & App.Minor
End Sub

Private Sub lblLink_Click()
    Hide
    frmMain.visitQMLHomepage
    Unload Me
End Sub
