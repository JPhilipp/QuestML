VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Find in Stations"
   ClientHeight    =   1068
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4188
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1068
   ScaleWidth      =   4188
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1212
   End
   Begin VB.TextBox txtFind 
      Height          =   288
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3012
   End
   Begin VB.Label Label1 
      Caption         =   "Find:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public find As String, suggestedText As String

Private Sub Form_Load()
    txtFind.text = suggestedText
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.text)
End Sub

Private Sub cmdCancel_Click()
    find = ""
    Hide
End Sub

Private Sub cmdOk_Click()
    find = Trim$(txtFind.text)
    Hide
End Sub

