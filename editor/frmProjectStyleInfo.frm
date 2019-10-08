VERSION 5.00
Begin VB.Form frmProjectStyleInfo 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Project Style - Info box"
   ClientHeight    =   3324
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5676
   Icon            =   "frmProjectStyleInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   5676
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1920
      TabIndex        =   11
      Top             =   2880
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4440
      TabIndex        =   10
      Top             =   2880
      Width           =   1092
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   372
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Information"
      Height          =   2652
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5412
      Begin VB.TextBox txtFontFamily 
         Height          =   288
         Left            =   1320
         TabIndex        =   21
         Top             =   2160
         Width           =   1332
      End
      Begin VB.TextBox txtFontSize 
         Height          =   288
         Left            =   3240
         TabIndex        =   20
         Top             =   2160
         Width           =   492
      End
      Begin VB.CommandButton cmdSetFontColor 
         Caption         =   "Set..."
         Height          =   252
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox txtFontColor 
         Height          =   288
         Left            =   3960
         TabIndex        =   16
         Top             =   1680
         Width           =   1332
      End
      Begin VB.CommandButton cmdSetBackgroundColor 
         Caption         =   "Set..."
         Height          =   252
         Left            =   2760
         TabIndex        =   15
         Top             =   1200
         Width           =   972
      End
      Begin VB.TextBox txtBackgroundColor 
         Height          =   288
         Left            =   3960
         TabIndex        =   14
         Top             =   1200
         Width           =   1332
      End
      Begin VB.TextBox txtInfoPadding 
         Height          =   288
         Left            =   4560
         TabIndex        =   12
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox txtInfoWidth 
         Height          =   288
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtInfoLeft 
         Height          =   288
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox txtInfoHeight 
         Height          =   288
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtInfoTop 
         Height          =   288
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   492
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5280
         Y1              =   1032
         Y2              =   1032
      End
      Begin VB.Label Label4 
         Caption         =   "Font-Family:"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "Size:"
         Height          =   252
         Left            =   2760
         TabIndex        =   22
         Top             =   2160
         Width           =   492
      End
      Begin VB.Label Label3 
         Caption         =   "Font-Color:"
         Height          =   252
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Back-Color:"
         Height          =   252
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   972
      End
      Begin VB.Shape shpFontColor 
         FillStyle       =   0  'Ausgefüllt
         Height          =   252
         Left            =   1320
         Top             =   1680
         Width           =   1332
      End
      Begin VB.Shape shpBackgroundColor 
         FillStyle       =   0  'Ausgefüllt
         Height          =   252
         Left            =   1320
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Padding:"
         Height          =   252
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   732
      End
      Begin VB.Image Image4 
         Height          =   384
         Left            =   120
         Picture         =   "frmProjectStyleInfo.frx":0442
         Top             =   480
         Width           =   384
      End
      Begin VB.Label Label6 
         Caption         =   "Width:"
         Height          =   252
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   732
      End
      Begin VB.Label Label11 
         Caption         =   "Left:"
         Height          =   252
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   732
      End
      Begin VB.Label Label12 
         Caption         =   "Height:"
         Height          =   252
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   732
      End
      Begin VB.Label Label13 
         Caption         =   "Top:"
         Height          =   252
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   732
      End
   End
End
Attribute VB_Name = "frmProjectStyleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
