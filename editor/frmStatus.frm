VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Status"
   ClientHeight    =   708
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   4416
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   708
   ScaleWidth      =   4416
   StartUpPosition =   1  'Fenstermitte
   Begin MSComctlLib.ProgressBar barProgress 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblTask 
      Caption         =   "Processing task..."
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4212
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

