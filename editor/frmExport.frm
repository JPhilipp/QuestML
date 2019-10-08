VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Export"
   ClientHeight    =   3708
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   3708
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   3708
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1080
      TabIndex        =   6
      Top             =   3240
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   1212
   End
   Begin VB.CommandButton cmdSelectFolder 
      Caption         =   "Select..."
      Height          =   372
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblProjectname 
      Alignment       =   2  'Zentriert
      Caption         =   "unnamed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3492
   End
   Begin VB.Label Label2 
      Caption         =   $"frmExport.frx":014A
      Height          =   1212
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3492
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Folder to export to:"
      Height          =   216
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3492
   End
   Begin VB.Label lblFolder 
      BorderStyle     =   1  'Fest Einfach
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3492
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Based on source by VBNet, Randy Birch
' http://www.mvps.org/vbnet/index.html?code/browse/browsefolders.htm

Option Explicit

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub cmdSelectFolder_Click()
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim path As String
    Dim pos As Integer
    
    lblFolder.Caption = ""
    bi.hOwner = Me.hwnd
    bi.pidlRoot = 0&
    bi.lpszTitle = "Select the directory to export to"
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(bi)
    path = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
       pos = InStr(path, Chr$(0))
       lblFolder.Caption = Left(path, pos - 1)
    End If

    Call CoTaskMemFree(pidl)
End Sub

Private Sub cmdCancel_Click()
    lblFolder.Caption = ""
    Unload Me
End Sub
