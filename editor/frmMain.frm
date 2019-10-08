VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "QML-Edit"
   ClientHeight    =   5880
   ClientLeft      =   132
   ClientTop       =   624
   ClientWidth     =   9348
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9348
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cboContext 
      Appearance      =   0  '2D
      Height          =   1872
      ItemData        =   "frmMain.frx":014A
      Left            =   3240
      List            =   "frmMain.frx":014C
      Sorted          =   -1  'True
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "cboContext"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2172
   End
   Begin MSComctlLib.Toolbar toolbar 
      Align           =   1  'Oben ausrichten
      Height          =   312
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   9348
      _ExtentX        =   16489
      _ExtentY        =   550
      ButtonWidth     =   529
      ButtonHeight    =   508
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imagesToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "search"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            Object.ToolTipText     =   "About Project"
            ImageKey        =   "about"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "style"
            Object.ToolTipText     =   "Project Style"
            ImageKey        =   "style"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "back"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "forward"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "View Help"
            ImageKey        =   "help"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imagesToolbar 
         Left            =   5256
         Top             =   -96
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         ImageWidth      =   18
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":014E
               Key             =   "new"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0592
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":09D6
               Key             =   "save"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0E1A
               Key             =   "about"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":125E
               Key             =   "style"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16A2
               Key             =   "image"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":19F6
               Key             =   "copy"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1AAE
               Key             =   "cut"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B56
               Key             =   "paste"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C1A
               Key             =   "search"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D66
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E02
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1EA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F4A
               Key             =   "back"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FEE
               Key             =   "forward"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2092
               Key             =   "backLo"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":212A
               Key             =   "forwardLo"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21C2
               Key             =   "help"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imagesSyntax 
      Left            =   5880
      Top             =   120
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":226A
            Key             =   "valueDefault"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":237E
            Key             =   "element"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2492
            Key             =   "attribute"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25A6
            Key             =   "value"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26BA
            Key             =   "if"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27CE
            Key             =   "text"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28E2
            Key             =   "media"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2982
            Key             =   "state"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A36
            Key             =   "station"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AD6
            Key             =   "inline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B6A
            Key             =   "valueUser"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   252
      Left            =   1680
      TabIndex        =   3
      Top             =   5160
      Width           =   492
   End
   Begin VB.TextBox txtStationName 
      Height          =   288
      Left            =   72
      TabIndex        =   2
      Top             =   5160
      Width           =   1548
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8280
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.ListBox lstStations 
      Height          =   4464
      ItemData        =   "frmMain.frx":2C06
      Left            =   60
      List            =   "frmMain.frx":2C08
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2112
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5172
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   9123
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   596
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "E&ditor"
      TabPicture(0)   =   "frmMain.frx":2C0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtStation"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmChoices"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Source"
      TabPicture(1)   =   "frmMain.frx":3094
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSource"
      Tab(1).Control(1)=   "txtValidate"
      Tab(1).Control(2)=   "cmdValidate"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "P&review"
      TabPicture(2)   =   "frmMain.frx":35AE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "webPreview"
      Tab(2).ControlCount=   1
      Begin RichTextLib.RichTextBox txtSource 
         Height          =   3852
         Left            =   -74880
         TabIndex        =   29
         Top             =   120
         Width           =   6732
         _ExtentX        =   11875
         _ExtentY        =   6795
         _Version        =   393217
         HideSelection   =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":3AC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame frmChoices 
         Caption         =   "Choices"
         Height          =   2532
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   6732
         Begin VB.TextBox txtChoice 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   3972
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Index           =   0
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1812
         End
         Begin VB.CommandButton cmdGo 
            Height          =   252
            Index           =   0
            Left            =   6120
            Picture         =   "frmMain.frx":3BA8
            Style           =   1  'Grafisch
            TabIndex        =   7
            Top             =   588
            Width           =   492
         End
         Begin VB.TextBox txtChoice 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   3972
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Index           =   1
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   960
            Width           =   1812
         End
         Begin VB.CommandButton cmdGo 
            Height          =   252
            Index           =   1
            Left            =   6120
            Picture         =   "frmMain.frx":3C34
            Style           =   1  'Grafisch
            TabIndex        =   10
            Top             =   960
            Width           =   492
         End
         Begin VB.TextBox txtChoice 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   3972
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Index           =   2
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1320
            Width           =   1812
         End
         Begin VB.CommandButton cmdGo 
            Height          =   252
            Index           =   2
            Left            =   6120
            Picture         =   "frmMain.frx":3CC0
            Style           =   1  'Grafisch
            TabIndex        =   13
            Top             =   1320
            Width           =   492
         End
         Begin VB.TextBox txtChoice 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   3972
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Index           =   3
            Left            =   4200
            TabIndex        =   15
            Top             =   1680
            Width           =   1812
         End
         Begin VB.CommandButton cmdGo 
            Height          =   252
            Index           =   3
            Left            =   6120
            Picture         =   "frmMain.frx":3D4C
            Style           =   1  'Grafisch
            TabIndex        =   16
            Top             =   1680
            Width           =   492
         End
         Begin VB.TextBox txtChoice 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   3972
         End
         Begin VB.ComboBox cboStation 
            Height          =   288
            Index           =   4
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   2040
            Width           =   1812
         End
         Begin VB.CommandButton cmdGo 
            Height          =   252
            Index           =   4
            Left            =   6120
            Picture         =   "frmMain.frx":3DD8
            Style           =   1  'Grafisch
            TabIndex        =   19
            Top             =   2040
            Width           =   492
         End
         Begin VB.Label lblChoicesText 
            Caption         =   "Text"
            Height          =   252
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label lblChoicesStation 
            Caption         =   "Station"
            Height          =   252
            Left            =   4200
            TabIndex        =   27
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label lblChoicesGo 
            Caption         =   "Go"
            Height          =   252
            Left            =   6120
            TabIndex        =   26
            Top             =   360
            Width           =   492
         End
      End
      Begin VB.TextBox txtValidate 
         BackColor       =   &H8000000A&
         Height          =   612
         Left            =   -73800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   21
         Top             =   4080
         Width           =   5652
      End
      Begin VB.CommandButton cmdValidate 
         Caption         =   "&Validate"
         Enabled         =   0   'False
         Height          =   252
         Left            =   -74880
         TabIndex        =   20
         Top             =   4080
         Width           =   972
      End
      Begin SHDocVwCtl.WebBrowser webPreview 
         Height          =   4572
         Left            =   -74880
         TabIndex        =   22
         Top             =   120
         Width           =   6732
         ExtentX         =   11874
         ExtentY         =   8064
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
         Location        =   ""
      End
      Begin VB.TextBox txtStation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1932
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   4
         Top             =   120
         Width           =   6732
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Station content"
      Height          =   252
      Left            =   2280
      TabIndex        =   24
      Top             =   120
      Width           =   2292
   End
   Begin VB.Label Label4 
      Caption         =   "Stations overview"
      Height          =   252
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1932
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &as..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export..."
      End
      Begin VB.Menu mnuFilePrintview 
         Caption         =   "&Print view..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileValidateWithDTD 
         Caption         =   "Validate with &DTD"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileShowXML 
         Caption         =   "Show &XML at validation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreateHelp 
         Caption         =   "&Create Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuEditForward 
         Caption         =   "F&orward"
      End
      Begin VB.Menu mnuEditBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditContinueSearch 
         Caption         =   "&Continue Search"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWrapWithTag 
         Caption         =   "&Wrap Tag..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEditSelectTag 
         Caption         =   "Select tag"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &all"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuProjectStyle 
         Caption         =   "S&tyle..."
      End
      Begin VB.Menu mnuProjectSettings 
         Caption         =   "&Settings..."
      End
   End
   Begin VB.Menu mnuStation 
      Caption         =   "S&tation"
      Begin VB.Menu mnuStationSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuStationBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStationReferences 
         Caption         =   "&References"
      End
      Begin VB.Menu mnuStationFollowSelection 
         Caption         =   "&Follow selection"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuStationRename 
         Caption         =   "Re&name"
      End
      Begin VB.Menu mnuStationBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStationRevert 
         Caption         =   "&Revert"
      End
      Begin VB.Menu mnuStationDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuSourceInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuSourceInsertText 
         Caption         =   "&Text"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuInsertChoice 
         Caption         =   "&Choice"
         Begin VB.Menu mnuSourceInsertChoice 
            Caption         =   "&Choice"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuSourceInsertChoose 
            Caption         =   "C&hoose"
         End
         Begin VB.Menu mnuSourceInsertInput 
            Caption         =   "In&put"
         End
      End
      Begin VB.Menu mnuSourceInsertMedia 
         Caption         =   "&Media"
         Begin VB.Menu mnuSourceInsertImage 
            Caption         =   "&Image..."
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuSourceInsertImagemap 
            Caption         =   "I&mage Map"
            Shortcut        =   +{F6}
         End
         Begin VB.Menu mnuSourceInsertMusic 
            Caption         =   "M&usic..."
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu mnuInsertState 
         Caption         =   "&State"
         Begin VB.Menu mnuSourceInsertState 
            Caption         =   "&State"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuSourceInsertNumber 
            Caption         =   "&Number"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuSourceInsertString 
            Caption         =   "S&tring"
            Shortcut        =   ^{F8}
         End
      End
      Begin VB.Menu mnuSourceInsertIfElse 
         Caption         =   "If-&Else"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuInsertInclude 
         Caption         =   "&Include"
      End
      Begin VB.Menu mnuSourceInsertRandomize 
         Caption         =   "&Randomize"
      End
      Begin VB.Menu mnuInsertComment 
         Caption         =   "&Comment"
         Begin VB.Menu mnuSourceInsertComment 
            Caption         =   "&QML Comment"
         End
         Begin VB.Menu mnuSourceInsertXmlComment 
            Caption         =   "&XML Comment"
         End
      End
   End
   Begin VB.Menu mnuSyntax 
      Caption         =   "&Syntax"
      Visible         =   0   'False
      Begin VB.Menu mnuSyntaxInsert 
         Caption         =   "&Insert"
      End
      Begin VB.Menu mnuSyntaxHelp 
         Caption         =   "H&elp"
      End
   End
   Begin VB.Menu mnuStations 
      Caption         =   "Stations"
      Visible         =   0   'False
      Begin VB.Menu mnuStationsSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuStationsRevert 
         Caption         =   "Re&vert"
      End
      Begin VB.Menu mnuStationsBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStationsDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuProjectAnalyze 
         Caption         =   "&Analyze"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuToolsBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
         Begin VB.Menu mnuToolsOptionsAssistInline 
            Caption         =   "&Auto-complete"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "&View Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search Help"
      End
      Begin VB.Menu mnuHelpBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpVisitHomepage 
         Caption         =   "O&nline Help"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const runningFromIDE = True

Const questLocationHTMLNew = "project\new.htm"
Const defaultStartStation = "start"
Const windowMinWidth = 7284
Const windowMinHeight = 5208
Const treeviewChildOf = 4
Const g_programTitle = "QML-Edit 2"

Public objQuest As MSXML2.DOMDocument
Public qmlTopPath As String
Public browserStartURL As String
Public eventsOn As Boolean

Dim questLocationHTML As String
Dim questLocationXML As String
Dim beforeLastStationSelected As String
Dim lastStationSelected As String
Dim autoCreateChoice As Boolean
Dim autoHandleDebugMode As Boolean
Dim gTextToFind As String
Dim gTextChangeDidOccur As Boolean
Dim stationHistory As New clsHistory
Dim gExportFolder As String
Dim g_assistInline As Boolean

Private Sub Form_Load()
    Dim startedWithPath As String
    Dim applicationPath As String
    
    #If runningFromIDE Then
        applicationPath = _
                "D:\sites\questml\development"
    #Else
        applicationPath = App.path
    #End If
    
    loadAllSettings
    autoHandleDebugMode = True
    autoCreateChoice = True

    qmlTopPath = applicationPath
    eventsOn = True
    startedWithPath = Command$
    If startedWithPath <> "" Then
        loadQuestFile startedWithPath
    Else
        createNewFile
    End If
    webPreview.Navigate convertToURI(applicationPath & _
            "\project\blank.htm")
End Sub

Sub loadAllSettings()
    gExportFolder = GetSetting("qml", "settings", "exportFolder", "")
    g_assistInline = GetSetting("qml", "settings", "assistInline", True)
End Sub

Sub saveAllSettings()
    SaveSetting "qml", "settings", "exportFolder", gExportFolder
    SaveSetting "qml", "settings", "assistInline", g_assistInline
End Sub

Private Sub lstStations_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const rightClick = 2
    If Button = rightClick Then
        mnuStationFollowSelection.Visible = False
        PopupMenu mnuStation
        mnuStationFollowSelection.Visible = True
    End If
    
End Sub

Private Sub mnuEditBack_Click()
    goStationBack
End Sub

Private Sub mnuEditForward_Click()
    goStationForward
End Sub

Private Sub goStationBack()
    If updateFromActiveTab Then
        If stationHistory.back Then
            goToStation stationHistory.getValue, , False
        End If
    End If
End Sub

Private Sub goStationForward()
    If updateFromActiveTab Then
        If stationHistory.forward Then
            goToStation stationHistory.getValue, , False
        End If
    End If
End Sub

Private Sub goToStation(ByVal stationId As String, Optional ByVal doSelect As Boolean = True, Optional ByVal addToHistory As Boolean = True)
    If doSelect Then
        selectFromStationsOverview stationId
    End If
    loadStationIntoActiveTab stationId
    If addToHistory Then
        stationHistory.add stationId
    End If
    adaptBackForwardButtons
End Sub

Private Sub adaptBackForwardButtons()
    toolbar.Buttons("back").enabled = stationHistory.canGoBack
    toolbar.Buttons("forward").enabled = stationHistory.canGoForward
End Sub

Private Sub mnuFileExport_Click()
    exportQuest
End Sub

Private Sub exportQuest()
    Dim success As Boolean
    
    Select Case MsgBox("The file needs to be saved before exporting. Save now?", vbOKCancel, "Export")
        Case vbOK
            saveQuest success
        Case vbCancel
            success = False
    End Select
    
    If success Then
        doExportQuest
    Else
        MsgBox "Export cancelled.", , "Export"
    End If
End Sub

Private Sub doExportQuest()
    Dim dlgExport As frmExport
    Dim folderPath As String
    Dim projectName As String
    
    folderPath = ""
    projectName = nameOfFile(questLocationHTML, True)
    
    Set dlgExport = New frmExport
    With dlgExport
        .lblProjectname = projectName
        .lblFolder = gExportFolder
        .Show vbModal, Me
    
        If .lblFolder <> "" Then
            folderPath = .lblFolder
        End If
    End With

    Unload dlgExport
    If folderPath = "" Then
        MsgBox "Export cancelled", , "Export"
    Else
        gExportFolder = folderPath
        doExportQuestTo folderPath, projectName
    End If
End Sub

Private Sub doExportQuestTo(ByVal folderExportPath As String, ByVal projectName As String)
    Dim fileSystem As Object
    Dim folderCore As Object
    Dim fileSize As Long
    Dim projectFolder As String
    Dim doExport As Boolean
    Dim errorMessage As String
    
    errorMessage = ""
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    projectFolder = folderExportPath & "\" & projectName
    If fileSystem.folderExists(projectFolder) Then
        doExport = MsgBox("The project-folder exists. Overwrite?", vbOKCancel, "Export") = vbOK
    Else
        fileSystem.createFolder projectFolder
        doExport = True
    End If
    
    If doExport Then
        fileSystem.copyFolder qmlTopPath & "\tool\core", _
                projectFolder
        fileSystem.copyFile questLocationHTML, _
                projectFolder & "\" & projectName & ".htm"
        fileSystem.copyFile questLocationXML, _
                projectFolder & "\quest\" & projectName & ".xml"
        
        exportCopyScripts fileSystem, projectFolder

        errorMessage = exportMedia(projectFolder, fileSystem)
        
        fileSize = fileSystem.getFolder(projectFolder).Size
        fileSize = fileSize / 1024
        
        adaptReadme projectFolder
        
        MsgBox "Exported quest to " & projectFolder & _
                vbNewLine & "Project-Size: " & Str$(fileSize) & " KB" & _
                errorMessage, vbInformation, "Export finished"
    Else
        MsgBox "Export cancelled", , "Export"
    End If
End Sub

Sub exportCopyScripts(ByRef fileSystem As Object, ByVal projectFolder As String)
    Dim source As String
    Dim export As MSXML2.DOMDocument
    Dim fileNodes As IXMLDOMNodeList
    Dim fileNode As IXMLDOMElement
    
    Set export = getXml(qmlTopPath & "\tool\export.xml")
    Set fileNodes = export.selectNodes("//file")
    For Each fileNode In fileNodes
        source = fileNode.getAttribute("source")
        fileSystem.copyFile qmlTopPath & "\" & source, _
                projectFolder & "\" & source
    Next
End Sub

Private Sub adaptReadme(ByVal projectFolder As String)
    Dim text As String
    Dim readmePath As String
    
    readmePath = projectFolder & "\readme.txt"
    text = getFileText(readmePath)
    text = Replace(text, "[info]", getReadmeInfo)
    setFileText readmePath, text
End Sub

Private Function getReadmeInfo()
    Dim aboutNode As IXMLDOMElement
    Dim info As String
    info = ""
    
    Set aboutNode = objQuest.selectSingleNode("//about")
    If Not (aboutNode Is Nothing) Then
        info = info & getReadmeInfoPart(aboutNode, "title", "")
        info = info & getReadmeInfoPart(aboutNode, "author", "By ")
        info = info & vbNewLine
        info = info & getReadmeInfoPart(aboutNode, "intro", "")
        info = info & vbNewLine
        info = info & getReadmeInfoPart(aboutNode, "date", "Date of creation: ")
        info = info & getReadmeInfoPart(aboutNode, "homepage", "Homepage: ")
        info = info & getReadmeInfoPart(aboutNode, "email", "Email: ")
    End If
    
    getReadmeInfo = info
End Function

Private Function getReadmeInfoPart(ByRef aboutNode As IXMLDOMElement, ByVal nodeName As String, ByVal precede As String, Optional ByVal upperCase As Boolean = False) As String
    Dim infoNode As IXMLDOMElement
    Dim info As String
    Dim text As String
    info = ""
    
    Set infoNode = aboutNode.selectSingleNode(nodeName)
    If Not (infoNode Is Nothing) Then
        text = infoNode.text
        If upperCase Then text = UCase$(text)
        info = precede & text & vbNewLine
    End If
    
    getReadmeInfoPart = info
End Function

Private Function exportMedia(ByVal projectFolder As String, ByRef fileSystem As Object) As String
    Dim errorMessage As String
    errorMessage = ""
    
    frmMain.MousePointer = vbArrowHourglass
    frmMain.Refresh
    
    errorMessage = errorMessage & exportMediaStations(projectFolder, fileSystem)
    errorMessage = errorMessage & exportMediaBackground(projectFolder, fileSystem)
    errorMessage = errorMessage & exportMediaStyle(projectFolder, fileSystem)

    If errorMessage <> "" Then
        errorMessage = vbNewLine & vbNewLine & _
            "Some of the media could not be found. The cause " & _
            "can be either a wrong path in the source, or dynamically " & _
            "generated media using inline-values in source attributes." & _
            errorMessage
    End If

    frmMain.MousePointer = vbDefault
    
    exportMedia = errorMessage
End Function

Private Function exportMediaStations(ByVal projectFolder As String, ByRef fileSystem As Object) As String
    Const seperator = ";"
    Dim stationId As String
    Dim i As Long
    Dim errorMessage As String
    Dim source As String
    Dim xPath As String
    Dim sourceFrom As String
    Dim sourceTo As String
    Dim mediaElements As IXMLDOMNodeList
    Dim mediaElement As IXMLDOMElement
    Dim station As IXMLDOMElement
    Dim objStatus As frmStatus
    Dim loadedStatus As Boolean
    Dim handledMedia As String
    Dim fileHasBeenHandled As Boolean
    errorMessage = ""
    handledMedia = seperator
    loadedStatus = False
    
    xPath = "//image[@source] | //music[@source]"
    Set mediaElements = objQuest.selectNodes(xPath)
    
    If mediaElements.length > 0 Then
        Set objStatus = New frmStatus
        loadedStatus = True
        objStatus.barProgress.min = 0
        objStatus.barProgress.max = mediaElements.length
        objStatus.barProgress.Value = 0
        objStatus.lblTask = "Copying media..."
        objStatus.Show
        objStatus.Refresh
    End If
    
    For Each mediaElement In mediaElements
        i = i + 1
        objStatus.barProgress.Value = i
        
        source = mediaElement.getAttribute("source")
        source = Replace(source, "/", "\")
        
        fileHasBeenHandled = InStr(handledMedia, seperator & source & seperator) >= 1
        If Not fileHasBeenHandled Then
            sourceFrom = qmlTopPath & "\" & source
            
            If fileSystem.fileExists(sourceFrom) Then
                sourceTo = projectFolder & "\" & source
                createFoldersToCopyFile fileSystem, projectFolder, source
                fileSystem.copyFile sourceFrom, sourceTo
                
            Else
                If LCase$(source) <> "default" Then
                    Set station = mediaElement.parentNode
                    stationId = station.getAttribute("id")
                    errorMessage = errorMessage & vbNewLine & _
                            " - In station """ & stationId & """ the " & _
                            mediaElement.nodeName & " """ & source & """ "
                End If
            End If
            
            handledMedia = handledMedia & source & seperator
        End If
    Next
    
    If loadedStatus Then Unload objStatus
    
    exportMediaStations = errorMessage
End Function

Private Function exportMediaBackground(ByVal projectFolder As String, ByRef fileSystem As Object) As String
    Const urlStart = "url("
    Const urlEnd = ")"
    Dim errorMessage As String
    Dim source As String
    Dim xPath As String
    Dim sourceFrom As String
    Dim sourceTo As String
    Dim backgroundWithImage As IXMLDOMElement
    errorMessage = ""
        
    xPath = "//style/background[@image]"
    Set backgroundWithImage = objQuest.selectSingleNode(xPath)
    If Not (backgroundWithImage Is Nothing) Then
        source = backgroundWithImage.getAttribute("image")
        source = Replace(source, "/", "\")
        source = Replace(source, urlStart, "")
        source = Replace(source, urlEnd, "")
        sourceFrom = qmlTopPath & "\" & source
        sourceTo = projectFolder & "\" & source
        If fileSystem.fileExists(sourceFrom) Then
            sourceTo = projectFolder & "\" & source
            createFoldersToCopyFile fileSystem, projectFolder, source
            fileSystem.copyFile sourceFrom, sourceTo
        Else
            If LCase$(source) <> "default" Then
                errorMessage = errorMessage & vbNewLine & _
                        " - Background-image """ & sourceFrom & """"
            End If
        End If
    End If
    
    exportMediaBackground = errorMessage
End Function

Private Function exportMediaStyle(ByVal projectFolder As String, ByRef fileSystem As Object) As String
    Const urlStart = "url("
    Const urlEnd = ")"
    Dim errorMessage As String
    Dim classStyles As IXMLDOMNodeList
    Dim classStyle As IXMLDOMElement
    Dim startPosition As Integer
    Dim endPosition As Integer
    Dim styleString As String
    Dim source As String
    Dim xPath As String
    Dim sourceFrom As String
    Dim sourceTo As String
    errorMessage = ""
    
    xPath = "//style/class[@style]"
    Set classStyles = objQuest.selectNodes(xPath)
    For Each classStyle In classStyles
        styleString = classStyle.getAttribute("style")
        startPosition = InStr(LCase$(styleString), urlStart)
        If startPosition >= 1 Then
            startPosition = startPosition + Len(urlStart)
            endPosition = InStr(startPosition + 1, styleString, urlEnd)
            If endPosition >= 1 Then
                source = Mid$(styleString, startPosition, endPosition - startPosition)
                
                source = Replace(source, "/", "\")
                source = Replace(source, """", "")
                source = Replace(source, "'", "")
                
                sourceFrom = qmlTopPath & "\" & source
                sourceTo = projectFolder & "\" & source
                If fileSystem.fileExists(sourceFrom) Then
                    sourceTo = projectFolder & "\" & source
                    createFoldersToCopyFile fileSystem, projectFolder, source
                    fileSystem.copyFile sourceFrom, sourceTo
                Else
                    errorMessage = errorMessage & vbNewLine & _
                            " - Class-style media """ & sourceFrom & """"
                End If
                
            End If
        End If
    Next
    
    exportMediaStyle = errorMessage
End Function

Private Sub createFoldersToCopyFile(ByRef fileSystem As Object, ByVal absolutePath As String, ByVal relativePath As String)
    Dim subFolder As String
    Dim nextPosition As Integer
    Dim lastPosition As Integer
    Dim subFolderFound As Boolean
    Dim absoluteSubPath As String
    
    lastPosition = 0
    Do
        nextPosition = InStr(lastPosition + 1, relativePath, "\")
        If nextPosition >= 1 Then
            subFolderFound = True
            subFolder = Left$(relativePath, nextPosition - 1)
            absoluteSubPath = absolutePath & "\" & subFolder
            If Not fileSystem.folderExists(absoluteSubPath) Then
                fileSystem.createFolder absoluteSubPath
            End If
            lastPosition = nextPosition
        Else
            subFolderFound = False
        End If
    Loop Until Not subFolderFound
End Sub

Private Sub mnuFileShowXML_Click()
    mnuFileShowXML.Checked = Not mnuFileShowXML.Checked
End Sub

Private Sub mnuFileValidateWithDTD_Click()
    mnuFileValidateWithDTD.Checked = Not mnuFileValidateWithDTD.Checked
End Sub

Private Sub mnuHelpSearch_Click()
    Dim keyword As String
    Dim template As String
    Dim results As String
    Dim dlghelp As Form
    Dim filePath As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim subFile As Object
    Dim thisFile As String
    Dim thisLink As String
    Dim thisTitle As String
    Static lastKeyword As String
    
    results = ""
    keyword = InputBox("Enter keyword:", "Search help", lastKeyword)
    If keyword <> "" Then
        lastKeyword = keyword
    End If
    
    template = getFileText(qmlTopPath & _
            "\tool\qmledit\search_temp.tpl")
    filePath = qmlTopPath & "\help\search_temp.htm"
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.getFolder(qmlTopPath & "\help")
    For Each subFile In folder.Files
        If InStr(subFile.path, ".htm") >= 1 And _
                subFile.path <> filePath Then
            thisFile = getFileText(subFile.path)
            If InStr(LCase$(thisFile), keyword) >= 1 Then
                thisTitle = getTextBetween(thisFile, "<title>", "</title>")
                thisTitle = Replace$(thisTitle, " [QML Syntax]", "")
                thisLink = "<a href=""" & _
                        convertToURI(subFile.path) & """>" & _
                        thisTitle & "</a>"
                results = results & "<li>" & thisLink & "</li>"
            End If
        End If
    Next
    
    If results = "" Then
        results = "<li><em>No results found.</em></li>"
    End If
    results = "<ul>" & results & "</ul>"
            
    template = Replace$(template, "[keyword]", keyword)
    template = Replace$(template, "[results]", results)

    setFileText filePath, template
    Set dlghelp = New frmBrowser
    browserStartURL = convertToURI(filePath)
    dlghelp.Show
End Sub

Private Sub mnuHelpVisitHomepage_Click()
    visitQMLHomepage
End Sub

Private Sub mnuStationReferences_click()
    Dim dlgSyntax As Form
    Dim template As String
    Dim filePath As String
    Dim references As IXMLDOMNodeList
    Dim reference As IXMLDOMElement
    Dim stationId As String
    Dim station As IXMLDOMElement
    Dim referencesText As String
    Dim xPath(1 To 2) As String
    Dim xPath2 As String
    Dim i As Integer
    Dim thisText As String
    
    template = getFileText(qmlTopPath & _
            "\tool\qmledit\station_temp.tpl")
    
    template = Replace$(template, "[stationId]", lastStationSelected)
    
    referencesText = ""
    xPath(1) = "//choice[@station = '" & lastStationSelected & "']"
    xPath(2) = "//choose[@station = '" & lastStationSelected & "']"
    For i = 1 To 2
        
        Set references = objQuest.documentElement.selectNodes(xPath(i))
        For Each reference In references
            xPath2 = "ancestor(station)"
            Set station = reference.selectSingleNode(xPath2)
            stationId = station.getAttribute("id")
            If reference.nodeName = "choice" Then
                thisText = reference.text
            Else
                thisText = convertXMLtoText(reference.xml)
            End If
            referencesText = referencesText & _
                    "<li><strong>" & thisText & "</strong> <em>(from: " & stationId & ")</em></li>"
        Next
    Next
    
    If referencesText = "" Then
        referencesText = "<li><em>None</em></li>"
    End If
    template = Replace$(template, "[references]", referencesText)
    
    filePath = qmlTopPath & "\help\station_temp.htm"
    setFileText filePath, template
    
    Set dlgSyntax = New frmBrowser
    browserStartURL = convertToURI(filePath)
    dlgSyntax.Show
End Sub

Private Sub mnuStationRename_Click()
    Const dialogTitle = "QML Station Rename"
    Dim success As Boolean
    Dim newName As String
    Dim currentStation As IXMLDOMElement
    Dim xPath As String
    Dim oldName As String
    Dim adapt As Integer
    
    oldName = lastStationSelected
    newName = InputBox("Enter new name:", dialogTitle, oldName)
    If newName <> "" And newName <> oldName Then
        If stationExists(newName) Then
            MsgBox "A station by the name """ & newName & _
                    """ exists already.", vbCritical, dialogTitle
        Else
            adapt = MsgBox("Adapt choices leading to this station?", _
                    vbYesNoCancel, dialogTitle)
            
            If adapt <> vbCancel Then
                lstStations.List(lstStations.ListIndex) = newName
                lastStationSelected = newName

                xPath = "./station[@id = '" & oldName & "']"
                Set currentStation = objQuest.documentElement. _
                        selectSingleNode(xPath)
                currentStation.setAttribute "id", newName
                
                loadStationIntoActiveTab newName
                
                If adapt = vbYes Then
                    adaptToStationRename oldName, newName
                End If
                
            End If
        End If
        
    End If
End Sub

Sub adaptToStationRename(ByVal oldName As String, ByVal newName As String)
    Dim elements As IXMLDOMNodeList
    Dim element As IXMLDOMElement
    Dim xPath As String
    Dim visit As String
    Dim oldVisit As String
    Dim newVisit As String
    Dim oldText As String
    Dim child As IXMLDOMNode
        
    xPath = "//choice[@station = '" & oldName & "'] | " & _
            "//choose[@station = '" & oldName & "']"
    Set elements = objQuest.documentElement.selectNodes(xPath)

    For Each element In elements
        element.setAttribute "station", newName
    Next
    
    xPath = "//*[@number]"
    Set elements = objQuest.documentElement.selectNodes(xPath)
    For Each element In elements
        visit = "visits(" & oldName & ")"
        oldVisit = element.getAttribute("number")
        If InStr(1, oldVisit, visit, vbTextCompare) Then
            newVisit = Replace$(oldVisit, visit, "visits(" & newName & ")", , , vbTextCompare)
            element.setAttribute "number", newVisit
        End If
    Next

    xPath = "//text"
    Set elements = objQuest.documentElement.selectNodes(xPath)
    For Each element In elements
        For Each child In element.childNodes
            
            If child.nodeName = "#text" Then
                oldText = child.text
                If InStr(1, oldText, "[number visits(" & oldName & ")", vbTextCompare) >= 1 Then
                    child.text = Replace$(oldText, "[number visits(" & oldName & ")", _
                            "[number visits(" & newName & ")", , , vbTextCompare)
                End If
            
            End If
        Next
    Next
End Sub

Private Sub cmdAdd_Click()
    Dim success As Boolean
    
    txtStationName.text = txtStationName.text
    If txtStationName.text = "" Then
        MsgBox "You need to enter a station name."
    Else
        createStationAndJumpToIt txtStationName.text, success
        If success Then
            txtStationName.text = ""
            txtStationName.SetFocus
        Else
            selectAllText txtStationName
        End If
    End If
End Sub

Private Sub createStationAndJumpToIt(ByVal stationId As String, Optional ByRef success As Boolean = False)
    addStationToList stationId, True, success
    If success Then
        addStationToXML stationId
        If updateFromActiveTab Then
            goToStation stationId
            txtSource.SelStart = Len("<text>")
            txtSource.SelLength = Len(" ")
        End If
    End If
End Sub

Private Sub cmdGo_Click(index As Integer)
    Dim stationId As String
    Dim success As Boolean
    
    success = True
    'cboStation.Item(index).text = cboStation.Item(index).text
    stationId = cboStation.Item(index).text
    success = True
    
    If stationId = stationBackName Then
        stationId = beforeLastStationSelected
        If stationId <> "" Then
            If updateFromActiveTab Then
                goToStation stationId
            End If
        End If
    Else
        goCreateStation stationId, success
    End If
    
    If tabMain.Tab = tabType.tabEditor Then
        txtStation.SetFocus
    Else
        selectAllText cboStation.Item(index)
    End If
End Sub

Private Function goCreateStation(ByVal stationId As String, Optional ByVal success As Boolean = False)
    If stationExists(stationId) Then
        If updateFromActiveTab Then
            goToStation stationId
            success = True
        Else
            success = False
        End If
    ElseIf vbYes = MsgBox("Station doesn't exist. Create it?", vbYesNoCancel) Then
        createStationAndJumpToIt stationId
        success = True
    Else
        success = False
    End If
    
    goCreateStation = success
End Function

Private Sub cmdValidate_Click()
    Dim oldStart As Integer
    Dim oldsellength As Integer
    
    oldStart = txtSource.SelStart
    oldsellength = txtSource.SelLength
    If updateFromTabSource Then
        txtSource.SetFocus
        txtSource.SelStart = oldStart
        txtSource.SelLength = oldsellength
    End If
    
    cmdValidate.enabled = False
End Sub

Private Sub Form_Resize()
    If eventsOn Then
        If Me.WindowState <> vbMinimized Then
            eventsOn = False
            If Width < windowMinWidth Then
                Width = windowMinWidth
            End If
            If Height < windowMinHeight Then
                Height = windowMinHeight
            End If
            eventsOn = True
            
            resizeControlsDynamically
        End If
    End If
End Sub

Private Sub resizeControlsDynamically()
    Dim i As Integer
    Dim temp As Integer
    
    With Me
        .txtStationName.Top = Height - 924
        .cmdAdd.Top = .txtStationName.Top
        .lstStations.Height = .txtStationName.Top - 380
        
        .tabMain.Height = Height - 1032
        .cmdValidate.Top = Height - 2124
        .txtSource.Height = .cmdValidate.Top - 348
        .txtValidate.Top = .cmdValidate.Top
        
        .webPreview.Height = .tabMain.Height - 720
        
        .tabMain.Width = Width - 2448
        .txtStation.Width = .tabMain.Width - 240
        
        temp = .tabMain.Width - 240
        .txtSource.Width = temp
        
        temp = .tabMain.Width - 1070 - 240
        .txtValidate.Width = temp
        
        .webPreview.Width = .tabMain.Width - 240

        .frmChoices.Top = .tabMain.Height - 3012
        .frmChoices.Width = .tabMain.Width - 240
        .txtStation.Height = .frmChoices.Top - 268
        
        For i = .cboStation.LBound To .cboStation.UBound
            .txtChoice.Item(i).Width = .frmChoices.Width - 2760
            .cboStation.Item(i).Left = .frmChoices.Width - 2532
            .cmdGo.Item(i).Left = .frmChoices.Width - 612
        Next
        .lblChoicesStation.Left = _
                .cboStation.Item(.cboStation.LBound).Left
        .lblChoicesGo.Left = _
                .cmdGo.Item(.cmdGo.LBound).Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim success As Boolean
    
    Select Case _
        MsgBox("Do you want to save the Quest before exiting?", _
               vbYesNoCancel)
        Case vbYes
            success = True
            saveQuest success
            If Not success Then Cancel = 1
        Case vbNo
        Case vbCancel
            Cancel = 1
    End Select
    
    saveAllSettings
End Sub

Private Sub lstStations_Click()
    If eventsOn Then
        If updateFromActiveTab(True) Then
            goToStation getSelectedStation, False
        Else
            eventsOn = False
            selectFromStationsOverview lastStationSelected
            eventsOn = True
        End If
    End If
    beforeLastStationSelected = lastStationSelected
    lastStationSelected = getSelectedStation
End Sub

Private Sub lstStations_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        tryDeleteStation
    End If
End Sub

Private Sub mnuEditPaste_Click()
    selTextFromClipboard
End Sub

Private Sub mnuEditCopy_Click()
    selTextToClipboard
End Sub

Private Sub mnuEditCut_Click()
    selTextToClipboard
    selTextClear
End Sub

Private Sub selTextToClipboard()
    If TypeOf ActiveControl Is textBox Or TypeOf ActiveControl Is RichTextBox Then
        If ActiveControl.SelLength >= 1 Then
            Clipboard.SetText ActiveControl.SelText
        End If
    End If
End Sub

Private Sub selTextFromClipboard()
    If TypeOf ActiveControl Is textBox Or TypeOf ActiveControl Is RichTextBox Then
        ActiveControl.SelText = Clipboard.GetText
    End If
End Sub

Private Sub selTextClear()
    If TypeOf ActiveControl Is textBox Or TypeOf ActiveControl Is RichTextBox Then
        ActiveControl.SelText = ""
    End If
End Sub

Private Sub mnuEditSelectAll_Click()
    On Error Resume Next
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.text)
    On Error GoTo 0
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    If vbOK = MsgBox("Unsaved data will be lost.", vbOKCancel) Then
        createNewFile
    End If
End Sub

Private Sub mnuFileOpen_Click()
    If (vbOK = MsgBox("Unsaved data will be lost.", vbOKCancel)) Then
        openQuest
    End If
End Sub

Private Sub openQuest()
    With dlgCommon
        .dialogTitle = "Open quest"
        .Filter = "*.htm|*.htm"
        .InitDir = qmlTopPath
        .ShowOpen
        If .FileName <> "" Then
            loadQuestFile .FileName
        End If
    End With
End Sub

Private Sub saveQuest(Optional ByRef success As Boolean = False)
    success = False
    If updateFromActiveTab Then
        
        If questLocationXML = "" Then
            findFileToSaveTo success
        End If
        
        If questLocationXML <> "" Then
            handleDebugModeSave
            objQuest.save questLocationXML
            reformatQuestXMLfile questLocationXML
            questLocationXML = questLocationXML
            handleDebugModeLoad
            success = True
        Else
            success = False
        End If
    
    Else
        success = False
    End If
End Sub

Private Sub setFileSaveLocation()
    dlgCommon.Filter = "*.htm"
    dlgCommon.ShowOpen
    loadQuestFile dlgCommon.FileName
End Sub

Private Sub mnuFilePrintview_Click()
    Dim dlgPrintview As Form
    Dim template As String
    Dim filePath As String
    Dim titleElement As MSXML2.IXMLDOMElement
    Dim title As String

    If updateFromActiveTab Then
        title = ""
        Set titleElement = objQuest.selectSingleNode("//title")
        If Not (titleElement Is Nothing) Then
            title = titleElement.text
        End If
        template = getFileText(qmlTopPath & "\tool\print\index.tpl")
        template = Replace$(template, "[title]", title)
        template = Replace$(template, "[content]", getXhtmlForPrint)
        filePath = qmlTopPath & "\tool\print\index.htm"
        setFileText filePath, template

        Set dlgPrintview = New frmBrowser
        browserStartURL = convertToURI(filePath)
        dlgPrintview.Show
    End If
End Sub

Private Sub mnuFileSave_Click()
    saveQuest
End Sub

Private Sub mnuFileSaveAs_Click()
    findFileToSaveTo
    If dlgCommon.FileName <> "" Then
        saveQuest
    End If
End Sub

Private Sub findFileToSaveTo(Optional ByRef success As Boolean = False)
    Dim htmlFile As String
    Dim projectName As String

    With dlgCommon
        .DefaultExt = "htm"
        .dialogTitle = "Choose a place to save in the QML folder"
        .InitDir = qmlTopPath & "\"
        .FileName = ""
        .Filter = "*.htm"
        .ShowSave
    End With
    
    If dlgCommon.FileName <> "" Then
        If fileExists(dlgCommon.FileName) Then
            If Not (vbYes = MsgBox("Overwrite existing file?", vbYesNoCancel)) Then
                dlgCommon.FileName = ""
            End If
        End If
    End If
    
    If dlgCommon.FileName <> "" Then
        If InStr(dlgCommon.FileName, ".htm") < 1 Then
            dlgCommon.FileName = dlgCommon.FileName & ".htm"
        End If
        
        If LCase$(Left$(dlgCommon.FileName, InStrRev(dlgCommon.FileName, "\") - 1)) = LCase$(qmlTopPath) Then
            projectName = nameOfFile(dlgCommon.FileName, True)
            
            htmlFile = getFileText(qmlTopPath & "\tool\" & questLocationHTMLNew, success)
            htmlFile = Replace$(htmlFile, "[quest]", "quest/" & projectName)
            htmlFile = Replace$(htmlFile, "[station]", "start")
            
            If success Then
                setFileText dlgCommon.FileName, htmlFile, success
                If success Then
                    questLocationHTML = dlgCommon.FileName
                    updateMeCaption
                    questLocationXML = getXmlFilePath(htmlFile, questLocationHTML)
                End If
            End If
        Else
            MsgBox "You need to save in the program folder."
        End If
    End If
End Sub

Private Sub updateMeCaption()
    If questLocationHTML <> "" Then
        Caption = nameOfFile(questLocationHTML, True) & " - " & _
                g_programTitle
    End If
End Sub

Private Sub mnuEditSelectTag_Click()
    If Me.ActiveControl = Me.txtSource Then
        If isInTag Then
            selectSurround "<", ">"
        Else
            selectSurround ">", "<", True
        End If
    End If
End Sub

Private Function isInTag() As Boolean
    With txtSource
        If InStr(.SelStart + 1, .text, "<") <= _
                InStr(.SelStart + 1, .text, ">") Then
            isInTag = True
        Else
            isInTag = False
        End If
    End With
End Function

Private Sub selectSurround(ByVal char1 As String, ByVal char2 As String, Optional ByVal includeChars As Boolean = False)
    Dim lastOccBefore As Long, firstOccAfter As Long, _
            startAt As Long, lengthOf As Long

    With txtSource
        lastOccBefore = InStrRev(Mid$(.text, 1, _
                .SelStart), char2)
        If lastOccBefore >= 1 Then
            firstOccAfter = InStr(lastOccBefore, .text, _
                    char1, vbTextCompare)
            lengthOf = (firstOccAfter - lastOccBefore) - 1
            startAt = lastOccBefore
            If includeChars Then
                startAt = startAt - Len(char1)
                lengthOf = lengthOf + Len(char2) + 1
            End If
            If lengthOf >= 1 Then
                .SelStart = startAt
                .SelLength = lengthOf
            End If
        End If
    End With
End Sub

Private Sub mnuHelpAbout_Click()
    Dim dlgAbout As Form
    
    Set dlgAbout = New frmAbout
    dlgAbout.Show vbModal, Me
    Unload dlgAbout
End Sub

Private Sub createNewFile()
    loadQuestFile qmlTopPath & "\tool\" & questLocationHTMLNew
    questLocationHTML = ""
    questLocationXML = ""
    updateMeCaption
End Sub

Private Sub loadQuestFile(ByVal htmlPath As String)
    If setQuestObject(htmlPath) Then
        resetEditorFromQuestObject
    End If
End Sub

Private Sub resetEditorFromQuestObject()
    eventsOn = False
    resetEditor
    resetStationsOverview
    selectFromStationsOverview defaultStartStation
    stationHistory.add defaultStartStation
    eventsOn = True
    handleDebugModeLoad
    loadStationIntoActiveTab defaultStartStation
    menuSourceOnly tabMain.Tab = tabType.tabsource
    updateMeCaption
End Sub

Private Sub handleDebugModeLoad()
    If autoHandleDebugMode Then
        objQuest.documentElement.setAttribute "debug", "true"
    End If
End Sub

Private Sub handleDebugModeSave()
    If autoHandleDebugMode Then
        objQuest.documentElement.setAttribute "debug", "false"
    End If
End Sub

Private Sub resetEditor()
    Dim i As Integer
    
    stationHistory.clear
    beforeLastStationSelected = ""
    lastStationSelected = ""
    lstStations.clear
    txtStation.text = ""
    For i = cboStation.LBound To cboStation.UBound
        cboStation.Item(i).clear
        txtChoice.Item(i).text = ""
    Next
End Sub

Private Sub selectFromStationsOverview(ByVal stationId As String)
    Dim i As Integer, indexOfId, eventsOnWas As Boolean
    
    For i = 0 To lstStations.ListCount - 1
        If lstStations.List(i) = stationId Then
            indexOfId = i
            Exit For
        End If
    Next
    eventsOnWas = eventsOn
    eventsOn = False
    lstStations.ListIndex = indexOfId
    eventsOn = eventsOnWas
End Sub

Private Sub loadSelectedStationIntoActiveTab()
    loadStationIntoActiveTab getSelectedStation
End Sub

Private Function getSelectedStation() As String
    getSelectedStation = lstStations.List(lstStations.ListIndex)
End Function

Private Sub fillChoiceStations()
    Dim i As Integer, j As Integer
    For i = cboStation.LBound To cboStation.UBound
        For j = 0 To lstStations.ListCount - 1
            cboStation.Item(i).AddItem lstStations.List(j)
        Next
        cboStation.Item(i).AddItem stationBackName
    Next
End Sub

Private Sub fillChoices(ByRef objStation As IXMLDOMElement)
    Dim child As IXMLDOMNode
    Dim choiceI As Integer
    Dim childAsNode As IXMLDOMElement
    choiceI = 0
    
    For Each child In objStation.childNodes
        If child.nodeName = elementPath Then
            If child.hasChildNodes Then
                txtChoice(choiceI).text = child.firstChild.text
            End If
            Set childAsNode = child
            cboStation.Item(choiceI).text = childAsNode.getAttribute("station")
            choiceI = choiceI + 1
        End If
    Next
    
End Sub

Private Sub emptyChoicesAndStations()
    Dim i As Integer
    For i = cboStation.LBound To cboStation.UBound
        txtChoice.Item(i).text = ""
        Me.cboStation.Item(i).clear
    Next
End Sub

Private Function getTextFromStation(ByRef objStation As IXMLDOMNode) As String
    Dim child As IXMLDOMNode, textFromStation As String
    
    For Each child In objStation.childNodes
        If child.nodeName = elementText Then
            textFromStation = getTextForEditor(child.xml)
            Exit For
        End If
    Next
        
    getTextFromStation = textFromStation
End Function

Private Function getTextForEditor(ByVal text As String) As String
    text = Replace$(text, "<text>", "")
    text = Replace$(text, "</text>", "")
    text = Replace$(text, "<break type=""strong""/>", vbNewLine & vbNewLine)
    text = Replace$(text, "<break/>", vbNewLine)
    text = Replace$(text, "&lt;", "<")
    text = Replace$(text, "&gt;", ">")
    text = Replace$(text, "&amp;", "&")
    text = rTrimNewline(text)
    getTextForEditor = text
End Function

Public Function setStationFromSelectedId(ByRef objStation As IXMLDOMElement)
    Dim stationId As String
    stationId = getSelectedStation
    setStationFromId objStation, stationId
End Function

Private Function setStationFromId(ByRef objStation As IXMLDOMElement, ByVal id As String)
    Dim xPath As String
    xPath = "./station[@id = '" & id & "']"
    Set objStation = objQuest.documentElement.selectSingleNode(xPath)
End Function

Private Sub resetStationsOverview()
    Dim stations As IXMLDOMNodeList
    Dim station As IXMLDOMElement
    Dim stationId As String
    
    Set stations = objQuest.documentElement.selectNodes("./station")
    For Each station In stations
        stationId = station.getAttribute("id")
        addStationToList stationId, False
    Next
End Sub

Private Function addStationToList(ByVal stationId As String, Optional ByVal checkUniqueness As Boolean = True, Optional ByRef success As Boolean = False)
    Dim isUnique As Boolean
    Dim eventsOnWas As Boolean
    
    If Len(stationId) >= 1 And Len(stationId) <= 255 Then
        If stationId <> stationBackName Then
            
            If checkUniqueness Then
                isUnique = Not stationExists(stationId)
            Else
                isUnique = True
            End If
            
            eventsOnWas = eventsOn
            eventsOn = False
            
            If isUnique Then
                lstStations.AddItem stationId
                success = True
            Else
                MsgBox "Station ID must be unique." & vbNewLine & _
                       "Adding cancelled.", vbExclamation
                success = False
            End If
            
            eventsOn = eventsOnWas
            
        Else
            MsgBox "The word """ & stationBackName & """ is reserved to take the user" & vbNewLine & _
                    "to the last visited station.", vbExclamation
        End If
    Else
        MsgBox "Station-IDs must have between 1 and 255 letters.", vbExclamation
        success = False
    End If
    
    addStationToList = success
End Function

Public Function stationExists(ByVal stationId As String) As Boolean
    Dim i As Integer, stationFound As Boolean
    For i = 0 To lstStations.ListCount - 1
        If lstStations.List(i) = stationId Then
            stationFound = True
            Exit For
        End If
    Next
    
    stationExists = stationFound
End Function

Private Function setQuestObject(ByVal htmlPath As String) As Boolean
    Dim htmlFile As String
    Dim xmlFilePath As String
    Dim success As Boolean
    
    htmlFile = getFileText(htmlPath, success)
    If success Then
        htmlFile = Replace$(htmlFile, "[quest]", "quest/" & nameOfFile(htmlPath, True))
        htmlFile = Replace$(htmlFile, "[station]", "")
        
        xmlFilePath = getXmlFilePath(htmlFile, htmlPath)
        
        success = validXML(xmlFilePath)
        If success Then
            loadXML xmlFilePath, objQuest
            questLocationXML = xmlFilePath
            questLocationHTML = htmlPath
        End If
    End If
    
    setQuestObject = success
End Function

Private Function getXmlFilePath(ByVal htmlFile As String, ByVal htmlPath As String) As String
    Dim xmlFileURL As String
    Dim xmlFilePath As String
    
    xmlFileURL = getTextBetween(htmlFile, "handleStation('", "'")
    xmlFilePath = Replace$(xmlFileURL, "/", "\")
    xmlFilePath = getPathOf(htmlPath) & xmlFilePath & ".xml"
    getXmlFilePath = xmlFilePath
End Function

Private Function loadXML(ByVal xmlPath As String, ByRef oXML As MSXML2.DOMDocument)
    Set oXML = New MSXML2.DOMDocument
    
    oXML.validateOnParse = True
    oXML.async = False
    oXML.Load xmlPath
End Function

Private Function validXML(ByVal xmlPath As String, Optional ByVal displayError As Boolean = True) As Boolean
    Dim oDoc As MSXML2.DOMDocument
    Dim success As Boolean
    
    Set oDoc = New MSXML2.DOMDocument
    
    oDoc.validateOnParse = True
    oDoc.async = False
    oDoc.Load xmlPath
    
    If oDoc.parseError.errorCode = 0 Then
        success = True
    Else
        success = False
        If displayError Then
            MsgBox oDoc.parseError.reason & vbCrLf & _
                oDoc.parseError.line & vbCrLf & _
                oDoc.parseError.srcText, vbCritical, "QML: Erronous XML (" & xmlPath & ")"
        End If
    End If
    validXML = success
End Function

Private Sub mnuHelpContent_Click()
    Dim dlghelp As Form
    
    Set dlghelp = New frmBrowser
    browserStartURL = convertToURI(frmMain.qmlTopPath) & _
            "/help/index.htm"
    dlghelp.Show
End Sub

Private Sub mnuEditContinueSearch_Click()
    If gTextToFind = "" Then
        settexttofind
    Else
        continueSearch
    End If
End Sub

Private Sub mnuEditFind_Click()
    settexttofind
End Sub

Private Sub settexttofind()
    Dim dlgFind As frmFind
    Set dlgFind = New frmFind
    If txtSource.SelText <> "" Then
        dlgFind.suggestedText = txtSource.SelText
    Else
        dlgFind.suggestedText = gTextToFind
    End If
    dlgFind.Show vbModal, Me
    If dlgFind.find <> "" Then
        gTextToFind = dlgFind.find
        continueSearch
    End If
    Unload dlgFind
End Sub

Private Sub continueSearch()
    Dim findPos As Integer
    With txtSource
        findPos = InStr(.SelStart + .SelLength + 1, _
                .text, gTextToFind, vbTextCompare)
    End With
    If findPos >= 1 Then
        txtSource.SelStart = findPos - 1
        txtSource.SelLength = Len(gTextToFind)
        txtSource.SetFocus
    Else
        If updateFromActiveTab Then
            searchInNextStation
        End If
    End If
End Sub

Private Sub searchInNextStation()
    If gotoNextStation Then
        loadSelectedStationIntoActiveTab
        txtSource.SelStart = 0
        txtSource.SetFocus
        continueSearch
    End If
End Sub

Private Function gotoNextStation() As Boolean
    Dim success As Boolean, newListIndex As Integer
    
    With lstStations
       If .ListIndex < .ListCount - 1 Then
            newListIndex = .ListIndex + 1
            success = True
        ElseIf vbYes = MsgBox("Restart search from first station?", _
                vbYesNoCancel) Then
            
            newListIndex = 0
            success = True
        Else
            success = False
        End If
        
        If success Then
            eventsOn = False
            .ListIndex = newListIndex
            eventsOn = True
        End If
    End With
    
    gotoNextStation = success
End Function

Private Sub mnuProjectAnalyze_Click()
    Dim graphFileName As String
    Dim dlghelp As Form
        
    If updateFromActiveTab Then
        graphFileName = prepareAnalysisPage
        Set dlghelp = New frmBrowser
        browserStartURL = convertToURI(graphFileName)
        dlghelp.Show vbModal
        Kill graphFileName
    End If
End Sub

Private Sub mnuProjectSettings_Click()
    Dim dlgProjectSettings As Form
    If updateFromActiveTab Then
        Set dlgProjectSettings = New frmProjectSettings
        dlgProjectSettings.Show vbModal, Me
        Unload dlgProjectSettings
        loadSelectedStationIntoActiveTab
    End If
End Sub

Private Sub mnuProjectAbout_Click()
    Dim dlgProjectAbout As Form
    If updateFromActiveTab Then
        Set dlgProjectAbout = New frmProjectAbout
        dlgProjectAbout.Show vbModal, Me
        Unload dlgProjectAbout
        loadSelectedStationIntoActiveTab
    End If
End Sub

Private Sub mnuProjectStyle_Click()
    Dim dlgProjectStyle As Form
    If updateFromActiveTab Then
        Set dlgProjectStyle = New frmProjectStyle
        dlgProjectStyle.Show vbModal, Me
        Unload dlgProjectStyle
        loadSelectedStationIntoActiveTab
    End If
End Sub

Private Sub mnuSourceInsertChoice_Click()
    insertToSource "<choice station="""">", "</choice>"
End Sub

Private Sub mnuInsertInclude_Click()
    insertToSource "<include>" & vbNewLine & _
            "   <in station=""", _
            """/>" & vbNewLine & "</include>"
End Sub

Private Sub mnuSourceInsertChoose_Click()
    insertToSource "<choose station=""", """/>"
End Sub

Private Sub mnuSourceInsertComment_Click()
    insertToSource "<comment>", "</comment>"
End Sub

Private Sub mnuSourceInsertXmlComment_Click()
    insertToSource "<!-- ", " -->"
End Sub

Private Sub mnuSourceInsertIfElse_Click()
    insertToSource "<if check=""", """>" & vbNewLine & "</if>" & vbNewLine & _
            "<else>" & vbNewLine & "</else>"
End Sub

Private Sub mnuSourceInsertImage_Click()
    Dim filePath As String
    filePath = getMediaFilePath("All image files|*.gif; *.jpg; *.png|" & _
            "*.gif|*.gif|" & _
            "*.jpg|*.jpg|" & _
            "*.png|*.png")
    If filePath <> "" Then
        insertToSource "<image source=""" & filePath & """ />", ""
    End If
End Sub

Private Sub mnuSourceInsertImagemap_Click()
    Dim useMap As String, choices As String
    useMap = Clipboard.GetText
    If InStr(useMap, "<map name=""imapa"">") >= 1 Then
        If InStr(useMap, "shape=rect") >= 1 Or _
                InStr(useMap, "shape=circle") >= 1 Then
            MsgBox "The image-map created with WebMap may contain " & _
                    "only polygons."
        Else
            On Error Resume Next
            choices = useMap
            choices = Replace$(choices, "<map name=""imapa"">", "")
            choices = Replace$(choices, "</map>", "")
            choices = Replace$(choices, "shape=poly", "")
            choices = Replace$(choices, "shape=poly", "")
            choices = Replace$(choices, "<area ", "<choice ")
            choices = Replace$(choices, " coords=", " area=")
            choices = Replace$(choices, " href=", " station=")
            choices = Replace$(choices, ">", ">Your text</choice>")
            On Error GoTo 0
            txtSource.SelText = choices
        End If
    Else
        MsgBox "You need to create the image-map with WebMap " & _
                "(in the tool folder) and save to clipboard."
    End If
End Sub

Private Sub mnuSourceInsertMusic_Click()
    Dim filePath As String
    filePath = getMediaFilePath("All music files|*.mid; *.wav|" & _
            "*.mid|*.mid|" & _
            "*.wav|*.wav")
    If filePath <> "" Then
        insertToSource "<music source=""" & filePath & """ />", ""
    End If
End Sub

Public Function getMediaFilePath(Optional ByVal filterStrng As String = "*.*|*.*") As String
    Dim filePath As String
    With dlgCommon
        .FileName = ""
        .Filter = filterStrng
        .InitDir = qmlTopPath & "\media"
        .ShowOpen
        filePath = .FileName
    End With
    filePath = toMediaURI(filePath)
    getMediaFilePath = filePath
End Function

Private Sub mnuSourceInsertInput_Click()
    insertToSource "<input name=""", """ station=""""/></input>"
End Sub

Private Sub mnuSourceInsertNumber_Click()
    insertToSource "<number name=""", """ value=""""/>"
End Sub

Private Sub mnuSourceInsertString_Click()
    insertToSource "<string name=""", """ value=""""/>"
End Sub

Private Sub mnuSourceInsertRandomize_Click()
    insertToSource "<randomize number=""", """ value=""""/>"
End Sub

Private Sub mnuSourceInsertState_Click()
    insertToSource "<state name=""", """/>"
End Sub

Private Sub mnuSourceInsertText_Click()
    insertToSource "<text>" & vbNewLine, vbNewLine & "</text>"
End Sub

Private Sub insertToSource(ByVal textBefore As String, ByVal textAfter As String, Optional ByVal overwriteOld As Boolean = False)
    Dim oldSelStart As Long
    Dim oldsellength As Long
    
    oldSelStart = txtSource.SelStart
    oldsellength = txtSource.SelLength
    
    If overwriteOld Then
        txtSource.SelText = textBefore & textAfter
        txtSource.SelStart = oldSelStart + Len(textBefore)
        txtSource.SelLength = 0
    Else
        txtSource.SelText = textBefore & txtSource.SelText & textAfter
        txtSource.SelStart = oldSelStart + Len(textBefore)
        txtSource.SelLength = oldsellength
    End If
    hiliteSource txtSource
End Sub

Private Sub mnuStationDelete_Click()
    tryDeleteStation
End Sub

Private Sub tryDeleteStation()
    Dim stationToDelete As String, stationId As String
    If lstStations.ListCount = 1 Then
        MsgBox "There must be at least a single station left.", vbExclamation
    ElseIf getSelectedStation = nameOfStartStation Then
        MsgBox "You cannot delete the start station.", vbExclamation
    Else
        stationId = lstStations.List(0)
        stationToDelete = getSelectedStation
        If stationId = stationToDelete Then
            stationId = lstStations.List(1)
        End If
        
        If vbOK = MsgBox("Station will be deleted.", vbOKCancel) Then
            eventsOn = False
            stationHistory.remove stationToDelete
            goToStation stationId
            removeStationByID stationToDelete
            lstStations.RemoveItem _
                getIndexOfListItem(lstStations, stationToDelete)
            eventsOn = True
        End If
    End If
End Sub

Private Sub mnuStationFollowSelection_Click()
    Dim stationId As String
    Dim success As Boolean
    
    With txtSource
        stationId = getAttributeValueOfText(.text, .SelStart)
    End With
    If InStr(stationId, "<") >= 1 Or _
            InStr(stationId, ">") >= 1 Or _
            InStr(stationId, vbNewLine) >= 1 Then
        stationId = ""
    End If
    
    If stationId = stationBackName Then
        stationId = beforeLastStationSelected
    End If
    
    If stationId <> "" And getSelectedStation <> stationId Then
        goCreateStation stationId, success
    End If
End Sub

Private Sub mnuStationRevert_Click()
    If vbOK = MsgBox("This will set back the current station." & vbNewLine & _
            "Changes since last validated editing are lost.", vbOKCancel) Then
        doRevertSource
    End If
End Sub

Private Sub doRevertSource()
    Dim objStation As IXMLDOMElement
    setStationFromSelectedId objStation
    loadStationIntoTabSource objStation
End Sub

Private Sub mnuStationSettings_Click()
    editStationSettings
End Sub

Private Sub editStationSettings()
    Dim dlgStationSettings As Form
    If updateFromActiveTab Then
        Set dlgStationSettings = New frmStationSettings
        dlgStationSettings.Show vbModal, Me
        Unload dlgStationSettings
        loadSelectedStationIntoActiveTab
    End If
End Sub

Private Sub mnuEditWrapWithTag_Click()
    Dim dlgWrapWithTag As Form
    Set dlgWrapWithTag = New frmWrapWithTag
    dlgWrapWithTag.Show vbModal
    Unload dlgWrapWithTag
End Sub

Private Sub mnuToolsOptionsAssistInline_Click()
    mnuToolsOptionsAssistInline.Checked = Not _
            mnuToolsOptionsAssistInline.Checked
    g_assistInline = mnuToolsOptionsAssistInline.Checked
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If eventsOn Then
        eventsOn = False
        If updateFromTab(PreviousTab) Then
            loadStationIntoActiveTab getSelectedStation
            
            Select Case tabMain.Tab
                Case tabType.tabEditor, tabType.tabPreview
                    menuSourceOnly False
                Case tabType.tabsource
                    menuSourceOnly True
            End Select
        End If
        eventsOn = True
    End If
   
End Sub

Private Sub menuSourceOnly(ByVal enabled As Boolean)
    mnuSourceInsert.enabled = enabled
    mnuEditSelectTag.enabled = enabled
    mnuStationRevert.enabled = enabled
    mnuStationFollowSelection.enabled = enabled
    mnuEditFind.enabled = enabled
    mnuEditContinueSearch.enabled = enabled
    mnuEditWrapWithTag.enabled = enabled
End Sub

Private Function updateFromTab(ByVal tabIndex As Integer) As Boolean
    Dim objStation As IXMLDOMElement, _
            eventsOnWas As Boolean, success As Boolean
    
    success = True
    Select Case tabIndex
        Case tabType.tabEditor
            setStationFromId objStation, getSelectedStation
            updateFromTabEditor objStation
        Case tabType.tabsource
            If Not updateFromTabSource Then
                eventsOnWas = eventsOn
                eventsOn = False
                tabMain.Tab = tabType.tabsource
                eventsOn = eventsOnWas
                success = False
            End If
        Case tabType.tabPreview
            updateFromTabPreview
    End Select
    
    updateFromTab = success
End Function

Private Sub updateFromTabPreview()
    menuForEditorSource True
    webPreview.Navigate convertToURI(qmlTopPath & _
            "\tool\project\blank.htm")
    removeQMLPreview
End Sub

Private Sub removeQMLPreview()
    Kill qmlTopPath & "\" & qmlPreviewName & ".htm"
    Kill qmlTopPath & "\quest\" & qmlPreviewName & ".xml"
End Sub

Private Sub updateFromTabEditor(ByRef objStation As IXMLDOMElement)
    tabEditorTextIntoXML objStation
    tabEditorChoicesIntoXML objStation
End Sub

Private Sub tabEditorChoicesIntoXML(ByRef objStation As IXMLDOMElement)
    Dim i As Integer
    
    removeChoices objStation
    For i = cboStation.LBound To cboStation.UBound
        If txtChoice(i).text <> "" And _
            cboStation.Item(i).text <> "" Then
            addChoiceToStation objStation, _
                getTextForXML(txtChoice(i).text), cboStation.Item(i).text
        End If
    Next
End Sub

Private Sub addChoiceToStation(ByRef objStation As IXMLDOMElement, ByVal text As String, ByVal toStation As String)
    addElmWithAttribAndText objQuest, objStation, _
        elementPath, "station", toStation, text
End Sub

Private Sub addStationToXML(ByRef stationId As String)
    Dim xPath
    
    xPath = "//station[@id = '" & stationId & "']"
    
    addElmWithAttribAndText objQuest, objQuest.documentElement, _
            "station", "id", stationId, ""
    addElmWithText objQuest, objQuest.selectSingleNode(xPath), _
            "text", " "
    If autoCreateChoice Then
        addElmWithAttribAndText objQuest, objQuest.selectSingleNode(xPath), _
                "choice", "station", "", ""
    End If
End Sub

Private Sub removeChoices(ByRef objStation As IXMLDOMElement)
    removeTopChildrenOf objStation, elementPath
End Sub

Private Sub tabEditorTextIntoXML(ByRef objStation As IXMLDOMElement)
    Dim child As IXMLDOMNode
    Dim splitted() As String
    Dim allText As String
    Dim i As Integer
    Dim textNode As IXMLDOMText
    Dim breakNode As IXMLDOMElement
    Dim newTextNode As IXMLDOMElement
    
    If txtStation.text = "" Then txtStation.text = " "
    
    splitted = Split(getTextForXML(txtStation.text), vbNewLine)
    Set newTextNode = objQuest.createElement("text")
    
    For i = LBound(splitted) To UBound(splitted)
        If splitted(i) = "" Then
            If newTextNode.childNodes.length >= 1 Then
                If newTextNode.lastChild.nodeName = "break" Then
                    Set breakNode = newTextNode.lastChild
                    breakNode.setAttribute "type", "strong"
                End If
            End If
        Else
            Set textNode = objQuest.createTextNode(splitted(i))
            newTextNode.appendChild textNode
            If i < UBound(splitted) Then
                Set breakNode = objQuest.createElement("break")
                newTextNode.appendChild breakNode
            End If
        End If
    Next
    
    For Each child In objStation.childNodes
        If child.nodeName = elementText Then
            objStation.replaceChild newTextNode, child
            Exit For
        End If
    Next
End Sub

Private Function getTextForXML(ByVal text As String) As String
    text = repeatedReplace(text, _
           vbNewLine & vbNewLine & vbNewLine, vbNewLine & vbNewLine)
    getTextForXML = text
End Function

Private Function getStationOpenTag() As String
    Dim objCheckStation As IXMLDOMElement
    Dim objTempStation As IXMLDOMElement
    Dim xPath As String
    Dim stationOpenTag As String

    xPath = "//station[@id = '" & lastStationSelected & "']"
    Set objTempStation = objQuest.documentElement.selectSingleNode(xPath)
    Set objCheckStation = objTempStation.cloneNode(False)
    stationOpenTag = objCheckStation.xml
    stationOpenTag = Replace$(stationOpenTag, "</station>", "")
    stationOpenTag = Replace$(stationOpenTag, "/>", ">")
    
    getStationOpenTag = stationOpenTag
End Function

Private Function updateFromTabSource()
    Dim userStation As MSXML2.DOMDocument
    Dim objStation As MSXML2.IXMLDOMElement
    Dim stationValid As Boolean
    Dim xmlTemp As String
    Dim sDtd As String
    
    Set userStation = New MSXML2.DOMDocument

    sDtd = ""
    If mnuFileValidateWithDTD.Checked Then
        sDtd = "<!DOCTYPE quest SYSTEM """ & getDoctypePath & """>" & vbNewLine
    End If

    xmlTemp = "<?xml version=""1.0"" ?>" & vbNewLine & _
            sDtd & _
            "<quest>" & vbNewLine & _
            "<about><title> </title><author> </author></about>" & vbNewLine & _
            getStationOpenTag & _
            txtSource.text & vbNewLine & _
            "</station>" & vbNewLine & "</quest>"

    If mnuFileShowXML.Checked Then
        MsgBox xmlTemp
    End If
    
    userStation.loadXML xmlTemp
   
    Me.txtValidate.ForeColor = 192
    If userStation.parseError.errorCode = 0 Then
        Set objStation = userStation.documentElement.selectSingleNode("//station")
        removeStationByID objStation.getAttribute("id")
        
        objQuest.documentElement.appendChild objStation
        Me.txtValidate.ForeColor = 32768
        Me.txtValidate.text = "Ok."
        stationValid = True
    Else
        Me.txtValidate.text = _
                rewriteParseErrorReason(userStation.parseError.reason)
        selectLineOfText txtSource, _
                userStation.parseError.line, _
                userStation.parseError.linepos
        stationValid = False
    End If
    
    updateFromTabSource = stationValid
End Function

Private Function rewriteParseErrorReason(ByVal oldErrorReason As String)
    Dim text As String
    text = oldErrorReason
    
    text = Replace$(text, "DTD/Schema", "QML syntax")
    text = Replace$(text, "#PCDATA", "[text]")
    
    rewriteParseErrorReason = text
End Function

Private Function getDoctypePath()
    getDoctypePath = convertToURI(qmlTopPath & "\script\quest.dtd")
End Function

Sub removeStationByID(ByVal stationId As String)
    Dim objStation As IXMLDOMElement
    setStationFromId objStation, stationId
    objQuest.documentElement.removeChild objStation
End Sub

Private Function updateFromActiveTab(Optional ByVal useLastStation As Boolean = False)
    Dim objStation As IXMLDOMElement, stationName As String, _
            success As Boolean
    success = True
    If useLastStation Then
        stationName = lastStationSelected
    Else
        stationName = getSelectedStation
    End If
    
    Select Case tabMain.Tab
        Case tabType.tabEditor
            setStationFromId objStation, stationName
            updateFromTabEditor objStation
        Case tabType.tabsource
            success = updateFromTabSource
    End Select
    
    updateFromActiveTab = success
End Function

Private Sub loadStationIntoActiveTab(ByVal stationId As String)
    Dim objStation As IXMLDOMElement
    Dim stationOkForTabEditor As Boolean
    Dim eventsWere As Boolean
    setStationFromId objStation, stationId
    
    If tabMain.Tab = tabType.tabEditor Then
        stationOkForTabEditor = getStationOkForTabEditor(objStation)
        If Not stationOkForTabEditor Then
            eventsWere = eventsOn
            eventsOn = False
            tabMain.Tab = tabType.tabsource
            eventsOn = eventsWere
        End If
    End If
    
    Select Case tabMain.Tab
        Case tabType.tabEditor
            loadStationIntoTabEditor objStation
        Case tabType.tabsource
            loadStationIntoTabSource objStation
        Case tabType.tabPreview
            loadStationIntoTabPreview stationId
    End Select
End Sub

Private Sub loadStationIntoTabEditor(ByRef objStation As IXMLDOMElement)
    txtStation = getTextFromStation(objStation)
    emptyChoicesAndStations
    fillChoiceStations
    fillChoices objStation
End Sub

Private Sub loadStationIntoTabSource(ByRef objStation As IXMLDOMElement)
    Dim sourceToShow As String
    Dim stationInnerXml As String
    
    txtValidate.text = ""
    cmdValidate.enabled = False
    stationInnerXml = getInnerXml(objStation)
    sourceToShow = reformatQuestXMLString(stationInnerXml)
    setSourceAndHilite txtSource, trimNewline(sourceToShow)
End Sub

Private Sub loadStationIntoTabPreview(ByVal stationId As String)
    Dim htmlLocation As String

    menuForEditorSource False
    savePreview htmlLocation
    webPreview.Navigate convertToURI(htmlLocation)
    webPreview.SetFocus
End Sub

Private Sub savePreview(Optional ByRef htmlLocation As String, Optional ByRef xmlLocation As String)
    Dim qmlpreview As String
    
    qmlpreview = getFileText(qmlTopPath & "\tool\project\new.htm")
    qmlpreview = Replace$(qmlpreview, "[quest]", "quest/" & qmlPreviewName)
    qmlpreview = Replace$(qmlpreview, "[station]'", getSelectedStation & "'")
    htmlLocation = qmlTopPath & "\" & qmlPreviewName & ".htm"
    setFileText htmlLocation, qmlpreview
    xmlLocation = qmlTopPath & "\quest\" & qmlPreviewName & ".xml"
    objQuest.save xmlLocation
End Sub

Private Sub menuForEditorSource(ByVal showMenu As Boolean)
    mnuEdit.enabled = showMenu
End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "new"
            mnuFileNew_Click
        Case "open"
            mnuFileOpen_Click
        Case "save"
            mnuFileSave_Click
        
        Case "cut"
            mnuEditCut_Click
        Case "copy"
            mnuEditCopy_Click
        Case "paste"
            mnuEditPaste_Click
        Case "find"
            mnuEditFind_Click
        
        Case "about"
            mnuProjectAbout_Click
        Case "style"
            mnuProjectStyle_Click
            
        Case "back"
            goStationBack
        Case "forward"
            goStationForward
    
        Case "help"
            mnuHelpContent_Click
    
    End Select
End Sub

Private Sub txtSource_Change()
    If eventsOn Then
        gTextChangeDidOccur = True
        
        If Not cmdValidate.enabled Then
            cmdValidate.enabled = True
        End If
        If g_assistInline Then
            handleAssistInline
        End If
    End If
End Sub

Private Sub cboContext_KeyPress(KeyAscii As Integer)
    Dim text As String
    
    If KeyAscii = Asc("]") Or KeyAscii = 13 Then ' 13 = return
        text = cboContext.text
        text = text & "]"
        txtSource.SelText = text
        txtSource.SetFocus
        cboContext.Visible = False
        hiliteSource txtSource
    End If
End Sub

Private Sub cboContext_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim text As String
    
    If KeyCode = vbKeyEscape Then
        txtSource.SetFocus
        cboContext.Visible = False
    End If
End Sub

Sub handleAssistInline()
    Dim text As String
    Dim xPath As String
    Dim stateElement As MSXML2.IXMLDOMElement
    Dim stateElements As MSXML2.IXMLDOMNodeList
    Dim thisName As String
    Dim i As Long
    Dim names As String
    Dim wrappedName As String
    
    If txtSource.SelStart > 0 And _
            txtSource.SelStart < Len(txtSource.text) Then
        text = Mid$(txtSource.text, txtSource.SelStart, 1)
        
        If text = "[" Then
            names = ""
            cboContext.clear
            
            xPath = "//state | //number | //string"
            Set stateElements = objQuest.selectNodes(xPath)
            For Each stateElement In stateElements
                thisName = stateElement.getAttribute("name")
                wrappedName = "[" & thisName & "]"
                If Not InStr(names, wrappedName) >= 1 Then
                    names = names & wrappedName
                    cboContext.AddItem thisName
                End If
            Next
            cboContext.AddItem "qmlStation"
            cboContext.AddItem "qmlLastStation"
                        
            cboContext.Visible = True
            cboContext.SetFocus
        End If

    End If
End Sub

Private Sub txtSource_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 And Shift = 1 Or _
         KeyCode = 86 And Shift = 2 Then
        ' Shift+Insert / Ctrl+V
        KeyCode = 0
        mnuEditPaste_Click
    ElseIf KeyCode = 46 And Shift = 1 Or _
        KeyCode = 88 And Shift = 2 Then
        ' Shift+Delete / Ctrl+X
        KeyCode = 0
        mnuEditCut_Click
        
    ElseIf KeyCode = 45 And Shift = 2 Or _
        KeyCode = 67 And Shift = 2 Then
        ' Ctrl+Insert/ Ctrl+C
        KeyCode = 0
        mnuEditCopy_Click
        
    ElseIf KeyCode = 37 And Shift = 4 Then
        ' Alt+Left
        goStationBack
    ElseIf KeyCode = 39 And Shift = 4 Then
        ' Alt+Right
        goStationForward
        
    ElseIf KeyCode = 9 And Shift = 0 Then
        ' tab
        txtSource.SelText = vbTab
        KeyCode = 0
    ElseIf KeyCode = 9 And Shift = 1 Then
        ' back tab
        KeyCode = 0
    End If
End Sub

Private Sub txtSource_KeyUp(KeyCode As Integer, Shift As Integer)
    Static lastLine As Long
    Dim currentLine As Long
    
    If gTextChangeDidOccur Then
        currentLine = txtSource.GetLineFromChar(txtSource.SelStart)
        
        If currentLine <> lastLine Then
            hiliteSource txtSource
            gTextChangeDidOccur = False
            lastLine = currentLine
        End If

    End If
End Sub

Private Sub txtSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const rightClick = 2
    If Button = rightClick Then
        setMenuEdit False
        txtSource.enabled = False
        txtSource.enabled = True
        txtSource.SetFocus
        PopupMenu mnuEdit
        setMenuEdit True
    End If
End Sub

Private Sub setMenuEdit(ByVal state As Boolean)
        mnuEditWrapWithTag.Visible = state
        mnuEditBreak.Visible = state
        mnuEditBreak2.Visible = state
        mnuEditBack.Visible = state
        mnuEditForward.Visible = state
        mnuEditFind.Visible = state
        mnuEditContinueSearch.Visible = state
End Sub

Private Sub txtStation_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const rightClick = 2
    If Button = rightClick Then
        setMenuEdit False
        txtStation.enabled = False
        txtStation.enabled = True
        txtStation.SetFocus
        PopupMenu mnuEdit
        setMenuEdit True
    End If
End Sub

Private Sub txtStationName_KeyPress(KeyAscii As Integer)
    If eventsOn Then
        If KeyAscii = vbKeyReturn Then
            cmdAdd_Click
        End If
    End If
End Sub

Private Sub cboStation_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdGo_Click index
    End If
End Sub

Public Sub visitQMLHomepage()
    Dim dlgBrowser As Form
    
    Set dlgBrowser = New frmBrowser
    frmMain.browserStartURL = "http://questml.com"
    dlgBrowser.Show
End Sub
