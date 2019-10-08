VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProjectStyle 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Project Style"
   ClientHeight    =   4668
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5040
   Icon            =   "frmProjectStyle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4668
   ScaleWidth      =   5040
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1308
      TabIndex        =   15
      Top             =   4200
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3840
      TabIndex        =   17
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   372
      Left            =   2640
      TabIndex        =   16
      Top             =   4200
      Width           =   1092
   End
   Begin TabDlg.SSTab tabStyle 
      Height          =   3972
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   7006
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   420
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Page"
      TabPicture(0)   =   "frmProjectStyle.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Classes"
      TabPicture(1)   =   "frmProjectStyle.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmClasses"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frmClasses 
         BorderStyle     =   0  'Kein
         Height          =   3468
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4572
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   12
            Left            =   1188
            TabIndex        =   95
            Top             =   3240
            Width           =   864
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   11
            Left            =   1200
            TabIndex        =   94
            Top             =   3000
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   10
            Left            =   1200
            TabIndex        =   93
            Top             =   2760
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   9
            Left            =   1200
            TabIndex        =   92
            Top             =   2520
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   8
            Left            =   1200
            TabIndex        =   91
            Top             =   2280
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   7
            Left            =   1200
            TabIndex        =   90
            Top             =   2040
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   6
            Left            =   1200
            TabIndex        =   89
            Top             =   1800
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   5
            Left            =   1200
            TabIndex        =   88
            Top             =   1560
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   4
            Left            =   1200
            TabIndex        =   87
            Top             =   1320
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   3
            Left            =   1200
            TabIndex        =   86
            Top             =   1080
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   2
            Left            =   1200
            TabIndex        =   85
            Top             =   840
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   1
            Left            =   1200
            TabIndex        =   84
            Top             =   600
            Width           =   852
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   0
            Left            =   1200
            TabIndex        =   83
            Top             =   360
            Width           =   852
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   12
            Left            =   2148
            TabIndex        =   82
            Top             =   3228
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   11
            Left            =   2148
            TabIndex        =   81
            Top             =   2988
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   10
            Left            =   2148
            TabIndex        =   80
            Top             =   2760
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   9
            Left            =   2148
            TabIndex        =   79
            Top             =   2520
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   8
            Left            =   2148
            TabIndex        =   78
            Top             =   2280
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   7
            Left            =   2148
            TabIndex        =   77
            Top             =   2040
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   6
            Left            =   2148
            TabIndex        =   76
            Top             =   1800
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   5
            Left            =   2148
            TabIndex        =   75
            Top             =   1560
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   4
            Left            =   2148
            TabIndex        =   74
            Top             =   1308
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   3
            Left            =   2148
            TabIndex        =   73
            Top             =   1080
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   2
            Left            =   2148
            TabIndex        =   72
            Top             =   840
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   1
            Left            =   2148
            TabIndex        =   71
            Top             =   588
            Width           =   5652
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   0
            Left            =   2148
            TabIndex        =   70
            Top             =   360
            Width           =   5652
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   12
            Left            =   0
            TabIndex        =   69
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   11
            Left            =   0
            TabIndex        =   68
            Top             =   3000
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   10
            Left            =   0
            TabIndex        =   67
            Top             =   2760
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   9
            Left            =   0
            TabIndex        =   66
            Top             =   2520
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   8
            Left            =   0
            TabIndex        =   65
            Top             =   2280
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   7
            Left            =   0
            TabIndex        =   64
            Top             =   2040
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   6
            Left            =   0
            TabIndex        =   63
            Top             =   1800
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   5
            Left            =   0
            TabIndex        =   62
            Top             =   1560
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   4
            Left            =   0
            TabIndex        =   61
            Top             =   1320
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   3
            Left            =   0
            TabIndex        =   60
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   2
            Left            =   0
            TabIndex        =   59
            Top             =   840
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   1
            Left            =   0
            TabIndex        =   58
            Top             =   600
            Width           =   1212
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   0
            Left            =   0
            TabIndex        =   57
            Top             =   360
            Width           =   1212
         End
         Begin VB.CommandButton cmdSelectStyle 
            Caption         =   "Select..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2868
            TabIndex        =   56
            Top             =   108
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   13
            Left            =   0
            TabIndex        =   55
            Top             =   3516
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   13
            Left            =   2148
            TabIndex        =   54
            Top             =   3504
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   13
            Left            =   1200
            TabIndex        =   53
            Top             =   3516
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   14
            Left            =   0
            TabIndex        =   52
            Top             =   3768
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   14
            Left            =   2148
            TabIndex        =   51
            Top             =   3756
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   14
            Left            =   1200
            TabIndex        =   50
            Top             =   3768
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   15
            Left            =   0
            TabIndex        =   49
            Top             =   4020
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   15
            Left            =   2148
            TabIndex        =   48
            Top             =   4008
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   15
            Left            =   1200
            TabIndex        =   47
            Top             =   4020
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   16
            Left            =   0
            TabIndex        =   46
            Top             =   4284
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   16
            Left            =   2148
            TabIndex        =   45
            Top             =   4272
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   16
            Left            =   1200
            TabIndex        =   44
            Top             =   4284
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   17
            Left            =   0
            TabIndex        =   43
            Top             =   4548
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   17
            Left            =   2148
            TabIndex        =   42
            Top             =   4536
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   17
            Left            =   1200
            TabIndex        =   41
            Top             =   4548
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   18
            Left            =   0
            TabIndex        =   40
            Top             =   4788
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   18
            Left            =   2148
            TabIndex        =   39
            Top             =   4776
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   18
            Left            =   1200
            TabIndex        =   38
            Top             =   4788
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   19
            Left            =   0
            TabIndex        =   37
            Top             =   5064
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   19
            Left            =   2148
            TabIndex        =   36
            Top             =   5052
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00C0C0C0&
            Height          =   288
            Index           =   19
            Left            =   1200
            TabIndex        =   35
            Top             =   5064
            Width           =   852
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  '2D
            Height          =   288
            Index           =   20
            Left            =   0
            TabIndex        =   34
            Top             =   5340
            Width           =   1212
         End
         Begin VB.TextBox txtStyle 
            Appearance      =   0  '2D
            Height          =   300
            Index           =   20
            Left            =   2148
            TabIndex        =   33
            Top             =   5328
            Width           =   5652
         End
         Begin VB.TextBox txtInherits 
            Appearance      =   0  '2D
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Index           =   20
            Left            =   1200
            TabIndex        =   32
            Top             =   5340
            Width           =   852
         End
         Begin VB.CommandButton cmdClassExpand 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   4008
            TabIndex        =   31
            Top             =   108
            Width           =   492
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Undurchsichtig
            Height          =   5268
            Left            =   2040
            Top             =   360
            Width           =   132
         End
         Begin VB.Line lineTop 
            X1              =   0
            X2              =   4560
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   0
            Y1              =   360
            Y2              =   0
         End
         Begin VB.Label Label21 
            Caption         =   "Inherits"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   6.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1320
            TabIndex        =   98
            Top             =   120
            Width           =   612
         End
         Begin VB.Label Label20 
            Caption         =   "Style"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   2280
            TabIndex        =   97
            Top             =   120
            Width           =   1932
         End
         Begin VB.Line Line1 
            X1              =   1200
            X2              =   1200
            Y1              =   0
            Y2              =   360
         End
         Begin VB.Label Label19 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   96
            Top             =   120
            Width           =   972
         End
         Begin VB.Line Line2 
            X1              =   2040
            X2              =   2040
            Y1              =   0
            Y2              =   360
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Background"
         Height          =   972
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   4572
         Begin VB.CommandButton cmdSetBackgroundColor 
            Caption         =   "Set..."
            Height          =   252
            Left            =   2520
            TabIndex        =   1
            Top             =   240
            Width           =   972
         End
         Begin VB.TextBox txtBackgroundImage 
            Height          =   288
            Left            =   1440
            TabIndex        =   2
            Top             =   600
            Width           =   972
         End
         Begin VB.CommandButton cmdSetBackgroundImage 
            Caption         =   "Set..."
            Height          =   252
            Left            =   2520
            TabIndex        =   3
            Top             =   600
            Width           =   972
         End
         Begin VB.ComboBox cboBackgroundRepeat 
            Height          =   288
            ItemData        =   "frmProjectStyle.frx":047A
            Left            =   3600
            List            =   "frmProjectStyle.frx":048A
            Style           =   2  'Dropdown-Liste
            TabIndex        =   4
            Top             =   600
            Width           =   852
         End
         Begin VB.TextBox txtBackgroundColor 
            Height          =   288
            Left            =   1440
            TabIndex        =   0
            Top             =   240
            Width           =   972
         End
         Begin VB.Shape shpBackgroundColor 
            Height          =   252
            Left            =   3600
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Color:"
            Height          =   252
            Left            =   720
            TabIndex        =   29
            Top             =   240
            Width           =   612
         End
         Begin VB.Label Label2 
            Caption         =   "Image:"
            Height          =   252
            Left            =   720
            TabIndex        =   28
            Top             =   600
            Width           =   612
         End
         Begin VB.Image Image1 
            Height          =   384
            Left            =   120
            Picture         =   "frmProjectStyle.frx":04B5
            Top             =   360
            Width           =   384
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Font"
         Height          =   1332
         Left            =   -74880
         TabIndex        =   22
         Top             =   1440
         Width           =   4572
         Begin VB.CommandButton cmdSetFontColor 
            Caption         =   "Set..."
            Height          =   252
            Left            =   2520
            TabIndex        =   6
            Top             =   240
            Width           =   972
         End
         Begin VB.TextBox txtFontFamily 
            Height          =   288
            Left            =   1440
            TabIndex        =   7
            Top             =   600
            Width           =   972
         End
         Begin VB.TextBox txtFontSize 
            Height          =   288
            Left            =   3000
            TabIndex        =   8
            Top             =   600
            Width           =   492
         End
         Begin VB.ComboBox cboFontWeight 
            Height          =   288
            ItemData        =   "frmProjectStyle.frx":05A0
            Left            =   3600
            List            =   "frmProjectStyle.frx":05AA
            Style           =   2  'Dropdown-Liste
            TabIndex        =   9
            Top             =   600
            Width           =   852
         End
         Begin VB.ComboBox cboFontLinks 
            Height          =   288
            ItemData        =   "frmProjectStyle.frx":05BC
            Left            =   1440
            List            =   "frmProjectStyle.frx":05C6
            Style           =   2  'Dropdown-Liste
            TabIndex        =   10
            Top             =   960
            Width           =   972
         End
         Begin VB.TextBox txtFontColor 
            Height          =   288
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   972
         End
         Begin VB.Label Label3 
            Caption         =   "Color:"
            Height          =   252
            Left            =   720
            TabIndex        =   26
            Top             =   240
            Width           =   612
         End
         Begin VB.Shape shpFontColor 
            Height          =   252
            Left            =   3600
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label4 
            Caption         =   "Family:"
            Height          =   252
            Left            =   720
            TabIndex        =   25
            Top             =   600
            Width           =   732
         End
         Begin VB.Label Label5 
            Caption         =   "Size:"
            Height          =   252
            Left            =   2520
            TabIndex        =   24
            Top             =   600
            Width           =   492
         End
         Begin VB.Label Label7 
            Caption         =   "Links:"
            Height          =   252
            Left            =   720
            TabIndex        =   23
            Top             =   960
            Width           =   732
         End
         Begin VB.Image Image2 
            Height          =   384
            Left            =   120
            Picture         =   "frmProjectStyle.frx":05DE
            Top             =   480
            Width           =   384
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Content"
         Height          =   972
         Left            =   -74880
         TabIndex        =   18
         Top             =   2880
         Width           =   4572
         Begin VB.TextBox txtContentLeft 
            Height          =   288
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   492
         End
         Begin VB.TextBox txtContentTop 
            Height          =   288
            Left            =   3000
            TabIndex        =   12
            Top             =   240
            Width           =   492
         End
         Begin VB.TextBox txtContentWidth 
            Height          =   288
            Left            =   1440
            TabIndex        =   13
            Top             =   600
            Width           =   492
         End
         Begin VB.Label Label8 
            Caption         =   "Left:"
            Height          =   252
            Left            =   720
            TabIndex        =   21
            Top             =   240
            Width           =   732
         End
         Begin VB.Label Label9 
            Caption         =   "Top:"
            Height          =   252
            Left            =   2520
            TabIndex        =   20
            Top             =   240
            Width           =   492
         End
         Begin VB.Label Label10 
            Caption         =   "Width:"
            Height          =   252
            Left            =   720
            TabIndex        =   19
            Top             =   600
            Width           =   732
         End
         Begin VB.Image Image3 
            Height          =   384
            Left            =   120
            Picture         =   "frmProjectStyle.frx":06D8
            Top             =   360
            Width           =   384
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   3960
      Top             =   -120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
End
Attribute VB_Name = "frmProjectStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const fillStyleFilled = 0

Private Sub cmdClassExpand_Click()
    If cmdClassExpand.Caption = ">>" Then
        Width = 8400
        Height = 7200
        cmdClassExpand.Caption = "<<"
    Else ' If cmdClassExpand.Caption = "<<" then
        Width = 5112
        Height = 5004
        cmdClassExpand.Caption = ">>"
    End If
    
    lineTop.X2 = Width - (5112 - 4560)
    
    cmdClassExpand.Left = Width - (5112 - 4008)
    
    frmClasses.Width = Width - (5112 - 4548)
    frmClasses.Height = Height - (5004 - 3492)
    
    tabStyle.Width = Width - (5112 - 4800)
    tabStyle.Height = Height - (5004 - 3960)
    
    cmdCancel.Top = Height - (5004 - 4200)
    cmdOk.Top = Height - (5004 - 4200)
    cmdDefault.Top = Height - (5004 - 4200)
    
    cmdCancel.Left = Width - (5112 - 3840)
    cmdOk.Left = Width - (5112 - 1308)
    cmdDefault.Left = Width - (5112 - 2640)
End Sub

Private Sub cmdSelectStyle_Click()
    Dim dlgCSS As New frmCssBrowser
    dlgCSS.Show vbModal, Me
    Unload dlgCSS
End Sub

Private Sub Form_Load()
    Const tabPage = 0
    tabStyle.Tab = tabPage
    loadsettings
End Sub

Private Sub cmdSetBackgroundImage_Click()
    setTextBoxToWebImagePath txtBackgroundImage
End Sub

Private Sub cmdSetFontColor_Click()
    setColorShapeAndText shpFontColor, txtFontColor
End Sub

Private Sub cmdCancel_Click()
    exitForm
End Sub

Private Sub cmdDefault_Click()
    If vbOK = MsgBox("All styles are set" & vbNewLine & _
            "back to the browser default display.", vbOKCancel) Then
        removeStyleNode
        exitForm
    End If
End Sub

Private Sub removeStyleNode()
    Dim objStyle As IXMLDOMElement, child As IXMLDOMElement
    For Each child In frmMain.objQuest.documentElement.childNodes
        If child.nodeName = "style" Then
            frmMain.objQuest.documentElement.removeChild child
            Exit For
        End If
    Next
End Sub

Private Sub cmdOk_Click()
    saveSettings
    exitForm
End Sub

Private Sub exitForm()
    Unload Me
End Sub

Private Sub cmdSetBackgroundColor_Click()
    setColorShapeAndText shpBackgroundColor, txtBackgroundColor
End Sub

Private Sub setColorShapeAndText(ByRef colShape As Object, ByRef textBox As Object)
    Dim col As colorTriplet
    col = getColorTriplet(dlgColor)
    colShape.FillColor = getVBColor(col)
    colShape.FillStyle = fillStyleFilled
    textBox.text = colorTripletToCSS(col)
End Sub

Private Sub loadsettings()
    Dim objStyle As IXMLDOMElement, _
            objs As IXMLDOMNodeList
    Set objs = frmMain.objQuest.getElementsByTagName("style")
    If objs.length = 1 Then
        setDialogFromNode objs.Item(0)
    End If
End Sub

Private Sub setDialogFromNode(ByRef objStyle As IXMLDOMElement)
    setStyleFromNode objStyle
    setClassesFromNode objStyle
End Sub

Private Sub setClassesFromNode(ByRef objStyle As IXMLDOMElement)
    Dim classNode As IXMLDOMElement
    Dim classes As IXMLDOMNodeList
    Dim i As Integer
    
    Set classes = objStyle.selectNodes("./class")
    For Each classNode In classes
        txtName(i).text = classNode.getAttribute("name")
        If classNode.getAttribute("inherits") <> "" Then
            txtInherits(i).text = classNode.getAttribute("inherits")
        End If
        txtStyle(i).text = classNode.getAttribute("style")
        i = i + 1
        If i > txtName.UBound Then
            MsgBox "Too many classes to display", vbCritical
            Exit For
        End If
    Next
End Sub

Private Sub setStyleFromNode(ByRef objStyle As IXMLDOMElement)
    Dim child As IXMLDOMElement
    For Each child In objStyle.childNodes
        Select Case child.nodeName
            Case "background"
                If child.getAttribute("color") <> "default" Then
                    txtBackgroundColor.text = child.getAttribute("color")
                    shpBackgroundColor.FillColor = _
                            cssRGBToVBColor(txtBackgroundColor.text)
                End If
                shpBackgroundColor.FillStyle = fillStyleFilled
                If child.getAttribute("image") <> "default" Then
                    txtBackgroundImage.text = _
                            child.getAttribute("image")
                End If
                cboBackgroundRepeat.text = _
                        child.getAttribute("repeat")
            
            Case "font"
                If child.getAttribute("size") <> "default" Then
                    txtFontSize.text = child.getAttribute("size")
                End If
                If child.getAttribute("family") <> "default" Then
                    txtFontFamily.text = child.getAttribute("family")
                End If
                If child.getAttribute("color") <> "default" Then
                    txtFontColor.text = child.getAttribute("color")
                    shpFontColor.FillColor = _
                        cssRGBToVBColor(txtFontColor.text)
                    shpFontColor.FillStyle = fillStyleFilled
                End If
                If child.getAttribute("weight") <> "default" Then
                    cboFontWeight.text = child.getAttribute("weight")
                End If
                If child.getAttribute("links") <> "underlined" Then
                    cboFontLinks.text = child.getAttribute("links")
                End If
            
            Case "content"
                txtContentLeft.text = child.getAttribute("left")
                txtContentTop.text = child.getAttribute("top")
                txtContentWidth.text = child.getAttribute("width")
            
        End Select
    Next
End Sub

Private Sub tabStyle_Click(PreviousTab As Integer)
    If PreviousTab = 2 And cmdClassExpand.Caption = "<<" Then
        cmdClassExpand_Click
    End If
End Sub

Private Sub txtContentLeft_LostFocus()
    toPixelValue txtContentLeft
End Sub

Private Sub txtContentTop_LostFocus()
    toPixelValue txtContentTop
End Sub

Private Sub txtContentWidth_LostFocus()
    toPixelValue txtContentWidth
End Sub

Private Sub toPixelValue(ByRef textBox As Object)
    With textBox
        If Trim$(.text) <> "" Then
            If isNumber(Right$(.text, 1)) Then
                .text = .text & "px"
            End If
        End If
    End With
End Sub

Private Sub saveSettings()
    Dim objStyle As IXMLDOMElement
    Dim styleNode As IXMLDOMElement
    Dim onePlusInserted As Boolean
    Dim nodeBack As IXMLDOMElement
    Dim nodeFont As IXMLDOMElement
    Dim nodeContent As IXMLDOMElement
    Dim i As Integer
    Dim nodeClass As IXMLDOMElement

    Set objStyle = frmMain.objQuest.createElement("style")
    removeElementIfExists frmMain.objQuest.documentElement, "style"
    Set styleNode = _
            frmMain.objQuest.documentElement.insertBefore(objStyle, _
            frmMain.objQuest.documentElement.childNodes(1))
    
    If backgroundColorSet Then
        Set nodeBack = styleNode.appendChild( _
                frmMain.objQuest.createElement("background"))
        If txtBackgroundColor.text <> "" Then
            nodeBack.setAttribute "color", txtBackgroundColor.text
        End If
        If txtBackgroundImage.text <> "" Then
            nodeBack.setAttribute "image", txtBackgroundImage.text
        End If
        If cboBackgroundRepeat.text <> "" Then
            nodeBack.setAttribute "repeat", cboBackgroundRepeat.text
        End If
        onePlusInserted = True
    End If
    
    If fontSet Then
        Set nodeFont = styleNode.appendChild( _
                frmMain.objQuest.createElement("font"))
        If txtFontColor.text <> "" Then
            nodeFont.setAttribute "color", txtFontColor.text
        End If
        If txtFontFamily.text <> "" Then
            nodeFont.setAttribute "family", txtFontFamily.text
        End If
        If txtFontSize.text <> "" Then
            nodeFont.setAttribute "size", txtFontSize.text
        End If
        If cboFontWeight.text <> "" Then
            nodeFont.setAttribute "weight", cboFontWeight.text
        End If
        If cboFontLinks.text <> "" Then
            nodeFont.setAttribute "links", cboFontLinks.text
        End If
        onePlusInserted = True
    End If
    
    If contentSet Then
        Set nodeContent = styleNode.appendChild( _
                frmMain.objQuest.createElement("content"))
        
        If txtContentLeft.text <> "" Then
            nodeContent.setAttribute "left", txtContentLeft.text
        End If
        If txtContentTop.text <> "" Then
            nodeContent.setAttribute "top", txtContentTop.text
        End If
        If txtContentWidth.text <> "" Then
            nodeContent.setAttribute "width", txtContentWidth.text
        End If
        onePlusInserted = True
    End If
   
    For i = txtName.LBound To txtName.UBound
        If Not Trim$(txtName(i).text) = "" Then
            onePlusInserted = True
            Set nodeClass = styleNode.appendChild( _
                frmMain.objQuest.createElement("class"))
            nodeClass.setAttribute "name", Trim$(txtName(i).text)
            nodeClass.setAttribute "inherits", Trim$(txtInherits(i).text)
            nodeClass.setAttribute "style", Trim$(txtStyle(i).text)
        End If
    Next
    
    If Not onePlusInserted Then
        removeElementIfExists frmMain.objQuest.documentElement, "style"
    End If
End Sub

Private Function backgroundColorSet() As Boolean
    backgroundColorSet = txtBackgroundColor.text <> "" Or _
            txtBackgroundImage.text <> "" Or _
           cboBackgroundRepeat.text <> ""
End Function

Private Function fontSet() As Boolean
    fontSet = txtFontColor.text <> "" Or _
            txtFontFamily.text <> "" Or _
            txtFontSize.text <> "" Or _
            cboFontWeight.text <> "" Or _
            cboFontLinks.text <> ""
End Function

Private Function contentSet() As Boolean
    contentSet = txtContentLeft.text <> "" Or _
            txtContentTop.text <> "" Or _
            txtContentWidth.text <> ""
End Function

