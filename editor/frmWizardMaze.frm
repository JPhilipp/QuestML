VERSION 5.00
Begin VB.Form frmWizardMaze 
   Appearance      =   0  '2D
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Maze Wizard"
   ClientHeight    =   7800
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   7224
   Icon            =   "frmWizardMaze.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7224
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6000
      TabIndex        =   1
      Top             =   7308
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   4680
      TabIndex        =   0
      Top             =   7320
      Width           =   1212
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room5e"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   24
      Left            =   6168
      TabIndex        =   26
      Top             =   6168
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room4e"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   23
      Left            =   4728
      TabIndex        =   25
      Top             =   6168
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room3e"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   22
      Left            =   3288
      TabIndex        =   24
      Top             =   6168
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room2e"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   21
      Left            =   1848
      TabIndex        =   23
      Top             =   6168
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room1e"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   20
      Left            =   408
      TabIndex        =   22
      Top             =   6156
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room5d"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   19
      Left            =   6168
      TabIndex        =   21
      Top             =   4728
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room4d"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   18
      Left            =   4728
      TabIndex        =   20
      Top             =   4716
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room3d"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   17
      Left            =   3288
      TabIndex        =   19
      Top             =   4716
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room2d"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   16
      Left            =   1860
      TabIndex        =   18
      Top             =   4728
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room1d"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   15
      Left            =   408
      TabIndex        =   17
      Top             =   4716
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room5c"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   14
      Left            =   6180
      TabIndex        =   16
      Top             =   3288
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room4c"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   13
      Left            =   4728
      TabIndex        =   15
      Top             =   3276
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room3c"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   12
      Left            =   3300
      TabIndex        =   14
      Top             =   3288
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room2c"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   11
      Left            =   1860
      TabIndex        =   13
      Top             =   3276
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room1c"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   10
      Left            =   408
      TabIndex        =   12
      Top             =   3288
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room5b"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   9
      Left            =   6168
      TabIndex        =   11
      Top             =   1848
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room4b"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   8
      Left            =   4728
      TabIndex        =   10
      Top             =   1836
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room3b"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   7
      Left            =   3288
      TabIndex        =   9
      Top             =   1836
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room2b"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   6
      Left            =   1848
      TabIndex        =   8
      Top             =   1836
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room1b"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   5
      Left            =   408
      TabIndex        =   7
      Top             =   1836
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room5a"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   4
      Left            =   6168
      TabIndex        =   6
      Top             =   408
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room4a"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   3
      Left            =   4728
      TabIndex        =   5
      Top             =   396
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room3a"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   2
      Left            =   3288
      TabIndex        =   4
      Top             =   396
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room2a"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   1
      Left            =   1848
      TabIndex        =   3
      Top             =   396
      Width           =   624
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Room1a"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   0
      Left            =   408
      TabIndex        =   2
      Top             =   396
      Width           =   624
   End
   Begin VB.Image Image1 
      Height          =   576
      Left            =   240
      Picture         =   "frmWizardMaze.frx":0442
      Top             =   7200
      Width           =   360
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      X1              =   1440
      X2              =   1440
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   55
      X1              =   5400
      X2              =   5640
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   55
      X1              =   5640
      X2              =   5880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   55
      X1              =   5760
      X2              =   5760
      Y1              =   6600
      Y2              =   6360
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   55
      X1              =   5880
      X2              =   6120
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   55
      X1              =   5760
      X2              =   5760
      Y1              =   6360
      Y2              =   6120
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   55
      X1              =   5760
      X2              =   5760
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   54
      X1              =   3960
      X2              =   4200
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   54
      X1              =   4200
      X2              =   4440
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   54
      X1              =   4320
      X2              =   4320
      Y1              =   6600
      Y2              =   6360
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   54
      X1              =   4440
      X2              =   4680
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   54
      X1              =   4320
      X2              =   4320
      Y1              =   6360
      Y2              =   6120
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   54
      X1              =   4320
      X2              =   4320
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   53
      X1              =   2520
      X2              =   2760
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   53
      X1              =   2760
      X2              =   3000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   53
      X1              =   2880
      X2              =   2880
      Y1              =   6600
      Y2              =   6360
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   53
      X1              =   3000
      X2              =   3240
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   53
      X1              =   2880
      X2              =   2880
      Y1              =   6360
      Y2              =   6120
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   53
      X1              =   2880
      X2              =   2880
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   52
      X1              =   1080
      X2              =   1320
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   52
      X1              =   1320
      X2              =   1560
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   52
      X1              =   1440
      X2              =   1440
      Y1              =   6600
      Y2              =   6360
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   52
      X1              =   1560
      X2              =   1800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   52
      X1              =   1440
      X2              =   1440
      Y1              =   6360
      Y2              =   6120
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   52
      X1              =   1440
      X2              =   1440
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   51
      X1              =   6120
      X2              =   6360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   51
      X1              =   6360
      X2              =   6600
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   51
      X1              =   6480
      X2              =   6480
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   51
      X1              =   6600
      X2              =   6840
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   51
      X1              =   6480
      X2              =   6480
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   51
      X1              =   6480
      X2              =   6480
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   50
      X1              =   5400
      X2              =   5640
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   50
      X1              =   5640
      X2              =   5880
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   50
      X1              =   5760
      X2              =   5760
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   50
      X1              =   5880
      X2              =   6120
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   50
      X1              =   5760
      X2              =   5760
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   50
      X1              =   5760
      X2              =   5760
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   49
      X1              =   4680
      X2              =   4920
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   49
      X1              =   4920
      X2              =   5160
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   49
      X1              =   5040
      X2              =   5040
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   49
      X1              =   5160
      X2              =   5400
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   49
      X1              =   5040
      X2              =   5040
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   49
      X1              =   5040
      X2              =   5040
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   48
      X1              =   3960
      X2              =   4200
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   48
      X1              =   4200
      X2              =   4440
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   48
      X1              =   4320
      X2              =   4320
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   48
      X1              =   4440
      X2              =   4680
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   48
      X1              =   4320
      X2              =   4320
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   48
      X1              =   4320
      X2              =   4320
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   47
      X1              =   3240
      X2              =   3480
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   47
      X1              =   3480
      X2              =   3720
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   47
      X1              =   3600
      X2              =   3600
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   47
      X1              =   3720
      X2              =   3960
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   47
      X1              =   3600
      X2              =   3600
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   47
      X1              =   3600
      X2              =   3600
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   46
      X1              =   2520
      X2              =   2760
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   46
      X1              =   2760
      X2              =   3000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   46
      X1              =   2880
      X2              =   2880
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   46
      X1              =   3000
      X2              =   3240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   46
      X1              =   2880
      X2              =   2880
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   46
      X1              =   2880
      X2              =   2880
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   45
      X1              =   1800
      X2              =   2040
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   45
      X1              =   2040
      X2              =   2280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   45
      X1              =   2160
      X2              =   2160
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   45
      X1              =   2280
      X2              =   2520
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   45
      X1              =   2160
      X2              =   2160
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   45
      X1              =   2160
      X2              =   2160
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   44
      X1              =   1080
      X2              =   1320
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   44
      X1              =   1320
      X2              =   1560
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   44
      X1              =   1440
      X2              =   1440
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   44
      X1              =   1560
      X2              =   1800
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   44
      X1              =   1440
      X2              =   1440
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   44
      X1              =   1440
      X2              =   1440
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   43
      X1              =   360
      X2              =   600
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   43
      X1              =   600
      X2              =   840
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   43
      X1              =   720
      X2              =   720
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   43
      X1              =   840
      X2              =   1080
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   43
      X1              =   720
      X2              =   720
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   43
      X1              =   720
      X2              =   720
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   42
      X1              =   5400
      X2              =   5640
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   42
      X1              =   5640
      X2              =   5880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   42
      X1              =   5760
      X2              =   5760
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   42
      X1              =   5880
      X2              =   6120
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   42
      X1              =   5760
      X2              =   5760
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   42
      X1              =   5760
      X2              =   5760
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   41
      X1              =   3960
      X2              =   4200
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   41
      X1              =   4200
      X2              =   4440
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   41
      X1              =   4320
      X2              =   4320
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   41
      X1              =   4440
      X2              =   4680
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   41
      X1              =   4320
      X2              =   4320
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   41
      X1              =   4320
      X2              =   4320
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   40
      X1              =   2520
      X2              =   2760
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   40
      X1              =   2760
      X2              =   3000
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   40
      X1              =   2880
      X2              =   2880
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   40
      X1              =   3000
      X2              =   3240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   40
      X1              =   2880
      X2              =   2880
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   40
      X1              =   2880
      X2              =   2880
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   39
      X1              =   1080
      X2              =   1320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   39
      X1              =   1320
      X2              =   1560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   39
      X1              =   1440
      X2              =   1440
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   39
      X1              =   1560
      X2              =   1800
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   39
      X1              =   1440
      X2              =   1440
      Y1              =   4920
      Y2              =   4680
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   39
      X1              =   1440
      X2              =   1440
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   38
      X1              =   6120
      X2              =   6360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   38
      X1              =   6360
      X2              =   6600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   38
      X1              =   6480
      X2              =   6480
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   38
      X1              =   6600
      X2              =   6840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   38
      X1              =   6480
      X2              =   6480
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   38
      X1              =   6480
      X2              =   6480
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   37
      X1              =   5400
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   37
      X1              =   5640
      X2              =   5880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   37
      X1              =   5760
      X2              =   5760
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   37
      X1              =   5880
      X2              =   6120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   37
      X1              =   5760
      X2              =   5760
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   37
      X1              =   5760
      X2              =   5760
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   36
      X1              =   4680
      X2              =   4920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   36
      X1              =   4920
      X2              =   5160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   36
      X1              =   5040
      X2              =   5040
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   36
      X1              =   5160
      X2              =   5400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   36
      X1              =   5040
      X2              =   5040
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   36
      X1              =   5040
      X2              =   5040
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   35
      X1              =   3960
      X2              =   4200
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   35
      X1              =   4200
      X2              =   4440
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   35
      X1              =   4320
      X2              =   4320
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   35
      X1              =   4440
      X2              =   4680
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   35
      X1              =   4320
      X2              =   4320
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   35
      X1              =   4320
      X2              =   4320
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   34
      X1              =   3240
      X2              =   3480
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   34
      X1              =   3480
      X2              =   3720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   34
      X1              =   3600
      X2              =   3600
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   34
      X1              =   3720
      X2              =   3960
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   34
      X1              =   3600
      X2              =   3600
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   34
      X1              =   3600
      X2              =   3600
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   33
      X1              =   2520
      X2              =   2760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   33
      X1              =   2760
      X2              =   3000
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   33
      X1              =   2880
      X2              =   2880
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   33
      X1              =   3000
      X2              =   3240
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   33
      X1              =   2880
      X2              =   2880
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   33
      X1              =   2880
      X2              =   2880
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   32
      X1              =   1800
      X2              =   2040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   32
      X1              =   2040
      X2              =   2280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   32
      X1              =   2160
      X2              =   2160
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   32
      X1              =   2280
      X2              =   2520
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   32
      X1              =   2160
      X2              =   2160
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   32
      X1              =   2160
      X2              =   2160
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   31
      X1              =   1080
      X2              =   1320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   31
      X1              =   1320
      X2              =   1560
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   31
      X1              =   1440
      X2              =   1440
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   31
      X1              =   1560
      X2              =   1800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   31
      X1              =   1440
      X2              =   1440
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   31
      X1              =   1440
      X2              =   1440
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   30
      X1              =   360
      X2              =   600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   30
      X1              =   600
      X2              =   840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   30
      X1              =   720
      X2              =   720
      Y1              =   4440
      Y2              =   4200
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   30
      X1              =   840
      X2              =   1080
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   30
      X1              =   720
      X2              =   720
      Y1              =   4200
      Y2              =   3960
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   30
      X1              =   720
      X2              =   720
      Y1              =   4440
      Y2              =   4680
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   29
      X1              =   5400
      X2              =   5640
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   29
      X1              =   5640
      X2              =   5880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   29
      X1              =   5760
      X2              =   5760
      Y1              =   3720
      Y2              =   3480
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   29
      X1              =   5880
      X2              =   6120
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   29
      X1              =   5760
      X2              =   5760
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   29
      X1              =   5760
      X2              =   5760
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   28
      X1              =   3960
      X2              =   4200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   28
      X1              =   4200
      X2              =   4440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   28
      X1              =   4320
      X2              =   4320
      Y1              =   3720
      Y2              =   3480
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   28
      X1              =   4440
      X2              =   4680
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   28
      X1              =   4320
      X2              =   4320
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   28
      X1              =   4320
      X2              =   4320
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   27
      X1              =   2520
      X2              =   2760
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   27
      X1              =   2760
      X2              =   3000
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   27
      X1              =   2880
      X2              =   2880
      Y1              =   3720
      Y2              =   3480
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   27
      X1              =   3000
      X2              =   3240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   27
      X1              =   2880
      X2              =   2880
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   27
      X1              =   2880
      X2              =   2880
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   26
      X1              =   1080
      X2              =   1320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   26
      X1              =   1320
      X2              =   1560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   26
      X1              =   1440
      X2              =   1440
      Y1              =   3720
      Y2              =   3480
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   26
      X1              =   1560
      X2              =   1800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   26
      X1              =   1440
      X2              =   1440
      Y1              =   3480
      Y2              =   3240
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   26
      X1              =   1440
      X2              =   1440
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   25
      X1              =   6120
      X2              =   6360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   25
      X1              =   6360
      X2              =   6600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   25
      X1              =   6480
      X2              =   6480
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   25
      X1              =   6600
      X2              =   6840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   25
      X1              =   6480
      X2              =   6480
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   25
      X1              =   6480
      X2              =   6480
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   24
      X1              =   5400
      X2              =   5640
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   24
      X1              =   5640
      X2              =   5880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   24
      X1              =   5760
      X2              =   5760
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   24
      X1              =   5880
      X2              =   6120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   24
      X1              =   5760
      X2              =   5760
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   24
      X1              =   5760
      X2              =   5760
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   23
      X1              =   4680
      X2              =   4920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   23
      X1              =   4920
      X2              =   5160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   23
      X1              =   5040
      X2              =   5040
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   23
      X1              =   5160
      X2              =   5400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   23
      X1              =   5040
      X2              =   5040
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   23
      X1              =   5040
      X2              =   5040
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   22
      X1              =   3960
      X2              =   4200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   22
      X1              =   4200
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   22
      X1              =   4320
      X2              =   4320
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   22
      X1              =   4440
      X2              =   4680
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   22
      X1              =   4320
      X2              =   4320
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   22
      X1              =   4320
      X2              =   4320
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   21
      X1              =   3240
      X2              =   3480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   21
      X1              =   3480
      X2              =   3720
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   21
      X1              =   3600
      X2              =   3600
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   21
      X1              =   3720
      X2              =   3960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   21
      X1              =   3600
      X2              =   3600
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   21
      X1              =   3600
      X2              =   3600
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   20
      X1              =   2520
      X2              =   2760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   20
      X1              =   2760
      X2              =   3000
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   20
      X1              =   2880
      X2              =   2880
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   20
      X1              =   3000
      X2              =   3240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   20
      X1              =   2880
      X2              =   2880
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   20
      X1              =   2880
      X2              =   2880
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   19
      X1              =   1800
      X2              =   2040
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   19
      X1              =   2040
      X2              =   2280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   19
      X1              =   2160
      X2              =   2160
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   19
      X1              =   2280
      X2              =   2520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   19
      X1              =   2160
      X2              =   2160
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   19
      X1              =   2160
      X2              =   2160
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   18
      X1              =   1080
      X2              =   1320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   18
      X1              =   1320
      X2              =   1560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   18
      X1              =   1440
      X2              =   1440
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   18
      X1              =   1560
      X2              =   1800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   18
      X1              =   1440
      X2              =   1440
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   18
      X1              =   1440
      X2              =   1440
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   17
      X1              =   360
      X2              =   600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   17
      X1              =   600
      X2              =   840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   17
      X1              =   720
      X2              =   720
      Y1              =   3000
      Y2              =   2760
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   17
      X1              =   840
      X2              =   1080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   17
      X1              =   720
      X2              =   720
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   17
      X1              =   720
      X2              =   720
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   16
      X1              =   5400
      X2              =   5640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   16
      X1              =   5640
      X2              =   5880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   16
      X1              =   5760
      X2              =   5760
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   16
      X1              =   5880
      X2              =   6120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   16
      X1              =   5760
      X2              =   5760
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   16
      X1              =   5760
      X2              =   5760
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   15
      X1              =   3960
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   15
      X1              =   4200
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   15
      X1              =   4320
      X2              =   4320
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   15
      X1              =   4440
      X2              =   4680
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   15
      X1              =   4320
      X2              =   4320
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   15
      X1              =   4320
      X2              =   4320
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   14
      X1              =   2520
      X2              =   2760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   14
      X1              =   2760
      X2              =   3000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   14
      X1              =   2880
      X2              =   2880
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   14
      X1              =   3000
      X2              =   3240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   14
      X1              =   2880
      X2              =   2880
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   14
      X1              =   2880
      X2              =   2880
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   13
      X1              =   1080
      X2              =   1320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   13
      X1              =   1320
      X2              =   1560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   13
      X1              =   1440
      X2              =   1440
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   13
      X1              =   1560
      X2              =   1800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   13
      X1              =   1440
      X2              =   1440
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   13
      X1              =   1440
      X2              =   1440
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   12
      X1              =   6120
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   12
      X1              =   6360
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   12
      X1              =   6480
      X2              =   6480
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   12
      X1              =   6600
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   12
      X1              =   6480
      X2              =   6480
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   12
      X1              =   6480
      X2              =   6480
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   11
      X1              =   5400
      X2              =   5640
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   11
      X1              =   5640
      X2              =   5880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   11
      X1              =   5760
      X2              =   5760
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   11
      X1              =   5880
      X2              =   6120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   11
      X1              =   5760
      X2              =   5760
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   11
      X1              =   5760
      X2              =   5760
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   10
      X1              =   4680
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   10
      X1              =   4920
      X2              =   5160
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   10
      X1              =   5040
      X2              =   5040
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   10
      X1              =   5160
      X2              =   5400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   10
      X1              =   5040
      X2              =   5040
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   10
      X1              =   5040
      X2              =   5040
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   9
      X1              =   3960
      X2              =   4200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   9
      X1              =   4200
      X2              =   4440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   9
      X1              =   4320
      X2              =   4320
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   9
      X1              =   4440
      X2              =   4680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   9
      X1              =   4320
      X2              =   4320
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   9
      X1              =   4320
      X2              =   4320
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   8
      X1              =   3240
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   8
      X1              =   3480
      X2              =   3720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   8
      X1              =   3600
      X2              =   3600
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   8
      X1              =   3720
      X2              =   3960
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   8
      X1              =   3600
      X2              =   3600
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   8
      X1              =   3600
      X2              =   3600
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   7
      X1              =   2520
      X2              =   2760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   7
      X1              =   2760
      X2              =   3000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   7
      X1              =   2880
      X2              =   2880
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   7
      X1              =   3000
      X2              =   3240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   7
      X1              =   2880
      X2              =   2880
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   7
      X1              =   2880
      X2              =   2880
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   6
      X1              =   1800
      X2              =   2040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   6
      X1              =   2040
      X2              =   2280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   6
      X1              =   2160
      X2              =   2160
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   6
      X1              =   2280
      X2              =   2520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   6
      X1              =   2160
      X2              =   2160
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   6
      X1              =   2160
      X2              =   2160
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   5
      X1              =   1080
      X2              =   1320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   5
      X1              =   1320
      X2              =   1560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   5
      X1              =   1440
      X2              =   1440
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   5
      X1              =   1560
      X2              =   1800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   5
      X1              =   1440
      X2              =   1440
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   5
      X1              =   1440
      X2              =   1440
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   4
      X1              =   360
      X2              =   600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   600
      X2              =   840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   720
      X2              =   720
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   4
      X1              =   840
      X2              =   1080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   4
      X1              =   720
      X2              =   720
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   4
      X1              =   720
      X2              =   720
      Y1              =   1560
      Y2              =   1800
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   3
      X1              =   5400
      X2              =   5640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   5640
      X2              =   5880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   5760
      X2              =   5760
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   3
      X1              =   5880
      X2              =   6120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   3
      X1              =   5760
      X2              =   5760
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   3
      X1              =   5760
      X2              =   5760
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   2
      X1              =   3960
      X2              =   4200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   4200
      X2              =   4440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   4320
      X2              =   4320
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   2
      X1              =   4440
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   2
      X1              =   4320
      X2              =   4320
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   2
      X1              =   4320
      X2              =   4320
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   1
      X1              =   2520
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   2760
      X2              =   3000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   2880
      X2              =   2880
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   1
      X1              =   3000
      X2              =   3240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lnePathTop 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   1
      X1              =   2880
      X2              =   2880
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   1
      X1              =   2880
      X2              =   2880
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   24
      Left            =   6120
      Top             =   6120
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   23
      Left            =   4680
      Top             =   6120
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   22
      Left            =   3240
      Top             =   6120
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   21
      Left            =   1800
      Top             =   6120
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   20
      Left            =   360
      Top             =   6120
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   19
      Left            =   6120
      Top             =   4680
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   18
      Left            =   4680
      Top             =   4680
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   17
      Left            =   3240
      Top             =   4680
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   16
      Left            =   1800
      Top             =   4680
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   15
      Left            =   360
      Top             =   4680
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   14
      Left            =   6120
      Top             =   3240
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   13
      Left            =   4680
      Top             =   3240
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   12
      Left            =   3240
      Top             =   3240
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   11
      Left            =   1800
      Top             =   3240
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   10
      Left            =   360
      Top             =   3240
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   9
      Left            =   6120
      Top             =   1800
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   8
      Left            =   4680
      Top             =   1800
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   7
      Left            =   3240
      Top             =   1800
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   6
      Left            =   1800
      Top             =   1800
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   5
      Left            =   360
      Top             =   1800
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   4
      Left            =   6120
      Top             =   360
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   3
      Left            =   4680
      Top             =   360
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   2
      Left            =   3240
      Top             =   360
      Width           =   732
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   0
      Left            =   360
      Top             =   360
      Width           =   732
   End
   Begin VB.Line lnePathLeft 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      X1              =   1080
      X2              =   1320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorVertical 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   1320
      X2              =   1560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line lneConnectorHorizontal 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   1440
      X2              =   1440
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line lnePathRight 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      X1              =   1560
      X2              =   1800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape shpRoom 
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   732
      Index           =   1
      Left            =   1800
      Top             =   360
      Width           =   732
   End
   Begin VB.Line lnePathBottom 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      X1              =   1440
      X2              =   1440
      Y1              =   840
      Y2              =   1080
   End
   Begin VB.Shape shpBack 
      BackStyle       =   1  'Undurchsichtig
      Height          =   6972
      Left            =   120
      Top             =   120
      Width           =   6972
   End
End
Attribute VB_Name = "frmWizardMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tpePoint
    x As Integer
    y As Integer
End Type

Private Type tpeRectangle
    x1 As Integer
    y1 As Integer
    x2 As Integer
    y2 As Integer
End Type

Private Enum enmLocation
    room
    corridor
End Enum

Private Type tpeLocation
    id As String
    north As String
    east As String
    south As String
    west As String
    typeOf As enmLocation
End Type

Private Const lightGray = 12632256, black = 0

Dim mousePosition As tpePoint
Const mapMin = 1, mapMax = 9
Dim location(mapMin To mapMax, mapMin To mapMax) As tpeLocation

Private Sub Form_Load()
    Dim i As Integer
    For i = lneConnectorHorizontal.LBound To lneConnectorHorizontal.uBound
        lneConnectorHorizontal.Item(i).BorderColor = black
        lneConnectorVertical.Item(i).BorderColor = black
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    saveMapToQuest
    Unload Me
End Sub

Private Sub saveMapToQuest()
    Dim p As tpePoint, i As Integer
    
    For p.y = mapMin To mapMax
        For p.x = mapMin To mapMax
        
            If p.x = mapMin And p.y = mapMin Then
                
                With location(p.x, p.y)
                    .id = "room" & i
                    .typeOf = room
                    .north = ""
                    If lnePathLeft.Item(i).BorderColor = black Then
                        .east = "corridor" & i
                    End If
                    If lnePathTop.Item(i + mapMax).BorderColor = black Then
                        .south = "corridor" & i + mapMax
                    End If
                End With
            End If
        
            i = i + 1
        Next
    Next
    
End Sub

Private Sub Form_Click()
    Dim i As Integer
    For i = lnePathLeft.LBound To lnePathLeft.uBound
        handleSwitchLine lnePathLeft.Item(i)
        handleSwitchLine lnePathTop.Item(i)
        handleSwitchLine lnePathRight.Item(i)
        handleSwitchLine lnePathBottom.Item(i)
    Next
End Sub

Private Sub handleSwitchLine(ByRef lne As Object)
    Dim rectangle As tpeRectangle
    setRectangleFromLine rectangle, lne
    If pointInRectangle(mousePosition, rectangle) Then
        switchShapeBorderColor lne
    End If
End Sub

Private Sub switchShapeBorderColor(ByRef shp As Object)
    With shp
        If .BorderColor = black Then
            .BorderColor = lightGray
        Else
            .BorderColor = black
        End If
    End With
End Sub

Private Sub setRectangleFromLine(ByRef r As tpeRectangle, lne As Object)
    r.x1 = lne.x1
    r.y1 = lne.y1
    r.x2 = lne.x2
    r.y2 = lne.y2
    correctRectangle r
End Sub

Private Function pointInRectangle(ByRef p As tpePoint, ByRef r As tpeRectangle) As Boolean
    Dim xIn As Boolean, yIn As Boolean
    xIn = (p.x >= r.x1 And p.x <= r.x2)
    yIn = (p.y >= r.y1 And p.y <= r.y2)
    pointInRectangle = xIn And yIn
End Function

Private Function correctRectangle(ByRef r As tpeRectangle)
    correctRectangleOrder r
    setRectangleMargin r, 60
End Function

Private Sub setRectangleMargin(ByRef r As tpeRectangle, Optional ByVal margin As Integer = 100)
    r.x1 = r.x1 - margin
    r.x2 = r.x2 + margin
    r.y1 = r.y1 - margin
    r.y2 = r.y2 + margin
End Sub

Private Sub correctRectangleOrder(ByRef r As tpeRectangle)
    Dim temp As Integer
    If r.x1 > r.x2 Then
        temp = r.x1
        r.x1 = r.x2
        r.x2 = temp
    End If
    If r.y1 > r.y2 Then
        temp = r.y1
        r.y1 = r.y2
        r.y2 = temp
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mousePosition.x = x
    mousePosition.y = y
End Sub
