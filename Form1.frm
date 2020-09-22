VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "levymetals RS Client - "
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin l_RSSC.cmdButton cmdStopIM 
      Height          =   255
      Left            =   120
      TabIndex        =   79
      Top             =   9960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Stop IM Flood"
      Enabled         =   0   'False
   End
   Begin l_RSSC.cmdButton cmdIMFlood 
      Height          =   255
      Left            =   120
      TabIndex        =   78
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Start IM Flood"
   End
   Begin l_RSSC.cmdButton cmdGetLoad 
      Height          =   255
      Left            =   13560
      TabIndex        =   60
      Top             =   9000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Refresh"
   End
   Begin l_RSSC.cmdButton cmdShowMem 
      Height          =   375
      Left            =   14040
      TabIndex        =   42
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Members"
   End
   Begin l_RSSC.cmdButton cmdShowFree 
      Height          =   375
      Left            =   12120
      TabIndex        =   41
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Free"
   End
   Begin InetCtlsObjects.Inet inetUpdate 
      Left            =   6000
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   12600
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11400
      Top             =   0
   End
   Begin VB.Timer tmrIMFloodS 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   9600
   End
   Begin VB.Timer tmrIMFlood 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   9600
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   11520
      ScaleHeight     =   9495
      ScaleWidth      =   135
      TabIndex        =   21
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox pb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   11655
      TabIndex        =   20
      Top             =   9425
      Width           =   11655
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   5160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrGetLoad 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   11880
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrCount 
      Interval        =   1000
      Left            =   14760
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   9495
      Left            =   -85
      TabIndex        =   0
      Top             =   -25
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   16748
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
      Location        =   "http:///"
   End
   Begin VB.CheckBox chkAR 
      BackColor       =   &H00000000&
      Caption         =   "Auto Refresh"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12240
      TabIndex        =   18
      Top             =   9000
      Width           =   1455
   End
   Begin MSComctlLib.Slider sldrTrans 
      Height          =   375
      Left            =   11880
      TabIndex        =   19
      Top             =   9720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
   End
   Begin VB.Frame frameFree 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Free Worlds"
      ForeColor       =   &H0000FF00&
      Height          =   6615
      Left            =   12000
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 7"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 5"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 1"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 3"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 4"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 8"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 10"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 11"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 13"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   33
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 14"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   34
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 15"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   35
         Top             =   4320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 16"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   36
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 17"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   37
         Top             =   5040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 19"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   38
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 20"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   39
         Top             =   5760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 21"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   40
         Top             =   6120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 25"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   33
         Left            =   1080
         TabIndex        =   43
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 33"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   32
         Left            =   1080
         TabIndex        =   44
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 32"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   29
         Left            =   1080
         TabIndex        =   45
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 29"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   30
         Left            =   1080
         TabIndex        =   46
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 30"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   31
         Left            =   1080
         TabIndex        =   47
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 31"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   34
         Left            =   1080
         TabIndex        =   48
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 34"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   35
         Left            =   1080
         TabIndex        =   49
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 35"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   37
         Left            =   1080
         TabIndex        =   50
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 37"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   38
         Left            =   1080
         TabIndex        =   51
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 38"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   39
         Left            =   1080
         TabIndex        =   52
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 39"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   40
         Left            =   1080
         TabIndex        =   53
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 40"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   43
         Left            =   1080
         TabIndex        =   54
         Top             =   4320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 43"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   45
         Left            =   1080
         TabIndex        =   55
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 45"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   47
         Left            =   1080
         TabIndex        =   56
         Top             =   5040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 47"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   49
         Left            =   1080
         TabIndex        =   57
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 49"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   50
         Left            =   1080
         TabIndex        =   58
         Top             =   5760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 50"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   51
         Left            =   1080
         TabIndex        =   59
         Top             =   6120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 51"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   61
         Left            =   2040
         TabIndex        =   61
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 61"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   57
         Left            =   2040
         TabIndex        =   62
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 57"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   52
         Left            =   2040
         TabIndex        =   63
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 52"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   55
         Left            =   2040
         TabIndex        =   64
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 55"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   56
         Left            =   2040
         TabIndex        =   65
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 56"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   62
         Left            =   2040
         TabIndex        =   66
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 62"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   63
         Left            =   2040
         TabIndex        =   67
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 63"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   67
         Left            =   2040
         TabIndex        =   68
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 67"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   68
         Left            =   2040
         TabIndex        =   69
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 68"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   72
         Left            =   2040
         TabIndex        =   70
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 72"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   73
         Left            =   2040
         TabIndex        =   71
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 73"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   74
         Left            =   2040
         TabIndex        =   72
         Top             =   4320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 74"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   75
         Left            =   2040
         TabIndex        =   73
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 75"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   76
         Left            =   2040
         TabIndex        =   74
         Top             =   5040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 76"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   80
         Left            =   2040
         TabIndex        =   75
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 80"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   81
         Left            =   2040
         TabIndex        =   76
         Top             =   5760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 81"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   82
         Left            =   2040
         TabIndex        =   77
         Top             =   6120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 82"
      End
   End
   Begin VB.Frame frameMembers 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   11880
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   80
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 18"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   81
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 12"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 2"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   83
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 6"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   84
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 9"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   85
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 22"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   86
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 23"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   87
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 24"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   88
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 26"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   27
         Left            =   240
         TabIndex        =   89
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 27"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   44
         Left            =   1200
         TabIndex        =   90
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 44"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   42
         Left            =   1200
         TabIndex        =   91
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 42"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   28
         Left            =   1200
         TabIndex        =   92
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 28"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   36
         Left            =   1200
         TabIndex        =   93
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 36"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   41
         Left            =   1200
         TabIndex        =   94
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 41"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   46
         Left            =   1200
         TabIndex        =   95
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 46"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   48
         Left            =   1200
         TabIndex        =   96
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 48"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   53
         Left            =   1200
         TabIndex        =   97
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 53"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   54
         Left            =   1200
         TabIndex        =   98
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 54"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   58
         Left            =   1200
         TabIndex        =   99
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 58"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   70
         Left            =   2160
         TabIndex        =   100
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 70"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   69
         Left            =   2160
         TabIndex        =   101
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 69"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   59
         Left            =   2160
         TabIndex        =   102
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 59"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   64
         Left            =   2160
         TabIndex        =   103
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 64"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   66
         Left            =   2160
         TabIndex        =   104
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 66"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   71
         Left            =   2160
         TabIndex        =   105
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 71"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   77
         Left            =   2160
         TabIndex        =   106
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 77"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   79
         Left            =   2160
         TabIndex        =   107
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 79"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   83
         Left            =   2160
         TabIndex        =   108
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 83"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   84
         Left            =   2160
         TabIndex        =   109
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 84"
      End
      Begin l_RSSC.cmdButton World 
         Height          =   255
         Index           =   78
         Left            =   2160
         TabIndex        =   110
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "World 78"
      End
   End
   Begin VB.Label lblTrans 
      BackColor       =   &H00000000&
      Caption         =   "Transparency [currently 0%]:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11880
      TabIndex        =   22
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00000000&
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13560
      TabIndex        =   17
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Total Players Online:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   16
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "World:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12360
      TabIndex        =   15
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label lblServer 
      BackColor       =   &H00000000&
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13560
      TabIndex        =   14
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Players on world:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12240
      TabIndex        =   13
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblLoad 
      BackColor       =   &H00000000&
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13560
      TabIndex        =   12
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label secs 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   14520
      TabIndex        =   8
      Top             =   480
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   14400
      TabIndex        =   7
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   14040
      TabIndex        =   5
      Top             =   480
      Width           =   135
   End
   Begin VB.Label hrs2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13800
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Play Time:"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   12480
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label hrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13680
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label mins2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   14040
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label secs2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   14400
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label mins 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   14040
      TabIndex        =   9
      Top             =   480
      Width           =   375
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quality 
         Caption         =   "Quality"
         Begin VB.Menu mnuHigh 
            Caption         =   "High"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLow 
            Caption         =   "Low"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuAlwOnTop 
         Caption         =   "Always On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowScroll 
         Caption         =   "Show Scrollbars When Small"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Transparency %"
      End
      Begin VB.Menu mnuHideAdd 
         Caption         =   "Hide Adds From Free Worlds"
         Checked         =   -1  'True
      End
      Begin VB.Menu dash 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEditInstant 
         Caption         =   "Edit Instant-Message"
      End
      Begin VB.Menu mnuLogDet 
         Caption         =   "Edit Login Details"
      End
      Begin VB.Menu mnuDefWorld 
         Caption         =   "Set Default World"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Utilities 
      Caption         =   "Utilities"
      Begin VB.Menu Guides 
         Caption         =   "Guides"
         Begin VB.Menu mnuPriceGuide 
            Caption         =   "Price Guide"
         End
         Begin VB.Menu mnuQuestGuide 
            Caption         =   "Quest Guide"
         End
      End
   End
   Begin VB.Menu mnuSendLog 
      Caption         =   "Send Login"
   End
   Begin VB.Menu mnuSendInstant 
      Caption         =   "Send Instant Message"
   End
   Begin VB.Menu mnuScreenshot 
      Caption         =   "Take Screenshot"
   End
   Begin VB.Menu mnuSnip 
      Caption         =   "Snippets"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu mnuSendLogHow 
         Caption         =   "How to use Send Login"
      End
      Begin VB.Menu instantmesshelp 
         Caption         =   "How to use Instant Message"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Check for Updates"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'screenshot code not by levymetal, getting server players
'not by levymetal (but total players is).

Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public UserName As String
Public Password As String
Public qual As Integer
Public InstantMessage As String
Public server As Long

Public setWidth As Long
Public setHeight As Long

Public LastTrans As Long

Dim Parse1 As Integer
Dim strHTML As String
Dim strHTML2 As String

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_SNAPSHOT = &H2C

Dim myver As Integer

Private Sub chkAR_Click()
If chkAR = 1 Then
tmrGetLoad.Enabled = True
Else
tmrGetLoad.Enabled = False
End If
End Sub

Private Sub cmdGetload_Click()
If server = 0 Then
server = InputBox("No server is specified! What server are you on? (Number)")
Call GetLoad
Else
Call GetLoad
End If
End Sub

Private Sub cmdIMFlood_Click()
MsgBox "After you click OK, PLEASE click somewhere in runescape, within 2 seconds. Flooding will start after 2 seconds and may cause program to crash if you dont click in runescape!!!"
tmrIMFloodS.Enabled = True
cmdStopIM.Enabled = True
cmdIMFlood.Enabled = False
End Sub

Private Sub cmdShowFree_Click()
frameFree.Visible = True
frameMembers.Visible = False

mnuHideAdd.Enabled = True
End Sub

Private Sub cmdShowMem_Click()
frameFree.Visible = False
frameMembers.Visible = True

web.Top = -25
pb.Top = 9425
mnuHideAdd.Checked = False
mnuHideAdd.Enabled = False
End Sub

Public Sub StartTrans()

End Sub

Private Sub cmdStopIM_Click()
tmrIMFlood.Enabled = False
cmdStopIM.Enabled = False
cmdIMFlood.Enabled = True
End Sub

Private Sub cmdStopIM_MouseMove()
Call cmdStopIM_Click
End Sub

Private Sub Form_Load()
Dim ss() As String
Dim File
Dim tmpsrv As Integer
On Error Resume Next

myver = 1.9

Me.Show
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

Call sldrTrans_scroll
mnuHideAdd.Checked = False
mnuShowScroll.Checked = False

qual = 1
mnuAlwOnTop.Checked = False
setWidth = -160 - 80
setHeight = 450 - 25
Call Form_Resize

Open "c:\windows\lrsscset.ini" For Input As #1
Line Input #1, File
ss = Split(File, "*:.:*")
UserName = ss(0)
Password = ss(1)
qual = ss(2)
server = ss(3)
InstantMessage = ss(4)
Close #1

If UserName = "" Then UserName = "blank"
If Password = "" Then Password = "blank"
If server = 0 Then server = 3

If qual = 1 Then
mnuHigh.Checked = False
mnuLow.Checked = True
ElseIf qual = 0 Then
mnuHigh.Checked = True
mnuLow.Checked = False
End If

frameFree.Visible = True
frameMembers.Visible = False

tmpsrv = server
Call World_Click(tmpsrv)

Call Save

End Sub

Private Sub Form_LostFocus()
'tmrIMFlood.Enabled = False
tmrIMFloodS.Enabled = False
Call cmdStopIM_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
'web = 11655 x9495
'form = 11550x10035

If mnuHideAdd.Checked = False Then
If Form1.Height < 10035 Then web.Height = Form1.Height - setHeight + 25
Else
If Form1.Height < 10035 Then web.Height = Form1.Height - setHeight + 1395
End If
If Form1.Width < 11550 Then web.Width = Form1.Width - setWidth + 80

'If web.Width > 11655 Then web.Width = 11655
End Sub

Private Sub Form_Unload(Cancel As Integer)
Inet1.Cancel
Inet2.Cancel
End
End Sub

Private Sub instantmesshelp_Click()
MsgBox "Click File > Edit Instant-Message. Enter your message, then whenever you are in runescape just press 'Send Instant Message' and your message will be sent. Good for when you are selling/buying things. ", vbInformation
End Sub

Private Sub mnuAbout_Click()
MsgBox "Made by Christian, aka levymetal. Email = levymetal@gmail.com. levymetal's Runescape Switcher Client is freeware, and you can distribute it as much as you want as long as my name stays on it.", vbInformation
End Sub

Private Sub mnuAlwOnTop_Click()
If mnuAlwOnTop.Checked = False Then
Call Always_On_Top(Me.hwnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)
mnuAlwOnTop.Checked = True
Else
Call Always_On_Top(Me.hwnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, False)
mnuAlwOnTop.Checked = False
End If
End Sub

Private Sub mnuDefWorld_Click()
On Error GoTo e
Dim temp As Integer
temp = server
rep:
server = InputBox("Please enter the number of the world you want the program to load on startup:", , server)
If server > 84 Then
MsgBox "World 84 is the highest world! Please enter a valid world!"
GoTo rep
ElseIf server = 60 Or server = 65 Then
MsgBox "Servers 60 and 65 dont exist, please enter another server!"
GoTo rep
End If
Call Save
Exit Sub
e:
server = temp
End Sub

Private Sub mnuEditInstant_Click()
Dim temp As String
temp = InstantMessage
InstantMessage = InputBox("Please enter your Instant Message here:", , InstantMessage)
If InstantMessage = "" Then InstantMessage = temp
Call Save
End Sub

Private Sub mnuHideAdd_Click()
If mnuHideAdd.Checked = True Then
mnuHideAdd.Checked = False
web.Top = -25
pb.Top = 9425
Else
mnuHideAdd.Checked = True
web.Top = -1400
pb.Top = pb.Top - 1375
End If

Call Form_Resize
End Sub

Private Sub mnuHigh_Click()
If mnuHigh.Checked = True Then
Else
mnuHigh.Checked = True
mnuLow.Checked = False
qual = 0
End If
Call Save
End Sub

Private Sub mnuLogDet_Click()
Dim tempuser As String
Dim temppass As String

tempuser = UserName
temppass = Password

UserName = InputBox("Please enter your username", , UserName)
If UserName = "" Then UserName = tempuser
Password = InputBox("Please enter your password")
If Password = "" Then Password = temppass

Call Save
End Sub

Private Sub Save()
Open "c:\windows\lrsscset.ini" For Output As #1
Print #1, UserName & "*:.:*" & Password & "*:.:*" & qual & "*:.:*" & server & "*:.:*" & InstantMessage
Close #1
End Sub

Private Sub mnuLow_Click()
If mnuLow.Checked = True Then
Else
mnuLow.Checked = True
mnuHigh.Checked = False
qual = 1
End If
Call Save
End Sub

Private Sub mnuPriceGuide_Click()
LaunchURLInNewBrowser "http://www.runescapecommunity.com/lofiversion/index.php?t136820.html"
End Sub

Private Sub mnuQuestGuide_Click()
LaunchURLInNewBrowser "http://www.zybez.com/quests.php"
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub mnuScreenshot_Click()
Clipboard.Clear
Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
DoEvents
Picture1.Picture = Clipboard.GetData()
SavePicture Picture1, App.Path & "\img." & hrs & "." & mins & "." & secs & ".bmp"
End Sub

Private Sub mnuSendInstant_Click()
If InstantMessage = "" Then
Select Case MsgBox("You havent set an Instant Message yet! Would you like to set one now?", vbYesNo)
Case vbYes
Call mnuEditInstant_Click
Case vbNo
Exit Sub
End Select
End If
SendKeys "{tab}"
SendKeys InstantMessage
SendKeys "{enter}"
End Sub

Private Sub mnuSendLog_Click()
SendKeys "{tab}"
SendKeys Password
SendKeys "{tab}"
SendKeys UserName
End Sub

Private Sub mnuSendLogHow_Click()
MsgBox "click 'File > Edit login Details'. Enter your Username and Password. Then click on a server, and click Existing User. Then click Send Login, and your details are entered!" & vbCrLf & vbCrLf & "Note: You still must MANUALLY click login after this.", vbInformation
End Sub

Private Sub mnuShowScroll_Click()
If mnuShowScroll.Checked = True Then
setWidth = -160 - 80
setHeight = 450 - 25
mnuShowScroll.Checked = False
Else
setWidth = 90
setHeight = 705
mnuShowScroll.Checked = True
End If
Call Form_Resize
End Sub

Private Sub mnuSnip_Click()
Form2.Show
End Sub

Private Sub mnuTrans_Click()
On Error Resume Next
sldrTrans = InputBox("Please enter the % of transparency you want:", "Transparency", sldrTrans)
Call sldrTrans_scroll
End Sub

Private Sub mnuUpdate_Click()
Dim html As String
Dim ver() As String
Dim dat() As String
Dim size() As String
Dim info() As String

html = inetUpdate.OpenURL("http://70.85.146.146/~levy/l_rssc.html")
info = Split(html, ":")

dat = Split(info(4), "</body>")
ver = Split(info(3), "<br>")
size = Split(info(2), "<br>")

If ver(0) > myver Then
Select Case MsgBox("There is a newer version of l_RSSC, version" & ver(0) & vbCrLf & "Its dated" & dat(0) & vbCrLf & "Download is" & size(0) & ". Download now?", vbYesNo)
Case vbYes
web.Navigate "http://70.85.146.146/~levy/vbpro/l_rssc.exe"
End
Case vbNo
MsgBox "It is highly recommended that you update! Please update soon!"
End Select
Else
MsgBox "There is no newer version at the moment."
End If
End Sub

Private Sub Timer1_Timer()

If sldrTrans = 0 Then
Else
Call sldrTrans_scroll
End If

If Form1.Width > 11730 Then web.Width = 11655
If Form1.Height > 9840 Then web.Height = 9495

If secs = 0 Then secs = "00"
If secs = 1 Then secs = "01"
If secs = 2 Then secs = "02"
If secs = 3 Then secs = "03"
If secs = 4 Then secs = "04"
If secs = 5 Then secs = "05"
If secs = 6 Then secs = "06"
If secs = 7 Then secs = "07"
If secs = 8 Then secs = "08"
If secs = 9 Then secs = "09"

If mins = 0 Then mins = "00"
If mins = 1 Then mins = "01"
If mins = 2 Then mins = "02"
If mins = 3 Then mins = "03"
If mins = 4 Then mins = "04"
If mins = 5 Then mins = "05"
If mins = 6 Then mins = "06"
If mins = 7 Then mins = "07"
If mins = 8 Then mins = "08"
If mins = 9 Then mins = "09"

If hrs = 0 Then hrs = "00"
If hrs = 1 Then hrs = "01"
If hrs = 2 Then hrs = "02"
If hrs = 3 Then hrs = "03"
If hrs = 4 Then hrs = "04"
If hrs = 5 Then hrs = "05"
If hrs = 6 Then hrs = "06"
If hrs = 7 Then hrs = "07"
If hrs = 8 Then hrs = "08"
If hrs = 9 Then hrs = "09"
End Sub

Private Sub Timer2_Timer()
Call GetLoad
End Sub


Private Sub tmrCount_Timer()
secs = secs + 1
If secs = 60 Then
secs = 0
mins = mins + 1
End If
If mins = 60 Then
mins = 0
hrs = hrs + 1
End If
End Sub

Private Sub tmrGetLoad_Timer()
Call GetLoad
End Sub

Private Sub tmrIMFlood_Timer()
Call mnuSendInstant_Click
End Sub

Private Sub tmrIMFloodS_Timer()
tmrIMFlood.Enabled = True
tmrIMFloodS.Enabled = False
End Sub

Private Sub Web_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
If URL = "http://www.runescape.com/" Then Cancel = True
End Sub

Function server_players(World As Long, html As String)
World = CInt(World)
Select Case World
Case 1, 3, 4, 5, 7, 8, 9
Parse1 = 31
Case 2, 6, 10 To 79
Parse1 = 32
Case Else
server_players = "ERROR: please report"
Exit Function
End Select

If IsNumeric(Mid(html, (InStr(1, html, "World " & World)) + CStr(Len(World)) + Parse1, 4)) Then
server_players = Mid(html, (InStr(1, html, "World " & World)) + CStr(Len(World)) + Parse1, 4)
Else
server_players = "2000"
End If

End Function

Private Sub GetLoad()
On Error Resume Next
strHTML = Inet1.OpenURL("http://www.runescape.com/aff/runescape/serverlist.cgi?plugin=0&lores.x=79&lores.y=42")
strHTML2 = Inet1.OpenURL("http://www.runescape.com/aff/runescape/title.html")

lblServer = server
lblLoad.Caption = server_players(server, strHTML)
lblTotal.Caption = total_server(strHTML2)
End Sub

Function total_server(html As String)
Dim ss() As String
Dim temp As String
ss = Split(html, "currently ")
temp = ss(1)
ss = Split(temp, " people")
total_server = ss(0)

End Function

Private Sub World_Click(Index As Integer)
Dim SName As String
Dim temp As String
Select Case Index
Case 1
SName = "ul2"
Case 2
SName = "ul4"
Case 3
SName = "po3"
Case 4
SName = "po4"
Case 5
SName = "po5"
Case 6
SName = "po6"
Case 7
SName = "above2"
Case 8
SName = "above3"
Case 9
SName = "above4"
Case 10
SName = "jolt7"
Case 11
SName = "jolt8"
Case 12
SName = "jolt9"
Case 13
SName = "nl3"
Case 14
SName = "nl4"
Case 15
SName = "uk2"
Case 16
SName = "uk3"
Case 17
SName = "tor1"
Case 18
SName = "tor2"
Case 19
SName = "cet1"
Case 20
SName = "cet2"
Case 21
SName = "cet5"
Case 22
SName = "nl1"
Case 23
SName = "uk4"
Case 24
SName = "uk5"
Case 25
SName = "cet6"
Case 26
SName = "ul5"
Case 27
SName = "nl5"
Case 28
SName = "nl6"
Case 29
SName = "ul6"
Case 30
SName = "po7"
Case 31
SName = "po8"
Case 32
SName = "ul1"
Case 33
SName = "at1"
Case 34
SName = "at2"
Case 35
SName = "at3"
Case 36
SName = "at4"
Case 37
SName = "tor3"
Case 38
SName = "planet1"
Case 39
SName = "planet2"
Case 40
SName = "planet3"
Case 41
SName = "planet4"
Case 42
SName = "po2"
Case 43
SName = "sl11"
Case 44
SName = "at6"
Case 45
SName = "planet5"
Case 46
SName = "planet6"
Case 47
SName = "above5"
Case 48
SName = "above6"
Case 49
SName = "ams1"
Case 50
SName = "ams2"
Case 51
SName = "ams3"
Case 52
SName = "ams4"
Case 53
SName = "ams5"
Case 54
SName = "ams6"
Case 55
SName = "ch1"
Case 56
SName = "cet4"
Case 57
SName = "ch3"
Case 58
SName = "ch4"
Case 59
SName = "ch5"
Case 60
'no such world
Case 61
SName = "se1"
Case 62
SName = "se2"
Case 63
SName = "se3"
Case 64
SName = "se4"
Case 65
'no such world
Case 66
SName = "se6"
Case 67
SName = "jolt10"
Case 68
SName = "jolt11"
Case 69
SName = "jolt12"
Case 70
SName = "sl10"
Case 71
SName = "uk7"
Case 72
SName = "sl1"
Case 73
SName = "sl2"
Case 74
SName = "sl3"
Case 75
SName = "sl4"
Case 76
SName = "sl5"
Case 77
SName = "sl6"
Case 78
SName = "sl7"
Case 79
SName = "sl8"
Case 80
SName = "jolt1"
Case 81
SName = "jolt2"
Case 82
SName = "jolt3"
Case 83
SName = "jolt4"
Case 84
SName = "jolt5"
End Select

web.Navigate ("http://" & SName & ".runescape.com:80/rs2.cgi?plugin=0&lowmem=" & qual & "&affiliate=runescape&randval=771846832")
server = Index
Form1.Caption = "levymetal's RS Client - World " & server

Call GetLoad
End Sub

Private Sub sldrTrans_scroll()
If sldrTrans > 90 Then sldrTrans = 90

If sldrTrans = 0 Then
SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And Not (WS_EX_LAYERED)
Else
If LastTrans = 0 Then
Dim NormalWindowStyle As Long
NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
Else
End If
End If

lblTrans.Caption = "Transparency [currently " & sldrTrans & "%]:"

SetLayeredWindowAttributes Me.hwnd, 0, 255 * (1 - (Val(sldrTrans) / 100)), LWA_ALPHA
LastTrans = sldrTrans
End Sub
