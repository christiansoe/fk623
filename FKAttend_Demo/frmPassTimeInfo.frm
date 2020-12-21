VERSION 5.00
Begin VB.Form frmPassTimeInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pass Time Info"
   ClientHeight    =   6840
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   9915
   Icon            =   "frmPassTimeInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDoorState 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmPassTimeInfo.frx":0442
      Left            =   240
      List            =   "frmPassTimeInfo.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   105
      Top             =   6000
      Width           =   1560
   End
   Begin VB.Frame Frame4 
      Caption         =   "Unlock Group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   5520
      TabIndex        =   74
      Top             =   4500
      Width           =   4200
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3240
         TabIndex        =   86
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2520
         TabIndex        =   85
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   84
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   83
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   82
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3240
         TabIndex        =   81
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   80
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   79
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   78
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtGroupMatch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   77
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdSetGroupMatch 
         Caption         =   "Set Match"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   2160
         TabIndex        =   76
         Top             =   1560
         Width           =   1800
      End
      Begin VB.CommandButton cmdGetGroupMatch 
         Caption         =   "Get Match"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   240
         TabIndex        =   75
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label lblLabel 
         Caption         =   "    6          7           8          9         10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   10
         Left            =   360
         TabIndex        =   99
         Top             =   860
         Width           =   3500
      End
      Begin VB.Label lblLabel 
         Caption         =   "    1          2           3          4          5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   9
         Left            =   360
         TabIndex        =   98
         Top             =   260
         Width           =   3500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Group Access"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   5520
      TabIndex        =   71
      Top             =   2250
      Width           =   4200
      Begin VB.TextBox txtGroupNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2110
         TabIndex        =   97
         Text            =   "1"
         Top             =   360
         Width           =   500
      End
      Begin VB.TextBox txtGroupPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   3320
         TabIndex        =   93
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.TextBox txtGroupPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   2110
         TabIndex        =   92
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.TextBox txtGroupPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   900
         TabIndex        =   91
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.CommandButton cmdGetGroupPassTime 
         Caption         =   "Get Group"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   240
         TabIndex        =   73
         Top             =   1560
         Width           =   1800
      End
      Begin VB.CommandButton cmdSetGroupPassTime 
         Caption         =   "Set Group"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   2160
         TabIndex        =   72
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label lblLabel 
         Caption         =   "Group (1~5)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   1440
         TabIndex        =   103
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   2920
         TabIndex        =   96
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1710
         TabIndex        =   95
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   500
         TabIndex        =   94
         Top             =   1000
         Width           =   400
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Access"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   5520
      TabIndex        =   62
      Top             =   0
      Width           =   4200
      Begin VB.TextBox txtUserPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   3700
         TabIndex        =   69
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.TextBox txtUserPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   2700
         TabIndex        =   68
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.TextBox txtUserPassTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   1800
         TabIndex        =   67
         Text            =   "0"
         Top             =   960
         Width           =   380
      End
      Begin VB.TextBox txtUserID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         TabIndex        =   66
         Text            =   "1"
         Top             =   360
         Width           =   860
      End
      Begin VB.TextBox txtGroupID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   720
         TabIndex        =   65
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdGetUserPasstime 
         Caption         =   "Get User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   240
         TabIndex        =   64
         Top             =   1560
         Width           =   1800
      End
      Begin VB.CommandButton cmdSetUserPassTime 
         Caption         =   "Set User "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   2160
         TabIndex        =   63
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3300
         TabIndex        =   90
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2300
         TabIndex        =   89
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label lblLabel 
         Caption         =   "TZ1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1400
         TabIndex        =   88
         Top             =   1000
         Width           =   400
      End
      Begin VB.Label lblLabel 
         Caption         =   "Group (1~5)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   120
         TabIndex        =   87
         Top             =   885
         Width           =   615
      End
      Begin VB.Label lblUserID 
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   70
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGetDoorState 
      Caption         =   "Get Door State"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3600
      TabIndex        =   2
      Top             =   6000
      Width           =   1710
   End
   Begin VB.CommandButton cmdSetDoorState 
      Caption         =   "Set Door State"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1830
      TabIndex        =   1
      Top             =   6000
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Zone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      Begin VB.CommandButton cmdWrite 
         Caption         =   "Write"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   102
         Top             =   3960
         Width           =   1125
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Read"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   101
         Top             =   3960
         Width           =   1245
      End
      Begin VB.TextBox txtPassTimeID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   840
         TabIndex        =   100
         Text            =   "1"
         Top             =   3960
         Width           =   860
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "0"
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "0"
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         DataField       =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "0"
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         DataField       =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   720
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "0"
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "0"
         Top             =   2470
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "0"
         Top             =   2470
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "0"
         Top             =   2070
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "0"
         Top             =   2070
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "0"
         Top             =   1670
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "0"
         Top             =   1670
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "0"
         Top             =   1270
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "0"
         Top             =   1270
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "0"
         Top             =   2470
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   720
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "0"
         Top             =   2470
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   720
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "0"
         Top             =   2070
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "0"
         Top             =   2070
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "0"
         Top             =   1670
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   720
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "0"
         Top             =   1670
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   720
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "0"
         Top             =   1270
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "0"
         Top             =   1270
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   720
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "0"
         Top             =   2880
         Width           =   630
      End
      Begin VB.TextBox txtStartHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   720
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "0"
         Top             =   3280
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "0"
         Top             =   2880
         Width           =   630
      End
      Begin VB.TextBox txtStartMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   3280
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   2880
         Width           =   630
      End
      Begin VB.TextBox txtEndHour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         Top             =   3280
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         Top             =   2880
         Width           =   630
      End
      Begin VB.TextBox txtEndMinute 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   3280
         Width           =   630
      End
      Begin VB.Label lblLabel 
         Caption         =   "   TZ (1~50)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   120
         TabIndex        =   104
         Top             =   3900
         Width           =   640
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   160
         TabIndex        =   61
         Top             =   920
         Width           =   500
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1500
         TabIndex        =   60
         Top             =   920
         Width           =   90
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   160
         TabIndex        =   59
         Top             =   3320
         Width           =   500
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   160
         TabIndex        =   58
         Top             =   2920
         Width           =   500
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   160
         TabIndex        =   57
         Top             =   2520
         Width           =   500
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Wen"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   160
         TabIndex        =   56
         Top             =   2120
         Width           =   500
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   160
         TabIndex        =   55
         Top             =   1720
         Width           =   500
      End
      Begin VB.Label lblWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   160
         TabIndex        =   54
         Top             =   1320
         Width           =   500
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   1500
         TabIndex        =   53
         Top             =   3320
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   4020
         TabIndex        =   52
         Top             =   3320
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   4020
         TabIndex        =   51
         Top             =   2920
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   4020
         TabIndex        =   50
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   4020
         TabIndex        =   49
         Top             =   2120
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   4020
         TabIndex        =   48
         Top             =   1720
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   4020
         TabIndex        =   47
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblEndSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   4020
         TabIndex        =   46
         Top             =   920
         Width           =   90
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   1500
         TabIndex        =   45
         Top             =   2920
         Width           =   90
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   1500
         TabIndex        =   44
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   1500
         TabIndex        =   43
         Top             =   2120
         Width           =   90
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   1500
         TabIndex        =   42
         Top             =   1720
         Width           =   90
      End
      Begin VB.Label lblStartSep 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   1500
         TabIndex        =   41
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   2640
         TabIndex        =   40
         Top             =   3195
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   2640
         TabIndex        =   39
         Top             =   2775
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   2640
         TabIndex        =   37
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   2640
         TabIndex        =   36
         Top             =   2355
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   2640
         TabIndex        =   35
         Top             =   1950
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   2640
         TabIndex        =   34
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label lblMidSep 
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   2640
         TabIndex        =   33
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label lblStartTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   975
         TabIndex        =   32
         Top             =   480
         Width           =   1070
      End
      Begin VB.Label lblEndTime 
         BackStyle       =   0  'Transparent
         Caption         =   "End Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3480
         TabIndex        =   31
         Top             =   480
         Width           =   1070
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   38
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmPassTimeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DataLen = 28
Private mlngTmpBuff(DataLen) As Long

Private mPassCtrlInfo As PASSCTRLTIME
Private mUserPassInfo As USERPASSINFO
Private mGroupPassInfo As GROUPPASSINFO
Private mGroupMatchInfo As GROUPMATCHINFO
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
    cmbDoorState.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub cmdSetDoorState_Click()
    Dim vnResultCode As Long
    Dim vnState As Long
    
    cmdSetDoorState.Enabled = False
    lblMessage.Caption = "Setting Door ..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetDoorState.Enabled = True
        Exit Sub
    End If

    Select Case cmbDoorState.ListIndex
        Case 0
            vnState = DOOR_CONROLRESET
        Case 1
            vnState = DOOR_OPEND
        Case 2
            vnState = DOOR_CLOSED
        Case 3
            vnState = DOOR_COMMNAD
    End Select
    vnResultCode = FK_SetDoorStatus(fnCommHandleIndex, vnState)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetDoorState.Enabled = True
End Sub

Private Sub cmdGetDoorState_Click()
    Dim vnResultCode As Long
    Dim vnState As Long

    cmdGetDoorState.Enabled = False
    lblMessage.Caption = "Getting Door State..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetDoorState.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_GetDoorStatus(fnCommHandleIndex, vnState)
    If vnResultCode = RUN_SUCCESS Then
        If vnState = DOOR_CONROLRESET Then
            lblMessage.Caption = "State Reset!"
        ElseIf vnState = DOOR_OPEND Then
            lblMessage.Caption = "Door Open!"
        ElseIf vnState = DOOR_CLOSED Then
            lblMessage.Caption = "Door Close!"
        ElseIf vnState = DOOR_COMMNAD Then
            lblMessage.Caption = "Command Open... Door Close!"
        Else
            lblMessage.Caption = "Door - Unknown!"
        End If
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetDoorState.Enabled = True
End Sub

Private Sub cmdGetGroupMatch_Click()
    Dim vnResultCode As Long

    cmdGetGroupMatch.Enabled = False
    lblMessage.Caption = "Getting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetGroupMatch.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_GetGroupMatch(fnCommHandleIndex, _
                                mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        CopyMemory mGroupMatchInfo, mlngTmpBuff(0), Len(mGroupMatchInfo)
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetGroupMatch.Enabled = True
End Sub

Private Sub cmdGetGroupPassTime_Click()
    Dim vGroupID As Long
    Dim vnResultCode As Long

    cmdGetGroupPassTime.Enabled = False
    lblMessage.Caption = "Getting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetGroupPassTime.Enabled = True
        Exit Sub
    End If

    vGroupID = Val(txtGroupNum.Text)
    vnResultCode = FK_GetGroupPassTime(fnCommHandleIndex, _
                                    vGroupID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        CopyMemory mGroupPassInfo, mlngTmpBuff(0), Len(mGroupPassInfo)
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetGroupPassTime.Enabled = True
End Sub

Private Sub cmdGetUserPassTime_Click()
    Dim vEnrollNumber As Long
    Dim vGroupID As Long
    Dim vnResultCode As Long

    cmdGetUserPasstime.Enabled = False
    lblMessage.Caption = "Getting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetUserPasstime.Enabled = True
        Exit Sub
    End If

    vEnrollNumber = Val(txtUserID.Text)
    vnResultCode = FK_GetUserPassTime(fnCommHandleIndex, _
                                    vEnrollNumber, vGroupID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        txtGroupID.Text = vGroupID
        lblMessage.Caption = "Success!"
        CopyMemory mUserPassInfo, mlngTmpBuff(0), Len(mUserPassInfo)
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetUserPasstime.Enabled = True
End Sub

Private Sub cmdRead_Click()
    Dim vnPassTimeID As Long
    Dim vnResultCode As Long

    cmdRead.Enabled = False
    lblMessage.Caption = ""
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdRead.Enabled = True
        Exit Sub
    End If

    vnPassTimeID = 1
    If IsNumeric(txtPassTimeID.Text) Then
        vnPassTimeID = Val(txtPassTimeID.Text)
    End If
    If vnPassTimeID < 1 Or vnPassTimeID > 50 Then
        vnPassTimeID = 1
    End If
    txtPassTimeID.Text = Trim(Str(vnPassTimeID))
    vnResultCode = FK_GetPassTime(fnCommHandleIndex, _
                                vnPassTimeID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        CopyMemory mPassCtrlInfo, mlngTmpBuff(0), Len(mPassCtrlInfo)
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdRead.Enabled = True
End Sub

Private Sub cmdSetGroupMatch_Click()
    Dim vnResultCode As Long

    cmdSetGroupMatch.Enabled = False
    lblMessage.Caption = "Setting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetGroupMatch.Enabled = True
        Exit Sub
    End If

    GetValue
    CopyMemory mlngTmpBuff(0), mGroupMatchInfo, Len(mGroupMatchInfo)
    
    vnResultCode = FK_SetGroupMatch(fnCommHandleIndex, _
                                    mlngTmpBuff(0), UBound(mlngTmpBuff))
    
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetGroupMatch.Enabled = True
End Sub

Private Sub cmdSetGroupPassTime_Click()
    Dim vGroupID As Long
    Dim vnResultCode As Long

    cmdSetGroupPassTime.Enabled = False
    lblMessage.Caption = "Setting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetGroupPassTime.Enabled = True
        Exit Sub
    End If

    GetValue
    CopyMemory mlngTmpBuff(0), mGroupPassInfo, Len(mGroupPassInfo)
    
    vGroupID = Val(txtGroupNum.Text)
    vnResultCode = FK_SetGroupPassTime(fnCommHandleIndex, _
                                    vGroupID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetGroupPassTime.Enabled = True
End Sub

Private Sub cmdSetUserPassTime_Click()
    Dim vEnrollNumber As Long
    Dim vGroupID As Long
    Dim vnResultCode As Long

    cmdSetUserPassTime.Enabled = False
    lblMessage.Caption = "Setting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetUserPassTime.Enabled = True
        Exit Sub
    End If

    GetValue
    CopyMemory mlngTmpBuff(0), mUserPassInfo, Len(mUserPassInfo)
    
    vEnrollNumber = Val(txtUserID.Text)
    vGroupID = Val(txtGroupID.Text)
    vnResultCode = FK_SetUserPassTime(fnCommHandleIndex, _
                                vEnrollNumber, vGroupID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
        ShowValue
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetUserPassTime.Enabled = True
End Sub

Private Sub cmdWrite_Click()
    Dim vnPassTimeID As Long
    Dim vnResultCode As Long

    cmdWrite.Enabled = False
    lblMessage.Caption = ""
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdWrite.Enabled = True
        Exit Sub
    End If

    GetValue
    CopyMemory mlngTmpBuff(0), mPassCtrlInfo, Len(mPassCtrlInfo)
    
    vnPassTimeID = 1
    If IsNumeric(txtPassTimeID.Text) Then
        vnPassTimeID = Val(txtPassTimeID.Text)
    End If
    If vnPassTimeID < 1 Or vnPassTimeID > 50 Then
        vnPassTimeID = 1
    End If
    txtPassTimeID.Text = Trim(Str(vnPassTimeID))
    vnResultCode = FK_SetPassTime(fnCommHandleIndex, _
                                vnPassTimeID, mlngTmpBuff(0), UBound(mlngTmpBuff))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Success!"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdWrite.Enabled = True
End Sub

Private Sub ShowValue()
    Dim vnii As Integer

    For vnii = 0 To MAX_PASSCTRL_COUNT - 1
        txtStartHour(vnii).Text = CStr(mPassCtrlInfo.mPassCtrlTime(vnii).StartHour)
        txtStartMinute(vnii).Text = CStr(mPassCtrlInfo.mPassCtrlTime(vnii).StartMinute)
        txtEndHour(vnii).Text = CStr(mPassCtrlInfo.mPassCtrlTime(vnii).EndHour)
        txtEndMinute(vnii).Text = CStr(mPassCtrlInfo.mPassCtrlTime(vnii).EndMinute)
    Next vnii

    For vnii = 0 To MAX_USERPASSINFO_COUNT - 1
        txtUserPassTime(vnii).Text = CStr(mUserPassInfo.UserPassID(vnii))
    Next vnii

    For vnii = 0 To MAX_GROUPPASSINFO_COUNT - 1
        txtGroupPassTime(vnii).Text = CStr(mGroupPassInfo.GroupPassID(vnii))
    Next vnii

    For vnii = 0 To MAX_GROUPMATCHINFO_COUNT - 1
        txtGroupMatch(vnii).Text = CStr(mGroupMatchInfo.GroupMatch(vnii))
    Next vnii

End Sub

Private Sub GetValue()
    Dim vnii As Integer

    For vnii = 0 To MAX_PASSCTRL_COUNT - 1
        mPassCtrlInfo.mPassCtrlTime(vnii).StartHour = CByte(txtStartHour(vnii).Text)
        mPassCtrlInfo.mPassCtrlTime(vnii).StartMinute = CByte(txtStartMinute(vnii).Text)
        mPassCtrlInfo.mPassCtrlTime(vnii).EndHour = CByte(txtEndHour(vnii).Text)
        mPassCtrlInfo.mPassCtrlTime(vnii).EndMinute = CByte(txtEndMinute(vnii).Text)
    Next vnii

    For vnii = 0 To MAX_USERPASSINFO_COUNT - 1
        mUserPassInfo.UserPassID(vnii) = CByte(txtUserPassTime(vnii).Text)
    Next vnii

    For vnii = 0 To MAX_GROUPPASSINFO_COUNT - 1
        mGroupPassInfo.GroupPassID(vnii) = CByte(txtGroupPassTime(vnii).Text)
    Next vnii

    For vnii = 0 To MAX_GROUPMATCHINFO_COUNT - 1
        mGroupMatchInfo.GroupMatch(vnii) = CInt(txtGroupMatch(vnii).Text)
    Next vnii
End Sub


