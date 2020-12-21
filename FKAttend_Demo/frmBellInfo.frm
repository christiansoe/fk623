VERSION 5.00
Begin VB.Form frmBellInfo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting Bell Time"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBellInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBellCount 
      Height          =   420
      Left            =   5040
      TabIndex        =   127
      Top             =   5040
      Width           =   630
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   510
      Left            =   6720
      TabIndex        =   125
      Top             =   5040
      Width           =   1500
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   0
      Left            =   1290
      TabIndex        =   122
      Top             =   1350
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   0
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   121
      Top             =   1290
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   23
      Left            =   7890
      TabIndex        =   118
      Top             =   4500
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   23
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   117
      Top             =   4440
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   23
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   116
      Top             =   4440
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   22
      Left            =   7890
      TabIndex        =   113
      Top             =   4050
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   22
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   112
      Top             =   3990
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   22
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   111
      Top             =   3990
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   21
      Left            =   7890
      TabIndex        =   108
      Top             =   3600
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   21
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   107
      Top             =   3540
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   21
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   106
      Top             =   3540
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   20
      Left            =   7890
      TabIndex        =   103
      Top             =   3150
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   20
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   102
      Top             =   3090
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   20
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   101
      Top             =   3090
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   19
      Left            =   7890
      TabIndex        =   98
      Top             =   2700
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   19
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   97
      Top             =   2640
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   19
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   96
      Top             =   2640
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   18
      Left            =   7890
      TabIndex        =   93
      Top             =   2250
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   18
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   92
      Top             =   2190
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   18
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   91
      Top             =   2190
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   17
      Left            =   7890
      TabIndex        =   88
      Top             =   1800
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   17
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   87
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   17
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   86
      Top             =   1740
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   16
      Left            =   7890
      TabIndex        =   83
      Top             =   1350
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   16
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   82
      Top             =   1290
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   16
      Left            =   9180
      MaxLength       =   2
      TabIndex        =   81
      Top             =   1290
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   15
      Left            =   4590
      TabIndex        =   78
      Top             =   4500
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   15
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   77
      Top             =   4440
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   15
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   76
      Top             =   4440
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   14
      Left            =   4590
      TabIndex        =   73
      Top             =   4050
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   14
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   72
      Top             =   3990
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   14
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   71
      Top             =   3990
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   13
      Left            =   4590
      TabIndex        =   68
      Top             =   3600
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   13
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   67
      Top             =   3540
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   13
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   66
      Top             =   3540
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   12
      Left            =   4590
      TabIndex        =   63
      Top             =   3150
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   12
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   62
      Top             =   3090
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   12
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   61
      Top             =   3090
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   11
      Left            =   4590
      TabIndex        =   58
      Top             =   2700
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   11
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   57
      Top             =   2640
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   11
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   56
      Top             =   2640
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   10
      Left            =   4590
      TabIndex        =   53
      Top             =   2250
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   10
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   52
      Top             =   2190
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   10
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   51
      Top             =   2190
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   9
      Left            =   4590
      TabIndex        =   48
      Top             =   1800
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   9
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   47
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   9
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   46
      Top             =   1740
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   8
      Left            =   4590
      TabIndex        =   43
      Top             =   1350
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   8
      Left            =   4980
      MaxLength       =   2
      TabIndex        =   42
      Top             =   1290
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   8
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   41
      Top             =   1290
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   7
      Left            =   1290
      TabIndex        =   38
      Top             =   4500
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   7
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   37
      Top             =   4440
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   7
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   36
      Top             =   4440
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   6
      Left            =   1290
      TabIndex        =   33
      Top             =   4050
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   6
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3990
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   6
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   31
      Top             =   3990
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   5
      Left            =   1290
      TabIndex        =   28
      Top             =   3600
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   5
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   27
      Top             =   3540
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   5
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   26
      Top             =   3540
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   4
      Left            =   1290
      TabIndex        =   23
      Top             =   3150
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   4
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   22
      Top             =   3090
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   4
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   21
      Top             =   3090
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   3
      Left            =   1290
      TabIndex        =   18
      Top             =   2700
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   3
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   17
      Top             =   2640
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   3
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   16
      Top             =   2640
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   2
      Left            =   1290
      TabIndex        =   13
      Top             =   2250
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   2
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2190
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   2
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2190
      Width           =   630
   End
   Begin VB.CheckBox chkValid 
      Caption         =   "Time1"
      Height          =   300
      Index           =   1
      Left            =   1290
      TabIndex        =   8
      Top             =   1800
      Width           =   210
   End
   Begin VB.TextBox txtHour 
      Height          =   420
      Index           =   1
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   1
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txtMinute 
      Height          =   420
      Index           =   0
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1290
      Width           =   630
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   510
      Left            =   8400
      TabIndex        =   0
      Top             =   5040
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bell Count :"
      Height          =   285
      Left            =   3705
      TabIndex        =   126
      Top             =   5130
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point   UseFlag   Start Time"
      Height          =   300
      Index           =   2
      Left            =   6870
      TabIndex        =   124
      Top             =   840
      Width           =   2760
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point   UseFlag   Start Time"
      Height          =   300
      Index           =   1
      Left            =   3570
      TabIndex        =   123
      Top             =   840
      Width           =   2760
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 24:"
      Height          =   300
      Index           =   23
      Left            =   6840
      TabIndex        =   120
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   23
      Left            =   9030
      TabIndex        =   119
      Top             =   4500
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 23:"
      Height          =   300
      Index           =   22
      Left            =   6840
      TabIndex        =   115
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   22
      Left            =   9030
      TabIndex        =   114
      Top             =   4050
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 22:"
      Height          =   300
      Index           =   21
      Left            =   6840
      TabIndex        =   110
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   21
      Left            =   9030
      TabIndex        =   109
      Top             =   3600
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 21:"
      Height          =   300
      Index           =   20
      Left            =   6840
      TabIndex        =   105
      Top             =   3150
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   20
      Left            =   9030
      TabIndex        =   104
      Top             =   3150
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 20:"
      Height          =   300
      Index           =   19
      Left            =   6840
      TabIndex        =   100
      Top             =   2700
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   19
      Left            =   9030
      TabIndex        =   99
      Top             =   2700
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 19:"
      Height          =   300
      Index           =   18
      Left            =   6840
      TabIndex        =   95
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   18
      Left            =   9030
      TabIndex        =   94
      Top             =   2250
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 18:"
      Height          =   300
      Index           =   17
      Left            =   6840
      TabIndex        =   90
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   17
      Left            =   9030
      TabIndex        =   89
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 17:"
      Height          =   300
      Index           =   16
      Left            =   6840
      TabIndex        =   85
      Top             =   1350
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   16
      Left            =   9030
      TabIndex        =   84
      Top             =   1350
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 16:"
      Height          =   300
      Index           =   15
      Left            =   3540
      TabIndex        =   80
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   15
      Left            =   5730
      TabIndex        =   79
      Top             =   4500
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 15:"
      Height          =   300
      Index           =   14
      Left            =   3540
      TabIndex        =   75
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   14
      Left            =   5730
      TabIndex        =   74
      Top             =   4050
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 14:"
      Height          =   300
      Index           =   13
      Left            =   3540
      TabIndex        =   70
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   13
      Left            =   5730
      TabIndex        =   69
      Top             =   3600
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 13:"
      Height          =   300
      Index           =   12
      Left            =   3540
      TabIndex        =   65
      Top             =   3150
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   12
      Left            =   5730
      TabIndex        =   64
      Top             =   3150
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 12:"
      Height          =   300
      Index           =   11
      Left            =   3540
      TabIndex        =   60
      Top             =   2700
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   11
      Left            =   5730
      TabIndex        =   59
      Top             =   2700
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 11:"
      Height          =   300
      Index           =   10
      Left            =   3540
      TabIndex        =   55
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   10
      Left            =   5730
      TabIndex        =   54
      Top             =   2250
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 10:"
      Height          =   300
      Index           =   9
      Left            =   3540
      TabIndex        =   50
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   9
      Left            =   5730
      TabIndex        =   49
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 9:"
      Height          =   300
      Index           =   8
      Left            =   3540
      TabIndex        =   45
      Top             =   1350
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   8
      Left            =   5730
      TabIndex        =   44
      Top             =   1350
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 8:"
      Height          =   300
      Index           =   7
      Left            =   240
      TabIndex        =   40
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   7
      Left            =   2430
      TabIndex        =   39
      Top             =   4500
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 7:"
      Height          =   300
      Index           =   6
      Left            =   240
      TabIndex        =   35
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   6
      Left            =   2430
      TabIndex        =   34
      Top             =   4050
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 6:"
      Height          =   300
      Index           =   5
      Left            =   240
      TabIndex        =   30
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   5
      Left            =   2430
      TabIndex        =   29
      Top             =   3600
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 5:"
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   3150
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   4
      Left            =   2430
      TabIndex        =   24
      Top             =   3150
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 4:"
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2700
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   3
      Left            =   2430
      TabIndex        =   19
      Top             =   2700
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 3:"
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   2
      Left            =   2430
      TabIndex        =   14
      Top             =   2250
      Width           =   90
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 2:"
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   1
      Left            =   2430
      TabIndex        =   9
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label lblSep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   300
      Index           =   0
      Left            =   2430
      TabIndex        =   5
      Top             =   1350
      Width           =   90
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point   UseFlag   Start Time"
      Height          =   300
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   840
      Width           =   2760
   End
   Begin VB.Label lblPoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Point 1:"
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1350
      Width           =   990
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
      Left            =   315
      TabIndex        =   2
      Top             =   255
      Width           =   9480
   End
End
Attribute VB_Name = "frmBellInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DataLen = (MAX_BELLCOUNT_DAY * 3)
Private mlngBellInfo(DataLen / 4 - 1) As Long
Private mBellCount As Long
Private mBellInfo As BellInfo
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub cmdRead_Click()
    Dim vnResultCode As Long

    cmdRead.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdRead.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_GetBellTime(fnCommHandleIndex, _
                                mBellCount, mlngBellInfo(0))
    lblMessage.Caption = ReturnResultPrint(vnResultCode)
    If vnResultCode = RUN_SUCCESS Then
        CopyMemory mBellInfo, mlngBellInfo(0), DataLen
        ShowValue
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdRead.Enabled = True
End Sub

Private Sub cmdWrite_Click()
    Dim vnResultCode As Long

    cmdWrite.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdWrite.Enabled = True
        Exit Sub
    End If

    GetValue
    CopyMemory mlngBellInfo(0), mBellInfo, DataLen
    vnResultCode = FK_SetBellTime(fnCommHandleIndex, _
                                mBellCount, mlngBellInfo(0))
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdWrite.Enabled = True
End Sub

Private Sub ShowValue()
On Error Resume Next
    Dim vnii As Long

    For vnii = 0 To MAX_BELLCOUNT_DAY - 1
        txtHour(vnii).Text = mBellInfo.mHour(vnii)
        txtMinute(vnii).Text = mBellInfo.mMinute(vnii)
        If mBellInfo.mValid(vnii) > 1 Then mBellInfo.mValid(vnii) = 0
        chkValid(vnii).Value = mBellInfo.mValid(vnii)
    Next vnii
    txtBellCount.Text = CStr(mBellCount)
End Sub

Private Sub GetValue()
On Error Resume Next
    Dim vnii As Long

    For vnii = 0 To MAX_BELLCOUNT_DAY - 1
        mBellInfo.mHour(vnii) = txtHour(vnii).Text
        mBellInfo.mMinute(vnii) = txtMinute(vnii).Text
        mBellInfo.mValid(vnii) = chkValid(vnii).Value
    Next vnii
    mBellCount = Val(txtBellCount.Text)
End Sub

