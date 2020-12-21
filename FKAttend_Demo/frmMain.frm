VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Menu"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   FillColor       =   &H008080FF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNetInfo 
      Caption         =   "Set Net Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   44
      Top             =   7920
      Width           =   3420
   End
   Begin VB.CommandButton cmdRealTimeInfo 
      Caption         =   "Set RealTime Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   43
      Top             =   7320
      Width           =   3420
   End
   Begin VB.CommandButton cmdSystemInfo 
      Caption         =   "System Info Management"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   42
      Top             =   3762
      Width           =   3420
   End
   Begin VB.CommandButton cmdLogData 
      Caption         =   "Log Data Management"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   41
      Top             =   3169
      Width           =   3420
   End
   Begin VB.CommandButton cmdEnrollData 
      Caption         =   "Enroll Data Management"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   40
      Top             =   2576
      Width           =   3420
   End
   Begin VB.CommandButton cmdProuctData 
      Caption         =   "Get Product Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   39
      Top             =   4355
      Width           =   3420
   End
   Begin VB.CommandButton cmdBellInfo 
      Caption         =   "Bell Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   38
      Top             =   4948
      Width           =   3420
   End
   Begin VB.CommandButton cmdUserInfo 
      Caption         =   "Set UserInfo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   37
      Top             =   5541
      Width           =   3420
   End
   Begin VB.CommandButton cmdCloseComm 
      Caption         =   "Close Comm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6090
      TabIndex        =   36
      Top             =   1980
      Width           =   1650
   End
   Begin VB.CommandButton cmdSetPassTime 
      Caption         =   "Set PassTime"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   35
      Top             =   6134
      Width           =   3420
   End
   Begin VB.CommandButton cmdSetAdjust 
      Caption         =   "Set AdjustInfo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   34
      Top             =   6727
      Width           =   3420
   End
   Begin VB.CommandButton cmdDeviceName 
      Caption         =   "..."
      Height          =   360
      Left            =   3570
      TabIndex        =   32
      Top             =   6555
      Width           =   390
   End
   Begin VB.TextBox txtDeviceName 
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
      Left            =   2115
      TabIndex        =   31
      Text            =   "FingerKeeper"
      Top             =   6075
      Width           =   1830
   End
   Begin VB.OptionButton optUSBDevice 
      BackColor       =   &H00FFC0C0&
      Caption         =   "USB Device"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   24
      Top             =   2040
      Width           =   2220
   End
   Begin VB.TextBox txtLicense 
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
      Height          =   420
      Left            =   5760
      TabIndex        =   18
      Text            =   "0"
      Top             =   1380
      Width           =   840
   End
   Begin VB.TextBox txtMachineNumber 
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
      Height          =   420
      Left            =   3120
      TabIndex        =   14
      Text            =   "1"
      Top             =   1380
      Width           =   600
   End
   Begin VB.OptionButton optSerialDevice 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Serial Device"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   9
      Top             =   2415
      Width           =   2220
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "    "
      Height          =   2685
      Left            =   315
      TabIndex        =   10
      Top             =   2430
      Width           =   3900
      Begin VB.ComboBox cmbComPort 
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
         ItemData        =   "frmMain.frx":0442
         Left            =   1740
         List            =   "frmMain.frx":0464
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtComTimeOut 
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
         Left            =   1740
         TabIndex        =   45
         Text            =   "3000"
         Top             =   1275
         Width           =   855
      End
      Begin VB.CheckBox chkUsingModem 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Using Modem"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Top             =   1755
         Width           =   1830
      End
      Begin VB.TextBox txtWaitDialTime 
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
         Left            =   2895
         TabIndex        =   26
         Text            =   "20"
         Top             =   2205
         Width           =   615
      End
      Begin VB.TextBox txtTelNumber 
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
         Left            =   2880
         TabIndex        =   25
         Text            =   "801"
         Top             =   1755
         Width           =   915
      End
      Begin VB.ComboBox cmbCommBaudRate 
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
         ItemData        =   "frmMain.frx":0487
         Left            =   1740
         List            =   "frmMain.frx":049A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   825
         Width           =   1575
      End
      Begin VB.Label lblComTimeOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   47
         Top             =   1305
         Width           =   1170
      End
      Begin VB.Label lblComTimeOutT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(ms)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   46
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label lblS 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
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
         Left            =   3570
         TabIndex        =   30
         Top             =   2205
         Width           =   180
      End
      Begin VB.Label lblWaitTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Wait Time For Dialing"
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
         Left            =   780
         TabIndex        =   29
         Top             =   2205
         Width           =   2115
      End
      Begin VB.Label lblTelphon 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   28
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label lblComBaudRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baudrate : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   17
         Top             =   885
         Width           =   1125
      End
      Begin VB.Label lblComPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ComPort : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   15
         Top             =   420
         Width           =   1200
      End
   End
   Begin VB.OptionButton optNetworkDevice 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Network Device"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   5295
      Width           =   2220
   End
   Begin VB.CommandButton cmdOpenComm 
      Caption         =   "Open Comm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4320
      TabIndex        =   2
      Top             =   1980
      Width           =   1650
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   2940
      Left            =   315
      TabIndex        =   4
      Top             =   5250
      Width           =   3900
      Begin VB.CheckBox chkUDPFlag 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   2595
         Width           =   975
      End
      Begin VB.TextBox txtTimeOut 
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
         Left            =   1800
         TabIndex        =   20
         Text            =   "5000"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   1800
         TabIndex        =   13
         Text            =   "0"
         Top             =   1710
         Width           =   855
      End
      Begin VB.TextBox txtPortNo 
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
         Left            =   1800
         TabIndex        =   6
         Text            =   "5005"
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtIPAddress 
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
         Left            =   1800
         TabIndex        =   5
         Text            =   "192.168.0.9"
         Top             =   400
         Width           =   1830
      End
      Begin VB.Label lblDeviceName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device Name :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   33
         Top             =   855
         Width           =   1500
      End
      Begin VB.Label lblTimeOutT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(ms)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Top             =   2205
         Width           =   555
      End
      Begin VB.Label lblTimeOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   2220
         Width           =   1650
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   1770
         Width           =   1650
      End
      Begin VB.Label lblPortNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label lblIPAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   460
         Width           =   1650
      End
   End
   Begin VB.Label lblComLicense 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "License :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   19
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label lblMachineNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Number :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   1440
      Width           =   1710
   End
   Begin VB.Label lblVer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FKAttend.dll (V2.8.8.609)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   8040
   End
   Begin VB.Label lbSubject 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "FKAttend.DLL Sample (v2.1.0.2)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   8040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gbOpenFlag As Boolean
Public gnCommHandleIndex As Long

Private Sub cmdDeviceName_Click()
    Dim m_DeviceName As String
    
    m_DeviceName = txtDeviceName.Text
    Trim (m_DeviceName)
    txtIPAddress.Text = GetIPAddressFromDNSName(m_DeviceName)
End Sub

Private Sub cmdRealTimeInfo_Click()
    Me.Visible = False
    frmRealTimeinfo.Visible = True
End Sub

Private Sub Form_Load()
    optSerialDevice.Value = False
    optNetworkDevice.Value = True
    optUSBDevice.Value = False
    OwnerEnableButtons False

    gbOpenFlag = False
    txtMachineNumber.Text = "1"
    cmbComPort.ListIndex = 0
    cmbCommBaudRate.ListIndex = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdCloseComm_Click
End Sub

Private Sub chkUsingModem_Click()
    optSerialDevice_Click
End Sub

Private Sub optSerialDevice_Click()
    If optSerialDevice.Value = True Then
        OwnerEnableItems 1
    End If
End Sub

Private Sub optNetworkDevice_Click()
    If optNetworkDevice.Value = True Then
        OwnerEnableItems 2
    End If
End Sub

Private Sub optUSBDevice_Click()
    If optUSBDevice.Value = True Then
        OwnerEnableItems 3
    End If
End Sub

Private Sub cmdOpenComm_Click()
    Dim vnMachineNumber As Long
    Dim vnCommPort As Long
    Dim vnCommBaudrate As Long
    Dim vstrTelNumber As String
    Dim vnWaitDialTime As Long
    Dim vnLicense As Long
    Dim vpszIPAddress As String
    Dim vpszNetPort As Long
    Dim vpszNetPassword As Long
    Dim vnTimeOut As Long
    Dim vnProtocolType As Long
    Dim vnResultCode As Long

    cmdOpenComm.Enabled = False
    vnMachineNumber = Val(txtMachineNumber.Text)
    vnLicense = Val(txtLicense.Text)
    
    If optSerialDevice.Value = True Then
        If chkUsingModem.Value = vbChecked Then
            vstrTelNumber = Trim(txtTelNumber.Text)
            vnWaitDialTime = Val(Trim(txtWaitDialTime.Text))
            If vnWaitDialTime < 10 And vnWaitDialTime > 60 Then
                vnWaitDialTime = 10
                txtWaitDialTime.Text = "10"
            End If
        Else
            vstrTelNumber = ""
            vnWaitDialTime = 0
        End If

        vnCommPort = Val(Trim(cmbComPort.Text))
        vnCommBaudrate = Val(Trim(cmbCommBaudRate.Text))
        vnTimeOut = Val(Trim(txtComTimeOut.Text))
        gnCommHandleIndex = FK_ConnectComm(vnMachineNumber, vnCommPort, vnCommBaudrate, vstrTelNumber, vnWaitDialTime, vnLicense, vnTimeOut)
        
    ElseIf optNetworkDevice.Value = True Then
        vpszIPAddress = Trim(txtIPAddress.Text)
        vpszNetPort = CLng(txtPortNo.Text)
        vpszNetPassword = CLng(txtPassword.Text)
        vnTimeOut = CLng(txtTimeOut.Text)
        If chkUDPFlag.Value = vbUnchecked Then
            vnProtocolType = PROTOCOL_TCPIP
        Else
            vnProtocolType = PROTOCOL_UDP
        End If
        gnCommHandleIndex = FK_ConnectNet(vnMachineNumber, vpszIPAddress, vpszNetPort, vnTimeOut, vnProtocolType, vpszNetPassword, vnLicense)
        
    ElseIf optUSBDevice.Value = True Then
        gnCommHandleIndex = FK_ConnectUSB(vnMachineNumber, vnLicense)
    End If
    
    If gnCommHandleIndex > 0 Then
        gbOpenFlag = True
        OwnerEnableButtons True
    Else
        vnResultCode = gnCommHandleIndex
        MsgBox ReturnResultPrint(vnResultCode), vbOKOnly, "error"
        cmdOpenComm.Enabled = True
    End If
End Sub

Private Sub cmdCloseComm_Click()
    If gbOpenFlag = True Then
        FK_DisConnect gnCommHandleIndex
        gbOpenFlag = False
        OwnerEnableButtons False
    End If
End Sub

Private Sub cmdEnrollData_Click()
    Me.Visible = False
    frmEnroll.Visible = True
End Sub

Private Sub cmdLogData_Click()
    Me.Visible = False
    frmLog.Visible = True
End Sub

Private Sub cmdSystemInfo_Click()
    Me.Visible = False
    frmSystemInfo.Visible = True
End Sub

Private Sub cmdProuctData_Click()
    Me.Visible = False
    frmProductData.Visible = True
End Sub

Private Sub cmdBellInfo_Click()
    Me.Visible = False
    frmBellInfo.Visible = True
End Sub

Private Sub cmdUserInfo_Click()
    Me.Visible = False
    frmUserInfo.Visible = True
End Sub

Private Sub cmdSetPassTime_Click()
    Me.Visible = False
    frmPassTimeInfo.Visible = True
End Sub

Private Sub cmdSetAdjust_Click()
    Me.Visible = False
    frmAdjust.Visible = True
End Sub

Private Sub cmdNetInfo_Click()
    Me.Visible = False
    frmNetInfo.Visible = True
End Sub

Private Sub OwnerEnableButtons(abEnableFlag As Boolean)
    cmdOpenComm.Enabled = Not abEnableFlag
    cmdCloseComm.Enabled = abEnableFlag
    cmdSystemInfo.Enabled = abEnableFlag
    cmdProuctData.Enabled = abEnableFlag
    cmdBellInfo.Enabled = abEnableFlag
    cmdUserInfo.Enabled = abEnableFlag
    cmdSetPassTime.Enabled = abEnableFlag
    cmdSetAdjust.Enabled = abEnableFlag
    cmdRealTimeInfo.Enabled = abEnableFlag
    cmdNetInfo.Enabled = abEnableFlag
    
    optSerialDevice.Enabled = Not abEnableFlag
    optNetworkDevice.Enabled = Not abEnableFlag
    optUSBDevice.Enabled = Not abEnableFlag
End Sub

Private Sub OwnerEnableItems(anEnableFlag As Long)
    lblComPort.Enabled = False
    cmbComPort.Enabled = False
    lblComBaudRate.Enabled = False
    cmbCommBaudRate.Enabled = False
    chkUsingModem.Enabled = False
    lblTelphon.Enabled = False
    txtTelNumber.Enabled = False
    lblWaitTime.Enabled = False
    txtWaitDialTime.Enabled = False
    lblS.Enabled = False
    txtComTimeOut.Enabled = False
    lblComTimeOut.Enabled = False
    lblComTimeOutT.Enabled = False

    lblIPAddress.Enabled = False
    txtIPAddress.Enabled = False
    
    lblDeviceName.Enabled = False
    txtDeviceName.Enabled = False
    
    lblPortNo.Enabled = False
    txtPortNo.Enabled = False
    lblPassword.Enabled = False
    txtPassword.Enabled = False
    lblTimeOut.Enabled = False
    txtTimeOut.Enabled = False
    lblTimeOutT.Enabled = False
    chkUDPFlag.Enabled = False
    
    Select Case anEnableFlag
        Case 1
            lblComPort.Enabled = True
            cmbComPort.Enabled = True
            lblComBaudRate.Enabled = True
            cmbCommBaudRate.Enabled = True
            txtComTimeOut.Enabled = True
            lblComTimeOut.Enabled = True
            lblComTimeOutT.Enabled = True
            chkUsingModem.Enabled = True
            If chkUsingModem.Value = vbChecked Then
                lblTelphon.Enabled = True
                txtTelNumber.Enabled = True
                lblWaitTime.Enabled = True
                txtWaitDialTime.Enabled = True
                lblS.Enabled = True
            End If
            
        Case 2
            lblIPAddress.Enabled = True
            txtIPAddress.Enabled = True
            lblDeviceName.Enabled = True
            txtDeviceName.Enabled = True
            lblPortNo.Enabled = True
            txtPortNo.Enabled = True
            lblPassword.Enabled = True
            txtPassword.Enabled = True
            lblTimeOut.Enabled = True
            txtTimeOut.Enabled = True
            lblTimeOutT.Enabled = True
            chkUDPFlag.Enabled = True
    End Select
End Sub

