VERSION 5.00
Begin VB.Form frmSystemInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage System Info"
   ClientHeight    =   4095
   ClientLeft      =   4995
   ClientTop       =   3105
   ClientWidth     =   7335
   Icon            =   "frmSytemInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSatus 
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
      Height          =   360
      ItemData        =   "frmSytemInfo.frx":0442
      Left            =   2880
      List            =   "frmSytemInfo.frx":0473
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2796
      Width           =   1320
   End
   Begin VB.CommandButton cmdSetDeviceTime 
      Caption         =   "SetDeviceTime"
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
      Left            =   585
      TabIndex        =   2
      Top             =   1860
      Width           =   1875
   End
   Begin VB.CommandButton cmdGetDeviceTime 
      Caption         =   "GetDeviceTime"
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
      Left            =   585
      TabIndex        =   1
      Top             =   1236
      Width           =   1875
   End
   Begin VB.CommandButton cmdPowerOn 
      Caption         =   "PowerOnDevice"
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
      Left            =   2715
      TabIndex        =   3
      Top             =   1236
      Width           =   1875
   End
   Begin VB.CommandButton PowerOffDevice 
      Caption         =   "PowerOffDevice"
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
      Left            =   2715
      TabIndex        =   4
      Top             =   1860
      Width           =   1875
   End
   Begin VB.CheckBox chkEnableDevice 
      Caption         =   "DisableDevice"
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
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1236
      Width           =   1875
   End
   Begin VB.CommandButton cmdGetDeviceStaus 
      Caption         =   "GetDeviceStatus"
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
      Left            =   2715
      TabIndex        =   7
      Top             =   3348
      Width           =   1875
   End
   Begin VB.CommandButton cmdGetDeviceInfo 
      Caption         =   "GetDeviceInfo"
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
      Left            =   585
      TabIndex        =   6
      Top             =   3348
      Width           =   1875
   End
   Begin VB.CommandButton cmdSetDeviceInfo 
      Caption         =   "SetDeviceInfo"
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
      Left            =   4950
      TabIndex        =   8
      Top             =   3348
      Width           =   1875
   End
   Begin VB.TextBox txtSetDevInfo 
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
      Left            =   5445
      TabIndex        =   0
      Top             =   2796
      Width           =   885
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
      Left            =   555
      TabIndex        =   11
      Top             =   480
      Width           =   6300
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status Paramerter:  Info Paramerter:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   828
      TabIndex        =   10
      Top             =   2688
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Status Value:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4668
      TabIndex        =   9
      Top             =   2688
      Width           =   660
   End
End
Attribute VB_Name = "frmSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
    cmbSatus.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub chkEnableDevice_Click()
    Dim vnFlag As Long
    Dim vnResultCode As Long

    lblMessage.Caption = "Working..."
    DoEvents

    If chkEnableDevice.Value = Unchecked Then
        vnFlag = 1
    Else
        vnFlag = 0
    End If

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, vnFlag)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    If chkEnableDevice.Value = Unchecked Then
        chkEnableDevice.Caption = "DisableDevice"
    Else
        chkEnableDevice.Caption = "EnableDevice"
    End If
End Sub

Private Sub cmdPowerOn_Click()
    FK_PowerOnAllDevice fnCommHandleIndex
End Sub

Private Sub PowerOffDevice_Click()
    Dim vnResultCode As Long

    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_PowerOffDevice(fnCommHandleIndex)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)
End Sub

Private Sub cmdGetDeviceTime_Click()
    Dim vdwDate
    Dim strDataTime As String
    Dim vnResultCode As Long

    vdwDate = "2011-3-15"

    cmdGetDeviceTime.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetDeviceTime.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_GetDeviceTime(fnCommHandleIndex, vdwDate)
    If vnResultCode = RUN_SUCCESS Then
        strDataTime = "Date = " & Format(vdwDate, "Long Date") & ", Time = " & Format(vdwDate, "Long Time")
        strDataTime = vdwDate
        lblMessage.Caption = strDataTime
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetDeviceTime.Enabled = True
End Sub

Private Sub cmdSetDeviceTime_Click()
    Dim vdwDate As Date
    Dim strDataTime As String
    Dim vnResultCode As Long

    cmdSetDeviceTime.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetDeviceTime.Enabled = True
        Exit Sub
    End If

    vdwDate = Now
    vnResultCode = FK_SetDeviceTime(fnCommHandleIndex, vdwDate)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetDeviceTime.Enabled = True
End Sub

Private Sub cmdGetDeviceStaus_Click()
    Dim vnStatusIndex As Long
    Dim vnValue As Long
    Dim vnResultCode As Long

    cmdGetDeviceStaus.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetDeviceStaus.Enabled = True
        Exit Sub
    End If

    vnStatusIndex = cmbSatus.ListIndex + 1
    vnResultCode = FK_GetDeviceStatus(fnCommHandleIndex, _
                                    vnStatusIndex, vnValue)
    If vnResultCode = RUN_SUCCESS Then
        Select Case vnStatusIndex
            Case GET_MANAGERS:  lblMessage.Caption = "Manager count = " & vnValue
            Case GET_USERS:  lblMessage.Caption = "User count = " & vnValue
            Case GET_FPS:  lblMessage.Caption = "Fp count = " & vnValue
            Case GET_PSWS:  lblMessage.Caption = "Password count = " & vnValue
            Case GET_SLOGS:  lblMessage.Caption = "SLog count = " & vnValue
            Case GET_GLOGS:  lblMessage.Caption = "GLog count = " & vnValue
            Case GET_ASLOGS:  lblMessage.Caption = "All SLog count = " & vnValue
            Case GET_AGLOGS:  lblMessage.Caption = "All GLog count = " & vnValue
            Case GET_CARDS:  lblMessage.Caption = "Card count = " & vnValue
            Case Else: lblMessage.Caption = "--"
        End Select
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetDeviceStaus.Enabled = True
End Sub

Private Sub cmdGetDeviceInfo_Click()
    Dim vnInfoIndex As Long
    Dim vnValue As Long
    Dim vnResultCode As Long

    cmdGetDeviceInfo.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetDeviceInfo.Enabled = True
        Exit Sub
    End If

    vnInfoIndex = cmbSatus.ListIndex + 1
    If vnInfoIndex = 11 Then
        vnInfoIndex = DI_VERIFY_KIND
    ElseIf vnInfoIndex = 12 Then
        vnInfoIndex = DI_MULTIUSERS
    ElseIf vnInfoIndex = 13 Then
        vnInfoIndex = DI_NETENABLE
    ElseIf vnInfoIndex = 14 Then
        vnInfoIndex = DI_ALARMDELAY
    ElseIf vnInfoIndex = 15 Then
        vnInfoIndex = DI_SENSORDELAY
    End If
 
    vnResultCode = FK_GetDeviceInfo(fnCommHandleIndex, _
                                    vnInfoIndex, vnValue)
    If vnResultCode = RUN_SUCCESS Then
        Select Case vnInfoIndex
            Case DI_MANAGERS:  lblMessage.Caption = "ManagerCount = " & vnValue
            Case DI_MACHINENUM:  lblMessage.Caption = "Machine Num = " & vnValue
            Case DI_LANGAUGE:  lblMessage.Caption = "Language = " & vnValue
            Case DI_POWEROFF_TIME:  lblMessage.Caption = "PowerOffTime = " & vnValue
            Case DI_LOCK_CTRL:  lblMessage.Caption = "LockOperate = " & vnValue
            Case DI_GLOG_WARNING:  lblMessage.Caption = "GLogWarning = " & vnValue
            Case DI_SLOG_WARNING:  lblMessage.Caption = "SLogWarning = " & vnValue
            Case DI_VERIFY_INTERVALS:  lblMessage.Caption = "ReVerifyTime = " & vnValue
            Case DI_RSCOM_BPS:  lblMessage.Caption = "Baudrate(" & vnValue & ") : "
                If vnValue = BPS_9600 Then
                    lblMessage.Caption = lblMessage.Caption & "9600"
                ElseIf vnValue = BPS_19200 Then
                    lblMessage.Caption = lblMessage.Caption & "19200"
                ElseIf vnValue = BPS_38400 Then
                    lblMessage.Caption = lblMessage.Caption & "38400"
                ElseIf vnValue = BPS_57600 Then
                    lblMessage.Caption = lblMessage.Caption & "57600"
                ElseIf vnValue = BPS_115200 Then
                    lblMessage.Caption = lblMessage.Caption & "115200"
                Else
                    lblMessage.Caption = lblMessage.Caption & "--"
                End If
            Case DI_VERIFY_KIND: lblMessage.Caption = "VerifyKind = "
                If vnValue = 0 Then
                    lblMessage.Caption = lblMessage.Caption & "F / P / C"
                ElseIf vnValue = 1 Then
                    lblMessage.Caption = lblMessage.Caption & "F + P"
                ElseIf vnValue = 2 Then
                    lblMessage.Caption = lblMessage.Caption & "F + C"
                ElseIf vnValue = 3 Then
                    lblMessage.Caption = lblMessage.Caption & "C"
                End If
            Case DI_DATE_SEPARATE: lblMessage.Caption = "DateSeperate = " & vnValue
            Case DI_MULTIUSERS: lblMessage.Caption = "MultiUsers = " & vnValue
            Case DI_NETENABLE:
                If vnValue = 1 Then
                    lblMessage.Caption = "Network Enabled."
                Else
                    lblMessage.Caption = "Network Disabled."
                End If
            Case DI_ALARMDELAY: lblMessage.Caption = "Alarm Delay = " & vnValue
            Case DI_SENSORDELAY: lblMessage.Caption = "Sensor Delay = " & vnValue
            Case Else: lblMessage.Caption = "--"
        End Select
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetDeviceInfo.Enabled = True
End Sub

Private Sub cmdSetDeviceInfo_Click()
    Dim vnInfoIndex As Long
    Dim vnValue As Long
    Dim vnResultCode As Long

    cmdSetDeviceInfo.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetDeviceInfo.Enabled = True
        Exit Sub
    End If

    vnInfoIndex = cmbSatus.ListIndex + 1
    vnValue = Val(txtSetDevInfo.Text)
    If vnInfoIndex = 11 Then
        vnInfoIndex = DI_VERIFY_KIND
    ElseIf vnInfoIndex = 12 Then
        vnInfoIndex = DI_MULTIUSERS
    ElseIf vnInfoIndex = 13 Then
        vnInfoIndex = DI_NETENABLE
    ElseIf vnInfoIndex = 14 Then
        vnInfoIndex = DI_ALARMDELAY
    ElseIf vnInfoIndex = 15 Then
        vnInfoIndex = DI_SENSORDELAY
    End If
     
    vnResultCode = FK_SetDeviceInfo(fnCommHandleIndex, _
                                    vnInfoIndex, vnValue)
    If vnInfoIndex = DI_MACHINENUM And vnResultCode = RUN_SUCCESS Then
        frmMain.txtMachineNumber.Text = vnValue
    End If
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetDeviceInfo.Enabled = True
End Sub
