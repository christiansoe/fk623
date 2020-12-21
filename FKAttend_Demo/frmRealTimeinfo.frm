VERSION 5.00
Begin VB.Form frmRealTimeinfo 
   Caption         =   "Setting RealTime Information"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSetTimeZone 
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
      Height          =   300
      Left            =   420
      TabIndex        =   30
      Top             =   2400
      Width           =   2340
   End
   Begin VB.OptionButton optSetWaitTime 
      Caption         =   "SendTime Property"
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
      Left            =   420
      TabIndex        =   29
      Top             =   840
      Width           =   2340
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   360
      TabIndex        =   22
      Top             =   825
      Width           =   4905
      Begin VB.TextBox txtAckTime 
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
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "0"
         Top             =   390
         Width           =   1470
      End
      Begin VB.TextBox txtWaitTime 
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
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "0"
         Top             =   810
         Width           =   1470
      End
      Begin VB.Label lblAckTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "AckTime :"
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
         Left            =   960
         TabIndex        =   28
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label lblAckTimeView 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(s)"
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
         Left            =   3735
         TabIndex        =   27
         Top             =   435
         Width           =   255
      End
      Begin VB.Label lblWaitTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "WaitTime :"
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
         Left            =   960
         TabIndex        =   26
         Top             =   855
         Width           =   1140
      End
      Begin VB.Label lblWaitTimeView 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(m)"
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
         Left            =   3720
         TabIndex        =   25
         Top             =   840
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   360
      TabIndex        =   3
      Top             =   2385
      Width           =   4905
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
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "0"
         Top             =   705
         Width           =   1470
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
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "0"
         Top             =   1920
         Width           =   1470
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
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "0"
         Top             =   1515
         Width           =   1470
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
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "0"
         Top             =   1125
         Width           =   1470
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
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   1470
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
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   1920
         Width           =   1470
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
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         Top             =   1530
         Width           =   1470
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
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         Top             =   1125
         Width           =   1470
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time 1 :"
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
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Top             =   750
         Width           =   1140
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
         Height          =   285
         Index           =   0
         Left            =   2805
         TabIndex        =   20
         Top             =   765
         Width           =   90
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time 4 :"
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
         Index           =   3
         Left            =   30
         TabIndex        =   19
         Top             =   1965
         Width           =   1140
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time 3 :"
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
         Index           =   2
         Left            =   30
         TabIndex        =   18
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Time 2 :"
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
         Index           =   1
         Left            =   30
         TabIndex        =   17
         Top             =   1170
         Width           =   1140
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
         Height          =   285
         Index           =   3
         Left            =   2805
         TabIndex        =   16
         Top             =   1965
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
         Height          =   285
         Index           =   2
         Left            =   2805
         TabIndex        =   15
         Top             =   1575
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
         Height          =   285
         Index           =   1
         Left            =   2805
         TabIndex        =   14
         Top             =   1170
         Width           =   90
      End
      Begin VB.Label lblStartHour 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Hour"
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
         Left            =   1365
         TabIndex        =   13
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lobStartMinute 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Minute"
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
         Left            =   3150
         TabIndex        =   12
         Top             =   330
         Width           =   1410
      End
   End
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
      Left            =   4155
      TabIndex        =   2
      Top             =   5100
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
      Left            =   2760
      TabIndex        =   1
      Top             =   5100
      Width           =   1245
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
      Left            =   375
      TabIndex        =   0
      Top             =   300
      Width           =   4845
   End
End
Attribute VB_Name = "frmRealTimeinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gOptionFlag As Long
Const gDataLen = 16
Private glngRealTimInfo(gDataLen / 4 - 1) As Long
Private gRealTimInfo As REALTIMEINFO
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    gOptionFlag = 1
    OwnerEnableItems gOptionFlag
    fnCommHandleIndex = frmMain.gnCommHandleIndex
End Sub

Private Sub cmdRead_Click()
    Dim vnResultCode As Long
    Dim vnii As Integer
    
    cmdRead.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdRead.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_GetRealTimeInfo(fnCommHandleIndex, _
                                glngRealTimInfo(0))
    lblMessage.Caption = ReturnResultPrint(vnResultCode)
    If vnResultCode = RUN_SUCCESS Then
        CopyMemory gRealTimInfo, glngRealTimInfo(0), gDataLen
    End If
    
    For vnii = 0 To MAX_REAL_TIME - 1
        txtStartHour(vnii).Text = gRealTimInfo.Hour(vnii)
        txtStartMinute(vnii).Text = gRealTimInfo.Minute(vnii)
    Next vnii
    
    txtAckTime.Text = gRealTimInfo.AckTime
    txtWaitTime.Text = gRealTimInfo.WaitTime
    gOptionFlag = gRealTimInfo.Valid
    If gOptionFlag = 0 Then gOptionFlag = 1
    
    OwnerEnableItems (gOptionFlag)
    
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
    
    gRealTimInfo.Valid = gOptionFlag
    
    For vnii = 0 To MAX_REAL_TIME - 1
        gRealTimInfo.Hour(vnii) = Val(Trim(txtStartHour(vnii).Text))
        gRealTimInfo.Minute(vnii) = Val(Trim(txtStartMinute(vnii).Text))
    Next vnii
    
    gRealTimInfo.AckTime = Val(Trim(txtAckTime.Text))
    gRealTimInfo.WaitTime = Val(Trim(txtWaitTime.Text))
    
    CopyMemory glngRealTimInfo(0), gRealTimInfo, gDataLen
    vnResultCode = FK_SetRealTimeInfo(fnCommHandleIndex, _
                                glngRealTimInfo(0))
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdWrite.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub OwnerEnableItems(anEnableFlag As Long)
    
    Dim vnii As Integer
    Frame2.Enabled = False
    For vnii = 0 To MAX_REAL_TIME - 1
        txtStartHour(vnii).Enabled = False
        txtStartMinute(vnii).Enabled = False
        lblStartSep(vnii).Enabled = False
        lblTime(vnii).Enabled = False
    Next vnii
    lblStartHour.Enabled = False
    lobStartMinute.Enabled = False
    
    Frame1.Enabled = False
    lblAckTime.Enabled = False
    txtAckTime.Enabled = False
    lblAckTimeView.Enabled = False
    lblWaitTime.Enabled = False
    txtWaitTime.Enabled = False
    lblWaitTimeView.Enabled = False
    
    Select Case anEnableFlag
        Case 1
            Frame1.Enabled = True
            lblAckTime.Enabled = True
            txtAckTime.Enabled = True
            lblAckTimeView.Enabled = True
            lblWaitTime.Enabled = True
            txtWaitTime.Enabled = True
            lblWaitTimeView.Enabled = True
            optSetWaitTime.Value = True
        Case 2
        
            Frame2.Enabled = True
            For vnii = 0 To MAX_REAL_TIME - 1
                txtStartHour(vnii).Enabled = True
                txtStartMinute(vnii).Enabled = True
                lblStartSep(vnii).Enabled = True
                lblTime(vnii).Enabled = True
            Next vnii
            lblStartHour.Enabled = True
            lobStartMinute.Enabled = True
            optSetTimeZone.Value = True
            
    End Select
End Sub

Private Sub optSetTimeZone_Click()
    If optSetTimeZone.Value = True Then
        OwnerEnableItems 2
        gOptionFlag = 2
    End If
End Sub

Private Sub optSetWaitTime_Click()
    If optSetWaitTime.Value = True Then
        OwnerEnableItems 1
        gOptionFlag = 1
    End If
End Sub
