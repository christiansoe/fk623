VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Log Data"
   ClientHeight    =   5685
   ClientLeft      =   4815
   ClientTop       =   3135
   ClientWidth     =   9585
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSave 
      Caption         =   "Save to file"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   840
      Width           =   1365
   End
   Begin VB.CommandButton cmdUsbGLogData 
      Caption         =   "Read USB GLogData"
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
      Left            =   7920
      TabIndex        =   6
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdUsbSLogData 
      Caption         =   "Read USB SLogData"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdEmptyGLogData 
      Caption         =   "Empty GLogData"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdEmptySLogData 
      Caption         =   "Empty SLogData"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdSLogData 
      Caption         =   "Read SLogData"
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
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdGLogData 
      Caption         =   "Read GLogData"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CheckBox chkReadMark 
      Caption         =   "ReadMark"
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
      Left            =   8040
      TabIndex        =   7
      Top             =   840
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid gridLogView 
      Height          =   3750
      Left            =   120
      TabIndex        =   8
      Top             =   1125
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   6
      Redraw          =   -1  'True
      GridLines       =   2
      AllowUserResizing=   1
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   8760
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEnrollData 
      AutoSize        =   -1  'True
      Caption         =   "Log Data :"
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
      Left            =   420
      TabIndex        =   10
      Top             =   795
      Width           =   960
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
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
      Left            =   1935
      TabIndex        =   9
      Top             =   795
      Width           =   570
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9330
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEF_MAX_LOGCOUNT = 200000     ' max log data count to support by device.
Const DEF_MAX_DISPCOUNT = 30000     ' max count to show on a grid.
Const DEF_MUL_TWIPS = 15

Private fnGridHeight As Long
Private fngrdIndex As Long
Private fgrdLogView() As MSFlexGrid
Private fnCommHandleIndex As Long


Private Sub Form_Load()
    Dim vnii As Long
    Dim vngrdNumber As Long

    fnCommHandleIndex = frmMain.gnCommHandleIndex
    chkReadMark.Value = vbChecked
    chkSave.Value = vbUnchecked
    fnGridHeight = gridLogView.Height

    vngrdNumber = DEF_MAX_LOGCOUNT / DEF_MAX_DISPCOUNT
    If vngrdNumber * DEF_MAX_DISPCOUNT < DEF_MAX_LOGCOUNT Then vngrdNumber = vngrdNumber + 1
    If vngrdNumber < 1 Then vngrdNumber = 1

    ReDim fgrdLogView(vngrdNumber)
    Set fgrdLogView(1) = gridLogView

    If vngrdNumber > 1 Then
        For vnii = 2 To vngrdNumber
            If Not fgrdLogView(vnii) Is Nothing Then
                Controls.Remove fgrdLogView(vnii)
                Set fgrdLogView(vnii) = Nothing
            End If

            Set fgrdLogView(vnii) = Controls.Add("MSFlexGridLib.MSFlexGrid", "FlexGrid" & vnii)
            With fgrdLogView(vnii)
                .Left = gridLogView.Left
                .Top = gridLogView.Top
                .Width = gridLogView.Width
                .Height = 0
                .GridLines = gridLogView.GridLines
                .Visible = False
            End With
        Next
    End If
    OwnerEnableButtons True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub funcGetSuperLogData(Optional abUSBFlag As Boolean = False)
    Dim vSEnrollNumber As Long
    Dim vGEnrollNumber As Long
    Dim vManipulation As Long
    Dim vBackupNumber As Long
    Dim vdwDate As Date
    Dim vnii As Long
    Dim vtObject
    Dim vstrLogItem As Variant
    Dim vstrFileName As String
    Dim vnReadMark As Long
    Dim vnResultCode As Long
    Dim vnYear
    Dim vnMonth
    Dim vnDay
    Dim vnHour
    Dim vnMinute
    Dim vnSecond
    
    lblMessage.Caption = "Waiting..."
    lblTotal.Caption = "Total : 0"
    DoEvents

    vstrLogItem = Array("", "SEnrollNo", "GEnrollNo", "Manipulation", "BackupNo", "DateTime")
    With fgrdLogView(1)
        .Redraw = False
        .Height = fnGridHeight
        .Cols = 6
        .Rows = 1
        .Clear
        .ColWidth(0) = 48 * DEF_MUL_TWIPS
        .Row = 0
        For vnii = 1 To .Cols - 1
            .Col = vnii
            .Text = vstrLogItem(vnii)
            .ColWidth(vnii) = 80 * DEF_MUL_TWIPS
            .ColAlignment(vnii) = 3
        Next vnii
        .ColWidth(3) = 140 * DEF_MUL_TWIPS
        .ColAlignment(3) = 2
        .ColWidth(5) = 140 * DEF_MUL_TWIPS
        .Redraw = True
    End With

    For Each vtObject In fgrdLogView
        If Not vtObject Is Nothing Then
            If vtObject.Name <> "gridLogView" Then
                vtObject.Height = 0
                vtObject.Visible = False
            End If
        End If
    Next

    If abUSBFlag = True Then
        dlgOpen.InitDir = CurDir
        dlgOpen.CancelError = False
        dlgOpen.Flags = cdlOFNHideReadOnly
        dlgOpen.Filter = "SLog Files (*.txt)|*.txt|All Files (*.*)|*.*"
        dlgOpen.FilterIndex = 1
        dlgOpen.InitDir = CurDir
        dlgOpen.ShowOpen
        vstrFileName = dlgOpen.FileName
        If vstrFileName = "" Then Exit Sub
    Else
        vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
        If vnResultCode <> RUN_SUCCESS Then
            lblMessage.Caption = gstrNoDevice
            Exit Sub
        End If
    End If

    MousePointer = vbHourglass
    DoEvents
    If abUSBFlag = True Then
        vnResultCode = FK_USBLoadSuperLogDataFromFile(fnCommHandleIndex, vstrFileName)
    Else
        If chkReadMark.Value = vbChecked Then
            vnReadMark = 1
        Else
            vnReadMark = 0
        End If
        vnResultCode = FK_LoadSuperLogData(fnCommHandleIndex, vnReadMark)
    End If

    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    Else
        lblMessage.Caption = "Getting..."
        DoEvents
        With fgrdLogView(1)
            .Redraw = False
            vnii = 1
            Do
                  vnResultCode = FK_GetSuperLogData_1(fnCommHandleIndex, _
                                                    vSEnrollNumber, _
                                                    vGEnrollNumber, _
                                                    vManipulation, _
                                                    vBackupNumber, _
                                                    vnYear, vnMonth, vnDay, vnHour, vnMinute, vnSecond)
                If vnResultCode <> RUN_SUCCESS Then
                    If vnResultCode = RUNERR_DATAARRAY_END Then
                        vnResultCode = RUN_SUCCESS
                    End If
                    Exit Do
                End If
                .AddItem (1)

                .Row = vnii
                .Col = 0
                .Text = vnii

                .Col = 1
                .Text = Trim(Str(vSEnrollNumber))

                .Col = 2
                .Text = Trim(Str(vGEnrollNumber))

                .Col = 3
                Select Case vManipulation
                    Case LOG_ENROLL_USER
                        .Text = "Enroll User"
                    Case LOG_ENROLL_MANAGER
                        .Text = "Enroll Manager"
                    Case LOG_ENROLL_DELFP
                        .Text = "Delete Fp Data"
                    Case LOG_ENROLL_DELPASS
                        .Text = "Delete Password"
                    Case LOG_ENROLL_DELCARD
                        .Text = "Delete Card Data"
                    Case LOG_LOG_ALLDEL
                        .Text = "Delete All LogData"
                    Case LOG_SETUP_SYS
                        .Text = "Modify System Info"
                    Case LOG_SETUP_TIME
                        .Text = "Modify System Time"
                    Case LOG_SETUP_LOG
                        .Text = "Modify Log Setting"
                    Case LOG_SETUP_COMM
                        .Text = "Modify Comm Setting"
                    Case LOG_PASSTIME
                        .Text = "Pass Time Set"
                    Case LOG_SETUP_DOOR
                        .Text = "Door Set Log"
                    Case Else
                        .Text = "--"
                End Select

                .Col = 4
                If vBackupNumber = BACKUP_PSW Then
                    .Text = "Password"
                ElseIf vBackupNumber = BACKUP_CARD Then
                    .Text = "Card"
                ElseIf vBackupNumber < BACKUP_PSW Then
                    .Text = "Fp-" & Trim(Str((vBackupNumber)))
                Else
                    .Text = "--"
                End If

                .Col = 5
                .Text = CStr(vnYear) & "/" & Format(vnMonth, "0#") & "/" & Format(vnDay, "0#") & _
                        " " & Format(vnHour, "0#") & ":" & Format(vnMinute, "0#") & ":" & Format(vnSecond, "0#")
             
                lblTotal.Caption = "Total : " & vnii
                DoEvents
                vnii = vnii + 1
            Loop
            .Redraw = True
        End With

        If vnResultCode = RUN_SUCCESS Then
            If abUSBFlag = True Then
                lblMessage.Caption = "USBReadSuperLogDataFromFile OK"
            Else
                lblMessage.Caption = "ReadAllSuperLogData OK"
            End If
        Else
            lblMessage.Caption = ReturnResultPrint(vnResultCode)
        End If
    End If

    MousePointer = vbDefault
    If abUSBFlag = False Then
        FK_EnableDevice fnCommHandleIndex, 1
    End If
End Sub

Private Sub cmdSLogData_Click()
    OwnerEnableButtons False
    funcGetSuperLogData
    OwnerEnableButtons True
End Sub

Private Sub cmdUsbSLogData_Click()
    OwnerEnableButtons False
    funcGetSuperLogData True
    OwnerEnableButtons True
End Sub

Private Sub cmdEmptySLogData_Click()
    Dim vnResultCode As Long

    cmdEmptySLogData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdEmptySLogData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_EmptySuperLogData(fnCommHandleIndex)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdEmptySLogData.Enabled = True
End Sub

Private Sub funcGeneralLogDataGridFormat()
    Dim vtObject
    Dim vnii As Long
    Dim vstrLogItem As Variant

    vstrLogItem = Array("", "EnrollNo", "VerifyMode", "InOutMode", "DateTime")
    With fgrdLogView(1)
        .Redraw = False
        .Height = fnGridHeight
        .Cols = 5
        .Rows = 1
        .Clear
        .ColWidth(0) = 48 * DEF_MUL_TWIPS
        .Row = 0
        For vnii = 1 To .Cols - 1
            .Col = vnii
            .Text = vstrLogItem(vnii)
            .ColWidth(vnii) = 80 * DEF_MUL_TWIPS
            .ColAlignment(vnii) = 3
        Next vnii
        .ColWidth(2) = 120 * DEF_MUL_TWIPS
        .ColWidth(4) = 140 * DEF_MUL_TWIPS
        .Redraw = True
        DoEvents
        .Redraw = False
    End With

    For Each vtObject In fgrdLogView
        If Not vtObject Is Nothing Then
            If vtObject.Name <> "gridLogView" Then
                With vtObject
                    .Redraw = False
                    .Left = fgrdLogView(1).Left
                    .Top = fgrdLogView(1).Top
                    .Width = fgrdLogView(1).Width
                    .Height = 0
                    .Cols = fgrdLogView(1).Cols
                    .Rows = 0
                    .Clear
                    For vnii = 0 To .Cols - 1
                        .ColWidth(vnii) = fgrdLogView(1).ColWidth(vnii)
                        .ColAlignment(vnii) = fgrdLogView(1).ColAlignment(vnii)
                    Next vnii
                    .Redraw = False
                    .Visible = False
                End With
            End If
        End If
    Next
    DoEvents
End Sub

Private Function funcShowGeneralLogDataToGrid(anCount As Long, aSEnrollNumber As Long, _
                                        aVerifyMode As Long, aInOutMode As Long, _
                                        adwDate As Date, abDrawFlag As Boolean) As Boolean

    Dim vnkk As Long
    Dim vnHeight As Long, vnTop As Long, vnPos As Long
    Dim vStr As String
    funcShowGeneralLogDataToGrid = True
    If anCount <= 1 Then
        fngrdIndex = 1
        fgrdLogView(1).Redraw = abDrawFlag
    Else
        If fngrdIndex * DEF_MAX_DISPCOUNT < anCount Then
            If abDrawFlag = False Then
                fgrdLogView(fngrdIndex).Redraw = True
            End If
            fngrdIndex = fngrdIndex + 1
            If fngrdIndex > UBound(fgrdLogView) Then
                funcShowGeneralLogDataToGrid = False
                Exit Function
            End If
            vnHeight = fnGridHeight
            vnTop = fgrdLogView(1).Top
            For vnkk = 1 To fngrdIndex
                fgrdLogView(vnkk).Top = vnTop + (vnkk - 1) * (vnHeight / fngrdIndex)
                fgrdLogView(vnkk).Height = vnHeight / fngrdIndex
            Next vnkk
            fgrdLogView(fngrdIndex).Redraw = abDrawFlag
            fgrdLogView(fngrdIndex).Visible = True
        End If
    End If
    vnPos = anCount - (fngrdIndex - 1) * DEF_MAX_DISPCOUNT
    If fngrdIndex > 1 Then vnPos = vnPos - 1

    With fgrdLogView(fngrdIndex)
        .AddItem (1)
        .Row = vnPos
        .Col = 0
        .Text = anCount
        .Col = 1
        .Text = aSEnrollNumber
        .Col = 2
        Select Case aVerifyMode Mod LOG_OPEN_DOOR
            Case LOG_FPVERIFY           '1
                vStr = "Fp"
            Case LOG_PASSVERIFY         '2
                vStr = "Password"
            Case LOG_CARDVERIFY         '3
                vStr = "Card"
            Case LOG_FPPASS_VERIFY      '4
                vStr = "Fp+Password"
            Case LOG_FPCARD_VERIFY      '5
                vStr = "Fp+Card"
            Case LOG_PASSFP_VERIFY      '6
                vStr = "Password+Fp"
            Case LOG_CARDFP_VERIFY      '7
                vStr = "Card+Fp"
            Case LOG_JOB_NO_VERIFY      '8
                vStr = "Job No"
            Case LOG_CARDPASS_VERIFY    '9
                vStr = "Card+Pass"
            Case LOG_CLOSE_DOOR         '10
                vStr = "Close Door"
            Case LOG_OPEN_HAND          '11
                vStr = "Hand Open"
            Case LOG_PROG_OPEN          '12
                vStr = "Prog Open"
            Case LOG_PROG_CLOSE         '13
                vStr = "PC Close"
            Case LOG_OPEN_IREGAL        '14
                vStr = "Iregal Open"
            Case LOG_CLOSE_IREGAL       '15
                vStr = "Iregal Close"
            Case LOG_OPEN_COVER         '16
                vStr = "Cover Open"
            Case LOG_CLOSE_COVER        '17
                vStr = "Cover Close"
            Case Else
                vStr = "--"
        End Select
        If aVerifyMode \ LOG_OPEN_THREAT = 1 Then
            vStr = vStr + " & Open Door as Threat"
        ElseIf aVerifyMode \ LOG_OPEN_DOOR = 1 Then
            vStr = vStr + " & Open Door"
        Else
            vStr = vStr
        End If
        .Text = vStr

        .Col = 3
        Select Case aInOutMode
            Case LOG_MODE_IO
                .Text = "General"
            Case LOG_MODE_IN1
                .Text = "IN1"
            Case LOG_MODE_IN2
                .Text = "IN2"
            Case LOG_MODE_IN3
                .Text = "IN3"
            Case LOG_MODE_OUT1
                .Text = "OUT1"
            Case LOG_MODE_OUT2
                .Text = "OUT2"
            Case LOG_MODE_OUT3
                .Text = "OUT3"
            Case Else
                .Text = "--"
        End Select

        .Col = 4
        .Text = CStr(Year(adwDate)) & "/" & Format(Month(adwDate), "0#") & "/" & Format(Day(adwDate), "0#") & _
                " " & Format(Hour(adwDate), "0#") & ":" & Format(Minute(adwDate), "0#") & ":" & Format(Second(adwDate), "0#")
        lblTotal.Caption = "Total : " & anCount
        DoEvents
    End With
End Function

Private Function funcShowGeneralLogDataToGrid_1(anCount As Long, aSEnrollNumber As Long, _
                                        aVerifyMode As Long, aInOutMode As Long, _
                                        astrDateTime As String, abDrawFlag As Boolean) As Boolean

    Dim vnkk As Long
    Dim vnHeight As Long, vnTop As Long, vnPos As Long
    Dim vStr As String
    funcShowGeneralLogDataToGrid_1 = True
    If anCount <= 1 Then
        fngrdIndex = 1
        fgrdLogView(1).Redraw = abDrawFlag
    Else
        If fngrdIndex * DEF_MAX_DISPCOUNT < anCount Then
            If abDrawFlag = False Then
                fgrdLogView(fngrdIndex).Redraw = True
            End If
            fngrdIndex = fngrdIndex + 1
            If fngrdIndex > UBound(fgrdLogView) Then
                funcShowGeneralLogDataToGrid_1 = False
                Exit Function
            End If
            vnHeight = fnGridHeight
            vnTop = fgrdLogView(1).Top
            For vnkk = 1 To fngrdIndex
                fgrdLogView(vnkk).Top = vnTop + (vnkk - 1) * (vnHeight / fngrdIndex)
                fgrdLogView(vnkk).Height = vnHeight / fngrdIndex
            Next vnkk
            fgrdLogView(fngrdIndex).Redraw = abDrawFlag
            fgrdLogView(fngrdIndex).Visible = True
        End If
    End If
    vnPos = anCount - (fngrdIndex - 1) * DEF_MAX_DISPCOUNT
    If fngrdIndex > 1 Then vnPos = vnPos - 1

    With fgrdLogView(fngrdIndex)
        .AddItem (1)
        .Row = vnPos
        .Col = 0
        .Text = anCount
        .Col = 1
        .Text = aSEnrollNumber
        .Col = 2
        Select Case aVerifyMode Mod LOG_OPEN_DOOR
            Case LOG_FPVERIFY           '1
                vStr = "Fp"
            Case LOG_PASSVERIFY         '2
                vStr = "Password"
            Case LOG_CARDVERIFY         '3
                vStr = "Card"
            Case LOG_FPPASS_VERIFY      '4
                vStr = "Fp+Password"
            Case LOG_FPCARD_VERIFY      '5
                vStr = "Fp+Card"
            Case LOG_PASSFP_VERIFY      '6
                vStr = "Password+Fp"
            Case LOG_CARDFP_VERIFY      '7
                vStr = "Card+Fp"
            Case LOG_JOB_NO_VERIFY      '8
                vStr = "Job No"
            Case LOG_CARDPASS_VERIFY    '9
                vStr = "Card+Pass"
            Case LOG_CLOSE_DOOR         '10
                vStr = "Close Door"
            Case LOG_OPEN_HAND          '11
                vStr = "Hand Open"
            Case LOG_PROG_OPEN          '12
                vStr = "Prog Open"
            Case LOG_PROG_CLOSE         '13
                vStr = "PC Close"
            Case LOG_OPEN_IREGAL        '14
                vStr = "Iregal Open"
            Case LOG_CLOSE_IREGAL       '15
                vStr = "Iregal Close"
            Case LOG_OPEN_COVER         '16
                vStr = "Cover Open"
            Case LOG_CLOSE_COVER        '17
                vStr = "Cover Close"
            Case Else
                vStr = "--"
        End Select
        If aVerifyMode \ LOG_OPEN_THREAT = 1 Then
            vStr = vStr + " & Open Door as Threat"
        ElseIf aVerifyMode \ LOG_OPEN_DOOR = 1 Then
            vStr = vStr + " & Open Door"
        Else
            vStr = vStr
        End If
        .Text = vStr

        .Col = 3
        Select Case aInOutMode
            Case LOG_MODE_IO
                .Text = "General"
            Case LOG_MODE_IN1
                .Text = "IN1"
            Case LOG_MODE_IN2
                .Text = "IN2"
            Case LOG_MODE_IN3
                .Text = "IN3"
            Case LOG_MODE_OUT1
                .Text = "OUT1"
            Case LOG_MODE_OUT2
                .Text = "OUT2"
            Case LOG_MODE_OUT3
                .Text = "OUT3"
            Case Else
                .Text = "--"
        End Select

        .Col = 4
        .Text = astrDateTime
              
        lblTotal.Caption = "Total : " & anCount
        DoEvents
    End With
End Function
Private Sub funcGetGeneralLogData(Optional abUSBFlag As Boolean = False)
    Dim vSEnrollNumber As Long
    Dim vVerifyMode As Long
    Dim vInOutMode As Long
    Dim vdwDate As Date
    Dim vnCount As Long
    Dim vstrFileName As String
    Dim vdBeginDate As Date
    Dim vdEndDate As Date
    Dim vstrTmp As String
    Dim vbRet As Boolean
    Dim vnReadMark As Long
    Dim vnFileNum As Integer
    Dim vstrFileData As String
    Dim vnResultCode As Long
    Dim vnYear As Long
    Dim vnMonth As Long
    Dim vnDay As Long
    Dim vnHour As Long
    Dim vnMinute As Long
    Dim vnSecond As Long
    
    lblMessage.Caption = "Waiting..."
    lblTotal.Caption = "Total : 0"
    DoEvents
    funcGeneralLogDataGridFormat

    If abUSBFlag = True Then
        dlgOpen.InitDir = CurDir
        dlgOpen.CancelError = False
        dlgOpen.Flags = cdlOFNHideReadOnly
        dlgOpen.Filter = "GLog Files (*.txt)|*.txt|All Files (*.*)|*.*"
        dlgOpen.FilterIndex = 1
        dlgOpen.InitDir = CurDir
        dlgOpen.ShowOpen
        vstrFileName = dlgOpen.FileName
        If vstrFileName = "" Then Exit Sub
    Else
        vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
        If vnResultCode <> RUN_SUCCESS Then
            lblMessage.Caption = gstrNoDevice
            Exit Sub
        End If
    End If

    MousePointer = vbHourglass
    DoEvents

    If abUSBFlag = True Then
        vnResultCode = FK_USBLoadGeneralLogDataFromFile(fnCommHandleIndex, vstrFileName)
    Else
        If chkReadMark.Value = vbChecked Then
            vnReadMark = 1
        Else
            vnReadMark = 0
        End If
        vnResultCode = FK_LoadGeneralLogData(fnCommHandleIndex, vnReadMark)
    End If
    
    'open file for save
    If chkSave.Value = Checked Then
        vnFileNum = FreeFile
        If vnReadMark = 0 Then
            vstrFileName = App.Path & "\AllLog.txt"
        Else
            vstrFileName = App.Path & "\Log.txt"
        End If
        If Dir(vstrFileName) <> "" Then Kill vstrFileName
        vstrFileData = "No." & vbTab & "EnrNo" & vbTab & "Verify" & vbTab & "InOut" & vbTab & "DateTime" + vbCrLf
        Open vstrFileName For Binary As #vnFileNum
        Put #vnFileNum, , vstrFileData
    End If

        
    If vnResultCode <> RUN_SUCCESS Then
       lblMessage.Caption = ReturnResultPrint(vnResultCode)
    Else
        lblMessage.Caption = "Getting..."
        DoEvents

        vnCount = 1
        Do
            vnResultCode = FK_GetGeneralLogData_1(fnCommHandleIndex, _
                                             vSEnrollNumber, _
                                             vVerifyMode, _
                                             vInOutMode, _
                                             vnYear, vnMonth, vnDay, vnHour, vnMinute, vnSecond)

            If vnResultCode <> RUN_SUCCESS Then
                If vnResultCode = RUNERR_DATAARRAY_END Then
                    vnResultCode = RUN_SUCCESS
                End If
                Exit Do
            End If

            vstrTmp = funcMakeDateTimeString(vnYear, vnMonth, vnDay, vnHour, vnMinute, vnSecond)
            If chkSave.Value = vbChecked Then
                vstrFileData = funcMakeGeneralLogFileData(vnCount, vSEnrollNumber, _
                                        vVerifyMode, vInOutMode, vstrTmp)
   
                Put #vnFileNum, , vstrFileData
            End If

            vbRet = funcShowGeneralLogDataToGrid_1(vnCount, vSEnrollNumber, vVerifyMode, vInOutMode, _
                                                   vstrTmp, False)

                               
            If vbRet = False Then Exit Do
            vnCount = vnCount + 1
        Loop
        If fngrdIndex > 0 Then
            fgrdLogView(fngrdIndex).Redraw = True
        End If

        If abUSBFlag = False And chkSave.Value = vbChecked Then
            Close #vnFileNum
        End If

        If vnResultCode = RUN_SUCCESS Then
            If abUSBFlag = True Then
                lblMessage.Caption = "USBReadGeneralLogDataFromFile OK"
            Else
                lblMessage.Caption = "ReadGeneralLogData OK"
            End If
        Else
            lblMessage.Caption = ReturnResultPrint(vnResultCode)
        End If
    End If
    
    MousePointer = vbDefault
    If abUSBFlag = False Then
        FK_EnableDevice fnCommHandleIndex, 1
    End If
End Sub

Private Sub cmdGlogData_Click()
    OwnerEnableButtons False
    funcGetGeneralLogData
    OwnerEnableButtons True
End Sub

Private Sub cmdUsbGLogData_Click()
    OwnerEnableButtons False
    funcGetGeneralLogData True
    OwnerEnableButtons True
End Sub

Private Sub cmdEmptyGLogData_Click()
    Dim vnResultCode As Long

    cmdEmptyGLogData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdEmptyGLogData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_EmptyGeneralLogData(fnCommHandleIndex)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice fnCommHandleIndex, 1
    cmdEmptyGLogData.Enabled = True
End Sub

Private Sub OwnerEnableButtons(abEnableFlag As Boolean)
    Dim vbFrmOpenFlag As Boolean

    vbFrmOpenFlag = frmMain.gbOpenFlag
    cmdSLogData.Enabled = vbFrmOpenFlag And abEnableFlag
    cmdEmptySLogData.Enabled = vbFrmOpenFlag And abEnableFlag
    cmdUsbSLogData.Enabled = abEnableFlag
    cmdGLogData.Enabled = vbFrmOpenFlag And abEnableFlag
    cmdEmptyGLogData.Enabled = vbFrmOpenFlag And abEnableFlag
    cmdUsbGLogData.Enabled = abEnableFlag
    DoEvents
End Sub

Private Function funcMakeGeneralLogFileData(ByVal anCount As Long, ByVal aSEnrollNumber As Long, _
                                        ByVal aVerifyMode As Long, ByVal aInOutMode As Long, _
                                        ByVal astrDateTime As String) As String
    Dim vstrData As String
    Dim vstrDTime As String
    
    vstrData = CStr(anCount) & vbTab & CStr(aSEnrollNumber) & vbTab & CStr(aVerifyMode) & vbTab & CStr(aInOutMode) & vbTab

    funcMakeGeneralLogFileData = vstrData & astrDateTime & vbCrLf
End Function

Private Function funcMakeDateTimeString(ByVal anYear As Long, ByVal anMonth As Long, ByVal anDay As Long, _
                                        ByVal anHour As Long, ByVal anMinute As Long, ByVal anSecond As Long)
    funcMakeDateTimeString = CStr(anYear & "/" & Format(anMonth, "0#") & "/" & Format(anDay, "0#") & _
                " " & Format(anHour, "0#") & ":" & Format(anMinute, "0#") & ":" & Format(anSecond, "0#"))
End Function



