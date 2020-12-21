VERSION 5.00
Begin VB.Form frmAdjust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjuste/Restore Info"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   Icon            =   "frmAdjust.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRestoredState 
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmAdjust.frx":0442
      Left            =   4380
      List            =   "frmAdjust.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1905
      Width           =   1245
   End
   Begin VB.TextBox txtRestoredHour 
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
      Left            =   2910
      MaxLength       =   32
      TabIndex        =   7
      Top             =   1905
      Width           =   495
   End
   Begin VB.TextBox txtAdjustedHour 
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
      Left            =   2910
      MaxLength       =   32
      TabIndex        =   2
      Top             =   1320
      Width           =   510
   End
   Begin VB.TextBox txtRestoredMinute 
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
      Left            =   3585
      MaxLength       =   32
      TabIndex        =   8
      Top             =   1905
      Width           =   495
   End
   Begin VB.TextBox txtAdjustedMinute 
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
      Left            =   3585
      MaxLength       =   32
      TabIndex        =   3
      Top             =   1320
      Width           =   510
   End
   Begin VB.CommandButton cmdSetAdjustInfo 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3900
      TabIndex        =   11
      Top             =   2625
      Width           =   1725
   End
   Begin VB.CommandButton cmdGetAdjustInfo 
      Caption         =   "Get"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2085
      TabIndex        =   10
      Top             =   2625
      Width           =   1725
   End
   Begin VB.TextBox txtAdjustedDay 
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
      Left            =   2265
      MaxLength       =   32
      TabIndex        =   1
      Top             =   1320
      Width           =   510
   End
   Begin VB.TextBox txtRestoredDay 
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
      Left            =   2265
      MaxLength       =   32
      TabIndex        =   6
      Top             =   1905
      Width           =   495
   End
   Begin VB.TextBox txtAdjustedMonth 
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
      Left            =   1575
      MaxLength       =   32
      TabIndex        =   0
      Top             =   1320
      Width           =   510
   End
   Begin VB.TextBox txtRestoredMonth 
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
      Left            =   1575
      MaxLength       =   32
      TabIndex        =   5
      Top             =   1905
      Width           =   495
   End
   Begin VB.ComboBox cmbAdjustedState 
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmAdjust.frx":0461
      Left            =   4380
      List            =   "frmAdjust.frx":046E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3450
      TabIndex        =   21
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3450
      TabIndex        =   20
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2100
      TabIndex        =   19
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2115
      TabIndex        =   18
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "HH:MM"
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
      Left            =   3045
      TabIndex        =   17
      Top             =   975
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "MM-DD"
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
      Left            =   1710
      TabIndex        =   16
      Top             =   975
      Width           =   915
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
      Left            =   210
      TabIndex        =   15
      Top             =   270
      Width           =   5415
   End
   Begin VB.Label lblChangeFlag 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change State"
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
      Left            =   4275
      TabIndex        =   14
      Top             =   975
      Width           =   1350
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjusted on"
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
      Left            =   225
      TabIndex        =   13
      Top             =   1380
      Width           =   1185
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restored on"
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
      Left            =   225
      TabIndex        =   12
      Top             =   1965
      Width           =   1230
   End
End
Attribute VB_Name = "frmAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fnMachineNumber As Long
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
    fnMachineNumber = Val(Trim(frmMain.txtMachineNumber.Text))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub cmbAdjustedState_Click()
    If cmbAdjustedState.ListIndex = 0 Then cmbRestoredState.ListIndex = 0
    If cmbAdjustedState.ListIndex = 1 Then cmbRestoredState.ListIndex = 2
    If cmbAdjustedState.ListIndex = 2 Then cmbRestoredState.ListIndex = 1
End Sub

Private Sub cmdGetAdjustInfo_Click()
    Dim vAdjustedState As Long
    Dim vAdjustedMonth As Long, vAdjustedDay As Long, vAdjustedHour As Long, vAdjustedMinute  As Long
    Dim vRestoredState As Long
    Dim vRestoredMonth As Long, vRestoredDay As Long, vRestoredHour As Long, vRestoredMinute  As Long
    Dim vbRet As Boolean
    Dim vErrorCode As Long

    lblMessage.Caption = ""
    DoEvents

    vbRet = FK_EnableDevice(fnCommHandleIndex, False)
    If vbRet = False Then
        lblMessage.Caption = gstrNoDevice
        Exit Sub
    End If

    vbRet = FK_GetAdjustInfo(fnCommHandleIndex, _
                             vAdjustedState, _
                             vAdjustedMonth, _
                             vAdjustedDay, _
                             vAdjustedHour, _
                             vAdjustedMinute, _
                             vRestoredState, _
                             vRestoredMonth, _
                             vRestoredDay, _
                             vRestoredHour, _
                             vRestoredMinute)
    If vbRet = True Then
        If vAdjustedState < 3 Then
            cmbAdjustedState.ListIndex = vAdjustedState
        Else
            cmbAdjustedState.ListIndex = 0
        End If
        txtAdjustedMonth.Text = vAdjustedMonth
        txtAdjustedDay.Text = vAdjustedDay
        txtAdjustedHour.Text = vAdjustedHour
        txtAdjustedMinute.Text = vAdjustedMinute
        
        If vRestoredState < 3 Then
            cmbRestoredState.ListIndex = vRestoredState
        Else
            cmbRestoredState.ListIndex = 0
        End If
        txtRestoredMonth.Text = vRestoredMonth
        txtRestoredDay.Text = vRestoredDay
        txtRestoredHour.Text = vRestoredHour
        txtRestoredMinute.Text = vRestoredMinute
        lblMessage.Caption = "Success!"
    Else
        lblMessage.Caption = "Faile!"
    End If

    FK_EnableDevice fnCommHandleIndex, True
End Sub

Private Sub cmdSetAdjustInfo_Click()
    Dim vAdjustedState As Long
    Dim vAdjustedMonth As Long, vAdjustedDay As Long, vAdjustedHour As Long, vAdjustedMinute  As Long
    Dim vRestoredState As Long
    Dim vRestoredMonth As Long, vRestoredDay As Long, vRestoredHour As Long, vRestoredMinute  As Long
    Dim vbRet As Boolean
    Dim vErrorCode As Long

    lblMessage.Caption = ""
    DoEvents

    vbRet = FK_EnableDevice(fnCommHandleIndex, False)
    If vbRet = False Then
        lblMessage.Caption = gstrNoDevice
        Exit Sub
    End If

    vAdjustedState = cmbAdjustedState.ListIndex
    vAdjustedMonth = txtAdjustedMonth.Text
    vAdjustedDay = txtAdjustedDay.Text
    vAdjustedHour = txtAdjustedHour.Text
    vAdjustedMinute = txtAdjustedMinute.Text
    
    vRestoredState = cmbRestoredState.ListIndex
    vRestoredMonth = txtRestoredMonth.Text
    vRestoredDay = txtRestoredDay.Text
    vRestoredHour = txtRestoredHour.Text
    vRestoredMinute = txtRestoredMinute.Text
    
    vbRet = FK_SetAdjustInfo(fnCommHandleIndex, _
                            vAdjustedState, _
                            vAdjustedMonth, _
                            vAdjustedDay, _
                            vAdjustedHour, _
                            vAdjustedMinute, _
                            vRestoredState, _
                            vRestoredMonth, _
                            vRestoredDay, _
                            vRestoredHour, _
                            vRestoredMinute)
    If vbRet = True Then
        lblMessage.Caption = "Success!"
    Else
        lblMessage.Caption = "Faile!"
    End If

    FK_EnableDevice fnCommHandleIndex, True
End Sub
