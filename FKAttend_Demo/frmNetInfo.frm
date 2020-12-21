VERSION 5.00
Begin VB.Form frmNetInfo 
   Caption         =   "Get && Set Net Information"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbSetNetInfo 
      Caption         =   "Set NetInfo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3315
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmbGetNetInfo 
      Caption         =   "Get NetInfo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CheckBox chkSerReq 
      Caption         =   "ServerRequest"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtServerPort 
      Height          =   405
      Left            =   2760
      MaxLength       =   32
      TabIndex        =   1
      Top             =   1560
      Width           =   2250
   End
   Begin VB.TextBox txtServerIP 
      Height          =   405
      Left            =   2760
      MaxLength       =   32
      TabIndex        =   0
      Top             =   960
      Width           =   2250
   End
   Begin VB.Label lblSerVerIP 
      AutoSize        =   -1  'True
      Caption         =   "Server IP Address : "
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
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Port : "
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
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
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
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmNetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
    txtServerIP.Text = "0.0.0.0"
    txtServerPort.Text = 0
End Sub

Private Sub cmbGetNetInfo_Click()
    
    Dim vnResultCode As Long
    Dim vstrServerIP As String
    Dim vnServerPort As Long
    Dim vnServerRequest As Long

    
    cmbGetNetInfo.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmbGetNetInfo.Enabled = True
        Exit Sub
    End If

    vstrServerIP = Space(256)
    vnResultCode = FK_GetServerNetInfo(fnCommHandleIndex, vstrServerIP, vnServerPort, vnServerRequest)
    
    If vnResultCode = RUN_SUCCESS Then
        txtServerIP.Text = vstrServerIP
        txtServerPort.Text = vnServerPort
        If vnServerRequest = 1 Then
            chkSerReq.Value = Checked
        Else
            chkSerReq.Value = Unchecked
        End If
    End If
    
    lblMessage.Caption = ReturnResultPrint(vnResultCode)
    FK_EnableDevice fnCommHandleIndex, 1
    cmbGetNetInfo.Enabled = True

End Sub

Private Sub cmbSetNetInfo_Click()
    Dim vnResultCode As Long
    Dim vstrServerIP As String
    Dim vnServerPort As Long
    Dim vnServerRequest As Long
    
    cmbSetNetInfo.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmbSetNetInfo.Enabled = True
        Exit Sub
    End If
    
    vstrServerIP = Trim(txtServerIP.Text)
    vnServerPort = Val(Trim(txtServerPort.Text))
    If chkSerReq.Value = Checked Then
        vnServerRequest = 1
    Else
        vnServerRequest = 0
    End If
    vnResultCode = FK_SetServerNetInfo(fnCommHandleIndex, vstrServerIP, vnServerPort, vnServerRequest)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)
    FK_EnableDevice fnCommHandleIndex, 0
    
    cmbSetNetInfo.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub
