VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   5025
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtUserMessage 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1815
      TabIndex        =   21
      Top             =   2670
      Width           =   4875
   End
   Begin VB.TextBox txtName 
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
      Left            =   1800
      TabIndex        =   20
      Top             =   990
      Width           =   2070
   End
   Begin VB.CommandButton cmdSetAllUserNews 
      Caption         =   "Set All News"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   1572
   End
   Begin VB.CommandButton cmdDeleteCompanyName 
      Caption         =   "Delete Company Name"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton cmdSetCompanyName 
      Caption         =   "Set Company Name"
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
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton cmdDeleteNewsID 
      Caption         =   "Delete News ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5040
      TabIndex        =   8
      Top             =   3840
      Width           =   1692
   End
   Begin VB.CommandButton cmdClearNewsID 
      Caption         =   "Clear News ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3516
      TabIndex        =   7
      Top             =   3840
      Width           =   1536
   End
   Begin VB.CommandButton cmdClearNews 
      Caption         =   "Clear News"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   6
      Top             =   3840
      Width           =   1692
   End
   Begin VB.TextBox txtMessageID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   16
      Top             =   2160
      Width           =   852
   End
   Begin VB.CommandButton cmdSetUserNews 
      Caption         =   "Set News ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3516
      TabIndex        =   4
      Top             =   3240
      Width           =   1536
   End
   Begin VB.CommandButton cmdGetUserNews 
      Caption         =   "Get  News ID "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5040
      TabIndex        =   5
      Top             =   3240
      Width           =   1692
   End
   Begin VB.CommandButton cmdSetNewsMessage 
      Caption         =   "Set News"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1572
   End
   Begin VB.CommandButton cmdGetNewsMessage 
      Caption         =   "Get News"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   1692
   End
   Begin VB.TextBox txtEnrollNumber 
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   11
      Top             =   1560
      Width           =   1440
   End
   Begin VB.CommandButton cmdGetUserName 
      Caption         =   "Get User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "Get EnrollData From Device"
      Top             =   1440
      Width           =   2052
   End
   Begin VB.CommandButton cmdSetUserName 
      Caption         =   "Set User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "Set EnrollData To Device"
      Top             =   960
      Width           =   2052
   End
   Begin VB.Label Label4 
      Caption         =   "( 1-50 User message count )"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2880
      TabIndex        =   18
      Top             =   2160
      Width           =   3852
   End
   Begin VB.Label Label2 
      Caption         =   "News  ID :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   17
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Message :"
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
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
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
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   6675
   End
   Begin VB.Label lblEnrollNum 
      AutoSize        =   -1  'True
      Caption         =   "Enroll Number :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   600
      TabIndex        =   12
      Top             =   1080
      Width           =   660
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NAMESIZE = 400

Private gGetState As Boolean
Private glngUserName(NAMESIZE) As Long
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    txtEnrollNumber.Text = "1"
    txtMessageID.Text = "1"
    txtUserMessage.Text = "Thank you"
    fnCommHandleIndex = frmMain.gnCommHandleIndex
    
    FK_SetFontName fnCommHandleIndex, CHINA_FONTNAME, 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub cmdClearNews_Click()
    Dim vnResultCode As Long
    
    cmdClearNews.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdClearNews.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_SetNewsMessage(fnCommHandleIndex, _
                                            255, _
                                            glngUserName(0))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Clear All New Message OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdClearNews.Enabled = True
End Sub

Private Sub cmdClearNewsID_Click()
    Dim vEnrollNumber As Long
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdClearNewsID.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdClearNewsID.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = 0
    vMessageNumber = 255
    vnResultCode = FK_SetUserNewsID(fnCommHandleIndex, _
                                         vEnrollNumber, _
                                         vMessageNumber)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Clear User News OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdClearNewsID.Enabled = True
End Sub

Private Sub cmdDeleteCompanyName_Click()
    Dim vnResultCode As Long
    
    cmdDeleteCompanyName.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdDeleteCompanyName.Enabled = True
        Exit Sub
    End If

    txtUserMessage.Text = Empty
    vnResultCode = FK_SetNewsMessage(fnCommHandleIndex, _
                                    0, txtUserMessage.Text)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Delete Company Name OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdDeleteCompanyName.Enabled = True
End Sub

Private Sub cmdDeleteNewsID_Click()
    Dim vEnrollNumber As Long
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdDeleteNewsID.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdDeleteNewsID.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = Val(txtEnrollNumber.Text)
    vMessageNumber = 255
    vnResultCode = FK_SetUserNewsID(fnCommHandleIndex, _
                                         vEnrollNumber, _
                                         vMessageNumber)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Set User News OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdDeleteNewsID.Enabled = True
End Sub

Private Sub cmdGetNewsMessage_Click()
    Dim vMessageNumber As Long
    Dim vNews As String
    Dim vnResultCode As Long
    
    cmdGetNewsMessage.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetNewsMessage.Enabled = True
        Exit Sub
    End If
    
    vMessageNumber = Val(txtMessageID.Text)
    vNews = Space(256)
    vnResultCode = FK_GetNewsMessage(fnCommHandleIndex, _
                                           vMessageNumber, _
                                           vNews)
    If vnResultCode = RUN_SUCCESS Then
        txtUserMessage.Text = vNews
        lblMessage.Caption = "Get News Message OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetNewsMessage.Enabled = True
End Sub

Private Sub cmdGetUserName_Click()
    Dim vEnrollNumber As Long
    Dim vName As String
    Dim vnResultCode As Long
    Dim vName1 As String
    
    cmdGetUserName.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetUserName.Enabled = True
        Exit Sub
    End If
    
    vName = Space(256)
    vEnrollNumber = Val(Trim(txtEnrollNumber.Text))
    
    vnResultCode = FK_GetUserName(fnCommHandleIndex, _
                                        vEnrollNumber, _
                                        vName)
'    vName1 = Left(vName, Len(Trim(vName)) - 1)
    If vnResultCode = RUN_SUCCESS Then
        txtName.Text = vName
        lblMessage.Caption = "GetUserName OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetUserName.Enabled = True
End Sub

Private Sub cmdGetUserNews_Click()
    Dim vEnrollNumber As Long
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdGetUserNews.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetUserNews.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = Val(Trim(txtEnrollNumber.Text))
    vnResultCode = FK_GetUserNewsID(fnCommHandleIndex, _
                                           vEnrollNumber, _
                                           vMessageNumber)
    If vnResultCode = RUN_SUCCESS Then
        txtMessageID.Text = vMessageNumber
        lblMessage.Caption = "Get User News ID OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetUserNews.Enabled = True
End Sub

Private Sub cmdSetCompanyName_Click()
    Dim vnResultCode As Long
    
    cmdSetCompanyName.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetCompanyName.Enabled = True
        Exit Sub
    End If
    
    vnResultCode = FK_SetNewsMessage(fnCommHandleIndex, _
                                        0, Trim(txtUserMessage.Text))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Set Company Name OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetCompanyName.Enabled = True
End Sub

Private Sub cmdSetNewsMessage_Click()
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdSetNewsMessage.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetNewsMessage.Enabled = True
        Exit Sub
    End If
    
    vMessageNumber = Val(Trim(txtMessageID.Text))
    vnResultCode = FK_SetNewsMessage(fnCommHandleIndex, _
                                            vMessageNumber, _
                                            Trim(txtUserMessage.Text))
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Set New Message OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetNewsMessage.Enabled = True
End Sub

Private Sub cmdSetUserName_Click()
    Dim vEnrollNumber As Long
    Dim vnResultCode As Long
    
    cmdSetUserName.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetUserName.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = Val(txtEnrollNumber.Text)
    vnResultCode = FK_SetUserName(fnCommHandleIndex, _
                                    vEnrollNumber, Trim(txtName.Text))
                                    
    
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "SetUserName OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetUserName.Enabled = True
End Sub

Private Sub cmdSetUserNews_Click()
    Dim vEnrollNumber As Long
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdSetUserNews.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetUserNews.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = Val(Trim(txtEnrollNumber.Text))
    vMessageNumber = Val(Trim(txtMessageID.Text))
    vnResultCode = FK_SetUserNewsID(fnCommHandleIndex, _
                                         vEnrollNumber, _
                                         vMessageNumber)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Set User News  OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetUserNews.Enabled = True
End Sub

Private Sub cmdSetAllUserNews_Click()
    Dim vEnrollNumber As Long
    Dim vMessageNumber As Long
    Dim vnResultCode As Long
    
    cmdSetAllUserNews.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetAllUserNews.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = 0
    vMessageNumber = Val(Trim(txtMessageID.Text))
    vnResultCode = FK_SetUserNewsID(fnCommHandleIndex, _
                                         vEnrollNumber, _
                                         vMessageNumber)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Set All User News OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice fnCommHandleIndex, 1
    cmdSetAllUserNews.Enabled = True
End Sub


