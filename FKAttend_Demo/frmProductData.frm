VERSION 5.00
Begin VB.Form frmProductData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Data"
   ClientHeight    =   3555
   ClientLeft      =   4995
   ClientTop       =   3105
   ClientWidth     =   6090
   Icon            =   "frmProductData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProductCode 
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
      Left            =   2565
      MaxLength       =   32
      TabIndex        =   3
      Top             =   2145
      Width           =   3210
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3975
      TabIndex        =   2
      Top             =   2665
      Width           =   1815
   End
   Begin VB.TextBox txtBackupNo 
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
      Left            =   2565
      MaxLength       =   32
      TabIndex        =   1
      Top             =   1545
      Width           =   3210
   End
   Begin VB.TextBox txtSerialNo 
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
      Left            =   2565
      MaxLength       =   32
      TabIndex        =   0
      Top             =   945
      Width           =   3210
   End
   Begin VB.Label lblBackuplNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Backup Number :"
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
      Left            =   510
      TabIndex        =   7
      Top             =   1605
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Product Code :"
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
      Left            =   510
      TabIndex        =   6
      Top             =   2205
      Width           =   1485
   End
   Begin VB.Label lblSerialNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Serial Number :"
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
      Left            =   510
      TabIndex        =   5
      Top             =   1005
      Width           =   1590
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
      Left            =   270
      TabIndex        =   4
      Top             =   345
      Width           =   5505
   End
End
Attribute VB_Name = "frmProductData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fnCommHandleIndex As Long

Private Sub Form_Load()
    fnCommHandleIndex = frmMain.gnCommHandleIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub cmdGetData_Click()
    Dim vstrData As String
    Dim vnResultCode As Long

    cmdGetData.Enabled = False
    txtSerialNo.Text = ""
    txtBackupNo.Text = ""
    txtProductCode.Text = ""
    lblMessage.Caption = "Waiting..."
    DoEvents

    vnResultCode = FK_EnableDevice(fnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FuncGetProductData(PRODUCT_SERIALNUMBER, vstrData)
    If vnResultCode = RUN_SUCCESS Then
        txtSerialNo.Text = vstrData

        vnResultCode = FuncGetProductData(PRODUCT_BACKUPNUMBER, vstrData)
        If vnResultCode = RUN_SUCCESS Then
            txtBackupNo.Text = vstrData

            vnResultCode = FuncGetProductData(PRODUCT_CODE, vstrData)
            If vnResultCode = RUN_SUCCESS Then
                txtProductCode.Text = vstrData
            End If
        End If
    End If

    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "GetProductData OK"
    End If
    
    FK_EnableDevice fnCommHandleIndex, 1
    cmdGetData.Enabled = True
End Sub

Private Function FuncGetProductData(anIndex As Long, astrItem As String) As Long
    Dim vstrData As String
    
    vstrData = Space(256)
    FuncGetProductData = FK_GetProductData(fnCommHandleIndex, anIndex, vstrData)
    If FuncGetProductData <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(FuncGetProductData)
        Exit Function
    End If
    astrItem = Trim(vstrData)
End Function

