VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEnroll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Enroll Data"
   ClientHeight    =   8490
   ClientLeft      =   3075
   ClientTop       =   1530
   ClientWidth     =   6990
   Icon            =   "frmEnroll.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUSBGetAllData_C 
      Caption         =   "Get All Enroll Data_C(USB)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   26
      ToolTipText     =   "Get All Enroll Data From USB file And Save To DataBase"
      Top             =   4395
      Width           =   3000
   End
   Begin VB.CommandButton cmdUSBSetAllData_C 
      Caption         =   "Set All Enroll Data_C(USB)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   25
      ToolTipText     =   "Load All Enroll Data From DataBase And Set To USB file"
      Top             =   4875
      Width           =   3000
   End
   Begin VB.CommandButton cmdBenumbManager 
      Caption         =   "Benumb All Managers"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   24
      Top             =   6795
      Width           =   3000
   End
   Begin VB.CommandButton cmdEmptyEnrollData 
      Caption         =   "Empty Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   14
      ToolTipText     =   "Clear all Enroll data Into Device"
      Top             =   7305
      Width           =   3000
   End
   Begin VB.CommandButton cmdGetEnrollInfo 
      Caption         =   "Get Enroll Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "Get All Enrolled User Info From Device"
      Top             =   5370
      Width           =   3000
   End
   Begin VB.CommandButton cmdGetAllEnrollData 
      Caption         =   "Get All Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Get All Enroll Data From Device And Save To DataBase"
      Top             =   2445
      Width           =   3000
   End
   Begin VB.CommandButton cmdSetAllEnrollData 
      Caption         =   "Set All Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "Load All Enroll Data From DataBase And Set To Device"
      Top             =   2910
      Width           =   3000
   End
   Begin VB.CommandButton cmdEnableUser 
      Caption         =   "Enable User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   11
      ToolTipText     =   "Set to enable status of user"
      Top             =   5835
      Width           =   1480
   End
   Begin VB.CommandButton cmdModifyPrivilege 
      Caption         =   "ModifyPrivilege"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "Modify privilege of user"
      Top             =   6315
      Width           =   3000
   End
   Begin VB.CommandButton cmdUSBSetAllData 
      Caption         =   "Set All Enroll Data(USB)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "Load All Enroll Data From DataBase And Set To USB file"
      Top             =   3930
      Width           =   3000
   End
   Begin VB.CommandButton cmdGetEnrollData 
      Caption         =   "Get Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Get EnrollData From Device"
      Top             =   840
      Width           =   3000
   End
   Begin VB.CommandButton cmdDeleteEnrollData 
      Caption         =   "Delete Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Delete Enroll Data Into Device"
      Top             =   1785
      Width           =   3000
   End
   Begin VB.CommandButton cmdSetEnrollData 
      Caption         =   "Set Enroll Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Set EnrollData To Device"
      Top             =   1305
      Width           =   3000
   End
   Begin VB.CommandButton cmdClearData 
      Caption         =   "Clear All Data(E,GL,SL) "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   15
      ToolTipText     =   "Clear all Enroll data and all Log data Into Device"
      Top             =   7785
      Width           =   3000
   End
   Begin VB.CommandButton cmdDisableUser 
      Caption         =   "Disable User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5235
      TabIndex        =   12
      ToolTipText     =   "Set to disable status of user"
      Top             =   5835
      Width           =   1480
   End
   Begin VB.CommandButton cmdUSBGetAllData 
      Caption         =   "Get All Enroll Data(USB)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Get All Enroll Data From USB file And Save To DataBase"
      Top             =   3450
      Width           =   3000
   End
   Begin VB.ComboBox cmbBackupNumber 
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
      ItemData        =   "frmEnroll.frx":0442
      Left            =   2280
      List            =   "frmEnroll.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1410
      Width           =   1215
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
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrivilege 
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
      ItemData        =   "frmEnroll.frx":0446
      Left            =   2280
      List            =   "frmEnroll.frx":0448
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1995
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete DB"
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
      Left            =   2190
      TabIndex        =   1
      ToolTipText     =   "Delete All Saved Data From DataBase"
      Top             =   7080
      Width           =   1245
   End
   Begin VB.ListBox lstEnrollData 
      Height          =   3765
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   3240
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   5760
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoEnroll 
      Height          =   405
      Left            =   120
      Top             =   7118
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "0/0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoFind 
      Height          =   330
      Left            =   720
      Top             =   7200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Height          =   285
      Left            =   255
      TabIndex        =   23
      Top             =   945
      Width           =   1440
   End
   Begin VB.Label lblBackupNumber 
      AutoSize        =   -1  'True
      Caption         =   "Backup Number :"
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
      Left            =   255
      TabIndex        =   22
      Top             =   1515
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Privilege :"
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
      Left            =   255
      TabIndex        =   21
      Top             =   2100
      Width           =   870
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total : "
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
      Left            =   2070
      TabIndex        =   17
      Top             =   2700
      Width           =   630
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
      Left            =   210
      TabIndex        =   16
      Top             =   150
      Width           =   6555
   End
   Begin VB.Label lblEnrollData 
      AutoSize        =   -1  'True
      Caption         =   "Enroll Data :"
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
      Left            =   150
      TabIndex        =   3
      Top             =   2700
      Width           =   1125
   End
End
Attribute VB_Name = "frmEnroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const fcstrCnn40 = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="

Private mnCommHandleIndex As Long
Private mbGetState As Boolean

Private mlngPasswordData As Long

Const PWD_DATA_SIZE = 40
Const FP_DATA_SIZE = 1680
Const FACE_DATA_SIZE = 20080
Const VEIN_DATA_SIZE = 3080
Private mbytEnrollData(FACE_DATA_SIZE - 1) As Byte

Private ftBackupNumber As Variant

Private Sub Form_Load()
    Dim vstrDBPath As String
    Dim vnii As Long
    
    mnCommHandleIndex = frmMain.gnCommHandleIndex
    mbGetState = False
    ftBackupNumber = Array(BACKUP_FP_0, "Fp-0", _
                        BACKUP_FP_1, "Fp-1", _
                        BACKUP_FP_2, "Fp-2", _
                        BACKUP_FP_3, "Fp-3", _
                        BACKUP_FP_4, "Fp-4", _
                        BACKUP_FP_5, "Fp-5", _
                        BACKUP_FP_6, "Fp-6", _
                        BACKUP_FP_7, "Fp-7", _
                        BACKUP_FP_8, "Fp-8", _
                        BACKUP_FP_9, "Fp-9", _
                        BACKUP_PSW, "Pass", _
                        BACKUP_CARD, "Card", _
                        BACKUP_FACE, "Face", _
                        BACKUP_VEIN_0, "Vein")

    lblMessage.Caption = ""
    txtEnrollNumber.Text = "1"
    cmbBackupNumber.Clear
    For vnii = 1 To UBound(ftBackupNumber) Step 2
        cmbBackupNumber.AddItem (ftBackupNumber(vnii))
    Next vnii
    cmbBackupNumber.ListIndex = 0
    cmbPrivilege.Clear
    cmbPrivilege.AddItem ("User")
    cmbPrivilege.AddItem ("Manager")
    cmbPrivilege.AddItem ("Registrar")
    cmbPrivilege.ListIndex = 0
    lstEnrollData.Clear
      
On Error GoTo lp_end
    DBWithItemEnable False

    vstrDBPath = App.Path & "\datEnrollDat.mdb"
    If Dir(vstrDBPath) = "" Then
        dlgOpen.InitDir = CurDir
        dlgOpen.Flags = cdlOFNHideReadOnly
        dlgOpen.Filter = "DB Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
        dlgOpen.FilterIndex = 1
        dlgOpen.ShowOpen
        vstrDBPath = dlgOpen.FileName
        dlgOpen.FileName = Empty
        If Len(vstrDBPath) <= 0 Or Dir(vstrDBPath) = "" Then
            GoTo lp_end
        End If
    End If

    With AdoEnroll
        .ConnectionString = fcstrCnn40 & vstrDBPath
        .RecordSource = "select * from tblEnroll"
        .ConnectionTimeout = 30
        .Refresh
        If .Recordset.RecordCount > 0 Then
            .Recordset.MoveLast
            .Recordset.MoveFirst
        End If
    End With

    AdoFind.ConnectionString = AdoEnroll.ConnectionString
    DBWithItemEnable True
lp_end:
    If frmMain.gbOpenFlag = False Then
        DisableButtons
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Visible = True
End Sub

Private Sub AdoEnroll_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim vnPos As Long
    Dim vEnrollNumber As Long
    Dim vEnrollName As String
    Dim vBackupNumber As Long
    Dim vPrivilege As Long

    If mbGetState = True Then Exit Sub
    With AdoEnroll.Recordset
        vnPos = .AbsolutePosition
        If vnPos < 0 Then vnPos = 0
        AdoEnroll.Caption = "  " & vnPos & "/" & .RecordCount
        If .RecordCount >= 1 Then FuncReadFromToDB vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName
    End With
End Sub

Private Sub AdoFind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    If ErrorNumber = 3704 Then
        fCancelDisplay = True
    End If
End Sub

Private Sub cmdDel_Click()
    cmdDel.Enabled = False
    DoEvents
    AdoFind.RecordSource = "delete * from tblEnroll"
On Error Resume Next
    AdoFind.Refresh
    AdoEnroll.Refresh
    cmdDel.Enabled = True
End Sub

Private Sub cmdGetEnrollData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vnResultCode As Long
   
    cmdGetEnrollData.Enabled = False
    lstEnrollData.Clear
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetEnrollData.Enabled = True
        Exit Sub
    End If
    
    vEnrollNumber = Val(txtEnrollNumber.Text)
    vBackupNumber = FuncGetBackupNumberFromItem
    
    vnResultCode = FK_GetEnrollData(mnCommHandleIndex, _
                                          vEnrollNumber, _
                                          vBackupNumber, _
                                          vPrivilege, _
                                          mbytEnrollData(0), _
                                          mlngPasswordData)
    If vnResultCode = RUN_SUCCESS Then
        If vPrivilege = MP_NONE Then
            cmbPrivilege.ListIndex = 0
        ElseIf vPrivilege = MP_ALL Then
            cmbPrivilege.ListIndex = 1
        End If
        FuncDispToListBox vBackupNumber
        lblMessage.Caption = "GetEnrollData OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If
    
    FK_EnableDevice mnCommHandleIndex, 1
    cmdGetEnrollData.Enabled = True
End Sub

Private Sub cmdSetEnrollData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vnResultCode As Long

    cmdSetEnrollData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vEnrollNumber = Val(txtEnrollNumber.Text)
    vBackupNumber = FuncGetBackupNumberFromItem
    If cmbPrivilege.ListIndex = 0 Then
        vPrivilege = MP_NONE
    ElseIf cmbPrivilege.ListIndex = 1 Then
        vPrivilege = MP_ALL
    End If
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetEnrollData.Enabled = True
        Exit Sub
    End If
    
    vnResultCode = FK_PutEnrollData(mnCommHandleIndex, _
                                          vEnrollNumber, _
                                          vBackupNumber, _
                                          vPrivilege, _
                                          mbytEnrollData(0), _
                                          mlngPasswordData)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Saving..."
        DoEvents
        vnResultCode = FK_SaveEnrollData(mnCommHandleIndex)
        If vnResultCode = RUN_SUCCESS Then
            lblMessage.Caption = "SetEnrollData OK"
        End If
    End If
        
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If
    
    FK_EnableDevice mnCommHandleIndex, 1
    cmdSetEnrollData.Enabled = True
End Sub

Private Sub cmdDeleteEnrollData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vnResultCode As Long

    cmdDeleteEnrollData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdDeleteEnrollData.Enabled = True
        Exit Sub
    End If

    vEnrollNumber = Val(txtEnrollNumber.Text)
    vBackupNumber = FuncGetBackupNumberFromItem
    vnResultCode = FK_DeleteEnrollData(mnCommHandleIndex, _
                                            vEnrollNumber, _
                                            vBackupNumber)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "DeleteEnrollData OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice mnCommHandleIndex, 1
    cmdDeleteEnrollData.Enabled = True
End Sub

Private Sub cmdGetAllEnrollData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vEnrollName As String
    Dim vnEnableFlag As Long
    Dim vnMessRet As Long
    Dim vTitle As String
    Dim vnResultCode As Long
    
    cmdGetAllEnrollData.Enabled = False
    lstEnrollData.Clear
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetAllEnrollData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_ReadAllUserID(mnCommHandleIndex)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
        FK_EnableDevice mnCommHandleIndex, 1
        cmdGetAllEnrollData.Enabled = True
        Exit Sub
    End If
    
'---- Get Enroll data and save into database -------------
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
        mbGetState = True
        Do
FFF:
            vnResultCode = FK_GetAllUserID(mnCommHandleIndex, _
                                                 vEnrollNumber, _
                                                 vBackupNumber, _
                                                 vPrivilege, _
                                                 vnEnableFlag)
            If vnResultCode <> RUN_SUCCESS Then
                If vnResultCode = RUNERR_DATAARRAY_END Then
                    vnResultCode = RUN_SUCCESS
                End If
                Exit Do
            End If
EEE:
            vnResultCode = FK_GetEnrollData(mnCommHandleIndex, _
                                                  vEnrollNumber, _
                                                  vBackupNumber, _
                                                  vPrivilege, _
                                                  mbytEnrollData(0), _
                                                  mlngPasswordData)
           
            If vnResultCode <> RUN_SUCCESS Then
                vnMessRet = MsgBox(ReturnResultPrint(vnResultCode) & ": Continue ?", vbYesNoCancel, "GetEnrollData")
                If vnMessRet = vbYes Then
                    GoTo EEE
                ElseIf vnMessRet = vbCancel Then
                    Exit Do
                Else
                    GoTo FFF
                End If
            End If

            FuncSaveToDB vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName
            Me.Caption = Format(vEnrollNumber, "0000#")
            DoEvents
        Loop
        mbGetState = False
        DoEvents
        If .RecordCount > 1 Then
            .MoveFirst
            .MoveLast
        End If
            End With
    Me.Caption = vTitle
    MousePointer = vbDefault

    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "GetAllEnrollData OK"
    Else
        lblMessage.Caption = "GetAllEnrollData Error : " & ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice mnCommHandleIndex, 1
    cmdGetAllEnrollData.Enabled = True
End Sub

Private Sub cmdSetAllEnrollData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vEnrollName As String
    Dim vnMessRet As Long
    Dim vStr As String
    Dim vTitle As String
    Dim vnResultCode As Long
    Dim vbRet As Boolean
    
    cmdSetAllEnrollData.Enabled = False
    lstEnrollData.Clear
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdSetAllEnrollData.Enabled = True
        Exit Sub
    End If

    mbGetState = True
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
         If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While .EOF = False
FFF:
                vbRet = FuncReadFromToDB(vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName, False)
                If vbRet <> True Then
                    vStr = "SetAllEnrollData Error"
                    Exit Do
                End If
                vnResultCode = FK_PutEnrollData(mnCommHandleIndex, _
                                                      vEnrollNumber, _
                                                      vBackupNumber, _
                                                      vPrivilege, _
                                                      mbytEnrollData(0), _
                                                      mlngPasswordData)
                If vnResultCode <> RUN_SUCCESS Then
                    vStr = "SetAllEnrollData Error"
                    vnMessRet = MsgBox(ReturnResultPrint(vnResultCode) & ": Continue ?", vbYesNoCancel, "SetEnrollData")
                    If vnMessRet = vbYes Then GoTo FFF
                    If vnMessRet = vbCancel Then Exit Do
                End If
                lblMessage.Caption = "ID = " & Format(vEnrollNumber, "000#") & ", FpNo = " & vBackupNumber _
                                    & ", Count = " & .AbsolutePosition
                
                Me.Caption = .AbsolutePosition
                DoEvents
                .MoveNext
            Loop
        End If
    End With
    Me.Caption = vTitle
    MousePointer = vbDefault
    mbGetState = False
    
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "Saving..."
        DoEvents
        vnResultCode = FK_SaveEnrollData(mnCommHandleIndex)
        If vnResultCode = RUN_SUCCESS Then
            lblMessage.Caption = "SetAllEnrollData OK"
        Else
            lblMessage.Caption = ReturnResultPrint(vnResultCode)
        End If
    Else
        lblMessage.Caption = vStr & " : " & ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice mnCommHandleIndex, 1
    cmdSetAllEnrollData.Enabled = True
End Sub

Private Sub cmdUSBGetAllData_Click()
    Dim vEnrollNumber As Long
    Dim vEnrollName As String
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vnEnableFlag As Long
    Dim vTitle As String
    Dim vstrFileName As String
    Dim vnResultCode As Long
    Dim vnNewsKind As Long
    Dim vnModelKind As Long
        
    vnNewsKind = NEWS_STANDARD
    dlgOpen.CancelError = False
    dlgOpen.Flags = cdlOFNHideReadOnly
    dlgOpen.Filter = "DAT Files (*.dat)|*.dat|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.InitDir = CurDir
    dlgOpen.ShowOpen
    vstrFileName = dlgOpen.FileName
    If vstrFileName = "" Then Exit Sub

    cmdUSBGetAllData.Enabled = False
    lstEnrollData.Clear
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents
    vEnrollName = Space(256)
    
    'vnModelKind = FK625_FP3000
    'FK_SetUSBModel mnCommHandleIndex, vnModelKind
    'FK_SetUDiskFileFKModel mnCommHandleIndex, "FK625OF"
    
    vnResultCode = FK_USBReadAllEnrollDataFromFile(mnCommHandleIndex, vstrFileName)
    
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
        cmdUSBGetAllData.Enabled = True
        Exit Sub
    End If

'---- Get Enroll data and save into database -------------
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
        mbGetState = True
        Do
            vnResultCode = FK_USBGetOneEnrollData(mnCommHandleIndex, _
                                                  vEnrollNumber, _
                                                  vBackupNumber, _
                                                  vPrivilege, _
                                                  mbytEnrollData(0), _
                                                  mlngPasswordData, _
                                                  vnEnableFlag, _
                                                  vEnrollName)


            If vnResultCode <> RUN_SUCCESS Then
                If vnResultCode = RUNERR_DATAARRAY_END Then
                    vnResultCode = RUN_SUCCESS
                End If
                Exit Do
            End If

            FuncSaveToDB vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName
            Me.Caption = Format(vEnrollNumber, "0000#")
            DoEvents
        Loop
        mbGetState = False
        DoEvents
        If .RecordCount > 1 Then
            .MoveFirst
            .MoveLast
        End If
    End With
    Me.Caption = vTitle
    MousePointer = vbDefault

    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "GetAllEnrollData(USB) OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    cmdUSBGetAllData.Enabled = True
End Sub

Private Sub cmdUSBSetAllData_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vEnrollName As String
    Dim vnMessRet As Long
    Dim vStr As String
    Dim vTitle As String
    Dim vstrFileName As String
    Dim vnEnableFlag As Long
    Dim vnResultCode As Long
    Dim vbRet As Boolean
    Dim vnNewsKind As Long
    Dim vnModelKind As Long
    
    vnNewsKind = NEWS_STANDARD
    lstEnrollData.Clear

    dlgOpen.CancelError = False
    dlgOpen.Flags = cdlOFNHideReadOnly
    dlgOpen.Filter = "DAT Files (*.dat)|*.dat|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.InitDir = CurDir
    dlgOpen.ShowSave
    vstrFileName = dlgOpen.FileName
    If vstrFileName = "" Then Exit Sub
    
    cmdUSBSetAllData.Enabled = False
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents
    
    'vnModelKind = FK625_FP3000
    'FK_SetUSBModel mnCommHandleIndex, vnModelKind
    'FK_SetUDiskFileFKModel mnCommHandleIndex, "FK625OF"

    mbGetState = True
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
         If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While .EOF = False
FFF:
                vbRet = FuncReadFromToDB(vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName, False)
                If vbRet <> True Then
                    vStr = "SetAllEnrollData(USB) Error"
                    Exit Do
                End If

                vnEnableFlag = 1
                vnResultCode = FK_USBSetOneEnrollData(mnCommHandleIndex, _
                                                      vEnrollNumber, _
                                                      vBackupNumber, _
                                                      vPrivilege, _
                                                      mbytEnrollData(0), _
                                                      mlngPasswordData, _
                                                      vnEnableFlag, _
                                                      vEnrollName)
                
                If vnResultCode <> RUN_SUCCESS Then
                    vStr = "USBSetOneEnrollData Error"
                    vnMessRet = MsgBox(ReturnResultPrint(vnResultCode) & ": Continue ?", vbYesNoCancel, vStr)
                    If vnMessRet = vbYes Then GoTo FFF
                    If vnMessRet = vbCancel Then Exit Do
                End If
                lblMessage.Caption = "ID = " & Format(vEnrollNumber, "000#") & ", FpNo = " & vBackupNumber _
                                    & ", Count = " & .AbsolutePosition
                
                Me.Caption = .AbsolutePosition
                DoEvents
                .MoveNext
            Loop
        End If
    End With

    Me.Caption = vTitle
    MousePointer = vbDefault
    mbGetState = False
    
    If vnResultCode = RUN_SUCCESS Then
        vnResultCode = FK_USBWriteAllEnrollDataToFile(mnCommHandleIndex, vstrFileName)
        If vnResultCode = RUN_SUCCESS Then
            lblMessage.Caption = "USBWriteAllEnrollDataToFile OK"
        Else
            lblMessage.Caption = ReturnResultPrint(vnResultCode)
        End If
    Else
        lblMessage.Caption = vStr
    End If

    cmdUSBSetAllData.Enabled = True
End Sub

Private Sub cmdUSBGetAllData_C_Click()

    Dim vEnrollNumber As Long
    Dim vEnrollName As String
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vnEnableFlag As Long
    Dim vTitle As String
    Dim vstrFileName As String
    Dim vnResultCode As Long
    Dim vnNewsKind As Long
    Dim vnModelKind As Long
    
    vnNewsKind = NEWS_EXTEND
    dlgOpen.CancelError = False
    dlgOpen.Flags = cdlOFNHideReadOnly
    dlgOpen.Filter = "DAT Files (*.dat)|*.dat|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.InitDir = CurDir
    dlgOpen.ShowOpen
    vstrFileName = dlgOpen.FileName
    If vstrFileName = "" Then Exit Sub

    cmdUSBGetAllData_C.Enabled = False
    lstEnrollData.Clear
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents
    vEnrollName = Space(256)
    
    'vnModelKind = FK735_FP3000
    'FK_SetUSBModel mnCommHandleIndex, vnModelKind
    FK_SetUDiskFileFKModel mnCommHandleIndex, "FK735HS3"
    
    
    vnResultCode = FK_USBReadAllEnrollDataFromFile_Color(mnCommHandleIndex, vstrFileName)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
        cmdUSBGetAllData_C.Enabled = True
        Exit Sub
    End If

'---- Get Enroll data and save into database -------------
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
        mbGetState = True
        Do
            vnResultCode = FK_USBGetOneEnrollData_Color(mnCommHandleIndex, _
                                                  vEnrollNumber, _
                                                  vBackupNumber, _
                                                  vPrivilege, _
                                                  mbytEnrollData(0), _
                                                  mlngPasswordData, _
                                                  vnEnableFlag, _
                                                  vEnrollName, _
                                                  vnNewsKind)
            If vnResultCode <> RUN_SUCCESS Then
                If vnResultCode = RUNERR_DATAARRAY_END Then
                    vnResultCode = RUN_SUCCESS
                End If
                Exit Do
            End If

            FuncSaveToDB vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName
            Me.Caption = Format(vEnrollNumber, "0000#")
            DoEvents
        Loop
        mbGetState = False
        DoEvents
        If .RecordCount > 1 Then
            .MoveFirst
            .MoveLast
        End If
    End With
    Me.Caption = vTitle
    MousePointer = vbDefault

    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "GetAllEnrollDataColor(USB) OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    cmdUSBGetAllData_C.Enabled = True

End Sub

Private Sub cmdUSBSetAllData_C_Click()
   Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vEnrollName As String
    Dim vnMessRet As Long
    Dim vStr As String
    Dim vTitle As String
    Dim vstrFileName As String
    Dim vnEnableFlag As Long
    Dim vnResultCode As Long
    Dim vbRet As Boolean
    Dim vnNewsKind As Long
    Dim vnModelKind As Long
    
    vnNewsKind = NEWS_EXTEND
    lstEnrollData.Clear
    dlgOpen.CancelError = False
    dlgOpen.Flags = cdlOFNHideReadOnly
    dlgOpen.Filter = "DAT Files (*.dat)|*.dat|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.InitDir = CurDir
    dlgOpen.ShowSave
    vstrFileName = dlgOpen.FileName
    If vstrFileName = "" Then Exit Sub
    
    cmdUSBSetAllData_C.Enabled = False
    vTitle = Me.Caption
    lblMessage.Caption = "Working..."
    DoEvents
    
    'vnModelKind = FK735_FP3000
    'FK_SetUSBModel mnCommHandleIndex, vnModelKind
    FK_SetUDiskFileFKModel mnCommHandleIndex, "FK735HS3"

    mbGetState = True
    MousePointer = vbHourglass
    With AdoEnroll.Recordset
         If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do While .EOF = False
FFF:
                vbRet = FuncReadFromToDB(vEnrollNumber, vBackupNumber, vPrivilege, vEnrollName, False)
                If vbRet <> True Then
                    vStr = "SetAllEnrollData(USB) Error"
                    Exit Do
                End If

                vnEnableFlag = 1
                vnResultCode = FK_USBSetOneEnrollData_Color(mnCommHandleIndex, _
                                                      vEnrollNumber, _
                                                      vBackupNumber, _
                                                      vPrivilege, _
                                                      mbytEnrollData(0), _
                                                      mlngPasswordData, _
                                                      vnEnableFlag, _
                                                      vEnrollName, _
                                                      vnNewsKind)
                If vnResultCode <> RUN_SUCCESS Then
                    vStr = "USBSetOneEnrollDataColor Error"
                    vnMessRet = MsgBox(ReturnResultPrint(vnResultCode) & ": Continue ?", vbYesNoCancel, vStr)
                    If vnMessRet = vbYes Then GoTo FFF
                    If vnMessRet = vbCancel Then Exit Do
                End If
                lblMessage.Caption = "ID = " & Format(vEnrollNumber, "000#") & ", FpNo = " & vBackupNumber _
                                    & ", Count = " & .AbsolutePosition
                
                Me.Caption = .AbsolutePosition
                DoEvents
                .MoveNext
            Loop
        End If
    End With

    Me.Caption = vTitle
    MousePointer = vbDefault
    mbGetState = False
    
    If vnResultCode = RUN_SUCCESS Then
        vnResultCode = FK_USBWriteAllEnrollDataToFile_Color(mnCommHandleIndex, vstrFileName, vnNewsKind)
        If vnResultCode = RUN_SUCCESS Then
            lblMessage.Caption = "USBWriteAllEnrollDataColorToFile OK"
        Else
            lblMessage.Caption = ReturnResultPrint(vnResultCode)
        End If
    Else
        lblMessage.Caption = vStr
    End If

    cmdUSBSetAllData_C.Enabled = True
End Sub


Private Sub cmdGetEnrollInfo_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vstrBackupNumber As String
    Dim vPrivilege As Long
    Dim vstrPrivilege As String
    Dim vnEnableFlag As Long
    Dim vstrEnableFlag As String
    Dim vnii As Long
    Dim vnResultCode As Long
    
    cmdGetEnrollInfo.Enabled = False
    lblEnrollData = "User IDs"
    lstEnrollData.Clear
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdGetEnrollInfo.Enabled = True
        Exit Sub
    End If
    
    vnResultCode = FK_ReadAllUserID(mnCommHandleIndex)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
        FK_EnableDevice mnCommHandleIndex, 1
        cmdGetEnrollInfo.Enabled = True
        Exit Sub
    End If
    
'------ Show all enroll information ----------
    vnii = 0
    lstEnrollData.AddItem (" No.         EnNo           BkNo        Priv    Enable")
    Do
        vnResultCode = FK_GetAllUserID(mnCommHandleIndex, _
                                             vEnrollNumber, _
                                             vBackupNumber, _
                                             vPrivilege, _
                                             vnEnableFlag)
        If vnResultCode <> RUN_SUCCESS Then
            If vnResultCode = RUNERR_DATAARRAY_END Then
                vnResultCode = RUN_SUCCESS
            End If
            Exit Do
        End If

        If vPrivilege = MP_ALL Then
            vstrPrivilege = "Manager"
        ElseIf vPrivilege = MP_NONE Then
            vstrPrivilege = "User"
        End If

        vstrBackupNumber = FuncStringFromBackupNumber(vBackupNumber)
    
        If vnEnableFlag = 1 Then
            vstrEnableFlag = "E"
        Else
            vstrEnableFlag = "D"
        End If

        lstEnrollData.AddItem (Format(vnii, "000#") & "    " & _
                               Format(vEnrollNumber, "0000000#") & "      " & _
                               vstrBackupNumber & "     " & _
                               vstrPrivilege & "       " & _
                               vstrEnableFlag)

        vnii = vnii + 1
        lblTotal.Caption = "Total : " & vnii
        DoEvents
    Loop
    
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "GetEnrollInfo OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If
    
    FK_EnableDevice mnCommHandleIndex, 1
    cmdGetEnrollInfo.Enabled = True
End Sub

Private Sub FuncSetUserEnableStatus(abEnableFlag As Boolean)
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vnResultCode As Long

    lblMessage.Caption = "Working..."
    DoEvents

    vEnrollNumber = Val(txtEnrollNumber.Text)
    vBackupNumber = FuncGetBackupNumberFromItem
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        Exit Sub
    End If

    vnResultCode = FK_EnableUser(mnCommHandleIndex, _
                                vEnrollNumber, vBackupNumber, abEnableFlag)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice mnCommHandleIndex, 1
End Sub

Private Sub cmdEnableUser_Click()
    cmdEnableUser.Enabled = False
    FuncSetUserEnableStatus True
    cmdEnableUser.Enabled = True
End Sub

Private Sub cmdDisableUser_Click()
    cmdDisableUser.Enabled = False
    FuncSetUserEnableStatus False
    cmdDisableUser.Enabled = True
End Sub

Private Sub cmdModifyPrivilege_Click()
    Dim vEnrollNumber As Long
    Dim vBackupNumber As Long
    Dim vPrivilege As Long
    Dim vnResultCode As Long
    
    cmdModifyPrivilege.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vEnrollNumber = Val(txtEnrollNumber.Text)
    vBackupNumber = FuncGetBackupNumberFromItem
    If cmbPrivilege.ListIndex = 0 Then
        vPrivilege = MP_NONE
    ElseIf cmbPrivilege.ListIndex = 1 Then
        vPrivilege = MP_ALL
    End If
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdModifyPrivilege.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_ModifyPrivilege(mnCommHandleIndex, _
                                            vEnrollNumber, _
                                            vBackupNumber, _
                                            vPrivilege)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice mnCommHandleIndex, 1
    cmdModifyPrivilege.Enabled = True
End Sub

Private Sub cmdBenumbManager_Click()
    Dim vnResultCode As Long
    
    cmdBenumbManager.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdBenumbManager.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_BenumbAllManager(mnCommHandleIndex)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    FK_EnableDevice mnCommHandleIndex, 1
    cmdBenumbManager.Enabled = True
End Sub

Private Sub cmdEmptyEnrollData_Click()
    Dim vnResultCode As Long

    cmdEmptyEnrollData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents

    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdEmptyEnrollData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_EmptyEnrollData(mnCommHandleIndex)
    lblMessage.Caption = ReturnResultPrint(vnResultCode)

    cmdEmptyEnrollData.Enabled = True
    FK_EnableDevice mnCommHandleIndex, 1
End Sub

Private Sub cmdClearData_Click()
    Dim vnResultCode As Long
   
    cmdClearData.Enabled = False
    lblMessage.Caption = "Working..."
    DoEvents
    
    vnResultCode = FK_EnableDevice(mnCommHandleIndex, 0)
    If vnResultCode <> RUN_SUCCESS Then
        lblMessage.Caption = gstrNoDevice
        cmdClearData.Enabled = True
        Exit Sub
    End If

    vnResultCode = FK_ClearKeeperData(mnCommHandleIndex)
    If vnResultCode = RUN_SUCCESS Then
        lblMessage.Caption = "ClearKeeperData OK"
    Else
        lblMessage.Caption = ReturnResultPrint(vnResultCode)
    End If

    FK_EnableDevice mnCommHandleIndex, 1
    cmdClearData.Enabled = True
End Sub

Private Sub DBWithItemEnable(abEnableFlag As Boolean)
    AdoEnroll.Enabled = abEnableFlag
    cmdDel.Enabled = abEnableFlag
    cmdGetAllEnrollData.Enabled = abEnableFlag
    cmdSetAllEnrollData.Enabled = abEnableFlag
    cmdUSBGetAllData.Enabled = abEnableFlag
    cmdUSBSetAllData.Enabled = abEnableFlag
End Sub

Private Sub DisableButtons()
    cmdGetEnrollData.Enabled = False
    cmdSetEnrollData.Enabled = False
    cmdDeleteEnrollData.Enabled = False
    cmdGetAllEnrollData.Enabled = False
    cmdSetAllEnrollData.Enabled = False
    cmdGetEnrollInfo.Enabled = False
    cmdEnableUser.Enabled = False
    cmdDisableUser.Enabled = False
    cmdModifyPrivilege.Enabled = False
    cmdBenumbManager.Enabled = False
    cmdClearData.Enabled = False
    cmdEmptyEnrollData.Enabled = False
End Sub

Private Sub ShowByteArrayToListBox(aListBox As ListBox, aByteArray() As Byte, ByVal aLenToShow As Integer)
    Dim k As Long, n As Long, i As Long
    Dim strHex As String, strLine As String
    
    On Error GoTo errL_ShowByteArrayToListBox
    
    aListBox.Clear
    
    If aLenToShow > UBound(aByteArray) + 1 Then aLenToShow = UBound(aByteArray) + 1
    
    For k = 0 To (aLenToShow \ 8 + 1)
        strLine = ""
        For i = 0 To 7
            n = k * 8 + i
            If n >= aLenToShow Then Exit For
                
            strHex = Hex(aByteArray(n))
            If Len(strHex) = 1 Then strHex = "0" & strHex
            strLine = strLine & strHex & " "
        Next i
        aListBox.AddItem (strLine)
        If n >= aLenToShow Then Exit For
    Next
    Exit Sub
     
errL_ShowByteArrayToListBox:
    aListBox.AddItem ("Error to show data")
    Exit Sub
End Sub

Private Sub FuncDispToListBox(anBackupNumber As Long)
    Dim k As Long, n As Long, i As Long
    Dim vnLen As Long
    Dim strHex As String, strLine As String
    
    lstEnrollData.Clear
    lblEnrollData.Caption = "Enrolled Data :"
    lblTotal.Caption = ""
    
    If anBackupNumber = BACKUP_PSW Or anBackupNumber = BACKUP_CARD Then
        ShowByteArrayToListBox lstEnrollData, mbytEnrollData, PWD_DATA_SIZE
    ElseIf anBackupNumber >= BACKUP_FP_0 And anBackupNumber <= BACKUP_FP_9 Then
        ShowByteArrayToListBox lstEnrollData, mbytEnrollData, FP_DATA_SIZE
    ElseIf anBackupNumber = BACKUP_FACE Then
        ShowByteArrayToListBox lstEnrollData, mbytEnrollData, FACE_DATA_SIZE
    ElseIf anBackupNumber = BACKUP_VEIN_0 Then
        ShowByteArrayToListBox lstEnrollData, mbytEnrollData, VEIN_DATA_SIZE
    End If
End Sub

Private Sub ConvFpDataToSaveInDbForCompatibility(ByRef abytSrc() As Byte, ByRef abytDest() As Byte)
    Dim nTempLen As Long, lenConvFpData As Long
    Dim bytConvFpData() As Byte
    Dim k As Long, m As Long
    Dim bytTemp() As Byte

    nTempLen = (UBound(abytSrc) - LBound(abytSrc) + 1) / 4
    lenConvFpData = nTempLen * 5
    ReDim bytConvFpData(lenConvFpData)

    For k = 0 To nTempLen - 1
        ReDim bytTemp(3)
        bytTemp(0) = abytSrc(k * 4)
        bytTemp(1) = abytSrc(k * 4 + 1)
        bytTemp(2) = abytSrc(k * 4 + 2)
        bytTemp(3) = abytSrc(k * 4 + 3)

        'm = BitConverter.ToInt32(bytTemp, 0)
        CopyMemory m, bytTemp(0), 4
        bytConvFpData(k * 5) = 1
        If m < 0 Then
            If m = -2147483648# Then
                bytConvFpData(k * 5) = 2
                m = 2147483647
            Else
                bytConvFpData(k * 5) = 0
                m = -m
            End If
        End If
        'bytTemp = BitConverter.GetBytes(m)
        CopyMemory bytTemp(0), m, 4
        bytConvFpData(k * 5 + 1) = bytTemp(3)
        bytConvFpData(k * 5 + 2) = bytTemp(2)
        bytConvFpData(k * 5 + 3) = bytTemp(1)
        bytConvFpData(k * 5 + 4) = bytTemp(0)
    Next

    abytDest = bytConvFpData
End Sub

Private Sub ConvFpDataAfterReadFromDbForCompatibility(ByRef abytSrc() As Byte, ByRef abytDest() As Byte)
    Dim nTempLen As Long, lenConvFpData As Long
    Dim bytConvFpData() As Byte
    Dim k As Long, m As Long
    Dim bytTemp() As Byte

    nTempLen = (UBound(abytSrc) - LBound(abytSrc) + 1) / 5
    lenConvFpData = nTempLen * 4

    If lenConvFpData < FP_DATA_SIZE Then lenConvFpData = FP_DATA_SIZE
    ReDim bytConvFpData(lenConvFpData)

    For k = 0 To nTempLen - 1
        ReDim bytTemp(3)
        bytTemp(0) = abytSrc(k * 5 + 4)
        bytTemp(1) = abytSrc(k * 5 + 3)
        bytTemp(2) = abytSrc(k * 5 + 2)
        bytTemp(3) = abytSrc(k * 5 + 1)

        'm = BitConverter.ToInt32(bytTemp, 0)
        CopyMemory m, bytTemp(0), 4
        If abytSrc(k * 5) = 0 Then
            m = -m
        ElseIf abytSrc(k * 5) = 2 Then
            m = -2147483648#
        End If
        'bytTemp = BitConverter.GetBytes(m)
        CopyMemory bytTemp(0), m, 4

        bytConvFpData(k * 4 + 3) = bytTemp(3)
        bytConvFpData(k * 4 + 2) = bytTemp(2)
        bytConvFpData(k * 4 + 1) = bytTemp(1)
        bytConvFpData(k * 4 + 0) = bytTemp(0)
    Next

    abytDest = bytConvFpData
End Sub
    
Private Sub FuncSaveToDB(anEnrollNumber As Long, anBackupNumber As Long, anPrivilege As Long, anEnrollName As String)
    Dim vnii As Long
    Dim vnLen As Long
    Dim vbytEnrollData() As Byte
    Dim vbytConvFpData() As Byte

    With AdoEnroll.Recordset
        AdoFind.RecordSource = "select * from tblEnroll where EnrollNumber=" & CStr(anEnrollNumber) & _
              " and FingerNumber=" & CStr(anBackupNumber)
        AdoFind.Refresh
        If AdoFind.Recordset.RecordCount > 0 Then
            lblMessage.Caption = "Double ID : " & Format(anEnrollNumber, "0000#") & "-" & anBackupNumber
            lstEnrollData.AddItem (lblMessage.Caption)
        Else
            .AddNew
            !EMachineNumber = 0
            !EnrollNumber = anEnrollNumber
            !FingerNumber = anBackupNumber
            !Privilige = anPrivilege
            !EnrollName = Trim(anEnrollName)
            
            If anBackupNumber = BACKUP_PSW Or anBackupNumber = BACKUP_CARD Then
                ReDim vbytEnrollData(PWD_DATA_SIZE - 1)
                CopyMemory vbytEnrollData(0), mbytEnrollData(0), PWD_DATA_SIZE
                !FPdata = vbytEnrollData
            ElseIf anBackupNumber >= BACKUP_FP_0 And anBackupNumber <= BACKUP_FP_9 Then
                ReDim vbytEnrollData(FP_DATA_SIZE - 1)
                CopyMemory vbytEnrollData(0), mbytEnrollData(0), FP_DATA_SIZE
                ConvFpDataToSaveInDbForCompatibility vbytEnrollData, vbytConvFpData
                !FPdata = vbytConvFpData
            ElseIf anBackupNumber = BACKUP_FACE Then
                ReDim vbytEnrollData(FACE_DATA_SIZE - 1)
                CopyMemory vbytEnrollData(0), mbytEnrollData(0), FACE_DATA_SIZE
                !FPdata = vbytEnrollData
            ElseIf anBackupNumber = BACKUP_VEIN_0 Then
                ReDim vbytEnrollData(VEIN_DATA_SIZE - 1)
                CopyMemory vbytEnrollData(0), mbytEnrollData(0), VEIN_DATA_SIZE
                !FPdata = vbytEnrollData
            End If
            .Update
            
            lblMessage.Caption = Format(anEnrollNumber, "0000#") & "-" & anBackupNumber
            txtEnrollNumber.Text = Trim(Str(anEnrollNumber))
            cmbBackupNumber.ListIndex = FuncItemIndexFromBackupNumber(anBackupNumber)
            If anPrivilege = MP_NONE Then
                cmbPrivilege.ListIndex = 0
            ElseIf anPrivilege = MP_ALL Then
                cmbPrivilege.ListIndex = 1
            End If
        End If
    End With
End Sub

Private Function FuncReadFromToDB(anEnrollNumber As Long, anBackupNumber As Long, anPrivilege As Long, anEnrollName As String, Optional abdispFlag As Boolean = True) As Boolean
    Dim vbytConvFpData() As Byte
    Dim vbytEnrollData() As Byte
    
    FuncReadFromToDB = False
    With AdoEnroll.Recordset
        If .RecordCount <= 0 Then Exit Function
        If .AbsolutePosition <= 0 Then Exit Function
        If !EnrollNumber <= 0 Then Exit Function
        anEnrollNumber = !EnrollNumber
        txtEnrollNumber.Text = Trim(Str(anEnrollNumber))
        anBackupNumber = !FingerNumber
        cmbBackupNumber.ListIndex = FuncItemIndexFromBackupNumber(anBackupNumber)
        anPrivilege = !Privilige
        
        If IsNull(!EnrollName) Then
            anEnrollName = ""
        Else
            anEnrollName = !EnrollName
        End If
        
        If anPrivilege = MP_NONE Then
            cmbPrivilege.ListIndex = 0
        ElseIf anPrivilege = MP_ALL Then
            cmbPrivilege.ListIndex = 1
        End If
        
        ZeroMemory mbytEnrollData(0), UBound(mbytEnrollData) + 1
        If anBackupNumber = BACKUP_PSW Or anBackupNumber = BACKUP_CARD Then
            vbytEnrollData = !FPdata
            CopyMemory mbytEnrollData(0), vbytEnrollData(0), PWD_DATA_SIZE
        ElseIf anBackupNumber >= BACKUP_FP_0 And anBackupNumber <= BACKUP_FP_9 Then
            vbytConvFpData = !FPdata
            ConvFpDataAfterReadFromDbForCompatibility vbytConvFpData, vbytEnrollData
            CopyMemory mbytEnrollData(0), vbytEnrollData(0), FP_DATA_SIZE
        ElseIf anBackupNumber = BACKUP_FACE Then
            vbytEnrollData = !FPdata
            CopyMemory mbytEnrollData(0), vbytEnrollData(0), FACE_DATA_SIZE
        ElseIf anBackupNumber = BACKUP_VEIN_0 Then
            vbytEnrollData = !FPdata
            CopyMemory mbytEnrollData(0), vbytEnrollData(0), VEIN_DATA_SIZE
        End If
        
        If abdispFlag = True Then
            FuncDispToListBox anBackupNumber
        End If
        FuncReadFromToDB = True
    End With
End Function

Private Function FuncGetBackupNumberFromItem() As Long
    Dim vnIndex As Long

    vnIndex = cmbBackupNumber.ListIndex
    If vnIndex < 0 Then vnIndex = 0
    FuncGetBackupNumberFromItem = ftBackupNumber(vnIndex * 2)
End Function

Private Function FuncItemIndexFromBackupNumber(ByVal anBackupNumber As Long) As Long
    Dim vnii As Long

    FuncItemIndexFromBackupNumber = -1
    For vnii = 0 To UBound(ftBackupNumber) Step 2
        If ftBackupNumber(vnii) = anBackupNumber Then
            FuncItemIndexFromBackupNumber = vnii / 2
            Exit For
        End If
    Next vnii
End Function

Private Function FuncStringFromBackupNumber(ByVal anBackupNumber As Long) As String
    Dim vnii As Long

    FuncStringFromBackupNumber = "        "
    For vnii = 0 To UBound(ftBackupNumber) Step 2
        If ftBackupNumber(vnii) = anBackupNumber Then
            FuncStringFromBackupNumber = ftBackupNumber(vnii + 1)
            If anBackupNumber <> BACKUP_PSW And anBackupNumber <> BACKUP_CARD Then
                FuncStringFromBackupNumber = FuncStringFromBackupNumber & " "
            End If
            Exit For
        End If
    Next vnii
End Function

