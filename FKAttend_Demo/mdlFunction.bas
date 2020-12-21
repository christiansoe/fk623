Attribute VB_Name = "mdlPublic"
' ===============================================================================
' Win32 API Functions
' ===============================================================================
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)


' ===============================================================================
' FingerKeeper Interface Functions
' ===============================================================================

'// Connection
Public Declare Function FK_ConnectComm Lib "FKAttend" (ByVal nMachineNo As Long, ByVal nComPort As Long, ByVal nBaudRate As Long, ByVal pstrTelNumber As String, ByVal nWaitDialTime As Long, ByVal nLicense As Long, ByVal nComTimeOut As Long) As Long
Public Declare Function FK_ConnectNet Lib "FKAttend" (ByVal nMachineNo As Long, ByVal strIpAddress As String, ByVal nNetPort As Long, ByVal nTimeOut As Long, ByVal nProtocolType As Long, ByVal nNetPassword As Long, ByVal nLicense As Long) As Long
Public Declare Function FK_ConnectUSB Lib "FKAttend" (ByVal nMachineNo As Long, ByVal nLicense As Long) As Long
Public Declare Sub FK_DisConnect Lib "FKAttend" (ByVal nHandleIndex As Long)

'// Device Setting
Public Declare Function FK_EnableDevice Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnableFlag As Byte) As Long
Public Declare Sub FK_PowerOnAllDevice Lib "FKAttend" (ByVal nHandleIndex As Long)
Public Declare Function FK_PowerOffDevice Lib "FKAttend" (ByVal nHandleIndex As Long) As Long
Public Declare Function FK_GetDeviceStatus Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nStatusIndex As Long, ByRef pnValue As Long) As Long
Public Declare Function FK_GetDeviceTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnDateTime As Date) As Long
Public Declare Function FK_SetDeviceTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nDateTime As Date) As Long
Public Declare Function FK_GetDeviceInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nInfoIndex As Long, ByRef pnValue As Long) As Long
Public Declare Function FK_SetDeviceInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nInfoIndex As Long, ByVal nValue As Long) As Long
Public Declare Function FK_GetProductData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nDataIndex As Long, ByRef pstrValue As String) As Long

'// Log Data
Public Declare Function FK_LoadSuperLogData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nReadMark As Long) As Long
Public Declare Function FK_USBLoadSuperLogDataFromFile Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal astrFilePath As String) As Long
Public Declare Function FK_GetSuperLogData Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnSEnrollNumber As Long, ByRef pnGEnrollNumber As Long, ByRef nManipulation As Long, ByRef pnBackupNumber As Long, ByRef pnDateTime As Date) As Long
Public Declare Function FK_GetSuperLogData_1 Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnSEnrollNumber As Long, ByRef pnGEnrollNumber As Long, ByRef nManipulation As Long, ByRef pnBackupNumber As Long, ByRef apnYear As Long, ByRef apnMonth As Long, ByRef apnDay As Long, ByRef apnHour As Long, ByRef apnMinute As Long, ByRef apnSec As Long) As Long
Public Declare Function FK_EmptySuperLogData Lib "FKAttend" (ByVal nHandleIndex As Long) As Long

Public Declare Function FK_LoadGeneralLogData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nReadMark As Long) As Long
Public Declare Function FK_USBLoadGeneralLogDataFromFile Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal astrFilePath As String) As Long
Public Declare Function FK_GetGeneralLogData Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnEnrollNumber As Long, ByRef pnVerifyMode As Long, ByRef pnInOutMode As Long, ByRef pnDateTime As Date) As Long
Public Declare Function FK_GetGeneralLogData_1 Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnEnrollNumber As Long, ByRef pnVerifyMode As Long, ByRef pnInOutMode As Long, ByRef apnYear As Long, ByRef apnMonth As Long, ByRef apnDay As Long, ByRef apnHour As Long, ByRef apnMinute As Long, ByRef apnSec As Long) As Long
Public Declare Function FK_EmptyGeneralLogData Lib "FKAttend" (ByVal nHandleIndex As Long) As Long

'// Bell Time
Public Declare Function FK_GetBellTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnBellCount As Long, ByRef ptBellInfo As Any) As Long
Public Declare Function FK_SetBellTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nBellCount As Long, ByRef ptBellInfo As Any) As Long

'// Enroll Data
Public Declare Function FK_GetEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByRef pnMachinePrivilege As Long, ByRef pnEnrollData As Any, ByRef pnPassWord As Long) As Long
Public Declare Function FK_PutEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByVal nMachinePrivilege As Long, ByRef pnEnrollData As Any, ByVal nPassWord As Long) As Long
Public Declare Function FK_SaveEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long) As Long
Public Declare Function FK_DeleteEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long) As Long

Public Declare Function FK_SetUSBModel Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anModel As Long) As Long
Public Declare Function FK_SetUDiskFileFKModel Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal astrFKModel As String) As Long

Public Declare Function FK_USBReadAllEnrollDataFromFile Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal pstrFilePath As String) As Long
Public Declare Function FK_USBReadAllEnrollDataCount Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal pnValue As Long) As Long
Public Declare Function FK_USBGetOneEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnEnrollNumber As Long, ByRef pnBackupNumber As Long, ByRef pnMachinePrivilege As Long, ByRef pnEnrollData As Any, ByRef pnPassWord As Long, ByRef pnEnableFlag As Long, ByRef pnEnrollName As String) As Long
Public Declare Function FK_USBSetOneEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByVal nMachinePrivilege As Long, ByRef pnEnrollData As Any, ByVal nPassWord As Long, ByVal nEnableFlag As Long, ByVal pnEnrollName As String) As Long
Public Declare Function FK_USBWriteAllEnrollDataToFile Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal pstrFilePath As String) As Long

Public Declare Function FK_USBReadAllEnrollDataFromFile_Color Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal pstrFilePath As String) As Long
Public Declare Function FK_USBWriteAllEnrollDataToFile_Color Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal pstrFilePath As String, ByVal anNewsKind As Long) As Long
Public Declare Function FK_USBGetOneEnrollData_Color Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnEnrollNumber As Long, ByRef pnBackupNumber As Long, ByRef pnMachinePrivilege As Long, ByRef pnEnrollData As Any, ByRef pnPassWord As Long, ByRef pnEnableFlag As Long, ByRef pnEnrollName As String, ByVal anNewsKind As Long) As Long
Public Declare Function FK_USBSetOneEnrollData_Color Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByVal nMachinePrivilege As Long, ByRef pnEnrollData As Any, ByVal nPassWord As Long, ByVal nEnableFlag As Long, ByVal pnEnrollName As String, ByVal anNewsKind As Long) As Long

Public Declare Function FK_EnableUser Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByVal nEnableFlag As Long) As Long
Public Declare Function FK_ModifyPrivilege Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nBackupNumber As Long, ByVal nMachinePrivilege As Long) As Long
Public Declare Function FK_BenumbAllManager Lib "FKAttend" (ByVal nHandleIndex As Long) As Long
Public Declare Function FK_ReadAllUserID Lib "FKAttend" (ByVal nHandleIndex As Long) As Long
Public Declare Function FK_GetAllUserID Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef pnEnrollNumber As Long, ByRef pnBackupNumber As Long, ByRef pnMachinePrivilege As Long, ByRef pnEnable As Long) As Long
Public Declare Function FK_EmptyEnrollData Lib "FKAttend" (ByVal nHandleIndex As Long) As Long
Public Declare Function FK_ClearKeeperData Lib "FKAttend" (ByVal nHandleIndex As Long) As Long

Public Declare Function FK_SetFontName Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal aStrFontName As String, ByVal anFontType As Long) As Long

Public Declare Function FK_GetUserName Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByRef pstrUserName As String) As Long
Public Declare Function FK_SetUserName Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal pstrUserName As String) As Long
Public Declare Function FK_GetNewsMessage Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nNewsId As Long, ByRef pstrNews As String) As Long
Public Declare Function FK_SetNewsMessage Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nNewsId As Long, ByVal pstrNews As String) As Long
Public Declare Function FK_GetUserNewsID Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByRef pnNewsId As Long) As Long
Public Declare Function FK_SetUserNewsID Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal nEnrollNumber As Long, ByVal nNewsId As Long) As Long

'// Access Control
Public Declare Function FK_GetDoorStatus Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef apnStatusVal As Long) As Long
Public Declare Function FK_SetDoorStatus Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anStatusVal As Long) As Long
Public Declare Function FK_GetPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anPassTimeID As Long, ByRef apnPassTime As Any, ByVal anPassTimeSize As Long) As Long
Public Declare Function FK_SetPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anPassTimeID As Long, ByRef apnPassTime As Any, ByVal anPassTimeSize As Long) As Long
Public Declare Function FK_GetUserPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anEnrollNumber As Long, ByRef apnGroupID As Long, ByRef apnPassTimeID As Any, ByVal anPassTimeIDSize As Long) As Long
Public Declare Function FK_SetUserPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anEnrollNumber As Long, ByVal anGroupID As Long, ByRef apnPassTimeID As Any, ByVal anPassTimeIDSize As Long) As Long
Public Declare Function FK_GetGroupPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anGroupID As Long, ByRef apnPassTimeID As Any, ByVal anPassTimeIDSize As Long) As Long
Public Declare Function FK_SetGroupPassTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anGroupID As Long, ByRef apnPassTimeID As Any, ByVal anPassTimeIDSize As Long) As Long
Public Declare Function FK_GetGroupMatch Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef apnGroupMatch As Any, ByVal anGroupMatchSize As Long) As Long
Public Declare Function FK_SetGroupMatch Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef apnGroupMatch As Any, ByVal anGroupMatchSize As Long) As Long

'// Etc Functions
Public Declare Function FK_ConnectGetIP Lib "FKAttend" (ByRef apnComName As Any) As Long
Public Declare Function FK_GetAdjustInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef dwAdjustedState As Long, ByRef dwAdjustedMonth As Long, ByRef dwAdjustedDay As Long, ByRef dwAdjustedHour As Long, ByRef dwAdjustedMinute As Long, ByRef dwRestoredState As Long, ByRef dwRestoredMonth As Long, ByRef dwRestoredDay As Long, ByRef dwRestoredHour As Long, ByRef dwRestoredMinte As Long) As Long
Public Declare Function FK_SetAdjustInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal dwAdjustedState As Long, ByVal dwAdjustedMonth As Long, ByVal dwAdjustedDay As Long, ByVal dwAdjustedHour As Long, ByVal dwAdjustedMinute As Long, ByVal dwRestoredState As Long, ByVal dwRestoredMonth As Long, ByVal dwRestoredDay As Long, ByVal dwRestoredHour As Long, ByVal dwRestoredMinte As Long) As Long
Public Declare Function FK_GetAccessTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anEnrollNumber As Long, ByRef apnAccessTime As Long) As Long
Public Declare Function FK_SetAccessTime Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal anEnrollNumber As Long, ByVal anAccessTime As Long) As Long
Public Declare Function FK_GetRealTimeInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef apGetRealTime As Long) As Long
Public Declare Function FK_SetRealTimeInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef apSetRealTime As Long) As Long
Public Declare Function FK_GetServerNetInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByRef astrServerIPAddress As String, ByRef apServerPort As Long, ByRef apServerRequest As Long) As Long
Public Declare Function FK_SetServerNetInfo Lib "FKAttend" (ByVal nHandleIndex As Long, ByVal astrServerIPAddress As String, ByVal anServerPort As Long, ByVal apServerReques As Long) As Long


' ===============================================================================
' FingerKeeper Interface Constants & Structures
' ===============================================================================

Public Const CHINA_FONTNAME = "Arial"
Public Const JAPAN_FONTNAME = "MS PGothic"
Public Const THAI_FONTNAME = "AngsanaUPC"

'--- FK_SetUSBModel Parameters ---
Public Const FK625_FP1000 = 2001
Public Const FK625_FP2000 = 2002
Public Const FK625_FP3000 = 2003
Public Const FK625_FP5000 = 2004
Public Const FK625_FP10000 = 2005
Public Const FK625_FP30000 = 2006
Public Const FK625_ID30000 = 2007
Public Const FK635_FP700 = 3001
Public Const FK635_FP3000 = 3002
Public Const FK635_FP10000 = 3003
Public Const FK635_ID30000 = 3004
Public Const FK723_FP1000 = 4001
Public Const FK725_FP1000 = 5001
Public Const FK725_FP1500 = 5002
Public Const FK725_ID5000 = 5003
Public Const FK725_ID30000 = 5004
Public Const FK735_FP500 = 6001
Public Const FK735_FP3000 = 6002
Public Const FK735_ID30000 = 6003
Public Const FK925_FP3000 = 7001
Public Const FK935_FP3000 = 8001

'--- Bell Time ---
Public Const MAX_BELLCOUNT_DAY = 24
Type BELLINFO
    mValid(MAX_BELLCOUNT_DAY - 1) As Byte
    mHour(MAX_BELLCOUNT_DAY - 1) As Byte
    mMinute(MAX_BELLCOUNT_DAY - 1) As Byte
End Type             '72byte

Public Const MAX_PASSCTRLGROUP_COUNT = 50
Public Const MAX_PASSCTRL_COUNT = 7
'--- Pass Control Time ---
Type PASSTIME
    StartHour  As Byte
    StartMinute  As Byte
    EndHour  As Byte
    EndMinute  As Byte
End Type          '4byte

'--- Pass Control Time Infomation ---
Type PASSCTRLTIME
    mPassCtrlTime(MAX_PASSCTRL_COUNT - 1) As PASSTIME
End Type          '28byte

Public Const MAX_USERPASSINFO_COUNT = 3
Type USERPASSINFO
    UserPassID(MAX_USERPASSINFO_COUNT - 1) As Byte
End Type          '3byte

Public Const MAX_GROUPPASSKIND_COUNT = 5
Public Const MAX_GROUPPASSINFO_COUNT = 3
Type GROUPPASSINFO
    GroupPassID(MAX_GROUPPASSINFO_COUNT - 1) As Byte
End Type          '3byte

Public Const MAX_GROUPMATCHINFO_COUNT = 10
Type GROUPMATCHINFO
    GroupMatch(MAX_GROUPMATCHINFO_COUNT - 1) As Integer
End Type          '20byte


'--- Realtime Setting ---
Public Const MAX_REAL_TIME = 4
Type REALTIMEINFO
    Valid As Byte
    AckTime As Byte
    WaitTime As Byte
    Reserve As Byte
    SendPos As Long
    Hour(MAX_REAL_TIME - 1) As Byte
    Minute(MAX_REAL_TIME - 1) As Byte
End Type     ' 16 Byte


Public Const NEWS_EXTEND = 2
Public Const NEWS_STANDARD = 1


'//=============== Protocol Type ===============//
Public Const PROTOCOL_TCPIP = 0               ' TCP/IP
Public Const PROTOCOL_UDP = 1                 ' UDP

'//=============== Backup Number Constant ===============//
Public Const BACKUP_FP_0 = 0                  ' Finger 0
Public Const BACKUP_FP_1 = 1                  ' Finger 1
Public Const BACKUP_FP_2 = 2                  ' Finger 2
Public Const BACKUP_FP_3 = 3                  ' Finger 3
Public Const BACKUP_FP_4 = 4                  ' Finger 4
Public Const BACKUP_FP_5 = 5                  ' Finger 5
Public Const BACKUP_FP_6 = 6                  ' Finger 6
Public Const BACKUP_FP_7 = 7                  ' Finger 7
Public Const BACKUP_FP_8 = 8                  ' Finger 8
Public Const BACKUP_FP_9 = 9                  ' Finger 9
Public Const BACKUP_PSW = 10                  ' Password
Public Const BACKUP_CARD = 11                 ' Card
Public Const BACKUP_FACE = 12                 ' Face
Public Const BACKUP_VEIN_0 = 20               ' Vein 0

'//=============== Manipulation of SuperLogData ===============//
Public Const LOG_ENROLL_USER = 3              ' Enroll-User
Public Const LOG_ENROLL_MANAGER = 4           ' Enroll-Manager
Public Const LOG_ENROLL_DELFP = 5             ' FP Delete
Public Const LOG_ENROLL_DELPASS = 6           ' Pass Delete
Public Const LOG_ENROLL_DELCARD = 7           ' Card Delete
Public Const LOG_LOG_ALLDEL = 8               ' LogAll Delete
Public Const LOG_SETUP_SYS = 9                ' Setup Sys
Public Const LOG_SETUP_TIME = 10              ' Setup Time
Public Const LOG_SETUP_LOG = 11               ' Setup Log
Public Const LOG_SETUP_COMM = 12              ' Setup Comm
Public Const LOG_PASSTIME = 13                ' Pass Time Set
Public Const LOG_SETUP_DOOR = 14              ' Door Set Log

'//=============== VerifyMode of GeneralLogData ===============//
Public Const LOG_FPVERIFY = 1                 ' Fp Verify
Public Const LOG_PASSVERIFY = 2               ' Pass Verify
Public Const LOG_CARDVERIFY = 3               ' Card Verify
Public Const LOG_FPPASS_VERIFY = 4            ' Pass+Fp Verify
Public Const LOG_FPCARD_VERIFY = 5            ' Card+Fp Verify
Public Const LOG_PASSFP_VERIFY = 6            ' Pass+Fp Verify
Public Const LOG_CARDFP_VERIFY = 7            ' Card+Fp Verify
Public Const LOG_JOB_NO_VERIFY = 8            ' Job number Verify
Public Const LOG_CARDPASS_VERIFY = 9          ' Card+Pass Verify
Public Const LOG_CLOSE_DOOR = 10              ' Door Close
Public Const LOG_OPEN_HAND = 11               ' Hand Open
Public Const LOG_PROG_OPEN = 12               ' Open by PC
Public Const LOG_PROG_CLOSE = 13              ' Close by PC
Public Const LOG_OPEN_IREGAL = 14             ' Iregal Open
Public Const LOG_CLOSE_IREGAL = 15            ' Iregal Close
Public Const LOG_OPEN_COVER = 16              ' Cover Open
Public Const LOG_CLOSE_COVER = 17             ' Cover Close

Public Const LOG_OPEN_DOOR = 32               ' Door Open
Public Const LOG_OPEN_THREAT = 48             ' Door Open as threat
'//=============== IOMode of GeneralLogData ===============//
'Public Const LOG_IOMODE_IN = 0
'Public Const LOG_IOMODE_OUT = 1
'Public Const LOG_IOMODE_OVER_IN = 2    ' = LOG_IOMODE_IO
'Public Const LOG_IOMODE_OVER_OUT = 3

Public Const LOG_MODE_IO = 0 'General
Public Const LOG_MODE_IN1 = 1 'IN1
Public Const LOG_MODE_IN2 = 2 'IN2
Public Const LOG_MODE_IN3 = 3 'IN3
Public Const LOG_MODE_OUT1 = 4 'OUT1
Public Const LOG_MODE_OUT2 = 5 'OUT2
Public Const LOG_MODE_OUT3 = 6 'OUT3


'//=============== Machine Privilege ===============//
Public Const MP_NONE = 0                      ' General user
Public Const MP_ALL = 1                       ' Manager

'//=============== Index of  GetDeviceStatus ===============//
Public Const GET_MANAGERS = 1
Public Const GET_USERS = 2
Public Const GET_FPS = 3
Public Const GET_PSWS = 4
Public Const GET_SLOGS = 5
Public Const GET_GLOGS = 6
Public Const GET_ASLOGS = 7
Public Const GET_AGLOGS = 8
Public Const GET_CARDS = 9

'//=============== Index of  GetDeviceInfo ===============//
Public Const DI_MANAGERS = 1                  ' Numbers of Manager
Public Const DI_MACHINENUM = 2                ' Device ID
Public Const DI_LANGAUGE = 3                  ' Language
Public Const DI_POWEROFF_TIME = 4             ' Auto-PowerOff Time
Public Const DI_LOCK_CTRL = 5                 ' Lock Control
Public Const DI_GLOG_WARNING = 6              ' General-Log Warning
Public Const DI_SLOG_WARNING = 7              ' Super-Log Warning
Public Const DI_VERIFY_INTERVALS = 8          ' Verify Interval Time
Public Const DI_RSCOM_BPS = 9                 ' Comm Buadrate
Public Const DI_DATE_SEPARATE = 10            ' Date Separate Symbol
Public Const DI_VERIFY_KIND = 24              ' Verify Kind Symbol
Public Const DI_MULTIUSERS = 77               ' MultiUser
Public Const DI_NETENABLE = 14                ' Network Enable
Public Const DI_ALARMDELAY = 66               ' Alarm Output Delay Time
Public Const DI_SENSORDELAY = 67              ' Sensor Output Delay Time

'//=============== Baudrate = value of DI_RSCOM_BPS ===============//
Public Const BPS_9600 = 3
Public Const BPS_19200 = 4
Public Const BPS_38400 = 5
Public Const BPS_57600 = 6
Public Const BPS_115200 = 7

'//=============== Product Data Index ===============//
Public Const PRODUCT_SERIALNUMBER = 1    ' Serial Number
Public Const PRODUCT_BACKUPNUMBER = 2    ' Backup Number
Public Const PRODUCT_CODE = 3            ' Product code
Public Const PRODUCT_NAME = 4            ' Product name
Public Const PRODUCT_WEB = 5             ' Product web
Public Const PRODUCT_DATE = 6            ' Product date
Public Const PRODUCT_SENDTO = 7          ' Product sendto

'//=============== Door Status ===============//
Public Const DOOR_CONROLRESET = 0
Public Const DOOR_OPEND = 1
Public Const DOOR_CLOSED = 2
Public Const DOOR_COMMNAD = 3

'//=============== Error code ===============//
Public Const RUN_SUCCESS = 1
Public Const RUNERR_NOSUPPORT = 0
Public Const RUNERR_UNKNOWNERROR = -1
Public Const RUNERR_NO_OPEN_COMM = -2
Public Const RUNERR_WRITE_FAIL = -3
Public Const RUNERR_READ_FAIL = -4
Public Const RUNERR_INVALID_PARAM = -5
Public Const RUNERR_NON_CARRYOUT = -6
Public Const RUNERR_DATAARRAY_END = -7
Public Const RUNERR_DATAARRAY_NONE = -8
Public Const RUNERR_MEMORY = -9
Public Const RUNERR_MIS_PASSWORD = -10
Public Const RUNERR_MEMORYOVER = -11
Public Const RUNERR_DATADOUBLE = -12
Public Const RUNERR_MANAGEROVER = -14
Public Const RUNERR_FPDATAVERSION = -15


' ===============================================================================
' Error processing
' ===============================================================================
Public Const gstrNoDevice = "No Device"

Function ReturnResultPrint(anResultCode As Long) As String
   Select Case anResultCode
        Case RUN_SUCCESS
            ReturnResultPrint = "Successful!"
        Case RUNERR_NOSUPPORT
            ReturnResultPrint = "No support"
        Case RUNERR_UNKNOWNERROR
            ReturnResultPrint = "Unknown error"
        Case RUNERR_NO_OPEN_COMM
            ReturnResultPrint = "No Open Comm"
        Case RUNERR_WRITE_FAIL
            ReturnResultPrint = "Write Error"
        Case RUNERR_READ_FAIL
            ReturnResultPrint = "Read Error"
        Case RUNERR_INVALID_PARAM
            ReturnResultPrint = "Parameter Error"
        Case RUNERR_NON_CARRYOUT
            ReturnResultPrint = "execution of command failed"
        Case RUNERR_DATAARRAY_END
            ReturnResultPrint = "End of data"
        Case RUNERR_DATAARRAY_NONE
            ReturnResultPrint = "Nonexistence data"
        Case RUNERR_MEMORY
            ReturnResultPrint = "Memory Allocating Error"
        Case RUNERR_MIS_PASSWORD
            ReturnResultPrint = "License Error"
        Case RUNERR_MEMORYOVER
            ReturnResultPrint = "full enrolldata & can`t put enrolldata"
        Case RUNERR_DATADOUBLE
            ReturnResultPrint = "this ID is already  existed."
        Case RUNERR_MANAGEROVER
            ReturnResultPrint = "full manager & can`t put manager."
        Case RUNERR_FPDATAVERSION
            ReturnResultPrint = "mistake fp data version."
        Case Else
            ReturnResultPrint = "Unknown error"
    End Select
End Function
