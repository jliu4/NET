Option Strict Off
Option Explicit On
Option Compare Text

Module modRegConstants

    ' Global Constants

    Public Const gstrNULL As String = "" ' Empty string

    Public Const gstrSEP_DIR As String = "\" ' Directory separator character
    Public Const gstrSEP_REGKEY As String = "\" ' Registration key separator character.
    Public Const gstrSEP_DRIVE As String = ":" ' Driver separater character, e.g., C:\
    Public Const gstrSEP_DIRALT As String = "/" ' Alternate directory separator character
    Public Const gstrSEP_EXT As String = "." ' Filename extension separator character
    Public Const gstrSEP_PROGID As String = "."
    Public Const gstrSEP_FILE As String = "|" ' Use the character for delimiting filename lists because it is not a valid character in a filename.
    Public Const gstrSEP_LIST As String = "|"
    Public Const gstrSEP_URL As String = "://" ' Separator that follows HPPT in URL address
    Public Const gstrSEP_URLDIR As String = "/" ' Separator for dividing directories in URL addresses.

    Public Const gstrUNC As String = "\\" 'UNC specifier \\
    Public Const gstrCOLON As String = ":"
    Public Const gstrSwitchPrefix1 As String = "-"
    Public Const gstrSwitchPrefix2 As String = "/"
    Public Const gstrCOMMA As String = ","
    Public Const gstrDECIMAL As String = "."
    Public Const gstrQUOTE As String = """"
    Public Const gstrCCOMMENT As String = "//" ' Comment specifier used in C, etc.
    Public Const gstrASSIGN As String = "="
    Public Const gstrINI_PROTOCOL As String = "Protocol"
    Public Const gstrREMOTEAUTO As String = "RA"
    Public Const gstrDCOM As String = "DCOM"

    Public Const gintMAX_SIZE As Short = 255 ' Maximum buffer size
    Public Const gintMAX_PATH_LEN As Short = 260 ' Maximum allowed path length including path, filename,
    ' and command line arguments for NT (Intel) and Win95.
    Public Const gintMAX_GROUPNAME_LEN As Short = 30 ' Maximum length that we allow for an NT 3.51 group name.
    Public Const gintMIN_BUTTONWIDTH As Short = 1200
    Public Const gsngBUTTON_BORDER As Single = 1.4

    Public Const intDRIVE_REMOVABLE As Short = 2 ' Constants for GetDriveType
    Public Const intDRIVE_FIXED As Short = 3
    Public Const intDRIVE_REMOTE As Short = 4
    Public Const intDRIVE_CDROM As Short = 5
    Public Const intDRIVE_RAMDISK As Short = 6

    Public Const gintNOVERINFO As Short = 32767 ' flag indicating no version info

    'File names
    Public Const gstrFILE_SETUP As String = "SETUP.LST" ' Name of setup information file
    Public Const gstrTEMP_DIR As String = "TEMP"
    Public Const gstrTMP_DIR As String = "TMP"

    'Share type macros for files
    Public Const mstrPRIVATEFILE As String = ""
    Public Const mstrSHAREDFILE As String = "$(Shared)"

    'INI File keys
    Public Const gstrINI_SETUP As String = "Setup"
    Public Const gstrINI_APPNAME As String = "Title"
    Public Const gstrINI_APPDIR As String = "DefaultDir"
    Public Const gstrINI_APPEXE As String = "AppExe"
    Public Const gstrINI_APPTOUNINSTALL As String = "AppToUninstall"
    Public Const gstrINI_APPPATH As String = "AppPath"
    Public Const gstrINI_FORCEUSEDEFDEST As String = "ForceUseDefDir"
    Public Const gstrINI_DEFGROUP As String = "DefProgramGroup"

    Public Const gstrEXT_DEP As String = "DEP"

    'Setup information file macros
    Public Const gstrAPPDEST As String = "$(AppPath)"
    Public Const gstrWINDEST As String = "$(WinPath)"
    Public Const gstrWINSYSDEST As String = "$(WinSysPath)"
    Public Const gstrWINSYSDESTSYSFILE As String = "$(WinSysPathSysFile)"
    Public Const gstrPROGRAMFILES As String = "$(ProgramFiles)"
    Public Const gstrCOMMONFILES As String = "$(CommonFiles)"
    Public Const gstrCOMMONFILESSYS As String = "$(CommonFilesSys)"
    Public Const gstrDAODEST As String = "$(MSDAOPath)"
    Public Const gstrDONOTINSTALL As String = "$(DoNotInstall)"

    'Mouse Pointer Constants
    Public Const gintMOUSE_DEFAULT As Short = 0
    Public Const gintMOUSE_HOURGLASS As Short = 11

    'MsgError() Constants
    Public Const MSGERR_ERROR As Short = 1
    Public Const MSGERR_WARNING As Short = 2

    'MsgBox Constants
    Public Const MB_OK As Short = 0 ' OK button only
    Public Const MB_OKCANCEL As Short = 1 ' OK and Cancel buttons
    Public Const MB_ABORTRETRYIGNORE As Short = 2 ' Abort, Retry, Ignore buttons
    Public Const MB_YESNO As Short = 4 ' Yes and No buttons
    Public Const MB_RETRYCANCEL As Short = 5 ' Retry and Cancel buttons
    Public Const MB_ICONSTOP As Short = 16 ' Critical message
    Public Const MB_ICONQUESTION As Short = 32 ' Warning query
    Public Const MB_ICONEXCLAMATION As Short = 48 ' Warning message
    Public Const MB_ICONINFORMATION As Short = 64 ' Information message
    Public Const MB_DEFBUTTON1 As Short = 0 ' First button is default
    Public Const MB_DEFBUTTON2 As Short = 256 ' Second button is default
    Public Const MB_DEFBUTTON3 As Short = 512 ' Third button is default

    'MsgBox return values
    Public Const IDOK As Short = 1 ' OK button pressed
    Public Const IDCANCEL As Short = 2 ' Cancel button pressed
    Public Const IDABORT As Short = 3 ' Abort button pressed
    Public Const IDRETRY As Short = 4 ' Retry button pressed
    Public Const IDIGNORE As Short = 5 ' Ignore button pressed
    Public Const IDYES As Short = 6 ' Yes button pressed
    Public Const IDNO As Short = 7 ' No button pressed

    'Type Definitions

    Structure OFSTRUCT

        Dim cBytes As Byte
        Dim fFixedDisk As Byte
        Dim nErrCode As Short
        Dim nReserved1 As Short
        Dim nReserved2 As Short
        <VBFixedString(256), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=256)> Public szPathName() As Char

    End Structure


    Structure VERINFO ' Version FIXEDFILEINFO

        Dim strPad1 As Integer ' Pad out struct version
        Dim strPad2 As Integer ' Pad out struct signature
        Dim nMSLo As Short ' Low word of ver # MS DWord
        Dim nMSHi As Short ' High word of ver # MS DWord
        Dim nLSLo As Short ' Low word of ver # LS DWord
        Dim nLSHi As Short ' High word of ver # LS DWord
        <VBFixedArray(16)> Dim strPad3() As Byte ' Skip some of VERINFO struct (16 bytes)
        Dim FileOS As Integer ' Information about the OS this file is targeted for.
        <VBFixedArray(16)> Dim strPad4() As Byte ' Pad out the resto of VERINFO struct (16 bytes)

        Public Sub Initialize()

            ReDim strPad3(16)
            ReDim strPad4(16)

        End Sub

    End Structure


    Structure PROTOCOL
        Dim strName As String
        Dim strFriendlyName As String
    End Structure


    Structure OSVERSIONINFO ' for GetVersionEx API call
        Dim dwOSVersionInfoSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
        <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=128)> Public szCSDVersion() As Char
    End Structure


    Public Const OF_EXIST As Integer = &H4000
    Public Const OF_SEARCH As Integer = &H400
    Public Const HFILE_ERROR As Short = -1

    'Global Variables

    Public LF As String ' single line break
    Public LS As String ' Single line break
    Public CRLF As String ' Carriage Return/Line Feed

    ' Global variables used for silent and SMS installation

    Public gfSilent As Boolean ' Whether or not we are doing a silent install
    Public gstrSilentLog As String ' filename for output during silent install.
    Public gfSMS As Boolean ' Whether or not we are doing an SMS silent install
    Public gstrMIFFile As String ' status output file for SMS
    Public gfSMSStatus As Boolean ' status of SMS installation
    Public gstrSMSDescription As String ' description string written to MIF file for SMS installation
    Public gfNoUserInput As Boolean ' True if either gfSMS or gfSilent is True
    Public gfDontLogSMS As Boolean ' Prevents MsgFunc from being logged to SMS (e.g., for confirmation messasges)
    Public Const MAX_SMS_DESCRIP As Short = 255 ' SMS does not allow description strings longer than 255 chars.

    'List of available protocols

    Public gProtocol() As PROTOCOL
    Public gcProtocols As Short

    ' AXDist.exe and wint351.exe needed.  These are self extracting exes
    ' that install other files not installed by setup1.

    Public gfAXDist As Boolean
    Public Const gstrFILE_AXDIST As String = "AXDIST.EXE"
    Public gstrAXDISTInstallPath As String
    Public gfAXDistChecked As Boolean
    Public gfWINt351 As Boolean
    Public Const gstrFILE_WINT351 As String = "WINt351.EXE"
    Public gstrWINt351InstallPath As String
    Public gfWINt351Checked As Boolean

    'API/DLL Declarations for 32 bit SetupToolkit

    Declare Function DiskSpaceFree Lib "VB5STKIT.DLL" Alias "DISKSPACEFREE" () As Integer
    Declare Function SetTime Lib "VB5STKIT.DLL" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Short
    Declare Function AllocUnit Lib "VB5STKIT.DLL" () As Integer
    Declare Function GetWinPlatform Lib "VB5STKIT.DLL" () As Integer
    Declare Function fNTWithShell Lib "VB5STKIT.DLL" () As Boolean
    Declare Function FSyncShell Lib "VB5STKIT.DLL" Alias "SyncShell" (ByVal strCmdLine As String, ByVal intCmdShow As Integer) As Integer
    Declare Function DLLSelfRegister Lib "VB5STKIT.DLL" (ByVal lpDllName As String) As Short
    Declare Function GetClsidFromActXFile Lib "VB5STKIT.DLL" (ByVal pszFilename As String, ByVal pszProgID As String, ByVal pszClsid As String) As Integer
    Declare Function RegisterTLB Lib "VB5STKIT.DLL" (ByVal lpTLBName As String) As Short

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Declare Sub lmemcpy Lib "VB5STKIT.DLL" (ByRef strDest As Any, ByVal strSrc As Any, ByVal lBytes As Integer)

    Declare Function OSfCreateShellGroup Lib "VB5STKIT.DLL" Alias "fCreateShellFolder" (ByVal lpstrDirName As String) As Integer
    Declare Function OSfCreateShellLink Lib "VB5STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Integer
    Declare Function OSfRemoveShellLink Lib "VB5STKIT.DLL" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Integer
    Private Declare Function OSGetLongPathName Lib "VB5STKIT.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Integer) As Integer

    Declare Function OpenFile Lib "kernel32" (ByVal lpFilename As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal wStyle As Integer) As Integer
    Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Integer, ByVal lpBuffer As String, ByVal lpFilePart As String) As Integer

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Integer, ByVal lpFilename As String) As Integer

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Integer

    Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
    Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Integer
    Declare Function GetTempFilename32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Short, ByVal lpTempFilename As String) As Integer
    Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
    Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Public Const LB_FINDSTRINGEXACT As Short = &H1A2S
    Public Const LB_ERR As Short = (-1)

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags As Integer, ByVal SrcName As String, ByVal DestName As String, ByVal SrcDir As String, ByVal DestDir As String, ByVal CurrDir As Any, ByVal TmpName As String, ByRef lpTmpFileLen As Integer) As Integer

    Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFilename As String, ByRef lVerHandle As Integer) As Integer
    Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFilename As String, ByVal lVerHandle As Integer, ByVal lcbSize As Integer, ByRef lpvData As Byte) As Integer
    Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (ByRef lpvVerData As Byte, ByVal lpszSubBlock As String, ByRef lplpBuf As Integer, ByRef lpcb As Integer) As Integer
    Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Integer) As Integer
    Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Integer
End Module