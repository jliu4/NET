Option Strict Off
Option Explicit On
Option Compare Text

Module modRegistry

    ' Global Constants

    Public Const gstrNULL As String = "" ' Empty string
    Public Const gstrSEP_DIR As String = "\" ' Directory separator character
    Public Const gstrSEP_REGKEY As String = "\" ' Registration key separator character.

    'Type Definitions

    Structure DISKINFO ' Disk drive information
        Dim lAvail As Integer ' Bytes available on drive
        Dim lReq As Integer ' Bytes required for setup
        Dim lMinAlloc As Integer ' minimum allocation unit
    End Structure

    Structure DESTINFO ' save dest dir for certain files
        Dim strAppDir As String
        Dim strAUTMGR32 As String
        Dim strRACMGR32 As String
    End Structure

    Structure REGINFO ' save registration info for files
        Dim strFilename As String
        Dim strRegister As String
        'The following are used only for remote server registration
        Dim strNetworkAddress As String
        Dim strNetworkProtocol As String
        Dim intAuthentication As Short
        Dim fDCOM As Boolean ' True if DCOM, otherwise False
    End Structure

    'Global Variables

    Public gstrSETMSG As String
    Public gfRetVal As Short ' return value for form based functions
    Public gstrAppName As String ' name of app being installed
    Public gstrTitle As String ' "setup" name of app being installed
    Public gstrDefGroup As String ' Default name for group -- from setup.lst
    Public gstrDestDir As String ' dest dir for application files
    Public gstrAppExe As String ' name of app .EXE being installed
    Public gstrAppToUninstall As String ' Name of app exe/ocx/dll to be uninstalled.  Should be the same as gstrAppExe in most cases.
    Public gstrSrcPath As String ' path of source files
    Public gstrSetupInfoFile As String ' pathname of SETUP.LST file
    Public gstrWinDir As String ' windows directory
    Public gstrWinSysDir As String ' windows\system directory
    Public gsDiskSpace() As DISKINFO ' disk space for target drives
    Public gstrDrivesUsed As String ' dest drives used by setup
    Public glTotalCopied As Integer ' total bytes copied so far
    Public gintCurrentDisk As Short ' current disk number being installed
    Public gsDest As DESTINFO ' dest dirs for certain files
    Public gstrAppRemovalLog As String ' name of the app removal logfile
    Public gstrAppRemovalEXE As String ' name of the app removal executable
    Public gfAppRemovalFilesMoved As Boolean ' whether or not the app removal files have been moved to the application directory
    Public gfForceUseDefDest As Boolean ' If set to true, then the user will not be prompted for the destination directory
    Public fMainGroupWasCreated As Boolean ' Whether or not a main folder/group has been created
    Public gfRegDAO As Boolean ' If this gets set to true in the code, then
    ' we need to add some registration info for DAO
    ' to the registry.

    'Form/Module Constants

    'Possible ProgMan actions
    Const mintDDE_ITEMADD As Short = 1 ' AddProgManItem flag
    Const mintDDE_GRPADD As Short = 2 ' AddProgManGroup flag

    'Special file names
    Const mstrFILE_APPREMOVALLOGBASE As String = "ST5UNST" ' Base name of the app removal logfile
    Const mstrFILE_APPREMOVALLOGEXT As String = ".LOG" ' Default extension for the app removal logfile
    Const mstrFILE_AUTMGR32 As String = "AUTMGR32.EXE"
    Const mstrFILE_RACMGR32 As String = "RACMGR32.EXE"
    Const mstrFILE_CTL3D32 As String = "CTL3D32.DLL"
    Const mstrFILE_RICHED32 As String = "RICHED32.DLL"

    'Name of temporary file used for concatenation of split files
    Const mstrCONCATFILE As String = "VB5STTMP.CCT"

    'setup information file registration macros
    Const mstrDLLSELFREGISTER As String = "$(DLLSELFREGISTER)"
    Const mstrEXESELFREGISTER As String = "$(EXESELFREGISTER)"
    Const mstrTLBREGISTER As String = "$(TLBREGISTER)"
    Const mstrREMOTEREGISTER As String = "$(REMOTE)"
    Const mstrVBLREGISTER As String = "$(VBLREGISTER)" ' Bug 5-8039

    'Form/Module Variables

    Private msRegInfo() As REGINFO ' files to be registered
    Private mlTotalToCopy As Integer ' total bytes to copy
    Private mintConcatFile As Short ' handle of dest file for concatenation
    Private mlSpaceForConcat As Integer ' extra space required for concatenation
    Private mstrConcatDrive As String ' drive to use for concatenation
    Private mstrVerTmpName As String ' temp file name for VerInstallFile API

    ' Hkey cache (used for logging purposes)
    Private Structure HKEY_CACHE
        Dim hkey As Integer
        Dim strHkey As String
    End Structure

    Private hkeyCache() As HKEY_CACHE

    ' Registry manipulation API's (32-bit)
    Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
    Public Const HKEY_CURRENT_USER As Integer = &H80000001
    Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Public Const HKEY_USERS As Integer = &H80000003
    Const ERROR_SUCCESS As Short = 0
    Const ERROR_NO_MORE_ITEMS As Short = 259

    Const REG_SZ As Short = 1
    Const REG_BINARY As Short = 3
    Const REG_DWORD As Short = 4


    Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hkey As Integer) As Integer
    Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hkey As Integer, ByVal lpszSubKey As String, ByRef phkResult As Integer) As Integer
    Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hkey As Integer, ByVal lpszSubKey As String) As Integer
    Declare Function OSRegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hkey As Integer, ByVal iSubKey As Integer, ByVal lpszName As String, ByVal cchName As Integer) As Integer
    Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hkey As Integer, ByVal lpszSubKey As String, ByRef phkResult As Integer) As Integer

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function OSRegQueryValueExInteger Lib "advapi32" Alias "RegQueryValueExA" (ByVal hkey As Integer, ByVal lpszValueName As String, ByVal dwReserved As Integer, ByRef lpdwType As Integer, ByVal lpbData As Integer, ByRef cbData As Integer) As Integer
    Declare Function OSRegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hkey As Integer, ByVal lpszValueName As String, ByVal dwReserved As Integer, ByRef lpdwType As Integer, ByVal lpbData As String, ByRef cbData As Integer) As Integer

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function OSRegSetValueExInteger Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Integer, ByVal lpszValueName As String, ByVal dwReserved As Integer, ByVal fdwType As Integer, ByRef lpbData As Integer, ByVal cbData As Integer) As Integer
    Declare Function OSRegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Integer, ByVal lpszValueName As String, ByVal dwReserved As Integer, ByVal fdwType As Integer, ByRef lpbData As String, ByVal cbData As Integer) As Integer

    Declare Function GetCurrentProcessId Lib "kernel32" () As Integer

    ' FUNCTION: RegOpenKey
    '
    ' Opens an existing key in the system registry.
    '
    ' Returns: True if the key was opened OK, False otherwise
    '   Upon success, phkResult is set to the handle of the key.


    Function RegOpenKey(ByVal hkey As Integer, ByVal lpszSubKey As String, ByRef phkResult As Integer) As Boolean

        Dim lResult As Integer
        Dim strHkey As String

        On Error GoTo 0

        strHkey = strGetHKEYString(hkey)

        lResult = OSRegOpenKey(hkey, lpszSubKey, phkResult)
        If lResult = ERROR_SUCCESS Then
            RegOpenKey = True
            AddHkeyToCache(phkResult, strHkey & "\" & lpszSubKey)
        Else
            RegOpenKey = False
        End If

    End Function


    ' FUNCTION: RegPathWinCurrentVersion
    '
    ' Returns the name of the registry key
    ' "\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion"
    Function RegPathWinCurrentVersion() As String

        RegPathWinCurrentVersion = "SOFTWARE"

    End Function


    ' FUNCTION: RegQueryStringValue
    '
    ' Retrieves the string data for a named
    ' (strValueName = name) or unnamed (strValueName = "")
    ' value within a registry key.  If the named value
    ' exists, but its data is not a string, this function
    ' fails.
    '
    ' NOTE: For 16-bits, strValueName MUST be "" (but the
    ' NOTE: parameter is left in for source code compatability)
    '
    ' Returns: True on success, else False.
    '   On success, strData is set to the string data value

    Function RegQueryStringValue(ByVal hkey As Integer, ByVal strValueName As String, ByRef strData As String) As Boolean

        Dim lResult As Integer
        Dim lValueType As Integer
        Dim strBuf As String
        Dim lDataBufSize As Integer

        RegQueryStringValue = False

        On Error GoTo 0

        ' Get length/data type
        lResult = OSRegQueryValueExInteger(hkey, strValueName, 0, lValueType, 0, lDataBufSize)
        strBuf = New String(" ", lDataBufSize)
        lResult = OSRegQueryValueExString(hkey, strValueName, 0, 0, strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            If lValueType = REG_SZ Then
                strBuf = New String(" ", lDataBufSize)
                lResult = OSRegQueryValueExString(hkey, strValueName, 0, 0, strBuf, lDataBufSize)
                If lResult = ERROR_SUCCESS Then
                    RegQueryStringValue = True
                    strData = StripTerminator(strBuf)
                End If
            End If
        End If

    End Function


    'Adds or replaces an HKEY to the list of HKEYs in cache.
    'Note that it is not necessary to remove keys from
    'this list.

    Private Sub AddHkeyToCache(ByVal hkey As Integer, ByVal strHkey As String)

        Dim intIdx As Short

        intIdx = intGetHKEYIndex(hkey)

        If intIdx < 0 Then
            'The key does not already exist.  Add it to the end.
            On Error Resume Next
            ReDim Preserve hkeyCache(UBound(hkeyCache) + 1)
            If Err.Number Then
                'If there was an error, it means the cache was empty.
                On Error GoTo 0
                ReDim hkeyCache(0)
            End If
            On Error GoTo 0

            intIdx = UBound(hkeyCache)
        Else
            'The key already exists.  It will be replaced.
        End If

        hkeyCache(intIdx).hkey = hkey
        hkeyCache(intIdx).strHkey = strHkey

    End Sub


    'Given a predefined HKEY, return the text string representing that
    'key, or else return "".

    Private Function strGetPredefinedHKEYString(ByVal hkey As Integer) As String

        Select Case hkey
            Case HKEY_CLASSES_ROOT
                strGetPredefinedHKEYString = "HKEY_CLASSES_ROOT"
            Case HKEY_CURRENT_USER
                strGetPredefinedHKEYString = "HKEY_CURRENT_USER"
            Case HKEY_LOCAL_MACHINE
                strGetPredefinedHKEYString = "HKEY_LOCAL_MACHINE"
            Case HKEY_USERS
                strGetPredefinedHKEYString = "HKEY_USERS"
            Case Else
                strGetPredefinedHKEYString = ""
                'End Case
        End Select

    End Function


    'Given an HKEY, return the text string representing that
    'key.

    Private Function strGetHKEYString(ByVal hkey As Integer) As String

        Dim strKey As String

        'Is the hkey predefined?
        strKey = strGetPredefinedHKEYString(hkey)
        If strKey <> "" Then
            strGetHKEYString = strKey
            Exit Function
        End If

        'It is not predefined.  Look in the cache.
        Dim intIdx As Short
        intIdx = intGetHKEYIndex(hkey)
        If intIdx >= 0 Then
            strGetHKEYString = hkeyCache(intIdx).strHkey
        Else
            strGetHKEYString = ""
        End If

    End Function


    'Searches the cache for the index of the given HKEY.
    'Returns the index if found, else returns -1.

    Private Function intGetHKEYIndex(ByVal hkey As Integer) As Short

        Dim intUBound As Short

        On Error Resume Next
        intUBound = UBound(hkeyCache)
        If Err.Number Then
            'If there was an error accessing the ubound of the array,
            'then the cache is empty
            GoTo NotFound
        End If
        On Error GoTo 0

        Dim intIdx As Short
        For intIdx = 0 To intUBound
            If hkeyCache(intIdx).hkey = hkey Then
                intGetHKEYIndex = intIdx
                Exit Function
            End If
        Next intIdx

NotFound:
        intGetHKEYIndex = -1

    End Function


    'Get the path portion of a filename

    Function GetPathName(ByVal strFilename As String) As String

        Dim intPos As Short
        Dim strPathOnly As String
        Dim dirTmp As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Dim i As Short

        On Error Resume Next

        Err.Clear()

        intPos = Len(strFilename)

        'Change all '/' chars to '\'

        For i = 1 To Len(strFilename)
            If Mid(strFilename, i, 1) = gstrSEP_DIRALT Then
                Mid(strFilename, i, 1) = gstrSEP_DIR
            End If
        Next i

        If InStr(strFilename, gstrSEP_DIR) = intPos Then
            If intPos > 1 Then
                intPos = intPos - 1
            End If
        Else
            Do While intPos > 0
                If Mid(strFilename, intPos, 1) <> gstrSEP_DIR Then
                    intPos = intPos - 1
                Else
                    Exit Do
                End If
            Loop
        End If

        If intPos > 0 Then
            strPathOnly = Left(strFilename, intPos)
            If Right(strPathOnly, 1) = gstrCOLON Then
                strPathOnly = strPathOnly & gstrSEP_DIR
            End If
        Else
            strPathOnly = CurDir()
        End If

        If Right(strPathOnly, 1) = gstrSEP_DIR Then
            strPathOnly = Left(strPathOnly, Len(strPathOnly) - 1)
        End If

        GetPathName = UCase(strPathOnly)

        Err.Clear()

    End Function


    ' FUNCTION: StripTerminator
    '
    ' Returns a string without any zero terminator.  Typically,
    ' this was a string returned by a Windows API call.
    '
    ' IN: [strString] - String to remove terminator from
    '
    ' Returns: The value of the string passed in minus any
    '          terminating zero.

    Function StripTerminator(ByVal strString As String) As String

        Dim intZeroPos As Short

        intZeroPos = InStr(strString, Chr(0))
        If intZeroPos > 0 Then
            StripTerminator = Left(strString, intZeroPos - 1)
        Else
            StripTerminator = strString
        End If

    End Function
End Module