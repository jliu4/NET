Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Module modUtility

    ' The following typedefs and procedure declarations establish the infrastructure
    ' necessary to run the DOS Fortran programs from Visual Basic.  Simply using the
    ' SHELL function is problematic, as the commands issued using SHELL are run
    ' asynchronously, and thus the VB statements following SHELL will be run before
    ' the Fortran routines finish.  Using the Windows API functions CreateProcess()
    ' and WaitForSingleObject() allow execution to be suspended until the Fortran
    ' programs have completed.
    '
    ' This technique is described in the Microsoft Developer Network Knowledge Base
    ' (see Article ID Q129796).

    Private Structure STARTUPINFO
        Dim cb As Integer
        Dim lpReserved As String
        Dim lpDesktop As String
        Dim lpTitle As String
        Dim dwX As Integer
        Dim dwY As Integer
        Dim dwXSize As Integer
        Dim dwYSize As Integer
        Dim dwXCountChars As Integer
        Dim dwYCountChars As Integer
        Dim dwFillAttribute As Integer
        Dim dwFlags As Integer
        Dim wShowWindow As Short
        Dim cbReserved2 As Short
        Dim lpReserved2 As Integer
        Dim hStdInput As Integer
        Dim hStdOutput As Integer
        Dim hStdError As Integer
    End Structure

    Private Structure PROCESS_INFORMATION
        Dim hProcess As Integer
        Dim hThread As Integer
        Dim dwProcessID As Integer
        Dim dwThreadID As Integer
    End Structure

    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer

    Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Integer, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As Integer, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Integer

    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer

    Private Const NORMAL_PRIORITY_CLASS As Integer = &H20S
    Private Const INFINITE As Integer = -1
    Private Const STARTF_USESHOWWINDOW As Short = 1
    Private Const SW_HIDE As Short = 0
    Private Const SW_SHOWNORMAL As Short = 1
    Private Const SW_SHOWMINIMIZED As Short = 2
    Private Const SW_SHOWMAXIMIZED As Short = 3


    Public Sub ExecCmd(ByRef CmdLine As String, ByRef HideWindow As Boolean)

        Dim proc As PROCESS_INFORMATION
        Dim start As STARTUPINFO
        Dim ret As Integer

        'initialization
        start = Nothing

        '   Initialize the STARTUPINFO structure:
        '   Tells CreateProcess to use the wShowWindow parameter
        start.dwFlags = STARTF_USESHOWWINDOW

        '   Causes the process started to be hidden
        If HideWindow Then
            start.wShowWindow = SW_HIDE
        Else
            start.wShowWindow = SW_SHOWNORMAL
        End If
        start.cb = Len(start)

        '   Start the shelled application:
        ret = CreateProcessA(0, CmdLine, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, start, proc)

        '   Wait for the shelled application to finish:
        ret = WaitForSingleObject(proc.hProcess, INFINITE)
        ret = CloseHandle(proc.hProcess)

    End Sub


    Public Function PositiveNumber(ByRef FieldText As String) As Boolean

        If Not IsNumeric(FieldText) Then
            PositiveNumber = False
            Exit Function
        ElseIf Val(FieldText) < 0 Then
            PositiveNumber = False
            Exit Function
        End If
        PositiveNumber = True

    End Function


    Public Function CStrk(ByRef val1 As String) As VIVAMain.StrakesFlag

        If UCase(val1) = "YES" Then
            CStrk = VIVAMain.StrakesFlag.YES
        ElseIf UCase(val1) = "NO" Then
            CStrk = VIVAMain.StrakesFlag.NO
        Else
            Err.Raise(Number:=32200, Description:="Invalid argument")
        End If

    End Function


    Public Function CBcs(ByRef val1 As Object) As VIVAMain.BoundaryConditions

        Select Case CShort(val1)
            Case 0
                CBcs = VIVAMain.BoundaryConditions.Pinned
            Case 1
                CBcs = VIVAMain.BoundaryConditions.Fixed
            Case 2
                CBcs = VIVAMain.BoundaryConditions.Spring
            Case 3
                CBcs = VIVAMain.BoundaryConditions.Free
            Case Else
                Err.Raise(Number:=32201, Description:="Invalid value for conversion to boundary condition")
        End Select

    End Function


    Public Function CRst(ByRef val1 As Object) As VIVAMain.RiserTypes

        Select Case CShort(val1)
            Case 0
                CRst = VIVAMain.RiserTypes.Rigid
            Case 1
                CRst = VIVAMain.RiserTypes.SCRorLazyWave
            Case Else
                Err.Raise(Number:=32202, Description:="Invalid value for conversion to riser type")
        End Select

    End Function


    Public Function CIchv(ByRef val1 As Object) As VIVAMain.ICHARValues

        Select Case CShort(val1)
            Case 1
                CIchv = VIVAMain.ICHARValues.SmoothCylinder
            Case 2
                CIchv = VIVAMain.ICHARValues.VetcoRiser0
            Case 3
                CIchv = VIVAMain.ICHARValues.VetcoRiser30
            Case 4
                CIchv = VIVAMain.ICHARValues.VetcoRiser60
            Case 5
                CIchv = VIVAMain.ICHARValues.VetcoRiser90
            Case 6
                CIchv = VIVAMain.ICHARValues.VetcoRiser120
            Case 7
                CIchv = VIVAMain.ICHARValues.VetcoRiser150
            Case 8
                CIchv = VIVAMain.ICHARValues.StaggardBareBuoyant
            Case 9
                CIchv = VIVAMain.ICHARValues.TestedStrakes
            Case Else
                Err.Raise(Number:=32203, Description:="Invalid value for conversion to ICHARValue type")
        End Select

    End Function


    Public Function CStf(ByRef val1 As Object) As VIVAMain.StrakesFlag

        Select Case CBool(val1)
            Case True
                CStf = VIVAMain.StrakesFlag.YES
            Case False
                CStf = VIVAMain.StrakesFlag.NO
            Case Else
                Err.Raise(Number:=32204, Description:="Invalid value for conversion to StrakesFlag type")
        End Select

    End Function


    Public Function CFr(ByRef val1 As Object) As VIVAMain.FrequencyResponseValues

        Select Case CShort(val1)
            Case 1
                CFr = VIVAMain.FrequencyResponseValues.MultiFreq
            Case 2
                CFr = VIVAMain.FrequencyResponseValues.SingleFreq
            Case Else
                Err.Raise(Number:=32205, Description:="Invalid value for conversion to FrequencyResponseValues type")
        End Select

    End Function

End Module