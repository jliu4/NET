Option Strict Off
Option Explicit On
Module MARSConstants
	'--------------  conditional compilation flag -----------------
	
	Public Const DTCEL As Boolean = True
	
	'--------------------------------------------------------------
	
	Public Cancelled As Boolean ' for progress bar cancel action
	
	Public Const MaxNumPreFiles As Short = 4
	Public Const FileNumRes As Short = 99
	Public Const MaxCurrentPair As Short = 100
	
	Public Const OffsetOutput As String = "offset.out"
	Public Const AppVelOutput As String = "appvel.out"
	
	' Program constants
	Public Const mX As Short = 1
	'UPGRADE_NOTE: MY was upgraded to MY_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Const MY_Renamed As Short = 2
	Public Const mZ As Short = 3
	
	Public Const ObjectFileName As String = "ObjFile.dat"
	
	Public Const WarningColor As Short = 12
	Public Const DefaultColor As Short = 0
	
	' Physical constants
	Public Const PI As Double = 3.14159265358979
	Public Const Degrees2Radians As Double = PI / 180#
	Public Const Radians2Degrees As Double = 180# / PI
	
	Public Const Knots2Ftps As Double = 1.6878098571
	Public Const Ftps2Knots As Double = 0.5924838013
	
	Public Const RegistryName As String = "\App Paths\dodo.exe"
	Public Const DODODirDefault As String = "C:\Program Files\DODO"
	Public Const msgTitle As String = "DODO"
	Public Const IniFile As String = "dodo.ini"
	
	
	' "Baroness"        ' default vessel
	Public Const PontHeight As Double = 28.9
	Public Const PontLength As Double = 357
	Public Const PontWidth As Double = 28.9
	Public Const PLength As Double = 49
	Public Const PDistance As Double = 61
	Public Const hBraceRadius As Double = 8.8
	Public Const DeckElev As Double = 120
	
	
	Public DODODir As String
	Public MaxDynamicMotion As Double
	
	Public CurProj As Project
	Public Defaults As DODOIni
	Public CurVessel As Vessel
	Public NumTimeSteps As Short
	Public TimeStep As Single
	Public MaxTimeSteps As Short
	Public TransX() As Object
	Public TransY() As Object
	Public TransYaw() As Object
	Public TransTension() As Double

    Public IsMetricUnit As Boolean = False
    Public DiaFactor, FrcFactor, LFactor, StressFactor, MassFactor As Double
	Public VelFactor As Double
	Public FrcUnit, LUnit, VelUnit, DiameterUnit As String

    Public Sub RefreshUnitLabels(ByRef frm As frmTransient)

        On Error Resume Next ' ignore error if the form does not have the controls

        'frm.setLabel()

    End Sub

    Public Sub InitProject()
		
		CurProj = New Project
		Defaults = New DODOIni

        ' DODODir = GetDirFromRegistry(RegistryName) & "\"
        ' DODODir = "c:\program files\DODO\"
        IsMetricUnit = False

        Dim FS As Object

        FS = CreateObject("Scripting.FileSystemObject")
		'UPGRADE_WARNING: Couldn't resolve default property of object FS.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not FS.FolderExists(My.Application.Info.DirectoryPath & "\Data") Then
			'UPGRADE_WARNING: Couldn't resolve default property of object FS.CreateFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call FS.CreateFolder(My.Application.Info.DirectoryPath & "\Data")
		End If
		
		
		' check if data subdir exist
		DODODir = My.Application.Info.DirectoryPath & "\Data\"
		
	End Sub
	'------------------------------------------
End Module