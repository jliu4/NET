Option Strict Off
Option Explicit On
Module MARSConstants
	'--------------  conditional compilation flag -----------------
	
	Public Const DTCEL As Boolean = True
	
	'--------------------------------------------------------------
	
	Public Cancelled As Boolean ' for progress bar cancel action
    Public Const MaxCurrentPair As Short = 100

    Public Const MaxNumLines As Short = 24
	Public Const MaxNumSeg As Short = 24
	Public Const MaxNumSubSeg As Short = 200
	Public Const MaxNumWells As Short = 30
	Public Const MaxNumIter As Short = 100
	
	Public Const MaxNumPipes As Short = 10
	Public Const MaxNumPipeNodes As Short = 11
	Public Const MaxNumGrids As Short = 201
	
	Public Const MaxNumPreFiles As Short = 4
	Public Const FileNumRes As Short = 99

    Public Const OffsetOutput As String = "offset.out"
    Public Const AppVelOutput As String = "appvel.out"
    ' Program constants
    Public Const mX As Short = 1
    Public Const MY_Renamed As Short = 2
    Public Const mZ As Short = 3

    Public Const ObjectFileName As String = "ObjFile.dat"
    Public Const CatenaryFile As String = "catline.dat"

    Public Const WarningColor As Short = 12
    Public Const DefaultColor As Short = 0

    ' Physical constants

    Public Const RegistryName As String = "\App Paths\mars.exe"
    Public Const MarsDirDefault As String = "C:\Program Files\MARS"
    Public Const DODOsDirDefault As String = "C:\Program Files\MARS"
    Public Const msgTitle As String = "MARS"
    Public Const IniFile As String = "mars.ini"
    Public Const OnBoardFile As String = "onboard.dat"
    Public Const DerpDir As String = "C:\Riser Analysis\"
    Public Const DerpFile As String = "mars.off"

    Public Const PipeFile As String = "pipelines.inp"
    Public Const SeaBedFile As String = "seabed.inp"

    ' "Baroness"        ' default vessel
    Public Const PontHeight As Single = 28.9
    Public Const PontLength As Single = 357
    Public Const PontWidth As Single = 28.9
    Public Const PLength As Single = 49
    Public Const PDistance As Single = 61
    Public Const hBraceRadius As Single = 8.8
    Public Const DeckElev As Single = 120

    '   '  The following are constants to be used with a pontoon
    '   '  "Mar700"
    '    Public Const PontHeight As Single = 31.98
    '    Public Const PontLength As Single = 390.32
    '    Public Const PontWidth As Single = 39.36
    '    Public Const PLength As Single = 22.96
    '    Public Const PDistance As Single = 180.4
    '    Public Const hBraceRadius As Single = 3.28 * 3
    '    Public Const DeckElev As Single = 3.28 * 34.75

    Public MarsDir, DODODir As String
    Public VesselPointsFromFile As Boolean
    Public MaxDynamicMotion As Single

    Public DiaFactor, FrcFactor, LFactor, StressFactor, MassFactor As Single

    Public CurProj As Project
    'Public Defaults As MarsIni
    ' Public Defaults As DODOIni

    Public CurVessel As Vessel
    Public CurField As Wells
    Public NumTimeSteps As Short
    Public TimeStep As Single
    Public MaxTimeSteps As Short
    Public TransX() As Object
    Public TransY() As Object
    Public TransYaw() As Object
    Public TransTension() As Single
    'Unit conversions
    'length
    '   metric
    Public Const mm2M As Single = 0.001

    '   english
    Public Const In2Ft As Single = 1.0# / 12.0#

    '   english - metric
    Public Const Ft2M As Single = 0.3048

    'angle
    Public Const PI As Single = System.Math.PI '3.14159265358979
    Public Const Degrees2Radians As Double = System.Math.PI / 180.0#
    Public Const Radians2Degrees As Double = 180.0# / System.Math.PI

    'mass
    '   english - metric
    Public Const Lb2Kg As Single = 0.45359237

    'force
    '   g
    Public Const Gm As Single = 9.80665

    '   english - metric
    Public Const Lb2N As Single = Lb2Kg * Gm

    'time
    Public Const Sec2Year As Single = 1.0# / (365.0# * 24.0# * 60.0# * 60.0#)

    'volume
    '   english - metric
    Public Const Gal2M3 As Single = 0.003785412

    'velocity
    '   english - metric
    Public Const Kn2MPS As Single = 0.5144444
    Public Const Knots2Ftps As Single = 1.6878098571
    Public Const Ftps2Knots As Single = 0.5924838013


    Public Enum SegType '2.1.4
        Wire = 1
        Chain = 2
        PresetWire = 3
        PresetChain = 4
        PresetPoly = 5
        PresetBuoy = 6
    End Enum

    Public IsMetricUnit As Boolean
    Public DiameterFactor As Single
    Public VelFactor As Single
    Public FrcUnit, LUnit, VelUnit, DiameterUnit As String

    ' Public Sub RefreshUnitLabels(ByRef frm As System.Windows.Forms.Form)

    'End Sub

    'Public Sub InitProject()
    'On Error GoTo ErrHandler
    '   CurProj = New Project
    'Defaults = New DODOIni
    'Defaults = New MarsIni
    ' MarsDir = GetDirFromRegistry(RegistryName) & "\"
    ' MarsDir = "c:\program files\mars\"


    'Dim FS As Object

    '  FS = CreateObject("Scripting.FileSystemObject")
    'UPGRADE_WARNING: Couldn't resolve default property of object FS.FolderExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'If Not FS.FolderExists(My.Application.Info.DirectoryPath & "\Data") Then
    'UPGRADE_WARNING: Couldn't resolve default property of object FS.CreateFolder. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'Call FS.CreateFolder(My.Application.Info.DirectoryPath & "\Data")
    'End If


    ' check if data subdir exist
    '  MarsDir = My.Application.Info.DirectoryPath & "\Data\"
    '   DODODir = My.Application.Info.DirectoryPath & "\Data\"
    'Exit Sub

    'ErrHandler:
    'User pressed Cancel button
    '  MsgBox("InitProject")
    'Exit Sub
    'End Sub
    '------------------------------------------
End Module