Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Module VIVAMain
    ' Public Const damping_type_in As Short = 0

    Public Const risdyn_n_in As Short = 1
    Public Const risfat_in As Short = 2
    Public Const conditions_in As Short = 3
    Public Const uppers_in As Short = 4
    Public Const lowers_in As Short = 5
    Public Const vivo_n_in As Short = 6
    Public Const freq_in As Short = 7
    Public Const rispre_in As Short = 8
    Public Const modes_us_in As Short = 9
    Public Const curv_us_in As Short = 10
    Public Const visc_def_in As Short = 11

    Public Const summ1_out As Short = 12
    Public Const out_out As Short = 13
    Public Const out_mm_out As Short = 14
    Public Const bend_out As Short = 15
    Public Const bend_mm_out As Short = 16
    Public Const cdrag_out As Short = 17
    Public Const cdrag_mm_out As Short = 18
    Public Const fat1_out As Short = 19
    Public Const dyn_n_dat As Short = 20
    Public Const fat_out As Short = 21
    Public Const freq_out As Short = 22
    Public Const fat_mono_out As Short = 23
    Public Const fat_multi_out As Short = 24
    Public Const summ_out As Short = 25
    Public Const risout_dat As Short = 26
    Public Const risout_out As Short = 27
    Public Const summ_mm_out As Short = 28
    Public Const summ1_mm_out As Short = 29
    Public Const dynout_out As Short = 30
    Public Const strain_out As Short = 31
    Public Const strain_mm_out As Short = 32
    Public Const clv_out As Short = 33
    Public Const clv_dom_out As Short = 34
    Public Const fat1_h3_out As Short = 35
    Public Const fat_h3_out As Short = 36
    Public Const fat_mono_h3_out As Short = 37
    Public Const fat_multi_h3_out As Short = 38
    Public Const stres_hh_out As Short = 39
    Public Const modes_us_out As Short = 40
    'Public Const summ2_out As Short = 41
    'Public Const summmm11_mm_out As Short = 42

    Public Const summary_out As Short = 41 'this is always the last one
    '
    Public Const MaxCurrentPair As Short = 100
    Public Const MaxSegments As Short = 200
    Public Const MaxSSPoints As Integer = 6001
    Public Const MaxAuxLines As Short = 6
    Public Const NumRiserTypes As Short = 2
    Public Const NumBoundaryConditions As Short = 4
    Public Const MinIteration As Short = 30
    Public Const MinVIVACorePoints As Short = 200
    Public Const NumVIVACorePoints As Short = 6001
    Public Const MaxDOSOutLines As Short = 600
    Public Const MaxNumLatSupports As Short = 100
    Public Const MaxNumPreFiles As Short = 4
    Public Const NumOutputFiles As Short = summary_out + 1
    Public Const NumInputFiles As Short = 12
    'Files

    Public FatigueCurves As Collection
    Public PreviousFiles As PreFile

    Public VIVADOSPrograms(2, NumVIVADOSPrograms) As String
    Public OutputFiles(2, NumOutputFiles) As String

    Structure VIVADBFiles
        Public NumFiles As Integer
        Public FileNames() As String
        Public NumFrq() As Integer
        Public NumAmp() As Integer
        Public TypeNames() As String
        Public TypeID() As Integer

        Public Sub New(ByVal i As Integer)
            NumFiles = i
        End Sub
    End Structure

    'Unit conversions

    'length
    '   metric
    Public Const mm2M As Single = 0.001

    '   english
    Public Const In2Ft As Single = 1.0# / 12.0#

    '   english - metric
    Public Const Ft2M As Single = 0.3048

    'angle
    Public Const Pi As Single = 3.1415926535
    Public Const Deg2Rad As Single = Pi / 180.0#

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

    'Keyboard constants
    Public Const DotKey As Short = 190
    Public Const EqualKey As Short = 187
    Public Const MinusKey As Short = 189

    'VIVA Component Names (see also the file list, below)
    Public Enum Units

        Metric = 0
        English = 1

    End Enum

    Public Enum HiReData

        Use = 2
        NoUse = 1

    End Enum

    Public Enum DampingMethod

        ModalDamping = 1
        ConstantDamping = 0

    End Enum


    Public Enum RiserTypes

        Rigid = 0
        SCRorLazyWave = 1

    End Enum


    Public Enum BoundaryConditions

        Pinned = 0
        Fixed = 1
        Spring = 2
        Free = 3

    End Enum

    Public Enum ICHARValues

        SmoothCylinder = 1
        VetcoRiser0 = 2
        VetcoRiser30 = 3
        VetcoRiser60 = 4
        VetcoRiser90 = 5
        VetcoRiser120 = 6
        VetcoRiser150 = 7
        StaggardBareBuoyant = 8
        TestedHighRe = 9
        TestedStrakes = 10
        userAdded11 = 11
        userAdded12 = 12
        userAdded13 = 13
        userAdded14 = 14
        userAdded15 = 15
        userAdded16 = 16
        userAdded17 = 17
        userAdded18 = 18
    End Enum


    Public Enum ICHOICEValues

        RigidRiser = 1
        SCR = 2

    End Enum


    Public Enum StrakesFlag

        YES = True
        NO = False

    End Enum


    Public Enum FrequencyResponseValues

        MultiFreq = 1
        SingleFreq = 2

    End Enum


    Public CurProj As Project

    Public NumInputForms As Short
    Public HaveResults As Boolean
    Public UnitsChangeable As Boolean

    ' These arrays represent the results of calculations from the various
    ' FORTRAN programs.  The values computed by RISEPREP, RIS6, RIS8, and RISFAT
    ' are all computed in this Visual Basic program; the values computed by the
    ' FORTRAN program VIVACORE are read into this program after VIVACORE completes
    ' its execution.

    Public SumVSq As Single ' Sum of velocities squared
    Public Const NumVIVADOSPrograms As Short = 3

    Public VIVADIR As String
    Public ReferenceDir As String

    Public VIVACoreFiles As VIVADBFiles

    Public ManualFile(2) As String
    Public ReferenceFile(10) As String

    Public fso As Object
    Public Const RegAcroRd As String = "\AcroRd32.exe"
    Public Const AcroRdFile As String = "acrord32.exe"
    Public AcroDir As String


    Public Sub InitializeVariables()
        Dim i, j As Integer
        Dim oBuff As Object
        Dim iFileNum As Integer
        Dim app(2) As Char
        app(1) = "F"
        app(2) = "R"
        FatigueCurves = New Collection
        PreviousFiles = New PreFile
        fso = CreateObject("Scripting.FileSystemObject")


        'VIVADIR = GetVIVADirFromRegistry() 'TODO JLIU
        VIVADIR = My.Application.Info.DirectoryPath & "\"

        AcroDir = ""
        ManualFile(0) = "WinVIVA User Manual V8.2.pdf"
        ManualFile(1) = "DOS VIVA User Manual V8.4.pdf"

        ReferenceFile(0) = "3D VIV of Flexible Structures.pdf"
        'ReferenceFile(1) = "Definition of Coordinate Systems for Program VIVA.pdf"
        ReferenceFile(1) = "Fatigue Calculation in VIVA.pdf"
        ReferenceFile(3) = "OMAE2011-49820.pdf"
        ReferenceFile(4) = "OMAE2011-viva -final.pdf"
        'ReferenceFile(4) = "Inline VIVA.pdf"
        ReferenceFile(2) = "Modeling Strakes and Foil Sections in VIVA.pdf"
        ReferenceFile(5) = "Questions & Answers About VIVA.pdf"
        ReferenceFile(6) = "Vortex Induced Vibrations of Long Cyclindrical Structures.pdf"
        'ReferenceFile(8) = "Vortex-Induced Vibration Analysis Based on Dydrodynamic Databases(OMAE2011-50192).pdf"

        VIVACoreFiles = New VIVADBFiles(19)

        With VIVACoreFiles
            .FileNames = New String(.NumFiles) {}
            .NumFrq = New Integer(.NumFiles) {}
            .NumAmp = New Integer(.NumFiles) {}
            .TypeNames = New String(.NumFiles) {}
            .TypeID = New Integer(.NumFiles) {}
        End With

        i = 0
        With VIVACoreFiles
            .FileNames(i) = "basic_bare.db"
            .NumFrq(i) = 36
            .NumAmp(i) = 0
            .TypeNames(i) = "Basic Smooth"
            .TypeID(i) = i + 1
        End With

        For i = 1 To 6
            With VIVACoreFiles
                .FileNames(i) = "out" & Trim(CStr(i)) & ".db"
                .NumFrq(i) = 14
                .NumAmp(i) = 5
                .TypeNames(i) = "Vetco"
                .TypeID(i) = i + 1
            End With
        Next

        i = 7
        With VIVACoreFiles
            .FileNames(i) = "out" & Trim(CStr(i)) & ".db"
            .NumFrq(i) = 11
            .NumAmp(i) = 5
            .TypeNames(i) = "Vetco"
            .TypeID(i) = i + 1
        End With

        i = 8
        With VIVACoreFiles
            .FileNames(i) = "basic_bare-hr.db"
            .NumFrq(i) = 36
            .NumAmp(i) = 5
            .TypeNames(i) = "Basic High Reynolds"
            .TypeID(i) = i + 1
        End With

        i = 9
        With VIVACoreFiles
            .FileNames(i) = "out_s.db"
            .NumFrq(i) = 15
            .NumAmp(i) = 4
            .TypeNames(i) = "Straked P/D=17"
            .TypeID(i) = i + 1
        End With

        'Read no_files.in
        oBuff = Nothing
        iFileNum = FreeFile()
        If Dir(VIVADIR & "no_files.in", FileAttribute.Normal) = "" Then
            MsgBox("The file no_files.in is missing, or you are in the wrong working directory.", MsgBoxStyle.OkOnly, "WinVIVA 8.2 Error")
        Else
            FileOpen(iFileNum, VIVADIR & "no_files.in", OpenMode.Input, OpenAccess.Read)
            Input(iFileNum, oBuff)
            VIVACoreFiles.NumFiles = CLng(oBuff)

            For i = 0 To 9
                If EOF(iFileNum) Then
                    MsgBox("The file no_files.in appears an old version, or you are in the wrong working directory.", MsgBoxStyle.OkOnly, "WinVIVA 8.1 Error")
                Else
                    oBuff = LineInput(iFileNum)
                End If
            Next

            For i = 10 To VIVACoreFiles.NumFiles - 1 Step 1
                Input(iFileNum, oBuff)
                VIVACoreFiles.FileNames(i) = CStr(oBuff)
                Input(iFileNum, oBuff)
                VIVACoreFiles.NumFrq(i) = CLng(oBuff)
                Input(iFileNum, oBuff)
                VIVACoreFiles.NumAmp(i) = CLng(oBuff)
            Next
            FileClose(iFileNum)
        End If


        OutputFiles(0, rispre_in) = "rispre.in"
        OutputFiles(0, risdyn_n_in) = "risdyn-n.in"
        OutputFiles(0, risfat_in) = "risfat.in"
        OutputFiles(0, conditions_in) = "conditions.in"
        OutputFiles(0, uppers_in) = "uppers.in"
        OutputFiles(0, lowers_in) = "lowers.in"
        OutputFiles(0, vivo_n_in) = "vivo-n.in"
        OutputFiles(0, freq_in) = "freq.in"

        OutputFiles(0, modes_us_in) = "modes_us.in"
        OutputFiles(0, curv_us_in) = "curv_us.in"
        OutputFiles(0, visc_def_in) = "visc_def.in"
        OutputFiles(0, summ1_out) = "summ1.out"
        OutputFiles(0, out_out) = "out.out"
        OutputFiles(0, out_mm_out) = "out_mm.out"
        OutputFiles(0, bend_out) = "bend.out"
        OutputFiles(0, bend_mm_out) = "bend_mm.out"
        OutputFiles(0, cdrag_out) = "cdrag.out"
        OutputFiles(0, cdrag_mm_out) = "cdrag_mm.out"
        OutputFiles(0, fat1_out) = "fat1.out"
        OutputFiles(0, dyn_n_dat) = "dyn-n.dat"
        OutputFiles(0, fat_out) = "fat.out"
        OutputFiles(0, freq_out) = "freq.out"
        OutputFiles(0, fat_mono_out) = "fat-mono.out"
        OutputFiles(0, fat_multi_out) = "fat-multi.out"
        OutputFiles(0, summ_out) = "summ.out"
        OutputFiles(0, risout_dat) = "risout.dat"
        OutputFiles(0, risout_out) = "risout.out"
        OutputFiles(0, summ_mm_out) = "summ_mm.out"
        OutputFiles(0, summ1_mm_out) = "summ1_mm.out"
        OutputFiles(0, dynout_out) = "dynout.out"
        OutputFiles(0, strain_out) = "strain.out"
        OutputFiles(0, strain_mm_out) = "strain_mm.out"
        OutputFiles(0, clv_out) = "clv.out"
        OutputFiles(0, clv_dom_out) = "clv_dom.out"
        OutputFiles(0, fat1_h3_out) = "fat1_h3.out"
        OutputFiles(0, fat_h3_out) = "fat_h3.out"
        OutputFiles(0, fat_mono_h3_out) = "fat-mono_h3.out"
        OutputFiles(0, fat_multi_h3_out) = "fat-multi_h3.out"
        OutputFiles(0, stres_hh_out) = "stres_hh.out"
        OutputFiles(0, modes_us_out) = "modes_us.out"
        'OutputFiles(0, summ2_out) = "summ2.out"
        OutputFiles(0, summary_out) = "summary.out"
        'OutputFiles(0, summmm11_mm_out) = "summm11_mm.out"
        'damping_type.in has only one file for 1 riser or 2 risers
        For i = risdyn_n_in To summary_out
            For j = 1 To 2
                OutputFiles(j, i) = Left(OutputFiles(0, i), InStr(OutputFiles(0, i), ".") - 1)
                OutputFiles(j, i) = OutputFiles(j, i) & app(j) & Right(OutputFiles(0, i), Len(OutputFiles(0, i)) - Len(OutputFiles(j, i)))
            Next
        Next
   
        OutputFiles(1, stres_hh_out) = "stres_hh.out"
        OutputFiles(2, stres_hh_out) = "stres_hh.out"

        OutputFiles(1, risout_out) = "risoutF.out"
        OutputFiles(2, risout_out) = "risoutR.out"
        OutputFiles(1, risout_dat) = "risoutF.dat"
        OutputFiles(2, risout_dat) = "risoutR.dat"

        OutputFiles(1, dynout_out) = "dynoutF.out"
        OutputFiles(2, dynout_out) = "dynoutR.out"

        VIVADOSPrograms(1, 1) = "RISPREP-v8.EXE"
        VIVADOSPrograms(1, 2) = "RIS6-v8.EXE"
        VIVADOSPrograms(1, 3) = "VIVA-v84.EXE"
        VIVADOSPrograms(2, 1) = "RISPREP-vv81.EXE"
        VIVADOSPrograms(2, 2) = "RIS6-vv81.EXE"
        VIVADOSPrograms(2, 3) = "VIVARRAY-vv83.EXE"
        HaveResults = False
        NumInputForms = 0

        Dim FatCurv As FatigueCurve
        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "API X"
            .A = 1372900000.0#
            .B = 4.38
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "API X'"
            .A = 1911500000.0#
            .B = 3.74
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class B"
            .A = 2827900000.0#
            .B = 4.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class C"
            .A = 3919700000.0#
            .B = 3.5
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class D"
            .A = 5740800000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class E"
            .A = 5077300000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class F"
            .A = 4288500000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class F2"
            .A = 3673900000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class G"
            .A = 3130700000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class W"
            .A = 2705800000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

        FatCurv = New FatigueCurve
        With FatCurv
            .Name = "Class T"
            .A = 5653300000.0#
            .B = 3.0#
        End With
        FatigueCurves.Add(FatCurv)

ErrHandler:
        'User pressed Cancel button
        Exit Sub

    End Sub

End Module