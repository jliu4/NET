Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class Project

    Private mstrTitle, mstrDesc As String
    Private mstrFQFileName, mstrFileName As String
    Private mstrBakFileDirectory, mstrDataDirectory As String
    Private mblnSaved As Object
    Private mblnShowDOSBox As Boolean
    Private mintNumIterations, mintNumPoints As Short
    Private mFR As VIVAMain.FrequencyResponseValues
    Private mRisers As Short
    Private mRiserId As Short
    Private mSameRiserARiser As Boolean
    Private mRiser(2) As Riser
    Private mWater As Water
    Private mUnits As VIVAMain.Units
    Private mDampingMethod As VIVAMain.DampingMethod
    Private mRiserFLocY As Single
    Private mRiserFLocZ As Single
    Private mRiserRLocY As Single
    Private mRiserRLocZ As Single
    Private mBatchProcess As Boolean


    Public Property BatchProcess() As Boolean
        Get
            BatchProcess = mBatchProcess
        End Get

        Set(ByVal Value As Boolean)
            mBatchProcess = Value
        End Set

    End Property


    Public Property RiserRLocZ() As Single
        Get
            RiserRLocZ = mRiserRLocZ
        End Get

        Set(ByVal Value As Single)
            mRiserRLocZ = Value
        End Set

    End Property

    Public Property RiserRLocY() As Single
        Get
            RiserRLocY = mRiserRLocY
        End Get

        Set(ByVal Value As Single)
            mRiserRLocY = Value
        End Set

    End Property
    Public Property RiserFLocZ() As Single
        Get
            RiserFLocZ = mRiserFLocZ
        End Get

        Set(ByVal Value As Single)
            mRiserFLocZ = Value
        End Set

    End Property

    Public Property RiserFLocY() As Single
        Get
            RiserFLocY = mRiserFLocY
        End Get

        Set(ByVal Value As Single)
            mRiserFLocY = Value
        End Set

    End Property
    Public Property RiserId() As Short
        Get
            RiserId = mRiserId
        End Get

        Set(ByVal Value As Short)
            mRiserId = Value
        End Set

    End Property

       Public Property FileName() As String

        Get
            FileName = mstrFileName
        End Get

        Set(ByVal Value As String)
            mstrFileName = Value
        End Set

    End Property

    Public Property FQFileName() As String

        Get
            FQFileName = mstrFQFileName
        End Get

        Set(ByVal Value As String)
            Dim DirChrPos, Pos, NameLen As Short
            mstrFQFileName = Value
            NameLen = Len(mstrFQFileName)
            DirChrPos = NameLen
            For Pos = NameLen To 1 Step -1
                If Mid(mstrFQFileName, Pos, 1) = "\" Then
                    DirChrPos = Pos
                    Exit For
                End If
            Next Pos
            mstrFileName = Right(mstrFQFileName, NameLen - DirChrPos)
            mstrDataDirectory = Left(mstrFQFileName, DirChrPos)
        End Set

    End Property


    Public Property BakFileDirectory() As String

        Get
            BakFileDirectory = mstrBakFileDirectory
        End Get

        Set(ByVal Value As String)
            If Right(Trim(Value), 1) <> "\" Then Value = Trim(Value) & "\"
            If Len(Dir(Value, FileAttribute.Directory)) = 0 Then
                If MsgBox("Directory " & Value & " does not exist; create it?", MsgBoxStyle.YesNo, "WinVIVA - Directory Doesn't Exist") = MsgBoxResult.Yes Then
                    On Error GoTo MakeFailed
                    MkDir(Value)
                    On Error GoTo 0
                    mstrBakFileDirectory = Value
                Else
                    MsgBox("Directory set to " & VIVADIR, MsgBoxStyle.OkOnly, "WinVIVA - Default Directory Selected")
                    mstrBakFileDirectory = VIVADIR
                End If
            Else
                mstrBakFileDirectory = Value
            End If
            Exit Property

MakeFailed:
            MsgBox("Failure attempting to create " & Value & "; directory set to " & VIVADIR, MsgBoxStyle.OkOnly, "WinVIVA - Can't Create Directory")
            mstrBakFileDirectory = VIVADIR
            Exit Property
        End Set

    End Property

    Public Property DataDirectory() As String

        Get
            DataDirectory = mstrDataDirectory
        End Get

        Set(ByVal Value As String)
            mstrDataDirectory = Value
        End Set

    End Property


    Public Property Title() As String

        Get
            Title = mstrTitle
        End Get

        Set(ByVal Value As String)
            mstrTitle = Value
        End Set

    End Property


    Public Property Desc() As String

        Get
            Desc = mstrDesc
        End Get

        Set(ByVal Value As String)
            mstrDesc = Value
        End Set

    End Property


    Public Property Saved() As Boolean

        Get
            Saved = mblnSaved
        End Get

        Set(ByVal Value As Boolean)
            mblnSaved = Value
        End Set

    End Property


    Public Property NumIterations() As Short

        Get
            NumIterations = mintNumIterations
        End Get

        Set(ByVal Value As Short)
            mintNumIterations = Value
        End Set

    End Property


    Public Property NumPoints() As Short

        Get
            NumPoints = mintNumPoints
        End Get

        Set(ByVal Value As Short)
            mintNumPoints = Value
        End Set

    End Property


    Public Property ShowDOSBox() As Boolean

        Get
            ShowDOSBox = mblnShowDOSBox
        End Get

        Set(ByVal Value As Boolean)
            mblnShowDOSBox = Value
        End Set

    End Property


    Public Property FrequencyResponse() As VIVAMain.FrequencyResponseValues

        Get
            FrequencyResponse = mFR
        End Get

        Set(ByVal Value As VIVAMain.FrequencyResponseValues)
            mFR = Value
        End Set

    End Property


    Public Property Units() As VIVAMain.Units

        Get
            Units = mUnits
        End Get

        Set(ByVal Value As VIVAMain.Units)
            mUnits = Value
        End Set

    End Property


    Public ReadOnly Property Riser(ByVal mRiserId As Short) As Riser

        Get
            Riser = mRiser(mRiserId)
        End Get

    End Property


    Public ReadOnly Property Water() As Water

        Get
            Water = mWater
        End Get

    End Property

  
    Public Property nRisers() As Short

        Get
            nRisers = mRisers

        End Get

        Set(ByVal Value As Short)
            mRisers = Value
            If mRisers = 2 And Riser(2) Is Nothing Then
                mRiser(2) = New Riser()
            End If
        End Set

    End Property

    Public Property sameRiserARiser() As Boolean

        Get
            sameRiserARiser = mSameRiserARiser
        End Get

        Set(ByVal Value As Boolean)
            mSameRiserARiser = Value
        End Set

    End Property

    Public Property DampingMethod() As VIVAMain.DampingMethod

        Get
            DampingMethod = mDampingMethod
        End Get

        Set(ByVal Value As VIVAMain.DampingMethod)
            mDampingMethod = Value
        End Set

    End Property


    Public Sub New()

        MyBase.New()

        '   create riser and water objects
        If mRisers = 0 Then
            mRisers = 1
        End If

        If mRiserId = 0 Then
            mRiserId = 1
        End If

        If mRisers = 1 Then
            mRiser(1) = New Riser

        Else
            mRiser(1) = New Riser
            mRiser(2) = New Riser
        End If

        mWater = New Water

        '   initialize property variables
        mstrTitle = ""
        mstrDesc = ""
        mstrFileName = ""
        mstrFQFileName = ""
        mFR = VIVAMain.FrequencyResponseValues.MultiFreq
        mUnits = VIVAMain.Units.Metric
        mDampingMethod = VIVAMain.DampingMethod.ModalDamping
        mBatchProcess = False
        mblnSaved = True

    End Sub


    Public Sub ConvertUnits()

        '   Toggle data to a new set of units
        If mUnits = VIVAMain.Units.English Then
            mUnits = VIVAMain.Units.Metric
        ElseIf mUnits = VIVAMain.Units.Metric Then
            mUnits = VIVAMain.Units.English
        Else
            Err.Raise(Number:=vbObjectError + 32010, Description:="Invalid Units In Unit Conversion")
        End If

    End Sub

    Public Sub CopyRiser()
        mRiser(2) = mRiser(1)

    End Sub
    ' See if the data necessary to run a model are present
    Public Function Complete() As String

        Dim LastSSPoint As Short

        If mRisers = 1 Then
            '   check the riser.  The segments are checked for completetness on input,
            '   so the existence of at least one segment is all that is required
            If mRiser(1).Segments.Count = 0 Then
                Complete = "There must be at least one segment in the riser!"
                Exit Function
            End If

            '   If the riser is of the SCR (Lazy-wave) type, there must be at least one point
            '   on the static solution stored
            If mRiser(1).RiserType = VIVAMain.RiserTypes.SCRorLazyWave Then
                With mRiser(1).StaticSolution
                    'How many points are in the static solution?
                    LastSSPoint = .Count
                    If LastSSPoint = 0 Then
                        Complete = "SCR risers require a static solution to be provided!"
                        Exit Function
                    End If
                    If .Item(LastSSPoint).Distance <> 0.0# Then
                        Complete = "Last point in static solution must corespond to " & "the touch-down point with zero distance!"
                        Exit Function
                    End If
                End With
            End If
        Else
            '   check the riser.  The segments are checked for completetness on input,
            '   so the existence of at least one segment is all that is required
            If mRiser(2).Segments.Count = 0 Then
                Complete = "There must be at least one segment in the riser!"
                Exit Function
            End If

            '   If the riser is of the SCR (Lazy-wave) type, there must be at least one point
            '   on the static solution stored
            If mRiser(2).RiserType = VIVAMain.RiserTypes.SCRorLazyWave Then
                With mRiser(2).StaticSolution
                    'How many points are in the static solution?
                    LastSSPoint = .Count
                    If LastSSPoint = 0 Then
                        Complete = "SCR risers require a static solution to be provided!"
                        Exit Function
                    End If
                    If .Item(LastSSPoint).Distance <> 0.0# Then
                        Complete = "Last point in static solution must corespond to " & "the touch-down point with zero distance!"
                        Exit Function
                    End If
                End With
            End If


            '   If the riser is of the SCR (Lazy-wave) type, there must be at least one point
            '   on the static solution stored
            If mRiser(1).RiserType = VIVAMain.RiserTypes.SCRorLazyWave Then
                With mRiser(1).StaticSolution
                    'How many points are in the static solution?
                    LastSSPoint = .Count
                    If LastSSPoint = 0 Then
                        Complete = "SCR risers require a static solution to be provided!"
                        Exit Function
                    End If
                    If .Item(LastSSPoint).Distance <> 0.0# Then
                        Complete = "Last point in static solution must corespond to " & "the touch-down point with zero distance!"
                        Exit Function
                    End If
                End With
            End If
        End If

        '   There must be a current velocity function stored
        If mWater.CurrentProfile.Count = 0 Then
            Complete = "There must be a current profile for this case!"
            Exit Function
        End If

        '   The case must have been saved at least once
        If mstrFQFileName = "" Then
            Complete = "You must save the case prior to execution!"
            Exit Function
        End If

        '   If we've not exited the function yet, all is well
        Complete = ""

    End Function

End Class