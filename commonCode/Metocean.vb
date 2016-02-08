Option Strict Off
Option Explicit On
Friend Class Metocean
	
	' metocean criteria
	
	' properties
	' Name:         name
	' Heading:      set uniform enviroment heading to wind wave and current
	' Current:      current
	' Wave:         wave
	' Wind:         wind
	
	Private mstrName As String
	Private mclsCurrent As Current
	Private mclsWave As Wave
	Private mclsWind As Wind
	
	Public Sub New()
		MyBase.New()
        mclsWind = New Wind
        mclsWave = New Wave
        mclsCurrent = New Current
	End Sub
	

	Public Property Name() As String
		Get
			
			Name = mstrName
			
		End Get
		Set(ByVal Value As String)
			
			mstrName = Value
			
		End Set
	End Property
	
	Public WriteOnly Property Heading() As Single
		Set(ByVal Value As Single)
			
			mclsWind.Heading = Value
			mclsCurrent.Heading = Value
            mclsWave.Heading = Value
            mclsWave.SwellHeading = Value

        End Set
	End Property
	
	Public ReadOnly Property Current() As Current
		Get
			
			Current = mclsCurrent
			
		End Get
	End Property
	
	Public ReadOnly Property Wave() As Wave
		Get
			
			Wave = mclsWave
			
		End Get
	End Property

    Public ReadOnly Property Wind() As Wind
        Get

            Wind = mclsWind

        End Get
    End Property

    Public Function WriteData(ByVal fnum As Integer, ByVal UDEF As Boolean) As Object
        Dim i As Short

        If InStr(mclsWave.SpectrumName, "PSMZ") > 0 Then
            mclsWave.gamma = ""
        End If
        If InStr(mclsWave.SwellSpectrumName, "PSMZ") > 0 Then
            mclsWave.Swellgamma = ""
        End If
        If UDEF Then
            PrintLine(fnum, Format(mclsWind.Velocity, "#0.0000"), mclsWind.Heading, mclsWave.SpectrumName, mclsWave.Height, mclsWave.Period, mclsWave.gamma, mclsWave.Heading, mclsCurrent.Heading, Format(mclsCurrent.SurfaceVel, "#0.0000"), mclsCurrent.ProfileCount, mclsWave.SwellHeight, mclsWave.SwellPeriod, mclsWave.SwellSpectrumName, mclsWave.Swellgamma, mclsWave.SwellHeading)
        Else
            PrintLine(fnum, Format(mclsWind.Velocity, "#0.0000"), mclsWind.Heading, mclsWave.SpectrumName, mclsWave.Height, mclsWave.Period, mclsWave.gamma, mclsWave.Heading, mclsCurrent.Heading, Format(mclsCurrent.SurfaceVel, "#0.0000"), mclsCurrent.ProfileCount)
        End If
        If mclsCurrent.ProfileCount > 1 Then
            PrintLine(fnum, "WD" & Space(10) & "Current Vel")
            For i = 1 To mclsCurrent.ProfileCount
                PrintLine(fnum, mclsCurrent.Profile(i).Depth, Format(mclsCurrent.Profile(i).Velocity, "#0.0000"))
            Next i
        End If


    End Function

    Public Function ReadData(ByVal fnum As Integer, ByVal UDEF As Boolean) As Object
        Dim i, TmpCount, numPairs As Short
        Dim aline As String
        Dim Fields() As String

        '  file already opened
        aline = LineInput(fnum)

        Fields = Split_Renamed(aline, " ")

        mclsWind.Velocity = Fields(1) '  skip case no
        mclsWind.Heading = Fields(2)

        mclsWave.SpectrumName = Fields(3)
        mclsWave.Height = Fields(4)
        mclsWave.Period = Fields(5)
        If InStr(mclsWave.SpectrumName, "JONH") > 0 Then
            mclsWave.gamma = Fields(6)
            mclsWave.Heading = Fields(7)
            mclsCurrent.Heading = Fields(8)
            mclsCurrent.SurfaceVel = CDbl(Fields(9))
            TmpCount = CShort(Fields(10))
            If UDEF Then
                mclsWave.SwellHeight = Fields(11)
                mclsWave.SwellPeriod = Fields(12)
                mclsWave.SwellSpectrumName = Fields(13)
                mclsWave.SwellGamma = Fields(14)
                mclsWave.SwellHeading = Fields(15)
            End If
        Else
            mclsWave.gamma = 0
            mclsWave.Heading = Fields(6)
            mclsCurrent.Heading = Fields(7)
            mclsCurrent.SurfaceVel = CDbl(Fields(8))
            TmpCount = CShort(Fields(9))

            If UDEF Then
                mclsWave.SwellHeight = Fields(10)
                mclsWave.SwellPeriod = Fields(11)
                mclsWave.SwellSpectrumName = Fields(12)
                mclsWave.SwellGamma = Fields(13)
                mclsWave.SwellHeading = Fields(14)

            End If
        End If

        If TmpCount > 1 Then
            aline = LineInput(fnum) ' discard header WD Vel
            numPairs = mclsCurrent.ProfileCount

            For i = numPairs To 1 Step -1

                mclsCurrent.ProfileDelete((1))
            Next i
            For i = 1 To TmpCount
                aline = LineInput(fnum)
                Fields = Split_Renamed(aline, " ")
                mclsCurrent.ProfileAdd(CSng(Fields(0)), CSng(Fields(1)))
            Next i
        Else
            mclsCurrent.ProfileDelete((1))
            mclsCurrent.ProfileAdd(0, CSng(Fields(TmpCount)))
        End If

    End Function

End Class