Option Strict Off
Option Explicit On
Friend Class Current
	Implements System.Collections.IEnumerable
	
	' current properties
	
	' properties
	' Heading:      current heading (N clock-wise)
	' WaterDepth:   water depth (ft)
	' Profile:      current profile
	' NewEnum:      get id for each profile
	
	' methods
	' ProfileAdd:   add profile
	' ProfileCount: count profile
	' ProfileDelete:delete profile
	
	Private msngHeading As Single
	Private msngWaterDepth As Single
    Private mclsSurfVel As Single
    Private mcolProfile As Collection
	
	' initializing and terminating
	
	Public Sub New()
        MyBase.New()
        mcolProfile = New Collection
        msngHeading = -Math.PI
    End Sub

    ' properties
    Public Property SurfaceVel() As Single
        Get

            SurfaceVel = mclsSurfVel

        End Get
        Set(ByVal Value As Single)

            mclsSurfVel = Value
        End Set
    End Property

    Public Property Heading() As Single
		Get
			
			Heading = msngHeading
			
		End Get
		Set(ByVal Value As Single)
			
			msngHeading = Value
			
		End Set
	End Property
	
	
	Public Property WaterDepth() As Single
		Get
			
			WaterDepth = msngWaterDepth
			
		End Get
		Set(ByVal Value As Single)
			
			msngWaterDepth = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Profile(ByVal vntIndexKey As Object) As CurrentData
		Get
			
			Profile = mcolProfile.Item(vntIndexKey)
			
		End Get
	End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolProfile.GetEnumerator
    End Function

    Public ReadOnly Property ProfileCount() As Integer
        Get

            ProfileCount = mcolProfile.Count()

        End Get
    End Property

    ' methods

    Public Sub ProfileAdd(ByRef Depth As Single, ByRef Velocity As Single)

        ' Input
        '   Depth:      water depth (ft)
        '   Velocity:   current velocity (ft/s)

        Dim NewProfile As New CurrentData

        With NewProfile
            '       start from water surface
            If mcolProfile.Count() = 0 And Depth <> 0.0# Then
                .Depth = 0.0#
                .Velocity = Velocity
                mcolProfile.Add(NewProfile)
            End If

            .Depth = Depth
            .Velocity = Velocity
        End With

        mcolProfile.Add(NewProfile)

    End Sub

    Public Sub ProfileDelete(ByRef vntIndexKey As Object)

        mcolProfile.Remove(vntIndexKey)

    End Sub
End Class