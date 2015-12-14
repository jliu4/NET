Option Strict Off
Option Explicit On

Friend Class StaticSolutionPoint
	
	Private mintPointNum As Short
	Private msngDistance As Single
	Private msngAxialStaticTension, msngChangeInStaticTension As Single
    Private msngInPlaneAngle, msngOutPlaneAngle, msngDepth As Single
	Private msngUCurrentVelocity As Single
    Private msngVCurrentVelocity As Single
    Private msngYdist As Single
    Private msngZdist As Single

	
	
	Public Property PointNum() As Short

        Get
            PointNum = mintPointNum
        End Get

        Set(ByVal Value As Short)
            Static blnAlreadySet As Boolean

            If Not blnAlreadySet Then
                blnAlreadySet = True
                mintPointNum = Value
            Else
                Err.Raise(Number:=vbObjectError + 32112, Description:="Static Solution Point " & CStr(Value) & " already defined.")
            End If
        End Set

    End Property
	
	
	Public Property Distance() As Single

        Get
            Distance = msngDistance
        End Get

        Set(ByVal Value As Single)
            msngDistance = Value
        End Set

    End Property
	
	
	Public Property AxialStaticTension() As Single

        Get
            AxialStaticTension = msngAxialStaticTension
        End Get

        Set(ByVal Value As Single)
            msngAxialStaticTension = Value
        End Set

    End Property
	
	
	Public Property InPlaneAngle() As Single

        Get
            InPlaneAngle = msngInPlaneAngle
        End Get

        Set(ByVal Value As Single)
            msngInPlaneAngle = Value
        End Set

    End Property

    Public Property OutPlaneAngle() As Single

        Get
            OutPlaneAngle = msngOutPlaneAngle
        End Get

        Set(ByVal Value As Single)
            msngOutPlaneAngle = Value
        End Set

    End Property
	
	
	Public Property Depth() As Single

        Get
            Depth = msngDepth
        End Get

        Set(ByVal Value As Single)
            msngDepth = Value
        End Set

    End Property
	
	
	Public Property UCurrentVelocity() As Single

        Get
            UCurrentVelocity = msngUCurrentVelocity
        End Get

        Set(ByVal Value As Single)
            msngUCurrentVelocity = Value
        End Set

    End Property
	
	
	Public Property VCurrentVelocity() As Single

        Get
            VCurrentVelocity = msngVCurrentVelocity
        End Get

        Set(ByVal Value As Single)
            msngVCurrentVelocity = Value
        End Set

    End Property

    Public Property Ydist() As Single

        Get
            Ydist = msngYdist
        End Get

        Set(ByVal Value As Single)
            msngYdist = Value
        End Set

    End Property


    Public Property Zdist() As Single

        Get
            Zdist = msngZdist
        End Get

        Set(ByVal Value As Single)
            msngZdist = Value
        End Set

    End Property
	
	
    Public Property ChangeInStaticTension() As Single

        Get
            ChangeInStaticTension = msngChangeInStaticTension
        End Get

        Set(ByVal Value As Single)
            msngChangeInStaticTension = Value
        End Set

    End Property

End Class