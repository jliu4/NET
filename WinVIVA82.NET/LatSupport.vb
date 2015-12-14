Option Strict Off
Option Explicit On

Friend Class LatSupport
	
	Private msngStiffness, msngDist, msngDamping As Single
	
	
    Public Property Stiffness() As Single

        Get

            Stiffness = msngStiffness

        End Get

        Set(ByVal Value As Single)

            msngStiffness = Value

        End Set

    End Property
	
	
    Public Property Damping() As Single

        Get

            Damping = msngDamping

        End Get

        Set(ByVal Value As Single)

            msngDamping = Value

        End Set

    End Property
	
	
    Public Property Dist() As Single

        Get

            Dist = msngDist

        End Get
        Set(ByVal Value As Single)

            msngDist = Value

        End Set

    End Property

End Class