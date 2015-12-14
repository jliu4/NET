Option Strict Off
Option Explicit On

Friend Class CurrentPair

    Private mintPairNum As Short
    Private msngDepth As Single ', msngVelocity As Single
    Private msngVelU As Single
    Private msngVelV As Single
    Private msngVelA As Single


    Public Property PairNum() As Short

        Get
            PairNum = mintPairNum
        End Get

        Set(ByVal Value As Short)
            Static blnAlreadySet As Boolean

            If Not blnAlreadySet Then
                blnAlreadySet = True
                mintPairNum = Value
            Else
                Err.Raise(Number:=vbObjectError + 32107, Description:="Depth/Velocity" & CStr(Value) & " already defined.")
            End If
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


    Public Property VelU() As Single

        Get
            VelU = msngVelU
        End Get

        Set(ByVal Value As Single)
            msngVelU = Value
        End Set

    End Property


    Public Property VelV() As Single

        Get
            VelV = msngVelV
        End Get

        Set(ByVal Value As Single)
            msngVelV = Value
        End Set

    End Property


    Public Property VelA() As Single

        Get
            VelA = msngVelA
        End Get

        Set(ByVal Value As Single)
            msngVelA = Value
        End Set

    End Property

End Class