Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class FatigueCurve

    Private mstrName As String
    Private msngA, msngB As Single


    Public Property Name() As String

        Get
            Name = mstrName
        End Get

        Set(ByVal Value As String)
            mstrName = Value
        End Set

    End Property


    Public Property A() As Single

        Get
            A = msngA
        End Get

        Set(ByVal Value As Single)
            msngA = Value
            If msngA < 0.0# Then msngA = -msngA
        End Set

    End Property


    Public Property B() As Single

        Get
            B = msngB
        End Get

        Set(ByVal Value As Single)
            msngB = Value
            If msngB < 0.0# Then msngB = -msngB
        End Set

    End Property

End Class