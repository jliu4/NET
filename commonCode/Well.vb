Option Strict Off
Option Explicit On
Friend Class Well
	
	' well properties
	
	' properties
	' NameID:       well name and/or id
	' Xg:           x coordinate (ft)
	' Yg:           y coordinate (ft)
	' Depth:        water depth (ft)
	
	Private mstrNameID As String
    Private msngXg As Decimal
    Private msngYg As Decimal
    Private msngDepth As Decimal

    ' properties


    Public Property NameID() As String
        Get

            NameID = mstrNameID

        End Get
        Set(ByVal Value As String)

            mstrNameID = Value

        End Set
    End Property


    Public Property Xg() As Decimal
        Get

            Xg = msngXg

        End Get
        Set(ByVal Value As Decimal)

            msngXg = Value

        End Set
    End Property


    Public Property Yg() As Decimal
        Get

            Yg = msngYg

        End Get
        Set(ByVal Value As Decimal)

            msngYg = Value

        End Set
    End Property


    Public Property Depth() As Decimal
        Get

            Depth = msngDepth

        End Get
        Set(ByVal Value As Decimal)

            msngDepth = Value

        End Set
    End Property
End Class