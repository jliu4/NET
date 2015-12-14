Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class LMRP
	
	Private msngHeight, msngDragDiameter As Single
	Private msngDryWeight, msngWetWeight As Single
	
	
	Public Property Height() As Single

        Get
            Height = msngHeight
        End Get

        Set(ByVal Value As Single)
            msngHeight = Value
        End Set

    End Property
	
	
	Public Property DryWeight() As Single

        Get
            DryWeight = msngDryWeight
        End Get

        Set(ByVal Value As Single)
            msngDryWeight = Value
        End Set

    End Property
	
	
	Public Property WetWeight() As Single

        Get
            WetWeight = msngWetWeight
        End Get

        Set(ByVal Value As Single)
            msngWetWeight = Value
        End Set

    End Property
	
	
	Public Property DragDiameter() As Single

        Get
            DragDiameter = msngDragDiameter
        End Get

        Set(ByVal Value As Single)
            msngDragDiameter = Value
        End Set

    End Property
	

    Public ReadOnly Property MassPerLength() As Single

        Get
            With CurProj.Riser(CurProj.RiserId)
                MassPerLength = msngDryWeight / (msngHeight * Gm) + .Segments.Item(1).IntXSection * .ContentDensity
            End With
        End Get

    End Property
	

    Public ReadOnly Property AddedMassPerLength() As Single

        Get
            AddedMassPerLength = CurProj.Water.Density * msngDragDiameter ^ 2 * Pi / 4.0#
        End Get

    End Property
	

    Public ReadOnly Property WetWeightPerLength() As Single

        Get
            With CurProj
                WetWeightPerLength = msngWetWeight / msngHeight + .Riser(.RiserId).Segments.Item(1).IntXSection * (.Riser(.RiserId).ContentDensity - .Water.Density) * Gm
            End With
        End Get

    End Property


    Public Sub New()

        MyBase.New()
        msngHeight = 0
        msngDryWeight = 0
        msngWetWeight = 0
        msngDragDiameter = 0

    End Sub

End Class