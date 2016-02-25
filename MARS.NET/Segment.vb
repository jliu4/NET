Option Strict Off
Option Explicit On
Friend Class Segment

    ' mooring line segment data

    ' properties
    ' (input)
    ' SegType:      CHAIN, WIRE, POLYESTER, SPRING
    ' Length:       segment length (ft)
    ' Diameter:     segment diameter (in)
    ' TotalLength:  total available length (ft)
    ' BS:           breaking strength (lbs)
    ' E1:           elastic modulus (1st) (psi)
    ' E2:           elastic modulus (2nd, for nolinear segment) (psi)
    ' UnitDryWeight:weight in air (lbs/ft)
    ' UnitWetWeight:Weight in water (lbs/ft)
    ' Buoy:         net buoyancy of buoy at the end of segment (lbs)
    ' BuoyLength:   buoy length (ft)
    ' FrictionCoef: friction coefficient of segment to the sea bottom

    ' (calculated)
    ' TotalWeight:  total wet weight of the segment (lbs)
    ' LengthStr:    stretched length due to tension (ft)
    ' XLow:         lower end scope (ft)
    ' YLow:         lower end depth (ft)
    ' AngUpp:       upper end angle (to horizon) (rad)
    ' AngLow:       lower end angle (to horizon) (rad)
    ' TenUpp:       upper end tension (lbs)

    Private mstrSegType As String
	Private msngLength As Single
	Private msngTotalLength As Single
	Private msngDiameter As Single
	Private msngBS As Single
	Private msngE1 As Single
	Private msngE2 As Single
	Private msngUnitDryWeight As Single
	Private msngUnitWetWeight As Single
	Private msngBuoy As Single
	Private msngBuoyLength As Single
	Private msngFrictionCoef As Single
	
	Private msngLengthStr As Single
	Private msngXLow As Single
	Private msngYLow As Single
	Private msngAngUpp As Single
	Private msngAngLow As Single
	Private msngTenUpp As Single

    ' properties as input

    Public Property SegType() As String
        Get
            SegType = mstrSegType

        End Get
        Set(ByVal Value As String)
			
			mstrSegType = Value
			
		End Set
	End Property

    Public Property Length() As Single
        Get
            Length = msngLength

        End Get
        Set(ByVal Value As Single)
			
			msngLength = Value
			msngLengthStr = Value
			
		End Set
	End Property

    Public Property TotalLength() As Single
		Get
			
			If msngTotalLength = 0# Then msngTotalLength = msngLength
			TotalLength = msngTotalLength
			
		End Get
		Set(ByVal Value As Single)
			
			msngTotalLength = Value
			
		End Set
	End Property

    Public Property Diameter() As Single
		Get
			
			Diameter = msngDiameter
			
		End Get
		Set(ByVal Value As Single)
			
			msngDiameter = Value
			
		End Set
	End Property

    Public Property BS() As Single
        Get
            BS = msngBS

        End Get
        Set(ByVal Value As Single)
			
			msngBS = Value
			
		End Set
	End Property

    Public Property E1() As Single
        Get

            E1 = msngE1

        End Get
        Set(ByVal Value As Single)

            msngE1 = Value

        End Set
    End Property

    Public Property E2() As Single
        Get

            E2 = msngE2

        End Get
        Set(ByVal Value As Single)

            msngE2 = Value

        End Set
    End Property

    Public Property UnitDryWeight() As Single
		Get
			
			UnitDryWeight = msngUnitDryWeight
			
		End Get
		Set(ByVal Value As Single)
			
			msngUnitDryWeight = Value
			
		End Set
	End Property

    Public Property UnitWetWeight() As Single
        Get

            UnitWetWeight = msngUnitWetWeight

        End Get
        Set(ByVal Value As Single)

            msngUnitWetWeight = Value

        End Set
    End Property

    Public Property Buoy() As Single
        Get

            Buoy = msngBuoy

        End Get
        Set(ByVal Value As Single)

            msngBuoy = Value

        End Set
    End Property

    Public Property BuoyLength() As Single
		Get
			
			BuoyLength = msngBuoyLength
			
		End Get
		Set(ByVal Value As Single)
			
			msngBuoyLength = Value
			
		End Set
	End Property

    Public Property FrictionCoef() As Single
		Get
			
			FrictionCoef = msngFrictionCoef
			
		End Get
		Set(ByVal Value As Single)
			
			msngFrictionCoef = Value
			
		End Set
	End Property
	
	' properties as calculated
	
	Public ReadOnly Property XArea() As Single
        Get
            XArea = PI * msngDiameter ^ 2 / 4.0#

        End Get
    End Property
	
	Public ReadOnly Property TotalWeight() As Single
        Get
            TotalWeight = msngLength * msngUnitWetWeight

        End Get
    End Property
	
	
	Public Property LengthStr() As Single
        Get
            LengthStr = msngLengthStr

        End Get
        Set(ByVal Value As Single)
			
			msngLengthStr = Value
			'    If msngLengthStr < msngLength Then msngLengthStr = msngLength
			
		End Set
	End Property

    Public Property XLow() As Single
		Get
			
			XLow = msngXLow
			
		End Get
		Set(ByVal Value As Single)
			
			msngXLow = Value
			
		End Set
	End Property

    Public Property YLow() As Single
		Get
			
			YLow = msngYLow
			
		End Get
		Set(ByVal Value As Single)
			
			msngYLow = Value
			
		End Set
	End Property

    Public Property AngUpp() As Single
		Get
			
			AngUpp = msngAngUpp
			
		End Get
		Set(ByVal Value As Single)
			
			msngAngUpp = Value
			
		End Set
	End Property

    Public Property AngLow() As Single
		Get
			
			AngLow = msngAngLow
			
		End Get
		Set(ByVal Value As Single)
			
			msngAngLow = Value
			
		End Set
	End Property

    Public Property TenUpp() As Single
		Get
			
			TenUpp = msngTenUpp
			
		End Get
		Set(ByVal Value As Single)
			
			msngTenUpp = Value
			
		End Set
	End Property
End Class