Option Strict Off
Option Explicit On

Friend Class Water
	
    Private msngDensity As Single ', msngCM As Single
	Private mCurrentProfile As New CurrentProfile
	Private msngViscosity As Single 'jjx 07/19
	
	
	Public Property Density() As Single

        Get
            Density = msngDensity
        End Get

		Set(ByVal Value As Single)
            If Value < 0 Then
                Err.Raise(Number:=vbObjectError + 32102, Description:="Water density cannot be negative")
            Else
                msngDensity = Value
            End If
        End Set

	End Property
	

	

    'Public Property Get CM() As Single

    '   CM = msngCM

    'End Property

    'Public Property Let CM(ByVal NewValue As Single)

    '   If NewValue < 0 Then
    '      Err.Raise Number:=vbObjectError + 32103, _
    ''     Description:="CM cannot be negative"
    'Else
    '   msngCM = NewValue
    ' End If

    'End Property
	
	Public ReadOnly Property CurrentProfile() As CurrentProfile

        Get
            CurrentProfile = mCurrentProfile
        End Get

    End Property
	
	
	Public Property Viscosity() As Single

        Get
            Viscosity = msngViscosity
        End Get

		Set(ByVal Value As Single)
            If Value < 0 Then
                Err.Raise(Number:=vbObjectError + 32104, Description:="Water viscosity cannot be negative")
            Else
                msngViscosity = Value
            End If
        End Set

	End Property


    Public Sub New()

        MyBase.New()
        msngDensity = 1025.0#
        msngViscosity = 1.114 * 10 ^ (-6)
        ' msngCM = 2#

    End Sub

End Class