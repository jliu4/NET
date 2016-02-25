Option Strict Off
Option Explicit On
Friend Class RAOs
	
	' r.a.o.s (0 - 180 deg even spacing) in ship local system
	
	' properties
	' Draft:    vessel draft (ft)
	' RAO:      motion RAO
	' RAOsItem: motion RAOs at input frequency
	
	' methods
	' RAOsAdd:  add RAOs
	
	Private msngDraft As Single
	Private mcolRAOs As Collection
	
 
	Public Sub New()
		MyBase.New()
        mcolRAOs = New Collection
	End Sub

    ' properties	

    Public Property Draft() As Single
		Get
			
			Draft = msngDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngDraft = Value
			
		End Set
	End Property
	
	Public ReadOnly Property RAO(ByVal Frequency As Single, ByVal Direction As Single) As Motion
		Get
			
			' Input
			'   Frequency:  look-up wave frequency
			'   Direction:  look-up wave direction in vessel local system
			
			Dim N, i, Ns As Short
			Dim Rf As Single
			Dim Ry1, Rx1, Rr1 As Single
			Dim Ry2, Rx2, Rr2 As Single
			Dim r As New Motion
			
			N = mcolRAOs.Count()

            If N = 1 Then
                With r
                    .Surge = mcolRAOs.Item(1).RAOx.ForceCoef(Direction)
                    .Sway = mcolRAOs.Item(1).RAOy.ForceCoef(Direction)
                    .Yaw = mcolRAOs.Item(1).RAOr.ForceCoef(Direction)
                End With
            ElseIf Frequency <= mcolRAOs.Item(1).Freq Then
                Rf = Frequency / mcolRAOs.Item(1).Freq
                With mcolRAOs.Item(1)
                    r.Surge = .RAOx.ForceCoef(Direction) * Rf
                    r.Sway = .RAOy.ForceCoef(Direction) * Rf
                    r.Yaw = .RAOr.ForceCoef(Direction) * Rf
                End With
            ElseIf Frequency >= mcolRAOs.Item(N).Freq Then
                Rf = mcolRAOs.Item(N).Freq / Frequency
                With mcolRAOs.Item(N)
                    r.Surge = .RAOx.ForceCoef(Direction) * Rf
                    r.Sway = .RAOy.ForceCoef(Direction) * Rf
                    r.Yaw = .RAOr.ForceCoef(Direction) * Rf
                End With
			Else
				For i = 2 To N
                    If Frequency < mcolRAOs.Item(i).Freq Then
                        Ns = i
                        Exit For
                    End If
                Next 
				With mcolRAOs.Item(Ns - 1)
                    Rx1 = .RAOx.ForceCoef(Direction)
                    Ry1 = .RAOy.ForceCoef(Direction)
                    Rr1 = .RAOr.ForceCoef(Direction)
                End With
				With mcolRAOs.Item(Ns)
                    Rx2 = .RAOx.ForceCoef(Direction)
                    Ry2 = .RAOy.ForceCoef(Direction)
                    Rr2 = .RAOr.ForceCoef(Direction)
                End With
                Rf = (Frequency - mcolRAOs.Item(Ns - 1).Freq) / (mcolRAOs.Item(Ns).Freq - mcolRAOs.Item(Ns - 1).Freq)
                With r
					.Surge = Rx1 + (Rx2 - Rx1) * Rf
					.Sway = Ry1 + (Ry2 - Ry1) * Rf
					.Yaw = Rr1 + (Rr2 - Rr1) * Rf
				End With
			End If
			
			RAO = r
			
		End Get
	End Property
	
	Public ReadOnly Property RAOsItem(ByVal vntIndexKey As Object) As RAObyFreq
		Get
			
			RAOsItem = mcolRAOs.Item(vntIndexKey)
			
		End Get
	End Property
	
	' methods
	
	Public Sub RAOsAdd(ByRef Frequency As Single)
		
		' Input
		'   Frequency:  wave frequency of a set of new raos
		
		Dim NewRAOs As New RAObyFreq
		
		NewRAOs.Freq = Frequency
		
		mcolRAOs.Add(NewRAOs)
		
	End Sub
End Class