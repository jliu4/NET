Option Strict Off
Option Explicit On
Friend Class DrftForceCoef
	
	' drift force coefficient (0 - 180 deg even spacing) in ship local system
	
	' properties
	' Draft:        vessel draft (ft)
	' DrftFC:       drift force coefficient
	' DrftFCItem:   drift force coefficients at input frequency
	
	' methods
	' DrftFCAdd:    add drift force coefficients
	
	Private msngDraft As Single
	Private mcolDrftFC As Collection
	
	Public Sub New()
		MyBase.New()
        mcolDrftFC = New Collection
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
	
	Public ReadOnly Property DrftFC(ByVal Frequency As Single, ByVal Direction As Single) As Force
		Get
			
			' Input
			'   Frequency:  look-up wave frequency
			'   Direction:  look-up wave direction in vessel local system (rad)
			
			Dim N, i, Ns As Short
			Dim Rf As Single
			Dim Fy1, Fx1, Fm1 As Single
			Dim Fy2, Fx2, Fm2 As Single
			Dim FC As Force
			FC = New Force
			
			N = mcolDrftFC.Count()

            If N = 1 Then
                With FC
                    .Fx = mcolDrftFC.Item(1).FCx.ForceCoef(Direction)
                    .Fy = mcolDrftFC.Item(1).FCy.ForceCoef(Direction)
                    .MYaw = mcolDrftFC.Item(1).FCm.ForceCoef(Direction)
                End With
            ElseIf Frequency <= mcolDrftFC.Item(1).Freq Then
                Rf = Frequency / mcolDrftFC.Item(1).Freq
                With mcolDrftFC.Item(1)
                    FC.Fx = .FCx.ForceCoef(Direction) * Rf
                    FC.Fy = .FCy.ForceCoef(Direction) * Rf
                    FC.MYaw = .FCm.ForceCoef(Direction) * Rf
                End With
            ElseIf Frequency >= mcolDrftFC.Item(N).Freq Then
                Rf = mcolDrftFC.Item(N).Freq / Frequency
                With mcolDrftFC.Item(N)
                    FC.Fx = .FCx.ForceCoef(Direction) * Rf
                    FC.Fy = .FCy.ForceCoef(Direction) * Rf
                    FC.MYaw = .FCm.ForceCoef(Direction) * Rf
                End With
			Else
				For i = 2 To N
                    If Frequency < mcolDrftFC.Item(i).Freq Then
                        Ns = i
                        Exit For
                    End If
                Next 
				With mcolDrftFC.Item(Ns - 1)
                    Fx1 = .FCx.ForceCoef(Direction)
                    Fy1 = .FCy.ForceCoef(Direction)
                    Fm1 = .FCm.ForceCoef(Direction)
                End With
				With mcolDrftFC.Item(Ns)
                    Fx2 = .FCx.ForceCoef(Direction)
                    Fy2 = .FCy.ForceCoef(Direction)
                    Fm2 = .FCm.ForceCoef(Direction)
                End With
                Rf = (Frequency - mcolDrftFC.Item(Ns - 1).Freq) / (mcolDrftFC.Item(Ns).Freq - mcolDrftFC.Item(Ns - 1).Freq)
                With FC
					.Fx = Fx1 + (Fx2 - Fx1) * Rf
					.Fy = Fy1 + (Fy2 - Fy1) * Rf
					.MYaw = Fm1 + (Fm2 - Fm1) * Rf
				End With
			End If
			
			DrftFC = FC
			
		End Get
	End Property
	
	Public ReadOnly Property DrftFCItem(ByVal vntIndexKey As Object) As DrftFCbyFreq
		Get
			
			DrftFCItem = mcolDrftFC.Item(vntIndexKey)
			
		End Get
	End Property
	
	' methods
	
	Public Sub DrftFCAdd(ByRef Frequency As Single)
		
		' Input
		'   Frequency:  wave frequency of new set of drift coefficients
		
		Dim NewDrftFC As New DrftFCbyFreq
		
		NewDrftFC.Freq = Frequency
		
		mcolDrftFC.Add(NewDrftFC)
		
	End Sub
End Class