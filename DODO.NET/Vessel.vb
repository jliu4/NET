Option Strict Off
Option Explicit On
Friend Class Vessel
	
	' vessel properties and motion
	
	' properties
	' Name:         vessel name
	
	' ShipCurGlob:  current ship location in global system
	' ShipDesGlob:  design ship location in global system
	' ShipCurLocl:  current ship location in local system
	' ShipDraft:    ship current draft
	' ShipDraftSur: ship survival draft
	' ShipDraftOpr: ship operating draft
	
	' WaterDepth:   water depth
	
	' ShipMovGlob:  current ship movement in global system
	' ShipMovLocl:  current ship movement in local system
	
	' Riser:  riser system
	' EnvLoad:      environment load
	
	' SigLFM:       significant low frequency motion
	' SigWFM:       significant wave frequency motion
	
	Private mstrName As String
	
	Private mclsCriticalDamping As Motion
	Private mclsOriginalDampingPercent As Motion
	Private mclsDampingPercent As Motion
	Private mclsShipCurGlob As ShipGlobal
	Private mclsShipDesGlob As ShipGlobal
	Private mclsShipCurLocl As ShipLocal
	Private msngShipDraft As Single
	Private msngShipDraftSur As Single
	Private msngShipDraftOpr As Single
	
	Private msngWaterDepth As Single
	
	Private mclsRiser As Riser
	
	Private mclsEnvLoad As EnvLoad
	
	Private mcolShipMass As Collection
	Private mcolShipDamp As Collection
	Private mcolMotionRAOs As Collection
	Private mcolYawRateDrag As Collection
	
	Private Const DistTol As Single = 0.0001
	Private Const AngleTol As Single = 0.0001
	
	Private Const FreqS As Single = 0.02
	Private Const FreqE As Single = 2.4
	Private Const FreqD As Single = 0.02
	Private Const FreqE1 As Single = 0.2
	Private Const FreqE2 As Single = 1#
	Private Const FreqD1 As Single = 0.02
	Private Const FreqD2 As Single = 0.1
	
    Public Sub New()
        MyBase.New()
        mclsShipCurGlob = New ShipGlobal
        mclsShipDesGlob = New ShipGlobal
        mclsShipCurLocl = New ShipLocal

        mclsRiser = New Riser
        mclsEnvLoad = New EnvLoad

        mcolShipMass = New Collection
        mcolShipDamp = New Collection
        mcolMotionRAOs = New Collection
        mcolYawRateDrag = New Collection

        mclsDampingPercent = New Motion
        mclsOriginalDampingPercent = New Motion
        mclsCriticalDamping = New Motion
    End Sub
	
	
	Public Property Name() As String
		Get
			
			Name = mstrName
			
		End Get
		Set(ByVal Value As String)
			
			mstrName = Value
			
		End Set
	End Property
	
	Public ReadOnly Property ShipCurGlob() As ShipGlobal
		Get
			
			ShipCurGlob = mclsShipCurGlob
			
			mclsEnvLoad.ShipHead = mclsShipCurGlob.Heading
			
		End Get
	End Property
	
	Public ReadOnly Property ShipDesGlob() As ShipGlobal
		Get
			
			ShipDesGlob = mclsShipDesGlob
			
		End Get
	End Property
	
	Public ReadOnly Property ShipCurLocl() As ShipLocal
		Get
			
			ShipCurLocl = mclsShipCurLocl
			
		End Get
	End Property
	
	Public ReadOnly Property Riser() As Riser
		Get
			
			Riser = mclsRiser
			
		End Get
	End Property
	
	Public ReadOnly Property EnvLoad() As EnvLoad
		Get
			
			EnvLoad = mclsEnvLoad
			
		End Get
	End Property
	
	
	Public Property ShipDraft() As Single
		Get
			
			ShipDraft = msngShipDraft
			
		End Get
		Set(ByVal Value As Single)
			
			msngShipDraft = Value
			mclsEnvLoad.ShipDraft = Value
			
		End Set
	End Property
	
	
	Public Property ShipDraftSur() As Single
		Get
			
			ShipDraftSur = msngShipDraftSur
			
		End Get
		Set(ByVal Value As Single)
			
			msngShipDraftSur = Value
			
		End Set
	End Property
	
	
	Public Property ShipDraftOpr() As Single
		Get
			
			ShipDraftOpr = msngShipDraftOpr
			
		End Get
		Set(ByVal Value As Single)
			
			msngShipDraftOpr = Value
			
		End Set
	End Property
	
	
	Public Property WaterDepth() As Single
		Get
			
			WaterDepth = msngWaterDepth
			
		End Get
		Set(ByVal Value As Single)
			
			msngWaterDepth = Value
			
		End Set
	End Property
	
	Public ReadOnly Property ShipMass(ByVal Draft As Single) As Motion
		Get
			
			Dim N, i, Ns As Short
			Dim Rd As Single
			Dim My1, Mx1, Mm1 As Single
			Dim My2, Mx2, Mm2 As Single
			Dim mass As New Motion
			
			N = mcolShipMass.Count()
			
			If N = 1 Then
				With mass
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Surge = mcolShipMass.Item(1).VirMassSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Sway = mcolShipMass.Item(1).VirMassSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Yaw = mcolShipMass.Item(1).VirMassYaw
				End With
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Draft <= mcolShipMass.Item(1).Draft Then
					With mcolShipMass.Item(1)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx1 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My1 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm1 = .VirMassYaw
					End With
					With mcolShipMass.Item(2)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx2 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My2 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm2 = .VirMassYaw
					End With
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(2).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolShipMass.Item(1).Draft) / (mcolShipMass.Item(2).Draft - mcolShipMass.Item(1).Draft)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf Draft >= mcolShipMass.Item(N).Draft Then 
					With mcolShipMass.Item(N - 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx1 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My1 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm1 = .VirMassYaw
					End With
					With mcolShipMass.Item(N)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx2 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My2 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm2 = .VirMassYaw
					End With
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(N - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolShipMass.Item(N - 1).Draft) / (mcolShipMass.Item(N).Draft - mcolShipMass.Item(N - 1).Draft)
				Else
					For i = 2 To N
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(i).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Draft <= mcolShipMass.Item(i).Draft Then
							Ns = i
							Exit For
						End If
					Next 
					With mcolShipMass.Item(Ns - 1)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx1 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My1 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm1 = .VirMassYaw
					End With
					With mcolShipMass.Item(Ns)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mx2 = .VirMassSurge
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						My2 = .VirMassSway
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().VirMassYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Mm2 = .VirMassYaw
					End With
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(Ns - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass(Ns).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipMass().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolShipMass.Item(Ns - 1).Draft) / (mcolShipMass.Item(Ns).Draft - mcolShipMass.Item(Ns - 1).Draft)
				End If
				With mass
					.Surge = Mx1 + (Mx2 - Mx1) * Rd
					.Sway = My1 + (My2 - My1) * Rd
					.Yaw = Mm1 + (Mm2 - Mm1) * Rd
				End With
			End If
			
			ShipMass = mass
			
		End Get
	End Property
	
	Public ReadOnly Property DampingPercent() As Motion
		Get
			DampingPercent = mclsDampingPercent
		End Get
	End Property
	
	Public ReadOnly Property CriticalDamping() As Motion
		Get
			CriticalDamping = mclsCriticalDamping
		End Get
	End Property
	
	Public ReadOnly Property ShipDamp(ByVal Draft As Single) As Motion
		Get
			
			Dim Damp As New Motion
			
			If mclsDampingPercent.Surge > 0 And mclsDampingPercent.Sway > 0 And mclsDampingPercent.Yaw > 0 Then
				If mclsDampingPercent.Surge = mclsOriginalDampingPercent.Surge And mclsDampingPercent.Sway = mclsOriginalDampingPercent.Sway And mclsDampingPercent.Yaw = mclsOriginalDampingPercent.Yaw Then
					ShipDamp = GetOriginalDamping(Draft)
				Else
					Damp.Surge = mclsDampingPercent.Surge / 100 * mclsCriticalDamping.Surge
					Damp.Sway = mclsDampingPercent.Sway / 100 * mclsCriticalDamping.Sway
					Damp.Yaw = mclsDampingPercent.Yaw / 100 * mclsCriticalDamping.Yaw
					ShipDamp = Damp
				End If
			Else
				ShipDamp = GetOriginalDamping(Draft)
			End If
			
		End Get
	End Property
	
	Public ReadOnly Property ShipYawRateDrag(ByVal Draft As Single) As Single
		Get
			
			Dim N, i, Ns As Short
			Dim Rd As Single
			Dim YRD2, YRD1, YRD As Single
			
			
			N = mcolYawRateDrag.Count()
			
			
			If N = 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				YRD = mcolYawRateDrag.Item(1).YawRateDrag
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Draft <= mcolYawRateDrag.Item(1).Draft Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD1 = mcolYawRateDrag.Item(1).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD2 = mcolYawRateDrag.Item(2).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(2).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolYawRateDrag.Item(1).Draft) / (mcolYawRateDrag.Item(2).Draft - mcolYawRateDrag.Item(1).Draft)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf Draft >= mcolYawRateDrag.Item(N).Draft Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD1 = mcolYawRateDrag.Item(N - 1).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD2 = mcolYawRateDrag.Item(N).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(N - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolYawRateDrag.Item(N - 1).Draft) / (mcolYawRateDrag.Item(N).Draft - mcolYawRateDrag.Item(N - 1).Draft)
				Else
					For i = 2 To N
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(i).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Draft <= mcolYawRateDrag.Item(i).Draft Then
							Ns = i
							Exit For
						End If
					Next 
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD1 = mcolYawRateDrag.Item(Ns - 1).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().YawRateDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					YRD2 = mcolYawRateDrag.Item(Ns).YawRateDrag
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(Ns - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag(Ns).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolYawRateDrag().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Rd = (Draft - mcolYawRateDrag.Item(Ns - 1).Draft) / (mcolYawRateDrag.Item(Ns).Draft - mcolYawRateDrag.Item(Ns - 1).Draft)
				End If
				YRD = YRD1 + (YRD2 - YRD1) * Rd
			End If
			
			ShipYawRateDrag = YRD
			
		End Get
	End Property
	
	Public ReadOnly Property ShipMotion(ByVal InitialLoc As ShipGlobal, ByVal FinalLoc As ShipGlobal) As Motion
		Get
			
			Dim dy, Alpha, dx, Drz As Single
			Dim Move As Motion
			
			With InitialLoc
				dx = FinalLoc.Xg - .Xg
				dy = FinalLoc.Yg - .Yg
				Drz = -FinalLoc.Heading + .Heading
				If Drz >= PI Then Drz = Drz - PI * 2#
				If Drz <= -PI Then Drz = Drz + PI * 2#
				Alpha = PI / 2 - .Heading
			End With
			
			Move = New Motion
			With Move
				.Surge = System.Math.Cos(Alpha) * dx + System.Math.Sin(Alpha) * dy
				.Sway = -System.Math.Sin(Alpha) * dx + System.Math.Cos(Alpha) * dy
				.Yaw = Drz
			End With
			
			ShipMotion = Move
			
		End Get
	End Property
	
	Public Function GetSigWFM(ByRef Cancelled As Boolean, ByRef frmProgress As System.Windows.Forms.Form, Optional ByVal Heading As Single = 0) As Motion
		
		Dim MotionRAO As Motion
		Dim Direct, Freq, Sw As Single
		Dim Sway, Surge, Yaw As Single
		
		Dim IniProg As Short
		
		'UPGRADE_ISSUE: Control prgProgress could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        'IniProg = frmProgress.prgProgress.Value
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Heading) Then Heading = mclsShipCurGlob.Heading
		
		Direct = Heading - mclsEnvLoad.EnvCur.Wave.Heading
		Surge = 0#
		Sway = 0#
		Yaw = 0#
		
		For Freq = FreqS To FreqE Step FreqD
			Sw = mclsEnvLoad.EnvCur.Wave.Spectrum(Freq)
			MotionRAO = RAO(msngShipDraft, Freq, Direct)
			
			With MotionRAO
				Surge = Surge + .Surge ^ 2 * Sw * FreqD
				Sway = Sway + .Sway ^ 2 * Sw * FreqD
				Yaw = Yaw + .Yaw ^ 2 * Sw * FreqD
			End With
			
			'UPGRADE_NOTE: Object MotionRAO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			MotionRAO = Nothing
			System.Windows.Forms.Application.DoEvents()
			
			If Cancelled Then Exit For
			'UPGRADE_ISSUE: Control Progress could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            'frmProgress.Progress = IniProg + (Freq - FreqS) / (FreqE - FreqS) * 15
		Next Freq
		
		GetSigWFM = New Motion
		With GetSigWFM
			.Surge = 2# * System.Math.Sqrt(Surge)
			.Sway = 2# * System.Math.Sqrt(Sway)
			.Yaw = 2# * System.Math.Sqrt(Yaw)
		End With
		
	End Function
	
	Public Function GetSigLFM(ByRef StiffLocl As Force, ByRef Cancelled As Boolean, ByRef frmProgress As System.Windows.Forms.Form, Optional ByRef ShipLoc As ShipGlobal = Nothing) As Motion
		
		Dim Sf As Force
		Dim mass, Damp As Motion
		Dim Freq, Heading As Single
		Dim Freq0, FreqD0, Freq1 As Single
		Dim Sxy, Sxx, Sxz As Single
		Dim Sway, Surge, Yaw As Single
		
		Dim IniProg As Short
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(ShipLoc) Then ShipLoc = mclsShipCurGlob
		Heading = ShipLoc.Heading
		'UPGRADE_ISSUE: Control prgProgress could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        'IniProg = frmProgress.prgProgress.Value
		
		Surge = 0#
		Sway = 0#
		Yaw = 0#
		
		mass = ShipMass(msngShipDraft)
		Damp = ShipDamp(msngShipDraft)
		
		Freq = 0#
		Freq0 = 0#
		Do While Freq <= FreqE2
			If Freq < FreqE1 Then
				Freq1 = Freq + FreqD1
			Else
				Freq1 = Freq + FreqD2
			End If
			
			Sf = mclsEnvLoad.DrftFrSpec(Freq, Heading)
			Sxx = 1# / System.Math.Sqrt((mass.Surge * Freq ^ 2 - StiffLocl.Fx) ^ 2 + (Damp.Surge * Freq) ^ 2)
			Sxy = 1# / System.Math.Sqrt((mass.Sway * Freq ^ 2 - StiffLocl.Fy) ^ 2 + (Damp.Sway * Freq) ^ 2)
			Sxz = 1# / System.Math.Sqrt((mass.Yaw * Freq ^ 2 - StiffLocl.MYaw) ^ 2 + (Damp.Yaw * Freq) ^ 2)
			
			With Sf
				Surge = Surge + .Fx * Sxx ^ 2 * (Freq1 - Freq0) * 0.5
				Sway = Sway + .Fy * Sxy ^ 2 * (Freq1 - Freq0) * 0.5
				Yaw = Yaw + .MYaw * Sxz ^ 2 * (Freq1 - Freq0) * 0.5
			End With
			
			Freq0 = Freq
			Freq = Freq1
			
			'UPGRADE_NOTE: Object Sf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			Sf = Nothing
			System.Windows.Forms.Application.DoEvents()
			
			If Cancelled Then Exit Do
			'UPGRADE_ISSUE: Control Progress could not be resolved because it was within the generic namespace Form. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
            'frmProgress.Progress = IniProg + Freq / FreqE2 * 30
		Loop 
		
		GetSigLFM = New Motion
		With GetSigLFM
			.Surge = 2# * System.Math.Sqrt(Surge)
			.Sway = 2# * System.Math.Sqrt(Sway)
			.Yaw = 2# * System.Math.Sqrt(Yaw)
		End With
		
	End Function
	
	Public Function InputVsl(ByVal FileNum As Short) As Boolean
		
		Dim VslName As String
		Dim DesY, DesX, DesH As Single
		Dim CurY, CurX, CurH As Single
		Dim OprD, SurD, CurD As Single
		Dim WD As Single
		Dim mass, TopTen, Dia As Single
		
		InputVsl = False
		
		On Error GoTo ErrorHandler
		
		Input(FileNum, VslName)
		mstrName = VslName
		
		Input(FileNum, DesX)
		Input(FileNum, DesY)
		Input(FileNum, DesH)
		Input(FileNum, SurD)
		Input(FileNum, OprD)
		With mclsShipDesGlob
			.Xg = DesX
			.Yg = DesY
			.Heading = DesH
		End With
		msngShipDraftSur = SurD
		msngShipDraftOpr = OprD
		
		Input(FileNum, CurX)
		Input(FileNum, CurY)
		Input(FileNum, CurH)
		Input(FileNum, CurD)
		With mclsShipCurGlob
			.Xg = CurX
			.Yg = CurY
			.Heading = CurH
		End With
		ShipDraft = CurD
		
		Input(FileNum, WD)
		msngWaterDepth = WD
		mclsRiser.Length = WD
		
		Input(FileNum, mass)
		Input(FileNum, TopTen)
		Input(FileNum, Dia)
		With mclsRiser
			.TopTen = TopTen
			.mass = mass
			.Dia = Dia
		End With
		
		InputVsl = True
		Exit Function
ErrorHandler: 
		InputVsl = False
		MsgBox("Error reading vessel data: " & Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		
	End Function
	
	Public Function OutputVsl(ByVal FileNum As Short) As Boolean
		
		OutputVsl = False
		
		WriteLine(FileNum, mstrName)
		
		With mclsShipDesGlob
			WriteLine(FileNum, .Xg, .Yg, .Heading, msngShipDraftSur, msngShipDraftOpr)
		End With
		
		With mclsShipCurGlob
			WriteLine(FileNum, .Xg, .Yg, .Heading, msngShipDraft)
		End With
		
		WriteLine(FileNum, msngWaterDepth)
		
		With mclsRiser
			WriteLine(FileNum, .mass, .TopTen, .Dia)
		End With
		
		OutputVsl = True
		
	End Function
	
	Private Function RAO(ByVal Draft As Single, ByVal Frequency As Single, ByVal Direction As Single) As Motion
		
		Dim N, i, Ns As Short
		Dim Rd As Single
		Dim Ry1, Rx1, Rr1 As Single
		Dim Ry2, Rx2, Rr2 As Single
		Dim r As Motion
		
		N = mcolMotionRAOs.Count()
		
		If N = 1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			r = mcolMotionRAOs.Item(1).RAO(Frequency, Direction)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Draft <= mcolMotionRAOs.Item(1).Draft Then
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(1).RAO(Frequency, Direction)
				With r
					Rx1 = .Surge
					Ry1 = .Sway
					Rr1 = .Yaw
				End With
				'UPGRADE_NOTE: Object r may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				r = Nothing
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(2).RAO(Frequency, Direction)
				With r
					Rx2 = .Surge
					Ry2 = .Sway
					Rr2 = .Yaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(2).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolMotionRAOs.Item(1).Draft) / (mcolMotionRAOs.Item(2).Draft - mcolMotionRAOs.Item(1).Draft)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf Draft >= mcolMotionRAOs.Item(N).Draft Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(N - 1).RAO(Frequency, Direction)
				With r
					Rx1 = .Surge
					Ry1 = .Sway
					Rr1 = .Yaw
				End With
				'UPGRADE_NOTE: Object r may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				r = Nothing
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(N).RAO(Frequency, Direction)
				With r
					Rx2 = .Surge
					Ry2 = .Sway
					Rr2 = .Yaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(N - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolMotionRAOs.Item(N - 1).Draft) / (mcolMotionRAOs.Item(N).Draft - mcolMotionRAOs.Item(N - 1).Draft)
			Else
				For i = 2 To N
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(i).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Draft <= mcolMotionRAOs.Item(i).Draft Then
						Ns = i
						Exit For
					End If
				Next 
				
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(Ns - 1).RAO(Frequency, Direction)
				With r
					Rx1 = .Surge
					Ry1 = .Sway
					Rr1 = .Yaw
				End With
				'UPGRADE_NOTE: Object r may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				r = Nothing
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().RAO. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				r = mcolMotionRAOs.Item(Ns).RAO(Frequency, Direction)
				With r
					Rx2 = .Surge
					Ry2 = .Sway
					Rr2 = .Yaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(Ns - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs(Ns).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolMotionRAOs().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolMotionRAOs.Item(Ns - 1).Draft) / (mcolMotionRAOs.Item(Ns).Draft - mcolMotionRAOs.Item(Ns - 1).Draft)
			End If
			With r
				.Surge = Rx1 + (Rx2 - Rx1) * Rd
				.Sway = Ry1 + (Ry2 - Ry1) * Rd
				.Yaw = Rr1 + (Rr2 - Rr1) * Rd
			End With
		End If
		
		RAO = r
		
	End Function
	
	Public Function InputRAOs(ByVal FileNum As Short) As Boolean
		
		Dim kj, j, i, k, dk As Short
		Dim NumDir, NumDraft, NumFreq As Short
		Dim NewRAOs As RAOs
		Dim Freq, Draft, Dummy As Single
		Dim RAOy, RAOx, RAOr As Single
		Dim msg As String
		
		InputRAOs = False
		
		On Error GoTo ErrorHandler
		
		Input(FileNum, NumDraft)
		If NumDraft <> 2 Then GoTo ErrorHandler '  for now assume only two draft, first surv and then operating
		
		Do While mcolMotionRAOs.Count() > 0
			mcolMotionRAOs.Remove((1))
		Loop 
		
		Dim Fields() As String
		Dim TmpStr As String
		Dim FieldCount As Short
		
		For i = 1 To NumDraft
			Input(FileNum, Draft)
			Input(FileNum, NumDir)
			Input(FileNum, NumFreq)
			
			NewRAOs = New RAOs
			NewRAOs.Draft = Draft
			
			' This is for limiting certain users
			'        If i = 1 And Abs(Draft - 49.2) > 0.01 Then
			'            msg = "This version is for Noble Clyde Boudreaux only."
			'            MsgBox msg
			'            End
			'        End If
			'        If i = 2 And Abs(Draft - 66.3) > 0.01 Then
			'            msg = "This version is for Noble Clyde Boudreaux only."
			'            MsgBox msg
			'            End
			'        End If
			
			For j = 1 To NumDir
				Input(FileNum, Dummy)
				
				For k = 1 To NumFreq
					TmpStr = LineInput(FileNum)
					'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Fields = Split_Renamed(TmpStr, " ")
					FieldCount = UBound(Fields) - LBound(Fields) + 1
					
					If FieldCount <> 14 Then GoTo ErrorHandler
					
					Freq = CDbl(Fields(1))
					RAOx = CDbl(Fields(2))
					RAOy = CDbl(Fields(4))
					RAOr = CDbl(Fields(12))
					
					'   Input #FileNum, Dummy, Freq, RAOx, Dummy, RAOy, Dummy, _
					''       Dummy, Dummy, Dummy, Dummy, Dummy, Dummy, RAOr, Dummy
					
					RAOr = RAOr * Degrees2Radians
					
					If j = 1 Then NewRAOs.RAOsAdd(Freq)
					NewRAOs.RAOsItem(k).RAOx.ForceCoefAdd(RAOx)
					NewRAOs.RAOsItem(k).RAOy.ForceCoefAdd(RAOy)
					NewRAOs.RAOsItem(k).RAOr.ForceCoefAdd(RAOr)
				Next k
			Next j
			
			mcolMotionRAOs.Add(NewRAOs)
		Next i
		
		InputRAOs = True
		Exit Function
		
ErrorHandler: 
		If Len(Err.Description) > 0 Then
			MsgBox(Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		End If
	End Function
	
	Public Function ReadDampingPercent(ByVal fnum As Integer) As Object
		Dim tmp2, tmp1, tmp3 As Single
		Input(fnum, tmp1)
		Input(fnum, tmp2)
		Input(fnum, tmp3)
		If tmp1 > 10 And tmp2 > 10 And tmp3 > 10 Then
			With mclsDampingPercent
				.Surge = tmp1
				.Sway = tmp2
				.Yaw = tmp3
			End With
		Else
			With mclsDampingPercent
				.Surge = mclsOriginalDampingPercent.Surge
				.Sway = mclsOriginalDampingPercent.Sway
				.Yaw = mclsOriginalDampingPercent.Yaw
			End With
		End If
	End Function
	
	Public Function SaveDampingPercent(ByVal fnum As Integer) As Object
		With mclsDampingPercent
			WriteLine(fnum, .Surge, .Sway, .Yaw)
		End With
	End Function
	
	Private Function GetOriginalDamping(ByVal Draft As Single) As Motion
		Dim N, i, Ns As Short
		Dim Rd As Single
		Dim Dy1, Dx1, Dm1 As Single
		Dim Dy2, Dx2, Dm2 As Single
		Dim Damp As New Motion
		
		N = mcolShipDamp.Count()
		
		If N = 1 Then
			With Damp
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Surge = mcolShipDamp.Item(1).DampSurge
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Sway = mcolShipDamp.Item(1).DampSway
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Yaw = mcolShipDamp.Item(1).DampYaw
			End With
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Draft <= mcolShipDamp.Item(1).Draft Then
				With mcolShipDamp.Item(1)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx1 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy1 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm1 = .DampYaw
				End With
				With mcolShipDamp.Item(2)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx2 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy2 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm2 = .DampYaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(2).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolShipDamp.Item(1).Draft) / (mcolShipDamp.Item(2).Draft - mcolShipDamp.Item(1).Draft)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf Draft >= mcolShipDamp.Item(N).Draft Then 
				With mcolShipDamp.Item(N - 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx1 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy1 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm1 = .DampYaw
				End With
				With mcolShipDamp.Item(N)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx2 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy2 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm2 = .DampYaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(N - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(N).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolShipDamp.Item(N - 1).Draft) / (mcolShipDamp.Item(N).Draft - mcolShipDamp.Item(N - 1).Draft)
			Else
				For i = 2 To N
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(i).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Draft <= mcolShipDamp.Item(i).Draft Then
						Ns = i
						Exit For
					End If
				Next 
				With mcolShipDamp.Item(Ns - 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx1 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy1 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm1 = .DampYaw
				End With
				With mcolShipDamp.Item(Ns)
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSurge. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dx2 = .DampSurge
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampSway. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dy2 = .DampSway
					'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().DampYaw. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dm2 = .DampYaw
				End With
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(Ns - 1).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp(Ns).Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolShipDamp().Draft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Rd = (Draft - mcolShipDamp.Item(Ns - 1).Draft) / (mcolShipDamp.Item(Ns).Draft - mcolShipDamp.Item(Ns - 1).Draft)
			End If
			With Damp
				.Surge = Dx1 + (Dx2 - Dx1) * Rd
				.Sway = Dy1 + (Dy2 - Dy1) * Rd
				.Yaw = Dm1 + (Dm2 - Dm1) * Rd
			End With
		End If
		GetOriginalDamping = Damp
	End Function
	
	Public Function InputMDs(ByVal FileNum As Short) As Boolean
		' INPUT MASS, ADDED MASS AND DAMPING
		
		Dim i, j As Short
		Dim NumDraft As Short
		Dim NewMass As ShipMass
		Dim NewDamp As ShipDamp
		Dim NewYRD As ShipYawRate
		Dim IdStr As String
		Dim Draft As Single
		Dim MDz, MDx, MDy, YRD As Single
		
		InputMDs = False
		
		On Error GoTo ErrorHandler
		
		'    Input #FileNum, msngShipDraftSur, msngShipDraftOpr
		
		Input(FileNum, NumDraft)
		
		Do While mcolShipMass.Count() > 0
			mcolShipMass.Remove((1))
		Loop 
		Do While mcolShipDamp.Count() > 0
			mcolShipDamp.Remove((1))
		Loop 
		Do While mcolYawRateDrag.Count() > 0
			mcolYawRateDrag.Remove((1))
		Loop 
		
		For i = 1 To NumDraft
			Input(FileNum, Draft)
			
			For j = 1 To 3
				Input(FileNum, IdStr)
				
				'UPGRADE_ISSUE: Constant Default was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				Select Case IdStr
					Case "MASS"
						NewMass = New ShipMass
						NewMass.Draft = Draft
						
						Input(FileNum, MDx)
						Input(FileNum, MDy)
						Input(FileNum, MDz)
						With NewMass
							.MassSurge = MDx
							.MassSway = MDy
							.MassYaw = MDz
						End With
						
						Input(FileNum, MDx)
						Input(FileNum, MDy)
						Input(FileNum, MDz)
						With NewMass
							.AddMassSurge = MDx
							.AddMassSway = MDy
							.AddMassYaw = MDz
						End With
						
						mcolShipMass.Add(NewMass)
						
					Case "DAMP"
						NewDamp = New ShipDamp
						NewDamp.Draft = Draft
						
						Input(FileNum, MDx)
						Input(FileNum, MDy)
						Input(FileNum, MDz)
						With NewDamp
							.DampSurge = MDx
							.DampSway = MDy
							.DampYaw = MDz
						End With
						
						mcolShipDamp.Add(NewDamp)
						
					Case "YAWRATE"
						NewYRD = New ShipYawRate
						NewYRD.Draft = Draft
						
						Input(FileNum, YRD)
						NewYRD.YawRateDrag = YRD
						
						mcolYawRateDrag.Add(NewYRD)
						
                    Case Else
                        Exit Function

                End Select
			Next j
		Next i
		
		InputMDs = True
		
ErrorHandler: 
		
	End Function
	
	Public Function NewShipLoc(ByRef ShipCurLoc As ShipGlobal, ByRef ShipMotion As Motion) As ShipGlobal
		
		Dim Alpha As Single
		Dim NewLoc As ShipGlobal
		
		NewLoc = New ShipGlobal
		
		With ShipCurLoc
			Alpha = PI / 2# - .Heading
			NewLoc.Xg = .Xg + System.Math.Cos(Alpha) * ShipMotion.Surge - System.Math.Sin(Alpha) * ShipMotion.Sway
			NewLoc.Yg = .Yg + System.Math.Sin(Alpha) * ShipMotion.Surge + System.Math.Cos(Alpha) * ShipMotion.Sway
			NewLoc.Heading = .Heading - ShipMotion.Yaw
			'       heading direction is clockwise while yaw is counter-clockwise
		End With
		
		NewShipLoc = NewLoc
		
	End Function
	
	Public Sub NextPosnVel(ByVal dt As Single, ByRef Acc() As Motion, ByRef Vel() As Motion, ByRef Pos() As ShipGlobal)
		
		Dim V23 As Single
		Dim NewPos As ShipGlobal
		Dim Move As New Motion
		
		With Move
			V23 = Vel(1).Surge + (5# * Acc(1).Surge - Acc(0).Surge) * dt / 8#
			Vel(2).Surge = Vel(1).Surge + (3# * Acc(1).Surge - Acc(0).Surge) * dt / 2#
			.Surge = (Vel(1).Surge + 4# * V23 + Vel(2).Surge) * dt / 6
			
			V23 = Vel(1).Sway + (5# * Acc(1).Sway - Acc(0).Sway) * dt / 8#
			Vel(2).Sway = Vel(1).Sway + (3# * Acc(1).Sway - Acc(0).Sway) * dt / 2#
			.Sway = (Vel(1).Sway + 4# * V23 + Vel(2).Sway) * dt / 6
			
			V23 = Vel(1).Yaw + (5# * Acc(1).Yaw - Acc(0).Yaw) * dt / 8#
			Vel(2).Yaw = Vel(1).Yaw + (3# * Acc(1).Yaw - Acc(0).Yaw) * dt / 2#
			.Yaw = (Vel(1).Yaw + 4# * V23 + Vel(2).Yaw) * dt / 6
		End With
		
		NewPos = NewShipLoc(Pos(1), Move)
		With NewPos
			Pos(2).Xg = .Xg
			Pos(2).Yg = .Yg
			Pos(2).Heading = .Heading
		End With
		
		'UPGRADE_NOTE: Object NewPos may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		NewPos = Nothing
		'UPGRADE_NOTE: Object Move may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Move = Nothing
		
	End Sub
	
	Public Sub UpdatePosVelnAcc(ByRef Acc() As Motion, ByRef Vel() As Motion, ByRef Pos() As ShipGlobal)
		
		Dim i As Short
		
		For i = 0 To 1
			Acc(i).Surge = Acc(i + 1).Surge
			Vel(i).Surge = Vel(i + 1).Surge
			Pos(i).Xg = Pos(i + 1).Xg
			Acc(i).Sway = Acc(i + 1).Sway
			Vel(i).Sway = Vel(i + 1).Sway
			Pos(i).Yg = Pos(i + 1).Yg
			Acc(i).Yaw = Acc(i + 1).Yaw
			Vel(i).Yaw = Vel(i + 1).Yaw
			Pos(i).Heading = Pos(i + 1).Heading
		Next i
		
	End Sub
End Class