Option Strict Off
Option Explicit On
Friend Class Anchor
	
	' anchor location
	
	' properties
	' NeedUpdate:   coordinates changed
	' Xg:           x coordinate in global system (ft)
	' Yg:           y coordinate in global system (ft)
	' HoldCap:      anchor holding capacity
	' Model:        anchor model
	' Remark:       remark for anchor type and holding capacity
	
	Private mblnNeedUpdate As Boolean
	Private msngXg As Single
    Private msngYg As Single
    Private msngNode As Integer 'node in the aqwa model and default as 910i
    Private msngHoldCap As Single
	Private mstrModel As String
	Private mstrRemark As String

    Friend Property NeedUpdate() As Boolean
		Get
			
			NeedUpdate = mblnNeedUpdate
			
		End Get
		Set(ByVal Value As Boolean)
			
			mblnNeedUpdate = Value
			
		End Set
	End Property

    Public Property Xg() As Single
		Get
			
			Xg = msngXg
			
		End Get
		Set(ByVal Value As Single)
			
			msngXg = Value
			mblnNeedUpdate = True
			
		End Set
	End Property

    Public Property Yg() As Single
        Get

            Yg = msngYg

        End Get
        Set(ByVal Value As Single)

            msngYg = Value
            mblnNeedUpdate = True

        End Set
    End Property

    Public Property Node() As Integer
        Get

            Node = msngNode

        End Get
        Set(ByVal Value As Integer)

            msngNode = Value
            mblnNeedUpdate = True

        End Set
    End Property


    Public Property HoldCap() As Single
		Get
			
			HoldCap = msngHoldCap
			
		End Get
		Set(ByVal Value As Single)
			
			msngHoldCap = Value
			
		End Set
	End Property
	
	
	Public Property Model() As String
		Get
			
			Model = mstrModel
			
		End Get
		Set(ByVal Value As String)
			
			mstrModel = Value
			
		End Set
	End Property
	
	
	Public Property Remark() As String
		Get
			
			Remark = mstrRemark
			
		End Get
		Set(ByVal Value As String)
			
			mstrRemark = Value
			
		End Set
	End Property
End Class