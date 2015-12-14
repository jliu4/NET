Option Strict Off
Option Explicit On

Friend Class BottomPackage

    Private mLMRP As New LMRP
    Private mBOP As New BOP


    Public ReadOnly Property LMRP() As LMRP

        Get
            LMRP = mLMRP
        End Get

    End Property


    Public ReadOnly Property BOP() As BOP

        Get
            BOP = mBOP
        End Get

    End Property


    Public Sub New()

        MyBase.New()
        mLMRP = New LMRP
        mBOP = New BOP

    End Sub

End Class