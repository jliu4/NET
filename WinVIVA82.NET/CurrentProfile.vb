Option Strict Off
Option Explicit On

Friend Class CurrentProfile

    Implements System.Collections.IEnumerable

    Private intPairNum As Integer
    Private mcolCurrentProfile As New Collection
    Private mstrCurrentFile As String
    Private mstrProfileName As String


    Public Property CurrentFile() As String

        Get
            CurrentFile = mstrCurrentFile
        End Get

        Set(ByVal Value As String)
            mstrCurrentFile = Value
        End Set

    End Property


    Public Property ProfileName() As String

        Get
            ProfileName = mstrProfileName
        End Get

        Set(ByVal Value As String)
            mstrProfileName = Value
        End Set

    End Property


    Public Function Add(ByVal CDepth As Single, ByVal CVelU As Single, ByVal CVelV As Single, Optional ByVal CVelA As Single = 0) As CurrentPair

        Dim crpNew As New CurrentPair

        intPairNum = intPairNum + 1

        With crpNew
            .PairNum = intPairNum
            .Depth = CDepth
            '.Velocity = CVelocity
            .VelU = CVelU
            .VelV = CVelV
            .VelA = System.Math.Sqrt(.VelU ^ 2 + .VelV ^ 2)
            mcolCurrentProfile.Add(crpNew, CStr(.PairNum))
        End With

        Add = crpNew

    End Function


    Public Function Count() As Integer

        Count = mcolCurrentProfile.Count()

    End Function


    Public Sub Delete(ByVal index As Object)

        mcolCurrentProfile.Remove(index)

        If intPairNum > 0 Then
            intPairNum = intPairNum - 1
        Else
            intPairNum = 0
        End If

    End Sub


    Public Function Item(ByVal index As Object) As CurrentPair

        Item = mcolCurrentProfile.Item(index)

    End Function


    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        GetEnumerator = mcolCurrentProfile.GetEnumerator

    End Function


    ' This function interpolates the current velocity array
    ' to provide the current at a particular depth.  Simple
    ' linear interpolation is used.  Depth values outside the
    ' range of the array will be assigned the velocity of the
    ' appropriate endpoint (first or last)
    ' component : 1 for in-plane velocity, i.e., U
    '             2 for velocity normal to that plane,i.e., V
    Public Function VofD(ByVal Depth As Single, ByVal component As Single) As Single

        Dim Pair As CurrentPair
        Dim Matched As Boolean
        Dim LastDepth, LastVel As Single

        Matched = False

        LastDepth = mcolCurrentProfile.Item(1).Depth ' why index 1 -- begin from first pt

        If component = 1 Then
            LastVel = mcolCurrentProfile.Item(1).VelU ' Velocity be changed to VelU jjx 07/16
        Else
            LastVel = mcolCurrentProfile.Item(1).VelV
        End If

        If Depth > LastDepth Then
            For Each Pair In mcolCurrentProfile
                With Pair
                    If Depth >= LastDepth And Depth < .Depth Then
                        Matched = True
                        If .Depth = LastDepth Then
                            VofD = LastVel
                        Else
                            If component = 1 Then
                                VofD = LastVel + (.VelU - LastVel) * (Depth - LastDepth) / (.Depth - LastDepth)
                            Else
                                VofD = LastVel + (.VelV - LastVel) * (Depth - LastDepth) / (.Depth - LastDepth)
                            End If
                        End If
                        Exit For
                    Else
                        LastDepth = .Depth
                        If component = 1 Then
                            LastVel = .VelU
                        Else
                            LastVel = .VelV
                        End If
                    End If
                End With
            Next Pair
        End If

        If Not Matched Then
            VofD = LastVel
        End If

    End Function


    Public Sub New()

        MyBase.New()

        'added by jjx 07/15
        intPairNum = 0

    End Sub

End Class