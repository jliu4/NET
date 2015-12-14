Option Strict Off
Option Explicit On

' Version 8.1
' 2014, Copyright DTCEL, All Rights Reserved
' Jin Liu May 2014

Friend Class AuxLines

    Implements System.Collections.IEnumerable

    Private mcolAuxLines As New Collection
    Private mblnAuxInEICalcs As Boolean
    Private intAuxLineNum As Short


    Public Property AuxInEICalcs() As Boolean

        Get
            AuxInEICalcs = mblnAuxInEICalcs
        End Get

        Set(ByVal Value As Boolean)
            mblnAuxInEICalcs = Value
        End Set

    End Property


    Public ReadOnly Property MassPerLength() As Single

        Get
            Dim mAuxLine As AuxLine
            Dim sum As Single

            sum = 0.0#

            For Each mAuxLine In mcolAuxLines
                With mAuxLine
                    sum = sum + .ContentDensity * .ID ^ 2
                End With
            Next mAuxLine

            MassPerLength = sum * Pi / 4.0#
        End Get

    End Property

    Public ReadOnly Property AddedMassPerLength() As Single

        Get
            Dim mAuxLine As AuxLine
            Dim sum As Single

            sum = 0.0#

            For Each mAuxLine In mcolAuxLines
                With mAuxLine
                    sum = sum + (.OD) ^ 2
                End With
            Next mAuxLine

            AddedMassPerLength = sum * CurProj.Water.Density * Pi / 4.0#
        End Get

    End Property


    Public ReadOnly Property WetWeightPerLength() As Single

        Get
            Dim mAuxLine As AuxLine
            Dim sum As Single

            sum = 0.0#

            For Each mAuxLine In mcolAuxLines
                With mAuxLine
                    sum = sum + (.ContentDensity - CurProj.Water.Density) * .ID ^ 2
                End With
            Next mAuxLine

            WetWeightPerLength = sum * Gm * Pi / 4.0#
        End Get

    End Property


    Public ReadOnly Property EI() As Single

        Get
            Const ES As Single = 205800000000.0#
            Dim mAuxLine As AuxLine
            Dim sum As Single

            sum = 0.0#

            For Each mAuxLine In mcolAuxLines
                With mAuxLine
                    sum = sum + .OD ^ 4 - .ID ^ 4
                End With
            Next mAuxLine

            EI = sum * ES * Pi / 64.0#
        End Get

    End Property


    Public Function Add() As AuxLine

        Dim auxNew As New AuxLine

        intAuxLineNum = intAuxLineNum + 1

        If intAuxLineNum > MaxAuxLines Then
            Err.Raise(Number:=vbObjectError + 32109, Description:="Maximum number of AuxLines is " & MaxAuxLines)
            Add = Nothing
        Else
            auxNew.AuxLineNum = intAuxLineNum
            mcolAuxLines.Add(auxNew, CStr(auxNew.AuxLineNum))
            Add = auxNew
        End If

    End Function


    Public Function Count() As Integer

        Count = mcolAuxLines.Count()

    End Function


    Public Sub Delete(ByVal Index As Object)

        mcolAuxLines.Remove(Index)

        If intAuxLineNum > 0 Then
            intAuxLineNum = intAuxLineNum - 1
        Else
            intAuxLineNum = 0
        End If

    End Sub


    Public Function Item(ByVal Index As Object) As AuxLine

        Item = mcolAuxLines.Item(Index)

    End Function


    'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
    'Public Function NewEnum() As stdole.IUnknown
    '
    'NewEnum = mcolAuxLines.GetEnumerator
    '
    'End Function

    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator

        GetEnumerator = mcolAuxLines.GetEnumerator

    End Function


    Public Sub New()

        MyBase.New()
        intAuxLineNum = 0
        mblnAuxInEICalcs = False

    End Sub

End Class