Option Strict Off
Option Explicit On

Friend Class AuxLine

    Private mintAuxLineNum As Short
    Private msngOD, msngID, msngContentDensity As Single


    Public Property AuxLineNum() As Short

        Get
            AuxLineNum = mintAuxLineNum
        End Get

        Set(ByVal Value As Short)
            Static blnAlreadySet As Boolean

            If Not blnAlreadySet Then
                blnAlreadySet = True
                mintAuxLineNum = Value
            Else
                Err.Raise(Number:=vbObjectError + 32108, Description:="AuxLine" & CStr(Value) & " already defined.")
            End If
        End Set

    End Property


    Public Property ID() As Single

        Get
            ID = msngID
        End Get

        Set(ByVal Value As Single)
            msngID = Value
        End Set

    End Property


    Public Property OD() As Single

        Get
            OD = msngOD
        End Get

        Set(ByVal Value As Single)
            msngOD = Value
        End Set

    End Property


    Public Property ContentDensity() As Single

        Get
            ContentDensity = msngContentDensity
        End Get

        Set(ByVal Value As Single)
            msngContentDensity = Value
        End Set

    End Property

End Class