Option Strict Off
Option Explicit On

Friend Class PreFile

    ' PreFile       previous project files listed on menu
    ' Version 1.0
    ' 2003, Copyright DTCEL, All Rights Reserved


    ' methods
    ' AddPreFile    add previous project files
    ' CountPreFile  count the number of files in the list
    ' DelPreFile    delete a project file from the list
    ' PreFile       previous project files

    Private mcolPreFile As Collection


    ' initiation
    Public Sub New()
        MyBase.New()
        mcolPreFile = New Collection
    End Sub


    Protected Overrides Sub Finalize()

        mcolPreFile = Nothing

        MyBase.Finalize()

    End Sub


    ' methods
    Public Sub AddPreFile(ByVal FileName As String)

        Dim i, NumFiles As Short
        Dim Path, Name As String
        'Dim length As Short

        FileName = Trim(FileName)

        'initialization
        Path = ""
        Name = ""

        GetDirNFileName(FileName, Path, Name)

        If Name = "" Then Exit Sub

        NumFiles = mcolPreFile.Count()
        For i = 1 To NumFiles
            If PreFile(i) = FileName Then
                mcolPreFile.Remove(i)
                Exit For
            End If
        Next i

        NumFiles = mcolPreFile.Count()
        If NumFiles < 1 Then
            mcolPreFile.Add(FileName)
        Else
            mcolPreFile.Add(FileName, , 1)
        End If

        NumFiles = mcolPreFile.Count()
        If NumFiles > MaxNumPreFiles Then mcolPreFile.Remove(NumFiles)

    End Sub


    Public Function CountPreFile() As Integer

        CountPreFile = mcolPreFile.Count()

    End Function


    Public Sub DelPreFile(ByVal Index As Object)

        mcolPreFile.Remove(Index)

    End Sub


    Public Function PreFile(ByVal Index As Object) As String

        PreFile = mcolPreFile.Item(Index)

    End Function

End Class