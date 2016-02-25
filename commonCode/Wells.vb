Option Strict Off
Option Explicit On
Friend Class Wells
	Implements System.Collections.IEnumerable
    ' Wells     well collection
    ' properties
    ' CurWellNo current well no

    ' methods
    ' Add       add wave component
    ' Clear     clear all components
    ' Count     count component number
    ' Delete    delete wave component
    ' Item      find wave component item

    ' InputWS   input well site data from file
    ' OutputWS  output well site data to file

    Private mintCurWellNo As Short
	Private mcolWells As Collection

    Public Sub New()
        MyBase.New()
        mcolWells = New Collection
        DefaultWell()
        mintCurWellNo = 1

    End Sub

    ' properties

    Public Property CurWellNo() As Short
        Get
            CurWellNo = mintCurWellNo

        End Get
        Set(ByVal Value As Short)

            mintCurWellNo = Value

        End Set
    End Property

    ' methods
    Public Sub Add(ByVal NameID As String, ByVal Xg As Double, ByVal Yg As Double, ByVal Depth As Double)

        Dim WellNew As New Well

        With WellNew
            .NameID = NameID
            .Xg = Xg
            .Yg = Yg
            .Depth = Depth
        End With
        mcolWells.Add(WellNew)

        mintCurWellNo = mcolWells.Count()

    End Sub

    Public Sub Clear()
		
		Dim NumWells, i As Short
		
		NumWells = mcolWells.Count()
		For i = NumWells To 1 Step -1
			mcolWells.Remove(i)
		Next i
		
	End Sub
	
	Public Function Count() As Integer
		
		Count = mcolWells.Count()
		
	End Function
	
	Private Sub Delete(ByVal Index As Object)
		
		mcolWells.Remove(Index)
		
	End Sub
	
	Public Function Item(ByVal Index As Object) As Well
		
		Item = mcolWells.Item(Index)
		
	End Function

    Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        GetEnumerator = mcolWells.GetEnumerator
    End Function
	
	Public Function InputWS(ByVal FileNum As Short) As Boolean
		
		Dim i As Short
		Dim NumWells As Short
		Dim NameID As String
		Dim Y, X, Depth As Double

        'Dim NewMoorLine As MoorLine

        InputWS = False
		
		On Error GoTo ErrorHandler
		
		Input(FileNum, NumWells)
		
		Clear()
		If NumWells <= 0 Then
			DefaultWell()
		Else
			For i = 1 To NumWells
				Input(FileNum, X)
				Input(FileNum, Y)
				Input(FileNum, Depth)
				Input(FileNum, NameID)
				Add(NameID, X, Y, Depth)
			Next i
		End If
		mintCurWellNo = 1
		
		InputWS = True
		
ErrorHandler: 
		
	End Function
	
	Public Function OutputWS(ByVal FileNum As Short) As Boolean
		
		Dim i As Short
		Dim NumWells As Short
		
		OutputWS = False
		
		NumWells = mcolWells.Count()
		WriteLine(FileNum, NumWells)
		
		For i = 1 To NumWells
			With mcolWells.Item(i)
                WriteLine(FileNum, .Xg, .Yg, .Depth, .NameID)
            End With
		Next i
		
		OutputWS = True
		
	End Function
	
	Private Sub DefaultWell()
		
		Dim wellDefault As New Well
		
		With wellDefault
			.NameID = "DEFAULT"
			.Xg = 0#
			.Yg = 0#
			.Depth = 5000
		End With
		
		mcolWells.Add(wellDefault)
		
	End Sub
End Class