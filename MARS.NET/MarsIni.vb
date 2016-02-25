Option Strict Off
Option Explicit On
Friend Class MarsIni
    ' MarsIni       mars configuration data (mars.ini)

    ' properties
    ' WorkDir       working diretory

    ' methods
    ' AddPreFile    add previous project files
    ' CountPreFile  count the number of files in the list
    ' DelPreFile    delete a project file from the list
    ' PreFile       previous project files
    ' InputData     read from .ini file
    ' OutputData    write to .ini file

    Private mstrWorkDir As String
	Private mcolPreFile As Collection

    Public Sub New()
        MyBase.New()
        mstrWorkDir = MarsDir
        If mstrWorkDir = "" Then mstrWorkDir = MarsDirDefault & "\"

        mcolPreFile = New Collection
    End Sub

    ' properties

    Public Property WorkDir() As String
		Get
			
			WorkDir = mstrWorkDir
			
		End Get
		Set(ByVal Value As String)
			
			If Value <> "" Then mstrWorkDir = Value
			
		End Set
	End Property
	
	' methods
	
	Public Sub AddPreFile(ByVal FileName As String)
		
		Dim i, NumFiles As Short
		Dim Path, Name As String
		
		FileName = Trim(FileName)
		GetDirNFileName(FileName, Path, Name)
		If Name = "" Then Exit Sub
		
		NumFiles = mcolPreFile.Count()
		For i = 1 To NumFiles
			'    Debug.Print PreFile(i) = FileName, PreFile(i)
			If PreFile(i) = FileName Then
				mcolPreFile.Remove(i)
				Exit For
			End If
		Next i
		
		NumFiles = mcolPreFile.Count()
		If NumFiles < 1 Then
			mcolPreFile.Add(FileName)
		Else
			mcolPreFile.Add(FileName,  , 1)
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
	
	Public Function InputData(ByVal FileNum As Short) As Boolean
		
		Dim Keyword As String
		Dim InputVal As String
		Dim i, NumFiles As Short
		
		On Error GoTo ErrHandler
		
		Do Until EOF(FileNum)
			Keyword = LineInput(FileNum)
			Keyword = Trim(Keyword)
			
			Select Case Keyword
				Case "[Working Directory]"
					If Not EOF(FileNum) Then
						InputVal = LineInput(FileNum)
						mstrWorkDir = Trim(InputVal)
					End If
					
				Case "[Previous Project Files]"
					If Not EOF(FileNum) Then
						InputVal = LineInput(FileNum)
						NumFiles = CShort(Right(Trim(InputVal), 1))
						If NumFiles > MaxNumPreFiles Then NumFiles = MaxNumPreFiles
						
						For i = 1 To NumFiles
							If Not EOF(FileNum) Then
								InputVal = LineInput(FileNum)
								AddPreFile(Trim(InputVal))
							End If
						Next i
					End If
			End Select
		Loop 
		
		InputData = True
		Exit Function
		
ErrHandler: 
		
		InputData = False
		Exit Function
		
	End Function
	
	Public Function OutputData(ByVal FileNum As Short) As Boolean
		
		Dim i, NumFiles As Short
		
		On Error GoTo ErrHandler
		
		PrintLine(FileNum, "[Working Directory]")
		PrintLine(FileNum, mstrWorkDir)
		PrintLine(FileNum)
		
		PrintLine(FileNum, "[Previous Project Files]")
		NumFiles = mcolPreFile.Count()
		PrintLine(FileNum, "Num =" & NumFiles)
		For i = NumFiles To 1 Step -1
			PrintLine(FileNum, mcolPreFile.Item(i))
		Next i
		
		OutputData = True
		Exit Function
		
ErrHandler: 
		
		OutputData = False
		Exit Function
		
	End Function
End Class