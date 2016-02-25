Option Strict Off
Option Explicit On
Friend Class Project
    ' Project           project data
    ' Version 1.0
    ' 2001, Copyright DTCEL, All Rights Reserved

    ' properties
    ' Title             project title
    ' Description       project description
    ' Directory         project directory name
    ' FileName          project data file name
    ' Saved             data been saved
    ' Vessel            vessel data (including mooring)
    ' WellSites         wells location

    ' methods
    ' GetDirNFileName   get project directory and name
    ' ImportData        input data from file
    ' ExportData        output project data

    Private mstrTitle, mstrDescription As String
	Private mstrDirectory, mstrFileName As String
	Private mblnSaved As Boolean
    Private mclsVessel As Vessel
    Private mclsWellSites As Wells

    Public Sub New()
		MyBase.New()
        mstrTitle = ""
        mstrDescription = ""
        mstrDirectory = ""
        mstrFileName = ""
        mblnSaved = False

        mclsVessel = New Vessel
        mclsWellSites = New Wells
    End Sub

    ' properties
    Public Property Title() As String
        Get
            Title = mstrTitle

        End Get
        Set(ByVal Value As String)
			
			mstrTitle = Value
			
		End Set
	End Property

    Public Property Description() As String
		Get
			
			Description = mstrDescription
			
		End Get
		Set(ByVal Value As String)
			
			mstrDescription = Value
			
		End Set
	End Property

    Public Property Directory() As String
		Get
			
			Directory = mstrDirectory
			
		End Get
		Set(ByVal Value As String)
			
			Value = Trim(Value)
			If Right(Value, 1) <> "\" Then Value = Value & "\"
			mstrDirectory = Value
			
		End Set
	End Property

    Public Property FileName() As String
		Get
			
			FileName = mstrFileName
			
		End Get
		Set(ByVal Value As String)
			
			Value = Trim(Value)
			mstrFileName = Value
			
		End Set
	End Property

    Public Property Saved() As Boolean
		Get
			
			Saved = mblnSaved
			
		End Get
		Set(ByVal Value As Boolean)
			
			mblnSaved = Value
			
		End Set
	End Property
	
	Public ReadOnly Property Vessel() As Vessel
		Get
			
			Vessel = mclsVessel
			
		End Get
	End Property

    Public ReadOnly Property WellSites() As Wells
        Get

            WellSites = mclsWellSites

        End Get
    End Property

    ' methods
    Public Sub GetDirNFileName(ByVal NewFullName As String)
		
		Dim DirChrPos, Pos, NameLen As Short
		Dim Valid As Boolean
		
		NameLen = Len(NewFullName)
		DirChrPos = NameLen
		Valid = True
		For Pos = NameLen To 1 Step -1
			Select Case Mid(NewFullName, Pos, 1)
				Case "\"
					DirChrPos = Pos
					Exit For
				Case "/", ":", "*", "?", "<", ">", "|", Chr(34)
					Valid = False
			End Select
		Next Pos
		
		If Valid Then
			mstrFileName = Right(NewFullName, NameLen - DirChrPos)
		Else
			mstrFileName = ""
		End If
		mstrDirectory = Left(NewFullName, DirChrPos)
		
	End Sub
	
	Public Function ImportData(ByVal InputFile As String) As Boolean

        Dim Keyword, Entry As String
        Dim VslName As String
		Dim FileNum As Short
		
		On Error GoTo ErrorHandler
		
		ImportData = False


        FileNum = FreeFile
		
		FileOpen(FileNum, InputFile, OpenMode.Input, OpenAccess.Read)
		
		Do Until EOF(FileNum)
			Keyword = LineInput(FileNum)
			Keyword = Trim(Keyword)
			
			Select Case Keyword
				Case "[General]"
					If EOF(FileNum) Then Exit Function
					Input(FileNum, Entry)
					mstrTitle = Trim(Entry)
					If EOF(FileNum) Then Exit Function
					Input(FileNum, mstrDescription)

                Case "[Vessel]"
                    If Not mclsVessel.InputVsl(FileNum) Then Exit Function

                Case "[Riser]"
                    If Not mclsVessel.Riser.InputRiser(FileNum) Then Exit Function

                Case "[Mooring Lines]"
                    If Not mclsVessel.MoorSystem.InputML(FileNum) Then Exit Function

                Case "[Environment]"
                    If Not mclsVessel.EnvLoad.InputEnv(FileNum) Then Exit Function

                Case "[Well Sites]"
                    If Not mclsWellSites.InputWS(FileNum) Then Exit Function

                Case "[Damping Percent]"
					mclsVessel.ReadDampingPercent(FileNum)
				Case Else
			End Select
		Loop 
		
		FileClose(FileNum)


        ImportData = True
        Exit Function

ErrorHandler: 
		FileClose(FileNum)
		
	End Function
	
	Public Function ExportData(ByVal FileNum As Short) As Boolean
		
		ExportData = False
		
		PrintLine(FileNum, "[General]")
		WriteLine(FileNum, mstrTitle)
		WriteLine(FileNum, mstrDescription)

        PrintLine(FileNum, "[Vessel]") '  output currernt vessel station
		If Not mclsVessel.OutputVsl(FileNum) Then Exit Function

        PrintLine(FileNum, "[Riser]") '  output currernt vessel station
        If Not mclsVessel.Riser.OutputRiser(FileNum) Then Exit Function

        PrintLine(FileNum, "[Mooring Lines]")
        If Not mclsVessel.MoorSystem.OutputML(FileNum) Then Exit Function

        PrintLine(FileNum, "[Environment]")
		If Not mclsVessel.EnvLoad.OutputEnv(FileNum) Then Exit Function
		
		PrintLine(FileNum, "[Well Sites]")
		If Not mclsWellSites.OutputWS(FileNum) Then Exit Function

        PrintLine(FileNum, "[Damping Percent]")
        mclsVessel.SaveDampingPercent(FileNum)
		
		ExportData = True
		
	End Function
	
	Public Function VesselParticular(Optional ByVal sPath As String = "") As Boolean

        Dim VslName As String
        Dim InputOK, FileOpen_Renamed As Boolean
		Dim FS As Object
		
		On Error GoTo ErrorHandler
		
		FS = CreateObject("Scripting.FileSystemObject")
		
		FileOpen_Renamed = False
		VesselParticular = False
		
		With mclsVessel
			If Len(sPath) > 0 Then
                ' make a copy of the ship data files to MARSDir folder
                FS.CopyFile(sPath & "\" & .Name & ".*", mstrDirectory)
            End If

            VslName = mstrDirectory & .Name

            FileOpen(FileNumRes, VslName & ".rao", OpenMode.Input)
			FileOpen_Renamed = True
			InputOK = .InputRAOs(FileNumRes)
			FileClose(FileNumRes)
			FileOpen_Renamed = False
			If Not InputOK Then Exit Function
			
			FileOpen(FileNumRes, VslName & ".dam", OpenMode.Input)
			FileOpen_Renamed = True
			InputOK = .InputMDs(FileNumRes)
			FileClose(FileNumRes)
			FileOpen_Renamed = False
			If Not InputOK Then Exit Function
			
			FileOpen(FileNumRes, VslName & ".frc", OpenMode.Input)
			FileOpen_Renamed = True
			InputOK = .EnvLoad.InputFC(FileNumRes)
			FileClose(FileNumRes)
			FileOpen_Renamed = False
			If Not InputOK Then Exit Function
            'JLIU TODO comment out below, no example of .mor file 1/20/2016
            If 1 = 0 And .MoorSystem.MoorLineCount = 0 Then
                FileOpen(FileNumRes, VslName & ".mor", OpenMode.Input)
                FileOpen_Renamed = True
                InputOK = .InputFairleads(FileNumRes)
                FileClose(FileNumRes)
                FileOpen_Renamed = False
                If Not InputOK Then Exit Function
            End If
        End With
		
		VesselParticular = True
		Exit Function
		
ErrorHandler: 
		If FileOpen_Renamed Then FileClose(FileNumRes)
		
		MsgBox("Error reading RAO " & Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
		
	End Function
End Class