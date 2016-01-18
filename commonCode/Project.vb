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
	Private mclsEnvironment As EnvLoad
	Private mclsWellSites As Wells

    ' initialization and termination

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        mstrTitle = ""
        mstrDescription = ""
        mstrDirectory = ""
        mstrFileName = ""
        mblnSaved = False

        mclsVessel = New Vessel
        mclsWellSites = New Wells

    End Sub
    Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		
		'UPGRADE_NOTE: Object mclsVessel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsVessel = Nothing
		'UPGRADE_NOTE: Object mclsWellSites may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mclsWellSites = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
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
		
		Const NumCheck As Short = 4
		Dim i As Short
		Dim CheckInput(NumCheck) As Boolean
		Dim Keyword, Entry As String
		Dim VslName As String
		Dim FileNum As Short
		
		On Error GoTo ErrorHandler
		
		ImportData = False
		For i = 0 To NumCheck
			CheckInput(i) = False
		Next i
		
		
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
					CheckInput(0) = True

                Case "[Vessel]"
                    If mclsVessel.InputVsl(FileNum) Then CheckInput(1) = True

                Case "[Mooring Lines]"
                    If mclsVessel.MoorSystem.InputML(FileNum) Then CheckInput(2) = True

                Case "[Environment]"
					If mclsVessel.EnvLoad.InputEnv(FileNum) Then CheckInput(3) = True
					
				Case "[Well Sites]"
					If mclsWellSites.InputWS(FileNum) Then CheckInput(4) = True
					
				Case "[Damping Percent]"
					mclsVessel.ReadDampingPercent(FileNum)
				Case Else
			End Select
		Loop 
		
		FileClose(FileNum)
		
		For i = 0 To NumCheck
			If Not CheckInput(i) Then Exit Function
		Next i
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
		'    Write #FileNum, "Version " & App.Major & "." & App.Minor            '  This will break version compatibility
		
		PrintLine(FileNum, "[Vessel]") '  output currernt vessel station
		If Not mclsVessel.OutputVsl(FileNum) Then Exit Function
        '    Write #FileNum, mstrVesselDataPath     '  This will break version compatibility
        '    Instead of saving the RAO path, make a copy of Vessel Data to MARSDir

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
		
		'UPGRADE_NOTE: FileOpen was upgraded to FileOpen_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
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
				'UPGRADE_WARNING: Couldn't resolve default property of object FS.CopyFile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FS.CopyFile(sPath & "\" & .Name & ".*", mstrDirectory)
			End If
			
			VslName = mstrDirectory & "\" & .Name
			
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

            If .MoorSystem.MoorLineCount = 0 Then
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