Option Strict Off
Option Explicit On
Friend Class ExcelRunner
	
	Dim oxApp As Microsoft.Office.Interop.Excel.Application
	Dim oxBooks As Microsoft.Office.Interop.Excel.Workbooks
	Dim oxBook As Microsoft.Office.Interop.Excel.Workbook
	Dim oxTmpBook As Microsoft.Office.Interop.Excel.Workbook
	Dim AFfname1, ABfname, AFfname, AFfname2 As String
	
	Sub GetAQWAResults(ByVal sTitle As String, ByRef SubTitle As String, ByVal NumCases As Short, ByVal NumLines As Short, ByVal IsDataEnglishUnit As Boolean, ByVal IsDamaged As Short, ByRef BreakLine() As Short, ByVal WD As Single, ByRef VesselInitPos() As Single, ByRef preT() As Double, ByRef BS() As Double, ByVal WorkDir As String, ByVal blnPFLH As Boolean)
		Dim nrowcop As Object
		Dim nrowd As Object
		Dim NLA As Object
		Dim DamLineCur As Object
		Dim nd As Object
		Dim vslzacc As Object
		Dim VslRzWF As Object
		Dim VslRyWF As Object
		Dim VslRxWF As Object
		Dim VslZWF As Object
		Dim VslYWF As Object
		Dim VslXWF As Object
		Dim NX As Object
		Dim VslRzLF As Object
		Dim VslRyLF As Object
		Dim VslRxLF As Object
		Dim VslZLF As Object
		Dim VslYLF As Object
		Dim VslXLF As Object
		Dim k As Object
		Dim Fields As Object
		Dim tmpVal As Object
		Dim nl As Object
		Dim ThruRz As Object
		Dim ThruY As Object
		Dim ThruX As Object
		Dim WaveRz As Object
		Dim WaveY As Object
		Dim WaveX As Object
		Dim WindRz As Object
		Dim WindY As Object
		Dim WindX As Object
		Dim CurrRz As Object
		Dim CurrY As Object
		Dim CurrX As Object
		Dim VslRz As Object
		Dim VslRy As Object
		Dim VslRx As Object
		Dim VslZ As Object
		Dim VslY As Object
		Dim VslX As Object
		Dim StiffRz As Object
		Dim StiffY As Object
		Dim stiffx As Object
		
		ABfname = "abrun.lis"
		AFfname = "afrun.lis"
		AFfname1 = "afrun1.lis"
		AFfname2 = "afrun2.lis"
		
		Dim brokenlineno(40, NumCases) As Short
		Dim TotalNumBroken(NumCases) As Short
		
		Dim Ten(40) As Single
		Dim TenLF(40) As Single
		Dim TenWF(40) As Single
		Dim NameDir As String
		Dim NumLineA As Short
		Dim NCase As Short
		Dim CaseNam As String
		Dim NameSheet As String
		Dim Row, Col As Short
		Dim WindVel, CurrVel, CurrDir, WindDir As Single
		Dim WaveTp, WaveHs, WaveDir As Single
		Dim SwellTp, SwellHs, SwellDir As Single
		
		
		Dim i, j As Short
		
		' for now assume only one line is damaged for each case
		' but post-processer will be able to handler multiple lines damaged in each case
		For i = 1 To NumCases
			TotalNumBroken(i) = 1
			For j = 1 To TotalNumBroken(i)
				brokenlineno(j, i) = BreakLine(i)
			Next j
		Next i
		
		NameDir = "Case" ' assume subdir name "case"
		
		oxBook = oxApp.Workbooks.Add(My.Application.Info.DirectoryPath & "\MooringResults.xlt")
        ' set intro-input values

        oxBook.Sheets("Intro-Input").Activate()
        oxApp.Range("MainTitle").Formula = sTitle
		oxApp.Range("SubTitle").Formula = SubTitle
		oxApp.Range("OperTension").Formula = preT(1)
		oxApp.Range("BreakStrength").Formula = BS(1)
		oxApp.Range("WD").Formula = WD
		oxApp.Range("xx").Formula = VesselInitPos(1)
		oxApp.Range("yy").Formula = VesselInitPos(2)
		oxApp.Range("zz").Formula = VesselInitPos(3)
		oxApp.Range("rx").Formula = VesselInitPos(4)
		oxApp.Range("ry").Formula = VesselInitPos(5)
		oxApp.Range("rz").Formula = VesselInitPos(6)
		oxApp.Visible = True
		
		Dim EndRow As Short
		For NCase = 1 To NumCases
			If IsDamaged = 1 Then
				CaseNam = NCase & "D1"
			ElseIf IsDamaged = 2 Then 
				CaseNam = NCase & "D2"
			Else
				CaseNam = CStr(NCase)
			End If
			
			NameSheet = NameDir & CaseNam
			CaseNam = NameDir & CaseNam
			
			If IsDamaged Then
				NumLineA = NumLines - TotalNumBroken(NCase)
			Else
				NumLineA = NumLines
			End If

            oxBook.Sheets("MOORING1").Copy(After:=oxBook.Sheets(NCase + 1))
            oxBook.Sheets("MOORING1 (2)").Name = CaseNam
            oxBook.ActiveSheet.Cells(72, 16) = CaseNam

            '--------------------------------------------------------------------------------------
            oxApp.Workbooks.OpenText(Filename:=WorkDir & "\" & CaseNam & "\" & ABfname, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object() {1, 1})

            With oxApp

                If Not .ActiveSheet.Cells.Find("11curr") Is Nothing Then
                    .ActiveSheet.Cells.Find("11curr").Activate()
                    .ActiveCell.Parse("[         11CURR][    3.38][   -112.50][    0.00   ][   0.00   ][   0.00   ][   0.00]")

                    Row = .ActiveCell.Row
                    ' MsgBox(.ActiveSheet.Cells(Row, 2).value)
                    CurrVel = .Cells(Row, 2).value
                    CurrDir = .Cells(Row, 3).value
                End If
                If Not .ActiveSheet.Cells.Find("11wind") Is Nothing Then
                    .ActiveSheet.Cells.Find("11wind").Activate()
                    .ActiveCell.Parse("[         11WIND][    1.30][    245.00][      0.00][      0.00][      0.00][      0.00]")
                    Row = .ActiveCell.Row
                    WindVel = .Cells(Row, 2).value
                    WindDir = .Cells(Row, 3).value
                End If
                If Not .ActiveSheet.Cells.Find("13curr") Is Nothing Then
                    .ActiveSheet.Cells.Find("13curr").Activate()
                    .ActiveCell.Parse("[         13CURR][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
                    Row = .ActiveCell.Row
                    CurrVel = .Cells(Row, 4).value
                    CurrDir = .Cells(Row, 5).value
                End If

                If Not .ActiveSheet.Cells.Find("13wind") Is Nothing Then

                    .ActiveSheet.Cells.Find("13wind").Activate()
                    .ActiveCell.Parse("[         13WIND][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
                    Row = .ActiveCell.Row
                    WindVel = .Cells(Row, 4).value
                    WindDir = .Cells(Row, 5).value
                End If


                If Not .Cells.Find("xswl") Is Nothing Then
					.Cells.Find("xswl").Activate()
					If InStr(1, .ActiveCell.Formula, "JONH") > 0 Then
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        SwellHs = .Cells(Row, 6).value
                        SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        SwellDir = .Cells(Row, 8).value
                    ElseIf InStr(1, .ActiveCell.Formula, "GAUS") > 0 Then 
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        SwellHs = .Cells(Row, 6).value
                        If .Cells.Find("GATP") Is Nothing Then
                            SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        Else
                            SwellTp = .Cells(Row, 7).value
                        End If
                        SwellDir = .Cells(Row, 8).value
                    Else
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        SwellHs = .Cells(Row, 6).value
                        SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        SwellDir = .Cells(Row, 8).value
                    End If
				End If
				
				.Cells.Find("spdn").Activate()
				.ActiveCell.Parse("[         13SPDN][   0][   0][    63.900][     0.000][     0.000][     0.000][     0.000][     0.000]")
				Row = .ActiveCell.Row
                WaveDir = .Cells(Row, 4).value

                If Not .Cells.Find(What:="13psmz") Is Nothing Then
					.Cells.Find(What:=("13psmz")).Activate()
					.ActiveCell.Parse("[         13CURR][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    WaveHs = .Cells(Row, 6).value
                    WaveTp = .Cells(Row, 7).value * 1.4
                End If
				If Not .Cells.Find(What:="13jonh") Is Nothing Then
					.Cells.Find(What:=("13jonh")).Activate()
					.ActiveCell.Parse("[         13CURR][  0][  0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    WaveHs = .Cells(Row, 7).value
                    WaveTp = 2 * Math.PI / .Cells(Row, 8).value '  Jonswap peak freq
                End If
				
				.Cells.Find("G L O B A L   S Y S T E M   S T I F F N E S S   M A T R I X").Activate()
				Row = .ActiveCell.Row + 11
                .Range(.Cells(Row, 1), .Cells(Row + 10, 1)).Parse("[                               X          ][   9.997E+03][   8.166E+01][  -1.613E+00][   7.336E+02][   2.999E+05][  -1.596E+05]")
                stiffx = .Cells(Row, 2).value
                StiffY = .Cells(Row + 2, 3).value
                StiffRz = .Cells(Row + 10, 7).value

                .Cells.Find("ITER NO.              CENTRE OF GRAVITY", SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious).Activate()
				Row = .ActiveCell.Row + 3
                .Range(.Cells(Row, 1), .Cells(Row + 8, 1)).Parse("[                      DRIFT                   ][   1.5612E+02][   2.7210E+04][   0.0000E+00][     0.0000E+00][   0.0000E+00][  -1.8411E+05]")
                VslX = .Cells(Row, 2).value
                VslY = .Cells(Row, 3).value
                VslZ = .Cells(Row, 4).value
                VslRx = .Cells(Row, 5).value
                VslRy = .Cells(Row, 6).value
                VslRz = .Cells(Row, 7).value
                Row = Row + 3
                CurrX = .Cells(Row, 2).value
                CurrY = .Cells(Row, 3).value
                CurrRz = .Cells(Row, 7).value
                Row = Row + 1
                WindX = .Cells(Row, 2).value
                WindY = .Cells(Row, 3).value
                WindRz = .Cells(Row, 7).value
                Row = Row + 1
                WaveX = .Cells(Row, 2).value
                WaveY = .Cells(Row, 3).value
                WaveRz = .Cells(Row, 7).value
                Row = Row + 2
                ThruX = .Cells(Row, 2).value
                ThruY = .Cells(Row, 3).value
                ThruRz = .Cells(Row, 7).value

                'Row = Row + 8 'JLIU 11/4/2014 two nods
                .Cells.Find("FORCE LINE", SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious).Activate() 'JLIU not depending print out nodes anymore.
				Row = .ActiveCell.Row - NumLineA + 1
				For nl = 1 To NumLineA
                    If InStr(1, .Cells(Row + nl - 1, 1).value, "TOT TENSION") > 0 Then
                        .Cells(Row + nl - 1, 1).Parse("[                      TENSION LINE      1        1.3861E+06   2.1844E+06  -3.1547E+06     TOT TENSION][ 4.0798E+06 ][AL  1.90E+06 ][LL     0.000]")
                    Else
                        .Cells(Row + nl - 1, 1).Parse("[                      TENSION LINE      1        1.3861E+06   2.1844E+06  -3.1547E+06   TOT][ 4.0798E+06 ][AL  1.90E+06 ][LL     0.000]")
                    End If
                    '      Next i
                    '      For NL = 1 To NumLineA
                    Ten(nl) = .Cells(Row + nl - 1, 2).value
                Next nl
				
				
				.Range("A1").Select()
				If Not .Cells.Find("MODE") Is Nothing Then
                    tmpVal = .Cells.Find("PERIOD       (SECONDS)").FormulaR1C1
                    Fields = Split_Renamed(tmpVal, " ")
                    For k = 1 To UBound(Fields)
                        oxBook.ActiveSheet.Cells(90 + k, 15).FormulaR1C1 = Fields(k)
                    Next k
                    tmpVal = .Cells.Find("DAMPING(PER CENT CRIT)").FormulaR1C1
                    Fields = Split_Renamed(tmpVal, " ")
                    For k = 1 To UBound(Fields) - 2
                        oxBook.ActiveSheet.Cells(90 + k, 16).FormulaR1C1 = Fields(k + 2)
                    Next k
                    tmpVal = .Cells.Find("TYPE OF MOTION").FormulaR1C1
                    Fields = Split_Renamed(tmpVal, " ")
                    For k = 1 To UBound(Fields) - 1
                        oxBook.ActiveSheet.Cells(90 + k, 17).FormulaR1C1 = Fields(k + 1)
                    Next k
				End If
				.CutCopyMode = False
				.ActiveWorkbook.Close(SaveChanges:=False)
				
				If blnPFLH Then

                    '  Open Fer file
                    '--------------------------------------------------------------------------
                    .Workbooks.OpenText(Filename:=WorkDir & "\" & CaseNam & "\" & AFfname, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object() {1, 1})

                    .Cells.Find("L / F   S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					
					Row = .ActiveCell.Row + 14
                    .Cells(Row, 1).Parse("[                       1            1][        0.004][      0.008][      0.000][      0.000][      0.000][      0.004]")
                    '                              [                       1            1][        0.004][      0.008][      0.000][      0.000][      0.000][      0.004]"
                    '[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]
                    VslXLF = .Cells(Row, 2).value
                    VslYLF = .Cells(Row, 3).value
                    VslZLF = .Cells(Row, 4).value
                    VslRxLF = .Cells(Row, 5).value
                    VslRyLF = .Cells(Row, 6).value
                    VslRzLF = .Cells(Row, 7).value

                    NX = 1
                    nl = 0
                    Do While (NX - 1) * 10 < NumLineA
                        .Cells.Find("L / F   S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        '    Row = .ActiveCell.Row
                        .Cells.Find("MOORING LINE NUMBER", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        '   Row = .ActiveCell.Row
                        Row = .ActiveCell.Row + 7 '+ (NX - 1) * 21    'added last term
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")

                        For i = 1 To 10
                            nl = nl + 1
                            If nl <= NumLineA Then TenLF(nl) = .Cells(Row, i + 1).value
                        Next i
                        NX = NX + 1
                    Loop

                    .Cells.Find(" H / F   S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					Row = .ActiveCell.Row + 14
                    '.Cells(Row, 1).Parse "[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]"
                    .Cells(Row, 1).Parse("[                       1            1][        0.516][      0.387][      0.518][      0.566][      0.830][      0.665]")

                    VslXWF = .Cells(Row, 2).value
                    VslYWF = .Cells(Row, 3).value
                    VslZWF = .Cells(Row, 4).value
                    VslRxWF = .Cells(Row, 5).value
                    VslRyWF = .Cells(Row, 6).value
                    VslRzWF = .Cells(Row, 7).value


                    .Cells.Find("H / F   S I G N I F I C A N T   A C C E L E R A T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					Row = .ActiveCell.Row + 14
                    '.Cells(Row, 1).Parse "[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]"
                    .Cells(Row, 1).Parse("[                       1            1][         0.144][      0.107][      0.152][      0.182][      0.268][      0.199]")

                    vslzacc = .Cells(Row, 4).value

                    NX = 1
                    nl = 0

                    Do While (NX - 1) * 10 < NumLineA
                        .Cells.Find("H / F   S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        '   Row = .ActiveCell.Row

                        .Cells.Find("MOORING LINE NUMBER", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        '     Row = .ActiveCell.Row
                        Row = .ActiveCell.Row + 7 '+ (NX - 1) * 21    'added last term
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
                            nl = nl + 1
                            If nl <= NumLineA Then TenWF(nl) = .Cells(Row, i + 1).value
                        Next i
                        NX = NX + 1
                    Loop

                    .CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
					
				Else

                    '--------------------------------------------------------------------------------------
                    .Workbooks.OpenText(Filename:=WorkDir & "\" & CaseNam & "\" & AFfname1, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object() {1, 1})

                    .Cells.Find("S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    VslXLF = .Cells(Row, 2).value
                    VslYLF = .Cells(Row, 3).value
                    VslZLF = .Cells(Row, 4).value
                    VslRxLF = .Cells(Row, 5).value
                    VslRyLF = .Cells(Row, 6).value
                    VslRzLF = .Cells(Row, 7).value

                    NX = 1
                    nl = 0
                    .Cells.Find("S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S").Activate()
                    Do While (NX - 1) * 10 < NumLineA
                        .Cells.Find("MOORING LINE NUMBER", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        Row = .ActiveCell.Row + 7 + (NX - 1) * 21 'added last term
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
                            nl = nl + 1
                            If nl <= NumLineA Then TenLF(nl) = .Cells(Row, i + 1).value
                        Next i
                        NX = NX + 1

                    Loop

                    .CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
                    '--------------------------------------------------------------------------------------

                    .Workbooks.OpenText(Filename:=WorkDir & "\" & CaseNam & "\" & AFfname2, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object() {1, 1})

                    .Cells.Find("S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    VslXWF = .Cells(Row, 2).value
                    VslYWF = .Cells(Row, 3).value
                    VslZWF = .Cells(Row, 4).value
                    VslRxWF = .Cells(Row, 5).value
                    VslRyWF = .Cells(Row, 6).value
                    VslRzWF = .Cells(Row, 7).value

                    .Cells.Find("S I G N I F I C A N T   A C C E L E R A T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    vslzacc = .Cells(Row, 4).value

                    NX = 1
                    nl = 0
                    .Cells.Find("S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S").Activate()
                    Do While (NX - 1) * 10 < NumLineA
                        .Cells.Find("MOORING LINE NUMBER", After:= .ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
                        Row = .ActiveCell.Row + 7 + (NX - 1) * 21 'added last term
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
                            nl = nl + 1
                            If nl <= NumLineA Then TenWF(nl) = .Cells(Row, i + 1).value
                        Next i
                        NX = NX + 1
                    Loop

                    .CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
					
				End If

                .Cells(6, 3) = WindDir
                .Cells(6, 7) = WaveDir
                .Cells(6, 8) = SwellDir
                .Cells(8, 7) = WaveTp
                .Cells(8, 8) = SwellTp
                .Cells(6, 12) = CurrDir

                .Cells(25, 6) = VslRx
                .Cells(25, 7) = VslRy
                .Cells(25, 8) = VslRz

                .Cells(26, 6) = VslRxLF
                .Cells(26, 7) = VslRyLF
                .Cells(26, 8) = VslRzLF

                .Cells(27, 6) = VslRxWF
                .Cells(27, 7) = VslRyWF
                .Cells(27, 8) = VslRzWF

                If IsDataEnglishUnit Then
                    .Cells(7, 3) = WindVel
                    .Cells(7, 7) = WaveHs
                    .Cells(7, 8) = SwellHs
                    .Cells(7, 12) = CurrVel

                    .Cells(14, 9) = ThruX
                    .Cells(14, 10) = ThruY
                    .Cells(14, 11) = ThruRz
                    .Cells(16, 9) = CurrX
                    .Cells(16, 10) = CurrY
                    .Cells(16, 11) = CurrRz
                    .Cells(17, 9) = WindX
                    .Cells(17, 10) = WindY
                    .Cells(17, 11) = WindRz
                    .Cells(18, 9) = WaveX
                    .Cells(18, 10) = WaveY
                    .Cells(18, 11) = WaveRz

                    .Cells(25, 3) = VslX
                    .Cells(25, 4) = VslY
                    .Cells(25, 5) = VslZ

                    .Cells(26, 3) = VslXLF
                    .Cells(26, 4) = VslYLF
                    .Cells(26, 5) = VslZLF

                    .Cells(27, 3) = VslXWF
                    .Cells(27, 4) = VslYWF
                    .Cells(27, 5) = VslZWF
                    .Cells(28, 5) = vslzacc

                    .Cells(14, 2) = stiffx
                    .Cells(15, 2) = StiffY
                    .Cells(16, 2) = StiffRz
                Else ' if results are in Metric unit, convert to English to fill in the excel template
                    .Cells(7, 3) = WindVel / 0.3048
                    .Cells(7, 7) = WaveHs / 0.3048
                    .Cells(7, 8) = SwellHs / 0.3048
                    .Cells(7, 12) = CurrVel / 0.3048

                    ' FORCES  KN  SO * 1000

                    ' Thrusters
                    .Cells(14, 9) = 1000 * ThruX / 9.80665 / 0.4536
                    .Cells(14, 10) = 1000 * ThruY / 9.80665 / 0.4536
                    .Cells(14, 11) = 1000 * ThruRz / 9.80665 / 0.4536 / 0.3048

                    ' Current Forces
                    .Cells(16, 9) = 1000 * CurrX / 9.80665 / 0.4536
                    .Cells(16, 10) = 1000 * CurrY / 9.80665 / 0.4536
                    .Cells(16, 11) = 1000 * CurrRz / 9.80665 / 0.4536 / 0.3048
                    ' Wind Forces
                    .Cells(17, 9) = 1000 * WindX / 9.80665 / 0.4536
                    .Cells(17, 10) = 1000 * WindY / 9.80665 / 0.4536
                    .Cells(17, 11) = 1000 * WindRz / 9.80665 / 0.4536 / 0.3048
                    ' Wave Forces
                    .Cells(18, 9) = 1000 * WaveX / 9.80665 / 0.4536
                    .Cells(18, 10) = 1000 * WaveY / 9.80665 / 0.4536
                    .Cells(18, 11) = 1000 * WaveRz / 9.80665 / 0.4536 / 0.3048

                    ' Motions
                    .Cells(25, 3) = VslX / 0.3048
                    .Cells(25, 4) = VslY / 0.3048
                    .Cells(25, 5) = VslZ / 0.3048

                    .Cells(26, 3) = VslXLF / 0.3048
                    .Cells(26, 4) = VslYLF / 0.3048
                    .Cells(26, 5) = VslZLF / 0.3048

                    .Cells(27, 3) = VslXWF / 0.3048
                    .Cells(27, 4) = VslYWF / 0.3048
                    .Cells(27, 5) = VslZWF / 0.3048

                    ' Acceleration
                    .Cells(28, 5) = vslzacc / 0.3048

                    ' Mooring Stiffness
                    .Cells(14, 2) = 1000 * stiffx / 9.80665 / 0.4536 * 0.3048
                    .Cells(15, 2) = 1000 * StiffY / 9.80665 / 0.4536 * 0.3048
                    .Cells(16, 2) = 1000 * StiffRz / 9.80665 / 0.4536

                End If
				
				
				For j = 1 To NumLines
					If IsDataEnglishUnit Then
                        .Cells(35 + j, 1) = BS(j)
                    Else
                        .Cells(35 + j, 1) = BS(j) / 9.80665 / 0.4536 * 1000
                    End If
				Next j
				
				
				Row = 35
                nl = 0 ' intact line counter
                nd = 0 ' damaged line counter
                DamLineCur = 0 ' current damaged line no.
                For NLA = 1 To NumLineA
					Row = Row + 1
                    nl = nl + 1
                    If IsDamaged > 0 Then
                        If brokenlineno(nd, NCase) > 0 And nd < TotalNumBroken(NCase) And nl > DamLineCur Then
                            nd = nd + 1
                            DamLineCur = brokenlineno(nd, NCase)
                        End If
                    End If
                    If nl = DamLineCur Then
                        For Col = 3 To 5
                            .Cells(Row, Col) = 0
                        Next Col
                        .Cells(Row + 9, 21) = "Dam"
                    Else
                        If IsDataEnglishUnit Then
                            .Cells(Row, 3) = Ten(NLA)
                            .Cells(Row, 4) = TenLF(NLA)
                            .Cells(Row, 5) = TenWF(NLA)
                        Else
                            .Cells(Row, 3) = Ten(NLA) / 9.80665 / 0.4536 * 1000
                            .Cells(Row, 4) = TenLF(NLA) / 9.80665 / 0.4536 * 1000
                            .Cells(Row, 5) = TenWF(NLA) / 9.80665 / 0.4536 * 1000
                        End If
					End If
				Next NLA
				
				If IsDamaged > 0 Then
					.Range("k24").Value = "Damaged"
				Else
					.Range("k24").Value = "Intact"
				End If

                .Range(.Cells(45 + NumLines, 14), .Cells(64, 29)).Delete(Shift:=Microsoft.Office.Interop.Excel.XlDirection.xlUp)

                If IsDamaged > 0 Then ' clear content for the damaged lines
					For nd = 1 To TotalNumBroken(NCase)
                        nrowd = 44 + brokenlineno(nd, NCase)
                        If brokenlineno(nd, NCase) < NumLines Then
                            nrowcop = 44 + NumLines - 1
                            ' move lines after the damaged line one row down
                            .Range("P" & nrowd & ":R" & nrowcop).Copy()
                            .Range("P" & nrowd + 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
                        End If
                        ' clean up the row for the damaged line in the tension table
                        .Range("P" & nrowd, "T" & nrowd).FormulaR1C1 = "  "
                        .Range("U" & nrowd).FormulaR1C1 = "DAMAGED"
                        .Range("X" & nrowd, "AB" & nrowd).FormulaR1C1 = "  "
                        .Range("AC" & nrowd).FormulaR1C1 = "DAMAGED"
                    Next nd
				End If


                '      Set Header and Title
                With .ActiveSheet.PageSetup
                    EndRow = 45 + NumLines
                    ' set printArea based on result Unit
                    If IsDataEnglishUnit Then
                        .PrintArea = "$O$1:$U$" & EndRow
                    Else
                        .PrintArea = "$V$1:$AC$" & EndRow
                    End If
                    .PrintTitleRows = ""
                    .PrintTitleColumns = ""
                    .LeftHeader = "&D"
                    .CenterHeader = " "
                    .RightHeader = "Prepared by DTCEL"
                    .CenterFooter = "David Tein Consulting Engineers, Ltd." & vbCrLf & "11777 Katy Freeway, Suite 434, Houston, TX 77079, 281-531-0888, FAX 281-531-5888, dtcel@dtcel.com"

                End With

            End With
		Next NCase
        ' hide template sheet
        oxApp.ActiveWorkbook.Sheets("MOORING1").Visible = False
        Call CreateSummary(oxApp, preT, BS, NumLines)
		' cleanup
		Call CleanupWindows()
	End Sub

    Function CleanupWindows() As Object
		On Error Resume Next ' ignore if window non-exist
		With oxApp
			'        .Windows(ABfname).WindowState = xlMinimized
			'        .Windows(ABfname).Close SaveChanges:=False
			'       .Windows(AFfname).WindowState = xlMinimized
			'     .Windows(AFfname).Close SaveChanges:=False
		End With
	End Function

    Private Function CreateSummary(ByRef oxApp As Microsoft.Office.Interop.Excel.Application, ByRef preT() As Double, ByRef BS() As Double, ByRef NumLines As Short) As Object
		Dim minFOS As Object
		Dim Col As Object
		Dim nPos As Object
		Dim i As Short
		Dim oxSheet, oSummarySheet As Microsoft.Office.Interop.Excel.Worksheet
		Dim j, k As Short
		
		On Error Resume Next
		
		Dim CaseNo As Short
		Dim MaxT, MaxGL, MeanT, MinT As Double
		Dim TmpCaseName As String
		Dim Payout, HorizT, Scope As Double
        Dim TmpStr As String
        Dim NCase As Short
		With oxApp.ActiveWorkbook
			oSummarySheet = .Sheets.Add(After:=.Worksheets(.Worksheets.Count))
			oSummarySheet.Name = "Summary"
			With oSummarySheet
				.Range("B2").Formula = .Application.Range("MainTitle").Value & ", " & .Application.Range("SubTitle").Value
				.Range("B1:B2").Font.Bold = True
				.Range("B1:B2").Font.Size = 16
				.Range("B4").Formula = "Metocean"
				.Range("B5").Formula = "Env to-Hdg CCW from Bow(deg)"
				.Range("B6").Formula = "Damaged Line No"
				.Range("B7").Formula = "Total Env Loads"
				.Range("B8").Formula = "Fx, Surge (kips)"
				.Range("B9").Formula = "Fy, Sway (kips)"
				.Range("B10").Formula = "Mz, Yaw (kips-ft)"
				.Range("B12").Formula = "Mean Offset (ft)"
				.Range("B13").Formula = "Surge Max (ft)"
				.Range("B14").Formula = "Surge Min (ft)"
				.Range("B15").Formula = "Surge Mean (ft)"
				.Range("B16").Formula = "Sway Max (ft)"
				.Range("B17").Formula = "Sway Min (ft)"
				.Range("B18").Formula = "Sway Mean (ft)"
				.Range("B19").Formula = "Yaw Max (deg)"
				.Range("B20").Formula = "Yaw Min (deg)"
				.Range("B21").Formula = "Yaw Mean (deg)"
				.Range("B22").Formula = "Total Offset, Max (ft)"
				.Range("B23").Formula = "Total Offset, Max (%WD)"
				
				.Range("B25").Formula = "2nd Max Tension (kips)"
				.Range("B26").Formula = "Max Tension (kips)"
				.Range("B27").Formula = "Min Factor of Safety"
				
				.Range("B29").Formula = "Line Factor of Safety"
				
				For j = 1 To NumLines
					.Range("B" & (29 + j)).Formula = "Line " & j
				Next j
				'    If IsDamaged Then
				'        .Range("B" & (29 + NumLines + 2)).Formula = ""
				'        .Range("B" & (29 + NumLines + 3)).Formula = ""
				'    Else
				.Range("B" & (29 + NumLines + 2)).Formula = "Most-Loaded Line No"
				.Range("B" & (29 + NumLines + 3)).Formula = "2nd-Loaded Line No"
				'    End If
				
				
				
				NCase = oxApp.Sheets.Count - 3
				ReDim Preserve L1Tmax(NCase)
				ReDim Preserve L2Tmax(NCase)
				
				For i = 1 To oxApp.Sheets.Count - 1
					CaseNo = i - 2
					oxSheet = oxApp.Sheets(i)
					If InStr(oxSheet.Name, "Case") > 0 Then '  to by-pass Input & Mooring1 template sheet
                        .Cells(4, i).Formula = oxSheet.Name
                        .Cells(5, i).Formula = oxSheet.Range("C6").Value
                        .Cells(6, i).Formula = GetDamagedLineNo(oxSheet, NumLines)
                        .Cells(6, i).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                        .Cells(8, i).Formula = VB6.Format(oxSheet.Range("P21").Value, "#0.0")
                        .Cells(9, i).Formula = VB6.Format(oxSheet.Range("R21").Value, "#0.0")
                        .Cells(10, i).Formula = Format(oxSheet.Range("T21").Value, "#0.0")

                        .Cells(12, i).FormulaR1C1 = Format(System.Math.Sqrt((oxSheet.Range("Q32").Value) ^ 2 + (oxSheet.Range("S32").Value) ^ 2), "0.0")
                        .Cells(13, i).Formula = oxSheet.Range("P37").Value
                        .Cells(14, i).Formula = oxSheet.Range("P38").Value
                        .Cells(15, i).Formula = oxSheet.Range("P32").Value
                        .Cells(16, i).Formula = oxSheet.Range("Q37").Value
                        .Cells(17, i).Formula = oxSheet.Range("Q38").Value
                        .Cells(18, i).Formula = oxSheet.Range("Q32").Value
                        .Cells(19, i).Formula = oxSheet.Range("U37").Value
                        .Cells(20, i).Formula = oxSheet.Range("U38").Value
                        .Cells(21, i).Formula = oxSheet.Range("U32").Value
                        .Cells(22, i).Formula = oxSheet.Range("R39").Value
                        .Cells(23, i).Formula = oxSheet.Range("U39").Value

                        .Cells(25, i).Formula = Format(oxSheet.Range("V13").Value, "0.00") '  2nd Loaded Line Tension
                        .Cells(26, i).Formula = Format(oxSheet.Range("T13").Value, "0.00") '  Most Loaded Line Tension
                        .Cells(27, i).Formula = Format(oxSheet.Range("T12").Value, "0.00") '  Min FOS

                        '    .Cells(26, i).Formula = "=MAX(R[+5]C:R[" & (NumLines + 5) & "]C)" '  most-loaded line tension"
                        '    .Cells(25, i).Formula = "=MAX(OFFSET(R[-" & (NumLines + 4) & "]C,R[2]C-1,0),OFFSET(R[-" & (NMoor + 4) & "]C,R[1]C+1,0))" ' 2nd-Loaded Line tension

                        For j = 1 To NumLines
                            If InStr(oxSheet.Range("U" & (44 + j)).Value, "DAMAGED") > 0 Then
                                .Cells(29 + j, i).Value = " "
                            Else
                                .Cells(29 + j, i).Value = oxSheet.Range("U" & (44 + j)).Value
                            End If
						Next j
						TmpStr = Find2MostLoadedLines(CaseNo, NumLines)

                        nPos = InStr(TmpStr, " ")
                        If nPos = 0 Then
                            Debug.Print("Error finding 2 Loaded Line Numbers: npos=0 tmpStr= " & TmpStr)
                        End If

                        If IsDamaged = 0 Then ' only save intact case most-loaded and second-Loaded
                            ' if damaged case find new most-loaded will flush intact most-loaded and 2nd loaded line numbers
                            ' most loaded line and 2nd loaded line numbers based on tension not FOS
                            .Cells(29 + NumLines + 2, i).Formula = Left(TmpStr, nPos - 1)
                            '    .Cells(29 + NumLines + 3, i).Formula = Right(TmpStr, Len(TmpStr) - nPos)

                            '    .Cells(29 + NumLines + 2, i).Formula = "=MATCH(R[-3]C,R[-" & (NumLines + 5) & "]C:R[-6]C,0)"
                            '    .Cells(29 + NumLines + 3, i).Formula = "=IF(OFFSET(R[-" & (NumLines + 3) & "]C,R[-1]C-1,0)<OFFSET(R[-" & (NumLines + 3) & "],R[-1]C+1,0),IF(R[-1]C-1<1,1,R[-1]C-1),IF(R[-1]C+1>" & NumLines & ",1,R[-1]C+1))"
                            .Cells(29 + NumLines + 3, i).Formula = Right(TmpStr, Len(TmpStr) - nPos)

                            L1Tmax(i - 2) = CShort(Left(TmpStr, nPos - 1))
                            L2Tmax(i - 2) = CShort(Right(TmpStr, Len(TmpStr) - nPos))
                            '      L1Tmax(i - 2) = .Cells(29 + NumLines + 2, i).Formula
                            '      L2Tmax(i - 2) = .Cells(29 + NumLines + 3, i).Formula
                        End If
					End If
				Next i
			End With
			
			With .ActiveSheet
                .Rows("13:98").NumberFormat = "0.00"
                .Columns("B:B").ColumnWidth = 26
                j = oxApp.Sheets.Count + 1

                Col = NCase + 4
                .Cells(21, Col).FormulaR1C1 = "Max Offset (%WD)"
                .Cells(22, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"
                .Cells(23, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"
                .Cells(25, Col).FormulaR1C1 = "FOS"
                .Cells(26, Col).FormulaR1C1 = "Min"
                .Cells(27, Col).FormulaR1C1 = "=MIN(RC[-" & NCase + 1 & "]:RC[-2])"
                .Cells(10, Col).FormulaR1C1 = "Max Mean (ft)"
                .Cells(11, Col).FormulaR1C1 = ""
                .Cells(12, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"

                .Cells(29, Col).FormulaR1C1 = "Min FOS"
                .Cells(29, Col + 1).FormulaR1C1 = "MaxT(k)"
                .Cells(29, Col + 2).FormulaR1C1 = "Case"
                .Cells(29, Col + 3).FormulaR1C1 = "Max FOS"
                .Cells(29, Col + 4).FormulaR1C1 = "Case"
                .Cells(29, Col + 5).FormulaR1C1 = ""
                .Cells(29, Col + 6).FormulaR1C1 = "MinT(k)"
                .Cells(28, Col + 7).FormulaR1C1 = "Approx"
                .Cells(29, Col + 7).FormulaR1C1 = "PreT(k)"
                .Range(.Cells(29, Col), .Cells(29, Col + 6)).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight

                For i = 1 To NumLines
                    .Cells(29 + i, Col + 7).FormulaR1C1 = preT(i)
                    ' intact FS
                    .Cells(29 + i, Col + 6).NumberFormat = "0"
                    .Cells(29 + i, Col + 7).NumberFormat = "0"
                    ' The FindMinFOS will return minFOS and CaseNo
                    minFOS = FindMinFOS(i, CaseNo)
                    .Cells(29 + i, Col).Value = minFOS
                    .Cells(29 + i, Col + 1).FormulaR1C1 = BS(i) / minFOS
                    .Cells(29 + i, Col + 2).Value = CaseNo
                    .Cells(29 + i, Col + 2).NumberFormat = "0"
                    ' The FindMaxFOS will return maxFOS and CaseNo
                    .Cells(29 + i, Col + 3).Value = FindMaxFOS(i, CaseNo)
                    .Cells(29 + i, Col + 4).Value = CaseNo
                    .Cells(29 + i, Col + 4).NumberFormat = "0"
                    .Cells(29 + i, Col + 1).NumberFormat = "0"
                    .Cells(29 + i, Col + 5).NumberFormat = "0"

                    If InStr(oxSheet.Name, "D1") > 0 Then
                        TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0") & "D1"
                    ElseIf InStr(oxSheet.Name, "D2") > 0 Then
                        TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0") & "D2"
                    Else
                        TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0")
                    End If
                    MaxT = IIf(oxApp.Sheets(TmpCaseName).Range("S" & 44 + i).Value > oxApp.Sheets(TmpCaseName).Range("T" & 44 + i).Value, oxApp.Sheets(TmpCaseName).Range("S" & 44 + i).value, oxApp.Sheets(TmpCaseName).Range("T" & 44 + i).value)

                    MeanT = oxApp.Sheets(TmpCaseName).Range("P" & 44 + i).value
                    MinT = MeanT - (MaxT - MeanT)
                    .Cells(29 + i, Col + 6).Value = MinT
                    '   .Cells(29 + i, Col + 7).Value = oxApp.Sheets(TmpCaseName).Range("T9")

                Next i
                ' format most-loaded and 2nd loaded line numbers
                .Range(.Cells(29 + NumLines + 2, 2), .Cells(29 + NumLines + 3, NCase + 2)).NumberFormat = "0"

                ' print environment
                .Cells(4, j + 3).FormulaR1C1 = "Wind"
                .Cells(5, j + 3).FormulaR1C1 = "Wave Hs"
                .Cells(6, j + 3).FormulaR1C1 = "            Tp"
                .Cells(7, j + 3).FormulaR1C1 = "Current"

                .Cells(4, j + 5).FormulaR1C1 = "  knots"
                .Cells(5, j + 5).FormulaR1C1 = "  ft"
                .Cells(6, j + 5).FormulaR1C1 = "  sec"
                .Cells(7, j + 5).FormulaR1C1 = "  knots"

                .Range("B1").Formula = "Water Depth = " & Format(oxApp.Sheets(TmpCaseName).Range("T6"), "0.0") & " ft, Pretension = " & oxApp.Sheets(TmpCaseName).Range("T9") & " kips"

                .Cells(4, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("P11").Value, "#0.00")
                .Cells(5, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("G7").Value, "#0.00")
                .Cells(6, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("G8").Value, "#0.00")
                .Cells(7, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("R11").Value, "#0.00")


            End With
		End With
	End Function
	
	Private Function FindMaxFOS(ByVal LineNo As Short, ByRef CaseNo As Short) As Object
		Dim oxSheet As Microsoft.Office.Interop.Excel.Worksheet
		Dim MaxVal As Double
		Dim i As Short
		MaxVal = -1000000
		CaseNo = -10
		With oxApp.ActiveWorkbook
			For i = 1 To .Sheets.Count
				oxSheet = .Sheets(i)
				If InStr(oxSheet.Name, "Case") > 0 Then
                    If IsNumeric(oxSheet.Range("U" & (44 + LineNo)).Value) Then '<> "DAMAGED" Then
                        If oxSheet.Range("U" & (44 + LineNo)).Value > MaxVal Then
                            MaxVal = oxSheet.Range("U" & (44 + LineNo)).Value
                            CaseNo = i - 2
                        End If
                    End If
                End If
			Next i
		End With
		FindMaxFOS = VB6.Format(MaxVal, "0.00")
	End Function
	
	Private Function FindMinFOS(ByVal LineNo As Short, ByRef CaseNo As Short) As Object
		Dim oxSheet As Microsoft.Office.Interop.Excel.Worksheet
		Dim MinVal As Double
		Dim i As Short
		MinVal = 1000000
		With oxApp.ActiveWorkbook
			For i = 1 To .Sheets.Count
				oxSheet = .Sheets(i)
                If InStr(oxSheet.Name, "Case") > 0 Then

                    If IsNumeric(oxSheet.Range("U" & (44 + LineNo)).Value) Then '<> "DAMAGED" Then

                        If Convert.ToDouble(oxSheet.Range("U" & (44 + LineNo)).Value) < MinVal Then
                            MinVal = Convert.ToDouble(oxSheet.Range("U" & (44 + LineNo)).Value)
                            CaseNo = i - 2
                        End If
                    End If
                End If
            Next i
		End With
		FindMinFOS = VB6.Format(MinVal, "0.00")
	End Function
	
	Private Function Find2MostLoadedLines(ByRef CaseNo As Short, ByRef NumLines As Short) As String
		Find2MostLoadedLines = ""
		Dim Min1, Min2 As Double
		Dim Line1No, Line2No As Short
		Dim i As Short
		Min1 = 1000000
		
		Line1No = -1
		Line2No = -1
		
		With oxApp.ActiveWorkbook.Sheets(CaseNo + 2)
			For i = 1 To NumLines
                Debug.Print(.Name & ", i=" & i & ", " & .Range("N" & (44 + i)).Value)

                If System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("T13").Value) < 0.09 And System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("V13").Value) < 0.09 Then
                    ' T1 close to T2
                    If Line1No = -1 Then Line1No = i
                    If Line1No <> i And Line2No = -1 Then Line2No = i
                ElseIf System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("T13").Value) < 0.09 And Line1No = -1 Then
                    Line1No = i
                ElseIf System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("V13").Value) < 0.19 And Line2No = -1 Then
                    If Line1No <> i Then Line2No = i
				End If
				If Line2No > 0 And Line1No > 0 Then Exit For
			Next i
		End With
		
		'        For i = 1 To NumLines
		'            If .Cells(29 + i, CaseNo).Value < Min1 Then
		'                Min1 = .Cells(29 + i, CaseNo).Value
		'                Line1No = i
		'            End If
		'        Next i
		'        Min2 = 1000000
		'        For i = 1 To NumLines
		'                If .Cells(29 + i, CaseNo).Value <> Min1 Then
		'                    If .Cells(29 + i, CaseNo).Value < Min2 Then
		'                       Min2 = .Cells(29 + i, CaseNo).Value
		'                        Line2No = i
		'                    End If
		'                End If
		'        Next i
		Find2MostLoadedLines = CStr(Line1No) & " " & CStr(Line2No)
	End Function
	
	Private Function GetDamagedLineNo(ByRef oxSheet As Microsoft.Office.Interop.Excel.Worksheet, ByRef NumLines As Short) As String
		Dim i As Short
		GetDamagedLineNo = "Intact"
		For i = 1 To NumLines
			If InStr(oxSheet.Range("U" & (44 + i)).Formula, "DAMAGED") > 0 Then
				GetDamagedLineNo = CStr(i)
			End If
		Next i
	End Function
	
	Function CleanupObjects() As Object
		On Error Resume Next ' ignore if window non-exist
        oxBook = Nothing
        oxBooks = Nothing
        oxApp = Nothing
    End Function
	
	Private Function DeleteWorksheet(ByRef strSheetName As String) As Boolean
		On Error Resume Next
		
		ExcelGlobal_definst.Application.DisplayAlerts = False
        ExcelGlobal_definst.ActiveWorkbook.Worksheets(strSheetName).Delete()
        ExcelGlobal_definst.Application.DisplayAlerts = True
		' Return True if no error occurred;
		' otherwise return False.
		DeleteWorksheet = Not CBool(Err.Number)
	End Function


    Public Sub New()
        MyBase.New()
        oxApp = New Microsoft.Office.Interop.Excel.Application
        oxBooks = oxApp.Workbooks
    End Sub

    Protected Overrides Sub Finalize()
        Call CleanupObjects()
        MyBase.Finalize()
	End Sub
End Class