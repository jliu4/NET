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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.Sheets().Activate. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.Sheets().Copy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oxBook.Sheets("MOORING1").Copy(After:=oxBook.Sheets(NCase + 1))
			'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.Sheets().Name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oxBook.Sheets("MOORING1 (2)").Name = CaseNam
			'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oxBook.ActiveSheet.Cells(72, 16) = CaseNam
			
			'--------------------------------------------------------------------------------------
			'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			oxApp.Workbooks.OpenText(FileName:=WorkDir & "\" & CaseNam & "\" & ABfname, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object(){1, 1})
			
			With oxApp
				
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not .ActiveSheet.Cells.Find("11curr") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ActiveSheet.Cells.Find("11curr").Activate()
                    .ActiveCell.Parse("[         11CURR][    3.38][   -112.50][    0.00   ][   0.00   ][   0.00   ][   0.00]")

                    Row = .ActiveCell.Row
                    ' MsgBox(.ActiveSheet.Cells(Row, 2).value)
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CurrVel = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CurrDir = .Cells(Row, 3).value
                End If
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not .ActiveSheet.Cells.Find("11wind") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ActiveSheet.Cells.Find("11wind").Activate()
					.ActiveCell.Parse("[         11WIND][    1.30][    245.00][      0.00][      0.00][      0.00][      0.00]")
					Row = .ActiveCell.Row
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WindVel = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WindDir = .Cells(Row, 3).value
                End If
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not .ActiveSheet.Cells.Find("13curr") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ActiveSheet.Cells.Find("13curr").Activate()
					.ActiveCell.Parse("[         13CURR][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CurrVel = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CurrDir = .Cells(Row, 5).value
                End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not .ActiveSheet.Cells.Find("13wind") Is Nothing Then
					
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ActiveSheet.Cells.Find("13wind").Activate()
					.ActiveCell.Parse("[         13WIND][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WindVel = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WindDir = .Cells(Row, 5).value
                End If
				
				
				If Not .Cells.Find("xswl") Is Nothing Then
					.Cells.Find("xswl").Activate()
					If InStr(1, .ActiveCell.Formula, "JONH") > 0 Then
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellHs = .Cells(Row, 6).value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellDir = .Cells(Row, 8).value
                    ElseIf InStr(1, .ActiveCell.Formula, "GAUS") > 0 Then 
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellHs = .Cells(Row, 6).value
                        If .Cells.Find("GATP") Is Nothing Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            SwellTp = .Cells(Row, 7).value
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellDir = .Cells(Row, 8).value
                    Else
						.ActiveCell.Parse("[     13XSWL][ JONH][          0.000 ][    0.000 ][    3.300 ][   33.700 ][    0.388 ][    2.000][    2.000]")
						Row = .ActiveCell.Row
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellHs = .Cells(Row, 6).value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellTp = 2 * Math.PI / .Cells(Row, 7).value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        SwellDir = .Cells(Row, 8).value
                    End If
				End If
				
				.Cells.Find("spdn").Activate()
				.ActiveCell.Parse("[         13SPDN][   0][   0][    63.900][     0.000][     0.000][     0.000][     0.000][     0.000]")
				Row = .ActiveCell.Row
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveDir = .Cells(Row, 4).value

                If Not .Cells.Find(What:="13psmz") Is Nothing Then
					.Cells.Find(What:=("13psmz")).Activate()
					.ActiveCell.Parse("[         13CURR][   0][   0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WaveHs = .Cells(Row, 6).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WaveTp = .Cells(Row, 7).value * 1.4
                End If
				If Not .Cells.Find(What:="13jonh") Is Nothing Then
					.Cells.Find(What:=("13jonh")).Activate()
					.ActiveCell.Parse("[         13CURR][  0][  0][     0.570][    90.000][     0.000][     0.000][     0.000][     0.000]")
					Row = .ActiveCell.Row
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WaveHs = .Cells(Row, 7).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    WaveTp = 2 * Math.PI / .Cells(Row, 8).value '  Jonswap peak freq
                End If
				
				.Cells.Find("G L O B A L   S Y S T E M   S T I F F N E S S   M A T R I X").Activate()
				Row = .ActiveCell.Row + 11
                .Range(.Cells(Row, 1), .Cells(Row + 10, 1)).Parse("[                               X          ][   9.997E+03][   8.166E+01][  -1.613E+00][   7.336E+02][   2.999E+05][  -1.596E+05]")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object stiffx. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                stiffx = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StiffY = .Cells(Row + 2, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StiffRz = .Cells(Row + 10, 7).value

                .Cells.Find("ITER NO.              CENTRE OF GRAVITY", SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious).Activate()
				Row = .ActiveCell.Row + 3
                .Range(.Cells(Row, 1), .Cells(Row + 8, 1)).Parse("[                      DRIFT                   ][   1.5612E+02][   2.7210E+04][   0.0000E+00][     0.0000E+00][   0.0000E+00][  -1.8411E+05]")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslX = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslY = .Cells(Row, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslZ = .Cells(Row, 4).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRx. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslRx = .Cells(Row, 5).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslRy = .Cells(Row, 6).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VslRz = .Cells(Row, 7).value
                Row = Row + 3
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrX = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrY = .Cells(Row, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                CurrRz = .Cells(Row, 7).value
                Row = Row + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindX = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindY = .Cells(Row, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WindRz = .Cells(Row, 7).value
                Row = Row + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveX = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveY = .Cells(Row, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                WaveRz = .Cells(Row, 7).value
                Row = Row + 2
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruX = .Cells(Row, 2).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruY = .Cells(Row, 3).value
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ThruRz = .Cells(Row, 7).value

                'Row = Row + 8 'JLIU 11/4/2014 two nods
                .Cells.Find("FORCE LINE", SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious).Activate() 'JLIU not depending print out nodes anymore.
				Row = .ActiveCell.Row - NumLineA + 1
				For nl = 1 To NumLineA
                    'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If InStr(1, .Cells(Row + nl - 1, 1).value, "TOT TENSION") > 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row + nl - 1, 1).Parse("[                      TENSION LINE      1        1.3861E+06   2.1844E+06  -3.1547E+06     TOT TENSION][ 4.0798E+06 ][AL  1.90E+06 ][LL     0.000]")
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row + nl - 1, 1).Parse("[                      TENSION LINE      1        1.3861E+06   2.1844E+06  -3.1547E+06   TOT][ 4.0798E+06 ][AL  1.90E+06 ][LL     0.000]")
                    End If
                    '      Next i
                    '      For NL = 1 To NumLineA
                    'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Ten(nl) = .Cells(Row + nl - 1, 2).value
                Next nl
				
				
				.Range("A1").Select()
				If Not .Cells.Find("MODE") Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells.Find().FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object tmpVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tmpVal = .Cells.Find("PERIOD       (SECONDS)").FormulaR1C1
					'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Fields = Split_Renamed(tmpVal, " ")
					For k = 1 To UBound(Fields)
						'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Fields(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oxBook.ActiveSheet.Cells(90 + k, 15).FormulaR1C1 = Fields(k)
					Next k
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells.Find().FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object tmpVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tmpVal = .Cells.Find("DAMPING(PER CENT CRIT)").FormulaR1C1
					'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Fields = Split_Renamed(tmpVal, " ")
					For k = 1 To UBound(Fields) - 2
						'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Fields(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oxBook.ActiveSheet.Cells(90 + k, 16).FormulaR1C1 = Fields(k + 2)
					Next k
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells.Find().FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object tmpVal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tmpVal = .Cells.Find("TYPE OF MOTION").FormulaR1C1
					'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Fields = Split_Renamed(tmpVal, " ")
					For k = 1 To UBound(Fields) - 1
						'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxBook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Fields(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						oxBook.ActiveSheet.Cells(90 + k, 17).FormulaR1C1 = Fields(k + 1)
					Next k
				End If
				.CutCopyMode = False
				.ActiveWorkbook.Close(SaveChanges:=False)
				
				If blnPFLH Then
					
					'  Open Fer file
					'--------------------------------------------------------------------------
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					.Workbooks.OpenText(FileName:=WorkDir & "\" & CaseNam & "\" & AFfname, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object(){1, 1})
					
					
					.Cells.Find("L / F   S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					
					Row = .ActiveCell.Row + 14
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1            1][        0.004][      0.008][      0.000][      0.000][      0.000][      0.004]")
                    '                              [                       1            1][        0.004][      0.008][      0.000][      0.000][      0.000][      0.004]"
                    '[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslXLF = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslYLF = .Cells(Row, 3).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslZLF = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRxLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRxLF = .Cells(Row, 5).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRyLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRyLF = .Cells(Row, 6).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRzLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRzLF = .Cells(Row, 7).value

                    'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NX = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nl = 0
					'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do While (NX - 1) * 10 < NumLineA
						.Cells.Find("L / F   S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'    Row = .ActiveCell.Row
						.Cells.Find("MOORING LINE NUMBER", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'   Row = .ActiveCell.Row
						Row = .ActiveCell.Row + 7 '+ (NX - 1) * 21    'added last term
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")

                        For i = 1 To 10
							'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nl = nl + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If nl <= NumLineA Then TenLF(nl) = .Cells(Row, i + 1).value
                        Next i
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						NX = NX + 1
					Loop 
					
					.Cells.Find(" H / F   S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					Row = .ActiveCell.Row + 14
                    '.Cells(Row, 1).Parse "[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1            1][        0.516][      0.387][      0.518][      0.566][      0.830][      0.665]")

                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslXWF = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslYWF = .Cells(Row, 3).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslZWF = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRxWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRxWF = .Cells(Row, 5).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRyWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRyWF = .Cells(Row, 6).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRzWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRzWF = .Cells(Row, 7).value


                    .Cells.Find("H / F   S I G N I F I C A N T   A C C E L E R A T I O N S   O F   C E N T R E   O F   G R A V I T Y", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
					Row = .ActiveCell.Row + 14
                    '.Cells(Row, 1).Parse "[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]"
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1            1][         0.144][      0.107][      0.152][      0.182][      0.268][      0.199]")

                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    vslzacc = .Cells(Row, 4).value

                    'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NX = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nl = 0
					
					'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do While (NX - 1) * 10 < NumLineA
						.Cells.Find("H / F   S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'   Row = .ActiveCell.Row
						
						.Cells.Find("MOORING LINE NUMBER", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'     Row = .ActiveCell.Row
						Row = .ActiveCell.Row + 7 '+ (NX - 1) * 21    'added last term
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
							'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nl = nl + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If nl <= NumLineA Then TenWF(nl) = .Cells(Row, i + 1).value
                        Next i
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						NX = NX + 1
					Loop 
					
					.CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
					
				Else
					
					'--------------------------------------------------------------------------------------
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					.Workbooks.OpenText(FileName:=WorkDir & "\" & CaseNam & "\" & AFfname1, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object(){1, 1})
					
					.Cells.Find("S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslXLF = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslYLF = .Cells(Row, 3).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslZLF = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRxLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRxLF = .Cells(Row, 5).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRyLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRyLF = .Cells(Row, 6).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRzLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRzLF = .Cells(Row, 7).value

                    'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NX = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nl = 0
					.Cells.Find("S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S").Activate()
					'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do While (NX - 1) * 10 < NumLineA
						.Cells.Find("MOORING LINE NUMBER", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Row = .ActiveCell.Row + 7 + (NX - 1) * 21 'added last term
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
							'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nl = nl + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If nl <= NumLineA Then TenLF(nl) = .Cells(Row, i + 1).value
                        Next i
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						NX = NX + 1
						
					Loop 
					
					.CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
					'--------------------------------------------------------------------------------------
					
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					.Workbooks.OpenText(FileName:=WorkDir & "\" & CaseNam & "\" & AFfname2, StartRow:=1, DataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=New Object(){1, 1})
					
					.Cells.Find("S I G N I F I C A N T   M O T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslXWF = .Cells(Row, 2).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslYWF = .Cells(Row, 3).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslZWF = .Cells(Row, 4).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRxWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRxWF = .Cells(Row, 5).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRyWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRyWF = .Cells(Row, 6).value
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslRzWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    VslRzWF = .Cells(Row, 7).value

                    .Cells.Find("S I G N I F I C A N T   A C C E L E R A T I O N S   O F   C E N T R E   O F   G R A V I T Y").Activate()
					Row = .ActiveCell.Row + 14
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(Row, 1).Parse("[                       1           1][       0.728][      0.369][      0.002][      0.010][      0.002][      0.137]")
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    vslzacc = .Cells(Row, 4).value

                    'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NX = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nl = 0
					.Cells.Find("S I G N I F I C A N T   T E N S I O N S   I N   T H E   M O O R I N G   L I N E S").Activate()
					'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Do While (NX - 1) * 10 < NumLineA
						.Cells.Find("MOORING LINE NUMBER", After:=.ActiveCell, LookIn:=Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlPart, SearchOrder:=Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, SearchDirection:=Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, MatchCase:=False).Activate()
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Row = .ActiveCell.Row + 7 + (NX - 1) * 21 'added last term
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells().Parse. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row, 1).Parse("[     1               ][ 6.6074E+04][ 7.6400E+04][ 8.9525E+04][ 1.0102E+05][ 1.0077E+05][ 8.3966E+04][ 7.1332E+04][ 5.8991E+04][ 3.3012E+03][ 3.3223E+03]")
                        For i = 1 To 10
							'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nl = nl + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If nl <= NumLineA Then TenWF(nl) = .Cells(Row, i + 1).value
                        Next i
						'UPGRADE_WARNING: Couldn't resolve default property of object NX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						NX = NX + 1
					Loop 
					
					.CutCopyMode = False
					.ActiveWorkbook.Close(SaveChanges:=False)
					
				End If


                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(6, 3) = WindDir
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(6, 7) = WaveDir
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(6, 8) = SwellDir
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(8, 7) = WaveTp
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(8, 8) = SwellTp
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(6, 12) = CurrDir

                'UPGRADE_WARNING: Couldn't resolve default property of object VslRx. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(25, 6) = VslRx
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(25, 7) = VslRy
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(25, 8) = VslRz

                'UPGRADE_WARNING: Couldn't resolve default property of object VslRxLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(26, 6) = VslRxLF
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRyLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(26, 7) = VslRyLF
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRzLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(26, 8) = VslRzLF

                'UPGRADE_WARNING: Couldn't resolve default property of object VslRxWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(27, 6) = VslRxWF
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRyWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(27, 7) = VslRyWF
                'UPGRADE_WARNING: Couldn't resolve default property of object VslRzWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(27, 8) = VslRzWF

                If IsDataEnglishUnit Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 3) = WindVel
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 7) = WaveHs
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 8) = SwellHs
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 12) = CurrVel

                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 9) = ThruX
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 10) = ThruY
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 11) = ThruRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 9) = CurrX
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 10) = CurrY
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 11) = CurrRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 9) = WindX
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 10) = WindY
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 11) = WindRz
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 9) = WaveX
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 10) = WaveY
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 11) = WaveRz

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 3) = VslX
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 4) = VslY
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 5) = VslZ

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 3) = VslXLF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 4) = VslYLF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 5) = VslZLF

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 3) = VslXWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 4) = VslYWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 5) = VslZWF
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(28, 5) = vslzacc

                    'UPGRADE_WARNING: Couldn't resolve default property of object stiffx. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 2) = stiffx
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(15, 2) = StiffY
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 2) = StiffRz
                Else ' if results are in Metric unit, convert to English to fill in the excel template
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 3) = WindVel / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 7) = WaveHs / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 8) = SwellHs / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(7, 12) = CurrVel / 0.3048

                    ' FORCES  KN  SO * 1000

                    ' Thrusters

                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 9) = 1000 * ThruX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 10) = 1000 * ThruY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object ThruRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 11) = 1000 * ThruRz / 9.80665 / 0.4536 / 0.3048

                    ' Current Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 9) = 1000 * CurrX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 10) = 1000 * CurrY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object CurrRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 11) = 1000 * CurrRz / 9.80665 / 0.4536 / 0.3048
                    ' Wind Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 9) = 1000 * WindX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 10) = 1000 * WindY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WindRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(17, 11) = 1000 * WindRz / 9.80665 / 0.4536 / 0.3048
                    ' Wave Forces
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 9) = 1000 * WaveX / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 10) = 1000 * WaveY / 9.80665 / 0.4536
                    'UPGRADE_WARNING: Couldn't resolve default property of object WaveRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(18, 11) = 1000 * WaveRz / 9.80665 / 0.4536 / 0.3048

                    ' Motions
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslX. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 3) = VslX / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 4) = VslY / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZ. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(25, 5) = VslZ / 0.3048

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 3) = VslXLF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 4) = VslYLF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZLF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(26, 5) = VslZLF / 0.3048

                    'UPGRADE_WARNING: Couldn't resolve default property of object VslXWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 3) = VslXWF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslYWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 4) = VslYWF / 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object VslZWF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(27, 5) = VslZWF / 0.3048

                    ' Acceleration
                    'UPGRADE_WARNING: Couldn't resolve default property of object vslzacc. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(28, 5) = vslzacc / 0.3048

                    ' Mooring Stiffness
                    'UPGRADE_WARNING: Couldn't resolve default property of object stiffx. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(14, 2) = 1000 * stiffx / 9.80665 / 0.4536 * 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(15, 2) = 1000 * StiffY / 9.80665 / 0.4536 * 0.3048
                    'UPGRADE_WARNING: Couldn't resolve default property of object StiffRz. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Cells(16, 2) = 1000 * StiffRz / 9.80665 / 0.4536

                End If
				
				
				For j = 1 To NumLines
					If IsDataEnglishUnit Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(35 + j, 1) = BS(j)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(35 + j, 1) = BS(j) / 9.80665 / 0.4536 * 1000
                    End If
				Next j
				
				
				Row = 35
				'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nl = 0 ' intact line counter
				'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				nd = 0 ' damaged line counter
				'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				DamLineCur = 0 ' current damaged line no.
				For NLA = 1 To NumLineA
					Row = Row + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					nl = nl + 1
					If IsDamaged > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If brokenlineno(nd, NCase) > 0 And nd < TotalNumBroken(NCase) And nl > DamLineCur Then
							'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nd = nd + 1
							'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							DamLineCur = brokenlineno(nd, NCase)
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object DamLineCur. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object nl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If nl = DamLineCur Then
						For Col = 3 To 5
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, Col) = 0
                        Next Col
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(Row + 9, 21) = "Dam"
                    Else
						If IsDataEnglishUnit Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, 3) = Ten(NLA)
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, 4) = TenLF(NLA)
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, 5) = TenWF(NLA)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, 3) = Ten(NLA) / 9.80665 / 0.4536 * 1000
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(Row, 4) = TenLF(NLA) / 9.80665 / 0.4536 * 1000
                            'UPGRADE_WARNING: Couldn't resolve default property of object NLA. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Cells(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
						'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						nrowd = 44 + brokenlineno(nd, NCase)
						'UPGRADE_WARNING: Couldn't resolve default property of object nd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If brokenlineno(nd, NCase) < NumLines Then
							'UPGRADE_WARNING: Couldn't resolve default property of object nrowcop. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							nrowcop = 44 + NumLines - 1
							' move lines after the damaged line one row down
							'UPGRADE_WARNING: Couldn't resolve default property of object nrowcop. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Range("P" & nrowd & ":R" & nrowcop).Copy()
							'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Range("P" & nrowd + 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone)
						End If
						' clean up the row for the damaged line in the tension table
						'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Range("P" & nrowd, "T" & nrowd).FormulaR1C1 = "  "
						'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Range("U" & nrowd).FormulaR1C1 = "DAMAGED"
						'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Range("X" & nrowd, "AB" & nrowd).FormulaR1C1 = "  "
						'UPGRADE_WARNING: Couldn't resolve default property of object nrowd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Range("AC" & nrowd).FormulaR1C1 = "DAMAGED"
					Next nd
				End If
				
				
				'      Set Header and Title
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				With .ActiveSheet.PageSetup
					EndRow = 45 + NumLines
					' set printArea based on result Unit
					If IsDataEnglishUnit Then
						'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.PrintArea = "$O$1:$U$" & EndRow
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.PrintArea = "$V$1:$AC$" & EndRow
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.PrintTitleRows = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.PrintTitleColumns = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.LeftHeader = "&D"
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CenterHeader = " "
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.RightHeader = "Prepared by DTCEL"
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveSheet.PageSetup. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CenterFooter = "David Tein Consulting Engineers, Ltd." & vbCrLf & "11777 Katy Freeway, Suite 434, Houston, TX 77079, 281-531-0888, FAX 281-531-5888, dtcel@dtcel.com"
					
				End With
				
			End With
		Next NCase
		' hide template sheet
		'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Visible. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
        Dim TmpStr As Single
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
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells().Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(4, i).Formula = oxSheet.Name
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(5, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(5, i).Formula = oxSheet.Range("C6").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(6, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(6, i).Formula = GetDamagedLineNo(oxSheet, NumLines)
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells().HorizontalAlignment. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(6, i).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(8, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(8, i).Formula = VB6.Format(oxSheet.Range("P21").Value, "#0.0")
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(9, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(9, i).Formula = VB6.Format(oxSheet.Range("R21").Value, "#0.0")
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(10, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(10, i).Formula = Format(oxSheet.Range("T21").Value, "#0.0")
                        TmpStr = System.Math.Sqrt((oxSheet.Range("Q32").Value) ^ 2 + (oxSheet.Range("S32").Value) ^ 2)
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(12, i).FormulaR1C1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(12, i).FormulaR1C1 = Format(TmpStr, "0.0")
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(13, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(13, i).Formula = oxSheet.Range("P37").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(14, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(14, i).Formula = oxSheet.Range("P38").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(15, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(15, i).Formula = oxSheet.Range("P32").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(16, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(16, i).Formula = oxSheet.Range("Q37").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(17, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(17, i).Formula = oxSheet.Range("Q38").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(18, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(18, i).Formula = oxSheet.Range("Q32").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(19, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(19, i).Formula = oxSheet.Range("U37").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(20, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(20, i).Formula = oxSheet.Range("U38").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(21, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(21, i).Formula = oxSheet.Range("U32").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(22, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(22, i).Formula = oxSheet.Range("R39").Value
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(23, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(23, i).Formula = oxSheet.Range("U39").Value

                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(25, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(25, i).Formula = Format(oxSheet.Range("V13").Value, "0.00") '  2nd Loaded Line Tension
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(26, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(26, i).Formula = Format(oxSheet.Range("T13").Value, "0.00") '  Most Loaded Line Tension
                        'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(27, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .Cells(27, i).Formula = Format(oxSheet.Range("T12").Value, "0.00") '  Min FOS

                        '    .Cells(26, i).Formula = "=MAX(R[+5]C:R[" & (NumLines + 5) & "]C)" '  most-loaded line tension"
                        '    .Cells(25, i).Formula = "=MAX(OFFSET(R[-" & (NumLines + 4) & "]C,R[2]C-1,0),OFFSET(R[-" & (NMoor + 4) & "]C,R[1]C+1,0))" ' 2nd-Loaded Line tension

                        For j = 1 To NumLines
                            If InStr(oxSheet.Range("U" & (44 + j)).Value, "DAMAGED") > 0 Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells().Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                .Cells(29 + j, i).Value = " "
                            Else
                                'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(29 + j, i).Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                .Cells(29 + j, i).Value = oxSheet.Range("U" & (44 + j)).Value
                            End If
						Next j
						TmpStr = Find2MostLoadedLines(CaseNo, NumLines)
						
						'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						nPos = InStr(TmpStr, " ")
						'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If nPos = 0 Then
							Debug.Print("Error finding 2 Loaded Line Numbers: npos=0 tmpStr= " & TmpStr)
						End If
						
						If IsDamaged = 0 Then ' only save intact case most-loaded and second-Loaded
                            ' if damaged case find new most-loaded will flush intact most-loaded and 2nd loaded line numbers
                            ' most loaded line and 2nd loaded line numbers based on tension not FOS
                            'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(29 + NumLines + 2, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(29 + NumLines + 2, i).Formula = Left(TmpStr, nPos - 1)
                            '    .Cells(29 + NumLines + 3, i).Formula = Right(TmpStr, Len(TmpStr) - nPos)

                            '    .Cells(29 + NumLines + 2, i).Formula = "=MATCH(R[-3]C,R[-" & (NumLines + 5) & "]C:R[-6]C,0)"
                            '    .Cells(29 + NumLines + 3, i).Formula = "=IF(OFFSET(R[-" & (NumLines + 3) & "]C,R[-1]C-1,0)<OFFSET(R[-" & (NumLines + 3) & "],R[-1]C+1,0),IF(R[-1]C-1<1,1,R[-1]C-1),IF(R[-1]C+1>" & NumLines & ",1,R[-1]C+1))"
                            'UPGRADE_WARNING: Couldn't resolve default property of object oSummarySheet.Cells(29 + NumLines + 3, i).Formula. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Cells(29 + NumLines + 3, i).Formula = Right(TmpStr, Len(TmpStr) - nPos)

                            'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            L1Tmax(i - 2) = CShort(Left(TmpStr, nPos - 1))
							'UPGRADE_WARNING: Couldn't resolve default property of object nPos. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							L2Tmax(i - 2) = CShort(Right(TmpStr, Len(TmpStr) - nPos))
							'      L1Tmax(i - 2) = .Cells(29 + NumLines + 2, i).Formula
							'      L2Tmax(i - 2) = .Cells(29 + NumLines + 3, i).Formula
						End If
					End If
				Next i
			End With
			
			With .ActiveSheet
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Rows. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Rows("13:98").NumberFormat = "0.00"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Columns. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Columns("B:B").ColumnWidth = 26
				j = oxApp.Sheets.Count + 1
				
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Col = NCase + 4
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(21, Col).FormulaR1C1 = "Max Offset (%WD)"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(22, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(23, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(25, Col).FormulaR1C1 = "FOS"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(26, Col).FormulaR1C1 = "Min"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(27, Col).FormulaR1C1 = "=MIN(RC[-" & NCase + 1 & "]:RC[-2])"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(10, Col).FormulaR1C1 = "Max Mean (ft)"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(11, Col).FormulaR1C1 = ""
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(12, Col).FormulaR1C1 = "=MAX(RC[-" & NCase + 1 & "]:RC[-2])"
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col).FormulaR1C1 = "Min FOS"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 1).FormulaR1C1 = "MaxT(k)"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 2).FormulaR1C1 = "Case"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 3).FormulaR1C1 = "Max FOS"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 4).FormulaR1C1 = "Case"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 5).FormulaR1C1 = ""
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 6).FormulaR1C1 = "MinT(k)"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(28, Col + 7).FormulaR1C1 = "Approx"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(29, Col + 7).FormulaR1C1 = "PreT(k)"
				'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Range(.Cells(29, Col), .Cells(29, Col + 6)).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
				
				
				
				For i = 1 To NumLines
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 7).FormulaR1C1 = preT(i)
					' intact FS
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 6).NumberFormat = "0"
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 7).NumberFormat = "0"
					' The FindMinFOS will return minFOS and CaseNo
					'UPGRADE_WARNING: Couldn't resolve default property of object FindMinFOS(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object minFOS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					minFOS = FindMinFOS(i, CaseNo)
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object minFOS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col).Value = minFOS
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object minFOS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 1).FormulaR1C1 = BS(i) / minFOS
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 2).Value = CaseNo
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 2).NumberFormat = "0"
					' The FindMaxFOS will return maxFOS and CaseNo
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object FindMaxFOS(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 3).Value = FindMaxFOS(i, CaseNo)
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 4).Value = CaseNo
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 4).NumberFormat = "0"
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 1).NumberFormat = "0"
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 5).NumberFormat = "0"
					
					If InStr(oxSheet.Name, "D1") > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0") & "D1"
					ElseIf InStr(oxSheet.Name, "D2") > 0 Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0") & "D2"
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TmpCaseName = "case" & VB6.Format(.Cells(29 + i, Col + 4).Value, "0")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MaxT = IIf(oxApp.Sheets(TmpCaseName).Range("S" & 44 + i).Value > oxApp.Sheets(TmpCaseName).Range("T" & 44 + i).Value, oxApp.Sheets(TmpCaseName).Range("S" & 44 + i), oxApp.Sheets(TmpCaseName).Range("T" & 44 + i))

                    'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    MeanT = oxApp.Sheets(TmpCaseName).Range("P" & 44 + i).value
                    MinT = MeanT - (MaxT - MeanT)
					'UPGRADE_WARNING: Couldn't resolve default property of object Col. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.Cells(29 + i, Col + 6).Value = MinT
					'   .Cells(29 + i, Col + 7).Value = oxApp.Sheets(TmpCaseName).Range("T9")
					
				Next i
				' format most-loaded and 2nd loaded line numbers
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Range(.Cells(29 + NumLines + 2, 2), .Cells(29 + NumLines + 3, NCase + 2)).NumberFormat = "0"
				
				' print environment
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(4, j + 3).FormulaR1C1 = "Wind"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(5, j + 3).FormulaR1C1 = "Wave Hs"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(6, j + 3).FormulaR1C1 = "            Tp"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(7, j + 3).FormulaR1C1 = "Current"
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(4, j + 5).FormulaR1C1 = "  knots"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(5, j + 5).FormulaR1C1 = "  ft"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(6, j + 5).FormulaR1C1 = "  sec"
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Cells(7, j + 5).FormulaR1C1 = "  knots"

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Range("B1").Formula = "Water Depth = " & Format(oxApp.Sheets(TmpCaseName).Range("T6"), "0.0") & " ft, Pretension = " & oxApp.Sheets(TmpCaseName).Range("T9") & " kips"

                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(4, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("P11").Value, "#0.00")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(5, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("G7").Value, "#0.00")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Cells(6, j + 4).FormulaR1C1 = Format(oxApp.Sheets(TmpCaseName).Range("G8").Value, "#0.00")
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.ActiveSheet.Cells. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range().Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
                    'Debug.Print(oxSheet.Range("U" & (44 + LineNo)).Value)
                    'MsgBox(oxSheet.Range("U" & (44 + LineNo)).Value)

                    If IsNumeric(oxSheet.Range("U" & (44 + LineNo)).Value) Then '<> "DAMAGED" Then

                        If Convert.ToDouble(oxSheet.Range("U" & (44 + LineNo)).Value) < MinVal Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object oxSheet.Range().Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets(CaseNo + 2).Name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Debug.Print(.Name & ", i=" & i & ", " & .Range("N" & (44 + i)).Value)
				
				'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("T13").Value) < 0.09 And System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("V13").Value) < 0.09 Then
					' T1 close to T2
					If Line1No = -1 Then Line1No = i
					If Line1No <> i And Line2No = -1 Then Line2No = i
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf System.Math.Abs(.Range("N" & (44 + i)).Value - .Range("T13").Value) < 0.09 And Line1No = -1 Then 
					Line1No = i
					'UPGRADE_WARNING: Couldn't resolve default property of object oxApp.ActiveWorkbook.Sheets().Range. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_NOTE: Object oxBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oxBook = Nothing
		'UPGRADE_NOTE: Object oxBooks may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oxBooks = Nothing
		'UPGRADE_NOTE: Object oxApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oxApp = Nothing
	End Function
	
	Private Function DeleteWorksheet(ByRef strSheetName As String) As Boolean
		On Error Resume Next
		
		ExcelGlobal_definst.Application.DisplayAlerts = False
		'UPGRADE_WARNING: Couldn't resolve default property of object ActiveWorkbook.Worksheets().Delete. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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