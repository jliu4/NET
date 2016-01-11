Option Strict Off
Option Explicit On
Module modMathFunctions
	
	
	Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
	
	Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	
	Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	
	' input handling
	Private Const DelimiterList As String = " ,=/;~!@#$%^&*()_+{}[]|\:<>?/" & vbTab
	Private Const MaxDelimiters As Short = 32
	Private Const PI As Double = 3.1415926
	
	' maximum points in spline fitting
	Private Const MaxPtNum As Short = 100
	
	'UPGRADE_NOTE: Split was upgraded to Split_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Split_Renamed(ByVal InputText As String, Optional ByVal Delimiter As String = "") As Object
		
		' This function splits the sentence in InputText into
		' words and returns a string array of the words. Each
		' element of the array contains one word.
		
		' This constant contains punctuation and characters
		' that should be filtered from the input string.
		Const CHARS As String = "!?,;:""'()[]{}" ' removed period dot from this
		Dim strReplacedText As String
		Dim intIndex As Short
		
		' Replace tab characters with space characters.
		strReplacedText = Trim(Replace(InputText, vbTab, " "))
		
		' Filter all specified characters from the string.
		For intIndex = 1 To Len(CHARS)
			strReplacedText = Trim(Replace(strReplacedText, Mid(CHARS, intIndex, 1), " "))
		Next intIndex
		
		' Loop until all consecutive space characters are
		' replaced by a single space character.
		Do While InStr(strReplacedText, "  ")
			strReplacedText = Replace(strReplacedText, "  ", " ")
		Loop 
		
		' Split the sentence into an array of words and return
		' the array. If a delimiter is specified, use it.
		'MsgBox "String:" & strReplacedText
		If Len(Delimiter) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Split_Renamed = Split(strReplacedText)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Split_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Split_Renamed = Split(strReplacedText, Delimiter)
		End If
	End Function
	
	Function IsActive(ByVal hProg As Integer) As Integer
		Dim hProc, RetVal As Integer
		hProc = OpenProcess(0, False, hProg)
		If hProc <> 0 Then GetExitCodeProcess(hProc, RetVal)
		IsActive = (RetVal = 259)
		CloseHandle(hProc)
	End Function
	
	
	Public Sub MatMul(ByRef A As Object, ByRef B As Object, ByRef c As Object)
		
		Dim i, j As Short
		Dim sum As Double
		
		For i = 1 To UBound(A)
			sum = 0#
			For j = 1 To UBound(A)
				'UPGRADE_WARNING: Couldn't resolve default property of object B(j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sum = A(i, j) * B(j) + sum
			Next j
			'UPGRADE_WARNING: Couldn't resolve default property of object c(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			c(i) = sum
		Next i
		
	End Sub
	
	Public Sub Invert(ByRef A As Object)
		
		Dim j, i, k As Short
		Dim Temp As Double
		Dim L, U As Short
		
		L = LBound(A)
		U = UBound(A)
		For k = L To U
			'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Temp = A(k, k)
			'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			A(k, k) = 1#
			For j = L To U
				'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				A(k, j) = A(j, k) / Temp
			Next j
			For i = L To U
				If i <> k Then
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Temp = A(i, k)
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(i, k) = 0#
					For j = L To U
						'UPGRADE_WARNING: Couldn't resolve default property of object A(k, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						A(i, j) = A(i, j) - Temp * A(k, j)
					Next j
				End If
			Next i
		Next k
		
	End Sub
	
	Public Sub SqMatMul(ByVal A As Object, ByVal B As Object, ByRef c As Object)
		
		Dim k, i, j, N As Short
		Dim sum As Double
		
		N = UBound(A)
		For i = 1 To N
			For j = 1 To N
				sum = 0#
				For k = 1 To N
					'UPGRADE_WARNING: Couldn't resolve default property of object B(k, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sum = A(i, k) * B(k, j) + sum
				Next k
				'UPGRADE_WARNING: Couldn't resolve default property of object c(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c(i, j) = sum
			Next j
		Next i
		
	End Sub
	
	' Public Sub GaussJordan(ByRef A, ByVal N As Integer, ByRef B, ByVal M As Integer)
	Public Sub GaussJordan(ByRef A As Object, ByRef N As Object, ByRef B As Object, ByRef M As Object)
		
		Dim L, j, i, k, ll As Short
		'UPGRADE_WARNING: Lower bound of array Rindex was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim c, R As Short
		Dim Rindex(50) As Short
		'UPGRADE_WARNING: Lower bound of array Pivot was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Pivot(50) As Double
		'UPGRADE_WARNING: Lower bound of array Cindex was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim Cindex(50) As Short
		Dim Dummy, Big, Pinverse As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For j = 1 To N
			Pivot(j) = 0
		Next j
		'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For i = 1 To N
			Big = 0#
			'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To N
				If Pivot(j) <> 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For k = 1 To N
						If Pivot(k) = 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If System.Math.Abs(A(j, k)) >= Big Then
								'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Big = System.Math.Abs(A(j, k))
								R = j
								c = k
							End If
						ElseIf Pivot(k) > 1 Then 
							Call MsgBox("Singular matrix inversion attempted", MsgBoxStyle.OKOnly, "MARS - Calculation Failure")
							Exit Sub
						End If
					Next k
				End If
			Next j
			Pivot(c) = Pivot(c) + 1
			If R <> c Then
				'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For L = 1 To N
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dummy = A(R, L)
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(R, L) = A(c, L)
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(c, L) = Dummy
				Next L
				'UPGRADE_WARNING: Couldn't resolve default property of object M. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For L = 1 To M
					'UPGRADE_WARNING: Couldn't resolve default property of object B(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dummy = B(R, L)
					'UPGRADE_WARNING: Couldn't resolve default property of object B(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					B(R, L) = B(c, L)
					'UPGRADE_WARNING: Couldn't resolve default property of object B(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					B(c, L) = Dummy
				Next L
			End If
			Rindex(i) = R
			Cindex(i) = c
			'UPGRADE_WARNING: Couldn't resolve default property of object A(c, c). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If A(c, c) = 0 Then
				MsgBox("Singular matrix inversion attempted", MsgBoxStyle.OKOnly, "MARS - Calculation Failure")
				Exit Sub
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Pinverse = 1# / A(c, c)
			'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			A(c, c) = 1#
			'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For L = 1 To N
				'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				A(c, L) = A(c, L) * Pinverse
			Next L
			'UPGRADE_WARNING: Couldn't resolve default property of object M. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For L = 1 To M
				'UPGRADE_WARNING: Couldn't resolve default property of object B(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				B(c, L) = B(c, L) * Pinverse
			Next L
			'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For ll = 1 To N
				If ll <> c Then
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dummy = A(ll, c)
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(ll, c) = 0#
					'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For L = 1 To N
						'UPGRADE_WARNING: Couldn't resolve default property of object A(c, L). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						A(ll, L) = A(ll, L) - A(c, L) * Dummy
					Next L
					'UPGRADE_WARNING: Couldn't resolve default property of object M. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For L = 1 To M
						'UPGRADE_WARNING: Couldn't resolve default property of object B(c, L). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object B(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						B(ll, L) = B(ll, L) - B(c, L) * Dummy
					Next L
				End If
			Next ll
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For L = N To 1 Step -1
			If Rindex(L) <> Cindex(L) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object N. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For k = 1 To N
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Dummy = A(k, Rindex(L))
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(k, Rindex(L)) = A(k, Cindex(L))
					'UPGRADE_WARNING: Couldn't resolve default property of object A(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					A(k, Cindex(L)) = Dummy
				Next k
			End If
		Next L
		
		Exit Sub
		
	End Sub
	
	
	' integer part with round-up
	Public Function Round(ByRef X As Object) As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Round = Fix(X + 0.5 * System.Math.Sign(X))
		
	End Function
	
	' minimum of two number
	Public Function Min(ByRef A As Object, ByRef B As Object) As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object B. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object A. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (A < B) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object A. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Min = A
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object B. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Min = B
		End If
		
	End Function
	
	' maximum of two number
	Public Function Max(ByRef A As Object, ByRef B As Object) As Double
		
		'UPGRADE_WARNING: Couldn't resolve default property of object B. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object A. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (A < B) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object B. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Max = B
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object A. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Max = A
		End If
		
	End Function
	
	' arcsine (-pi/2 to pi/2)
	Public Function Asin(ByRef X As Object) As Double
		
		If X = 1# Then
			Asin = PI / 2#
		ElseIf X = -1# Then 
			Asin = -PI / 2#
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Asin = System.Math.Atan(X / System.Math.Sqrt(-X * X + 1))
		End If
		
	End Function
	
	' arccosine (0 to pi)
	Public Function Acos(ByRef X As Object) As Double
		
		If X = 1# Then
			Acos = 0#
		ElseIf X = -1# Then 
			Acos = PI
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Acos = System.Math.Atan(-X / System.Math.Sqrt(-X * X + 1)) + PI / 2#
		End If
		
	End Function
	
	' angle determined by coordinate x and y (0 to 2pi) or (-pi to pi)
	Public Function Atan(ByRef X As Object, ByRef Y As Object, Optional ByRef NPi2Pi As Boolean = False) As Double
		
		If X > 0# Then
			If Y >= 0# Then
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Atan = System.Math.Atan(Y / X)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Atan = 2# * PI + System.Math.Atan(Y / X)
			End If
		ElseIf X < 0# Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Atan = PI + System.Math.Atan(Y / X)
		Else
			If Y > 0# Then
				Atan = PI * 0.5
			ElseIf Y < 0# Then 
				Atan = PI * 1.5
			Else
				Atan = 0#
			End If
		End If
		
		If NPi2Pi And Atan > PI Then Atan = Atan - 2 * PI
		
	End Function
	
	' determine coefficients of cubic spline
	'Invalid_string_refer_to_original_code
	Public Sub Spline3(ByRef N As Short, ByRef X() As Single, ByRef A() As Single, ByRef B() As Single, ByRef c() As Single, ByRef D() As Single, ByRef Clamped As Boolean)
		
		' Input
		'   n -- number of spline segments
		'   x -- series of x value
		'   a -- series of a value (=f(x))
		'   clamped -  flag clamped end
		' Output
		'   b -- series of b value
		'   c -- series of c value
		'   d -- series of d value
		
		Dim i As Short
		Dim H(MaxPtNum) As Single
		Dim L(MaxPtNum) As Single
		Dim Al(MaxPtNum) As Single
		Dim Mu(MaxPtNum) As Single
		Dim z(MaxPtNum) As Single
		Dim Fpo, Fpn As Single
		
		For i = 1 To N
			H(i) = X(i + 1) - X(i)
		Next 
		
		If Clamped Then
			Fpo = 0#
			Fpn = 0#
			Al(1) = 3# * ((A(2) - A(1)) / H(1) - Fpo)
			Al(N + 1) = 3# * (Fpn - (A(N + 1) - A(N)) / H(N))
			
			L(1) = 2# * H(1)
			Mu(1) = 0.5
			z(1) = Al(1) / L(1)
		Else
			L(1) = 1#
			Mu(1) = 0#
			z(1) = 0#
		End If
		
		For i = 2 To N
			Al(i) = 3# * ((A(i + 1) - A(i)) / H(i) - (A(i) - A(i - 1)) / H(i - 1))
		Next 
		
		For i = 2 To N
			L(i) = 2# * (H(i) + H(i - 1)) - H(i - 1) * Mu(i - 1)
			Mu(i) = H(i) / L(i)
			z(i) = (Al(i) - H(i - 1) * z(i - 1)) / L(i)
		Next 
		
		If Clamped Then
			L(N + 1) = H(N) * (2# - Mu(N))
			z(N + 1) = (Al(N + 1) - H(N) * z(N)) / L(N + 1)
			c(N + 1) = z(N + 1)
			B(N + 1) = 0#
			D(N + 1) = 0#
		Else
			L(N + 1) = 1#
			z(N + 1) = 0#
			c(N + 1) = z(N + 1)
			B(N + 1) = 0#
			D(N + 1) = 0#
		End If
		
		For i = N To 1 Step -1
			c(i) = z(i) - Mu(i) * c(i + 1)
			B(i) = (A(i + 1) - A(i)) / H(i) - H(i) * (c(i + 1) + 2# * c(i)) / 3#
			D(i) = (c(i + 1) - c(i)) / (3# * H(i))
		Next 
		
	End Sub
	
	
	
	Public Function Atn2(ByRef X As Object, ByRef Y As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If X = 0 Then ' On Y-axis
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Y > 0 Then ' Positive Y-axis
				'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Atn2 = PI / 2#
				'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf Y < 0 Then  ' Negative Y-axis
				'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Atn2 = -PI / 2#
			Else ' (0,0) Point
				'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Atn2 = 0#
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf X > 0 And Y >= 0 Then  ' First quadrant
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Atn2 = System.Math.Atan(Y / X)
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf X < 0 And Y >= 0 Then  ' Second quadrant
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Atn2 = PI - System.Math.Atan(-Y / X)
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf X < 0 And Y <= 0 Then  ' Third quadrant
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Atn2 = System.Math.Atan(Y / X) + PI
		Else ' Fourth quadrant
			'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Atn2. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Atn2 = 2# * PI - System.Math.Atan(-Y / X)
		End If
	End Function
	
	Public Function RadTo360(ByRef X As Double) As Double
		Dim A As Double
		Dim t As Short
		A = X * 180 / PI
		t = A \ 360#
		A = A - (360 * t)
		If A < 0 Then
			A = A + 360#
		End If
		If System.Math.Abs(System.Math.Abs(A) - 360#) < 0.01 Then
			A = 0#
		End If
		RadTo360 = A
	End Function
	Public Function ZeroToTwoPi(ByRef X As Single) As Single
		' Reduce an angle to the range 0-2*Pi (no negatives allowed!)
		Dim A As Double
		Dim t As Short
		If X < 0 And X > -2# * PI Then
			A = X + 2# * PI
		Else
			t = X \ (2# * PI)
			A = X - (2# * PI * t)
			If A < 0 Then A = A + 2 * PI
		End If
		ZeroToTwoPi = A
	End Function
End Module