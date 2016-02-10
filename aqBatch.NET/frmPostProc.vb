Option Strict Off
Option Explicit On
Friend Class frmPostProc
    Inherits System.Windows.Forms.Form
    Dim NLines As Short
    Dim JustEnterCell As Boolean
    Dim ExistingText As String
    Dim preT() As Double
    Dim preTmin, preTmax As Double
    Dim BS() As Double

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
        Dim InitPos(6) As Single
        Dim i As Short
        Dim msg As String

        msg = ValidateData()

        If Len(msg) > 0 Then GoTo ErrHandler

        ' if Metric, convert to English before sending  to Excel Template
        Dim FrcFactor, LFactor, FrcFactor2 As Double
        If optUnit(1).Checked Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            FrcFactor2 = 1000.0#

        Else
            LFactor = 1
            FrcFactor = 1
            FrcFactor2 = 1.0#

        End If

        For i = 1 To 6
            If i <= 3 Then
                InitPos(i) = CSng(Val(txtInitPos(i - 1).Text)) / LFactor
            Else
                InitPos(i) = CSng(Val(txtInitPos(i - 1).Text))
            End If
        Next

        With grdLinepreTBS
            For i = 1 To NLines
                'preT(i) = CDbl(.Rows(i - 1).Cells(0).Value)
                'BS(i) = CDbl(.Rows(i - 1).Cells(1).Value) / FrcFactor2
            Next i
        End With

        Dim oXR As New ExcelRunner

        On Error GoTo ErrHandler
        oXR.GetAQWAResults(txtAnalysisTitle.Text, txtSubTitle.Text, NumCases, CShort(Val(txtNumLines.Text)), optUnit(0).Checked, IsDamaged, BreakLine, CSng(Val(txtWaterDepth.Text) / LFactor), InitPos, preT, BS, WorkDir, frmMain.chkPFLH.CheckState = System.Windows.Forms.CheckState.Checked)
        oXR = Nothing
        Me.Close()
        Exit Sub
ErrHandler:
        msg = msg & vbCrLf & Err.Description & vbCrLf & err.Source
        MsgBox(msg, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "Error")
        oXR = Nothing
    End Sub

    Public Sub RefreshUnitLabels()
        On Error Resume Next ' ignore error if the form does not have the controls
        Dim i As Short
        For i = 0 To lblLengthUnit.Count - 1
            If optUnit(1).Checked Then
                lblLengthUnit(i).Text = "m"
            Else
                lblLengthUnit(i).Text = "ft"
            End If
        Next i
        For i = 0 To lblForceUnit.Count - 1
            If optUnit(1).Checked Then
                lblForceUnit(i).Text = "KN"
            Else
                lblForceUnit(i).Text = "kips"
            End If
        Next i
    End Sub

    Private Function ValidateData() As String
        Dim i As Short
        Dim msg As String
        If Val(txtNumLines.Text) = 0 Then
            msg = msg & vbCrLf & "Invalid number of mooring lines"
        Else
            If IsDamaged > 0 Then
                For i = 1 To NumCases
                    If UBound(BreakLine) <> NumCases Then
                        msg = msg & vbCrLf & "Must define all broken lines"
                        Exit For
                    Else
                        If Not BreakLine(i) > 0 Then
                            msg = msg & vbCrLf & "Must define all broken lines"
                            Exit For
                        End If
                    End If
                Next
            End If
        End If

        If WorkDir = "" Then
            msg = msg & vbCrLf & "Must define Target directory"
        End If

        If Len(txtAnalysisTitle.Text) = 0 Then
            msg = msg & vbCrLf & "Title must not be empty"
        End If

        ValidateData = msg
    End Function

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Public Function ReadBaseResults() As Boolean
        Dim q As Object
        Dim waterdeth As Object
        Dim tmpVal As Object
        Dim k As Object
        ReadBaseResults = False

        Dim FrcFactor, LFactor, FrcFactor2 As Double

        If optUnit(1).Checked Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            FrcFactor2 = 1000.0#
        Else
            LFactor = 1
            FrcFactor = 1
            FrcFactor2 = 1.0#

        End If

        ' read no environ base file to extract needed information
        Const ForReading As Short = 1
        Const ForWriting As Short = 2
        Const ForAppending As Short = 3
        Const TristateUseDefault As Short = -2
        Const TristateTrue As Short = -1
        Const TristateFalse As Short = 0
        Dim ts, fs, f, MyFile As Object
        Dim ss, s, TmpStr As String
        Dim i As Short

        fs = CreateObject("Scripting.FileSystemObject")

        On Error GoTo ErrHandler
        If Not fs.FileExists(frmMain.lblBaseDir.Text & "\abrun.lis") Then
            MsgBox("Your must run a base case with no environment first before reporting.")
            Exit Function
        End If
        ' read ABFile
        f = fs.GetFile(frmMain.lblBaseDir.Text & "\abrun.lis")
        ts = f.OpenAsTextStream(ForReading, TristateFalse)
        s = ts.Readall
        ts.Close()

        Dim s4, s2, s1, s3, s5 As String
        Dim pos2, pos1, pos3 As Integer

        'find title
        pos1 = InStr(1, s, "TITLE ") ' TITLE
        pos2 = InStr(pos1, s, vbCrLf) ' TITLE

        TmpStr = Trim(Mid(s, pos1 + 8, pos2 - pos1 - 8))

        ' Formulate TITLE
        If IsDamaged = 1 Then
            txtAnalysisTitle.Text = TmpStr & ", Damage Most Critical Line" ' title
        ElseIf IsDamaged = 2 Then
            txtAnalysisTitle.Text = TmpStr & ", Damage 2nd Critical Line" ' title
        Else
            txtAnalysisTitle.Text = TmpStr & ", Intact" ' title
        End If

        txtSubTitle.Text = frmMain.txtMetoceanName.Text

        ' get number of mooring lines
        Dim Fields() As String
        Dim Fields1() As String
        Dim MatchStr As String

        'If InStr(1, s, "14NLIN") > 0 Then 'JLIU 03/09/2015 assume mooring line always dynamic NLID and tension line is always staic "NLIN
        '   MatchStr = "14NLIN"
        'Else
        MatchStr = "14NLID"
        'End If

        Fields = Split_Renamed(s, MatchStr)

        On Error GoTo 0
        On Error Resume Next

        NLines = UBound(Fields)

        ReDim preT(NLines)
        ReDim BS(NLines)

        pos1 = InStr(1, s, "14COMP ") ' initial x

        ' determine how many mooring groups
        Dim BSLine(NLines) As Short
        Dim grpNo, Ngrp As Short

        Fields = Split_Renamed(s, "14COMP ")
        Ngrp = UBound(Fields)

        WaterDepth = 0
        For k = 1 To Ngrp
            Fields1 = Split_Renamed(Fields(k), " ")
            tmpVal = (CDbl(Fields1(4)) + CDbl(Fields1(5))) / 2
            'tmpVal = (Val(Trim(Mid(Fields(k), 9, 10))) + Val(Trim(Mid(Fields(k), 19, 10)))) / 2
            If tmpVal > waterdeth Then WaterDepth = tmpVal
        Next k

        txtWaterDepth.Text = CStr(WaterDepth)

        If Ngrp < NLines Then ' grouped
            'lblWarning.Visible = True
        Else
            'lblWarning.Visible = False
        End If

        Dim LastI As Short

        i = NLines
        LastI = NLines
        k = 0
        pos1 = InStrRev(s, MatchStr, Len(s))
        Do Until i <= 0
            pos2 = pos1
            pos1 = InStrRev(s, MatchStr, pos1 - 1)
            If pos2 - pos1 < 84 Then ' grouped lines
                k = k + 1
            Else ' go get ECAT BS
                pos2 = InStrRev(s, "14ECAT", pos2)
                BS(i) = CDbl(VB6.Format(Val(CStr(CDbl(Trim(Mid(s, pos2 + 56, 10))) / 1000.0#)), "#0.000")) 'JLIU
                LastI = i
                If k > 0 Then
                    For q = 0 To k
                        BS(LastI + q) = BS(i)
                    Next q
                End If
                k = 0
            End If
            If pos1 > 0 Then i = i - 1
        Loop

        '      ' search from last line backward for all line properties - ungrouped - each has own ECAT
        '      pos1 = InStrRev(s, MatchStr, Len(s))
        '
        '      For i = NLines To 1 Step -1
        '          ' check if group properties exist
        '          pos1 = InStrRev(s, "14ECAT", pos1)
        '          BS(i) = Format(Val(Trim(Mid(s, pos1 + 56, 10))) / 1000, "#0.0")
        '          pos1 = InStrRev(s, MatchStr, pos1 - 1)
        '      Next i


        pos1 = InStr(1, s, " X   = ") ' initial x
        pos2 = InStr(pos1, s, " Y   = ") ' initial y
        pos3 = InStr(pos2, s, " Z   = ") ' initial z

        txtInitPos(0).Text = Trim(Mid(s, pos1 + 7, 22))
        txtInitPos(1).Text = Trim(Mid(s, pos2 + 7, 22))
        txtInitPos(2).Text = Trim(Mid(s, pos3 + 7, 22))

        pos1 = InStr(1, s, " RX   = ") ' initial x
        pos2 = InStr(pos1, s, " RY   = ") ' initial y
        pos3 = InStr(pos2, s, " RZ   = ") ' initial z

        txtInitPos(3).Text = Trim(Mid(s, pos1 + 7, 22))
        txtInitPos(4).Text = Trim(Mid(s, pos2 + 7, 22))
        txtInitPos(5).Text = Trim(Mid(s, pos3 + 7, 22))


        pos1 = InStrRev(s, "ITER NO.              ")
        ' read pretension for each legs
        preTmin = 1.0E+16
        preTmax = -1.0E+16
        For i = 1 To NLines
            ' start from last found pos1 RX point, looking back for last iteration
            pos1 = InStr(pos1 + 5, s, "  TEN ") 'JLIU 3/9/2015 changed from TOT to TEN
            preT(i) = CDbl(VB6.Format(Val(CStr(CDbl(Trim(Mid(s, pos1 + 6, 10))) / 1000.0#)), "#0.000"))
            If preT(i) > preTmax Then
                preTmax = preT(i)
            End If
            If preT(i) < preTmin Then
                preTmin = preT(i)
            End If

            '      Debug.Print "i= " & i & " T= " & preT(i)
        Next i

        txtNumLines.Text = CStr(NLines)

        ReadBaseResults = True

        Exit Function
ErrHandler:
        MsgBox("Error: " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.OKOnly, "Error")
    End Function

    Private Sub grdLinepreTBS_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'txtEdit.Visible = False
    End Sub

    Private Sub optUnit_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUnit.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optUnit.GetIndex(eventSender)
            RefreshUnitLabels()
        End If
    End Sub

    Private Sub txtNumLines_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumLines.TextChanged
        Dim i As Object

        Dim FrcFactor, LFactor, FrcFactor2 As Double
        If optUnit(1).Checked Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            FrcFactor2 = 1000.0#
        Else
            LFactor = 1
            FrcFactor = 1
            FrcFactor2 = 1.0#
        End If

        NLines = Val(txtNumLines.Text)
        If NLines > 0 Then
            ReDim Preserve preT(NLines)
            ReDim Preserve BS(NLines)
        End If

        With grdLinePreTBS
            .ColumnCount = 2
            .RowCount = NLines
            .Columns(0).Width = 120
            .Columns(1).Width = 120
            .Columns(0).HeaderText = "Pretension " & lblForceUnit(0).Text
            .Columns(1).HeaderText = "Breaking Strength " & lblForceUnit(0).Text

            ' load current profile data from object
            For i = 1 To NLines
                .Rows(i - 1).HeaderCell = i
                .Rows(i - 1).Cells(0).Value = VB6.Format(preT(i) * FrcFactor2, "0")
                .Rows(i - 1).Cells(1).Value = VB6.Format(BS(i) * FrcFactor2, "0")

            Next i
        End With
    End Sub


End Class