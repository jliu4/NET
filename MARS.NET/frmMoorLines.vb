Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMoorLines
	Inherits System.Windows.Forms.Form
	
	Private Const InFile As Short = 1
	Private Const OutFile As Short = 2
	
	Private Const DraftStr1 As String = "Current Draft"
	Private Const DraftStr2 As String = "Survival Draft"
	Private Const DraftStr3 As String = "Operating Draft"
	
	Private ExistingTxt As String
	Private Changed As Boolean
	Private CheckingGrid As Boolean
	Private JEC As Boolean

    Private SegLabels(13) As String
    Private CurLine As Short
	Private NumLines As Short
	Private ShipDesLoc As ShipGlobal
	Private MoorLines As MoorSystem

    Private CurDraft As Single
    Private SurDraft As Single
    Private OprDraft As Single

    Private Updated As Boolean
	Private TenCalc As Boolean
	Private LocChanged As Boolean
	
	Public ChangeInCat As Boolean
	
	Private LUnit, FrcUnit As String
    Dim FrcFactor, LFactor, StressFactor As Single
    Dim DiameterFactor As Object

    Private Sub RefreshData()
		If IsMetricUnit Then
			LFactor = 0.3048 ' ft -> m
			FrcFactor = 4.448222 ' kips -> KN
			StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
			FrcUnit = "KN"
		Else
			LFactor = 1
			FrcFactor = 1
			StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
			FrcUnit = "kips"
		End If
		IniSegments()
		RefreshUnitLabels(Me)
		LoadMoorData(True)
	End Sub

    ' form load and unload

    Private Sub frmMoorLines_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Text = Text & " - " & CurProj.Title
        CheckingGrid = False
        ChangeInCat = False

        LoadMoorLines()

        If CurLine = 0 Or CurLine > NumLines Then CurLine = 1

        Updated = True
        InitiateTabs()

        RefreshData()

    End Sub

    Private Sub frmMoorLines_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus

        If ChangeInCat Then btnTopTension_Click(btnTopTension, New System.EventArgs())
        ChangeInCat = False

    End Sub

    Private Sub frmMoorLines_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        MoorLines = Nothing

    End Sub

    ' command buttons
    ' form level

    Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        With ShipDesLoc
            .Xg = CSng(txtVslSt(0).Text) / LFactor
            .Yg = CSng(txtVslSt(1).Text) / LFactor
            .Heading = CSng(txtVslSt(2).Text) * Degrees2Radians
        End With

        With CurVessel.ShipDesGlob
            .Xg = ShipDesLoc.Xg
            .Yg = ShipDesLoc.Yg
            .Heading = ShipDesLoc.Heading
        End With

        If btnAnchor.Enabled Then btnAnchor_Click(btnAnchor, New System.EventArgs())
        If btnScope.Enabled Then btnScope_Click(btnScope, New System.EventArgs())

        If Not Updated Then UpdateMoorData()
        Dim i As Short
        fraSegments.Visible = False
        For i = MoorLines.MoorLineCount To 1 Step -1
            Updated = False
            tabMoorLines.SelectedIndex = i - 1 'JLIU TODO
        Next i
        Updated = True
        tabMoorLines.SelectedIndex = CurLine - 1
        fraSegments.Visible = True

        SaveMoorLines()

        CurProj.Saved = False
        Cursor = System.Windows.Forms.Cursors.Default
        Me.Close()

    End Sub

    Private Sub grdAnchor_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdAnchor.DoubleClick

        'If grdAnchor.Rowindex = 0 Then
        'Select Case grdAnchor.Col
        'Case 2, 4, 5, 7
        'MSFlexGridEdit(grdAnchor, txtAnchEdit, System.Windows.Forms.Keys.F2)
        'End Select
        'End If

    End Sub

    Private Sub grdAnchor_EnterCell(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles grdAnchor.CellEnter

        If e.RowIndex = 0 Then
            Select Case e.ColumnIndex
                Case 2, 4, 5, 7
                    ExistingTxt = Trim(grdAnchor.Text)
                    JEC = True
            End Select
        End If

    End Sub

    'Private Sub grdAnchor_KeyDownEvent(ByVal eventSender As System.Object, ByVal e As KeyEventArgs) Handles grdAnchor.KeyDown

    'If e.RowIndex = 0 Then
    'Select Case grdAnchor.Col
    'Case 2, 4, 5, 7
    'AddHandler(grdAnchor, txtAnchEdit, e.KeyCode, e.Shift, JEC, ExistingTxt)
    'End Select
    'End If

    'End Sub

    Private Sub grdAnchor_LeaveCell(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles grdAnchor.CellLeave

        Dim Changed As Boolean

        ' If txtAnchEdit.Visible = True Then
        'grdAnchor.Text = txtAnchEdit.Text
        'txtAnchEdit.Visible = False
        'End If

        Changed = False
        If Not CheckingGrid Then
            If e.RowIndex = 0 Then
                Select Case e.ColumnIndex
                    Case 2, 5
                        If Trim(grdAnchor.Text) <> ExistingTxt Then Changed = True
                    Case 4, 7
                        If Trim(grdAnchor.Text) <> ExistingTxt Then
                            grdAnchor.Text = CheckData(grdAnchor.Text)
                            Changed = True
                        End If
                End Select
            End If
        End If

        If Changed Then Updated = False

    End Sub

    Private Sub grdAnchor_Leave(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles grdAnchor.Leave

        Dim Changed As Boolean

        Changed = False
        If Not CheckingGrid Then
            If e.RowIndex = 0 Then
                Select Case e.ColumnIndex
                    Case 2, 5
                        If Trim(grdAnchor.Text) <> ExistingTxt Then Changed = True
                    Case 4, 7
                        If Trim(grdAnchor.Text) <> ExistingTxt Then
                            grdAnchor.Text = CheckData(grdAnchor.Text)
                            Changed = True
                        End If
                End Select
            End If
        End If

        If Changed Then Updated = False

    End Sub

    Private Sub SetSegTypeCbo(ByRef SegType As String)

        Dim i As Short

        ' For i = 1 To cboSegType.Items.Count
        'If SegType = VB6.GetItemString(cboSegType, i - 1) Then Exit For
        'Next
        'If i > cboSegType.Items.Count Then i = 1
        'cboSegType.SelectedIndex = i - 1

    End Sub

    Private Sub grdSegments_CellValidating(
      ByVal sender As Object,
      ByVal e As DataGridViewCellValidatingEventArgs) _
      Handles grdSegments.CellValidating

        grdSegments.Rows(e.RowIndex).ErrorText = ""

        If grdSegments.Rows(e.RowIndex).IsNewRow Then Return

        If Not (e.ColumnIndex = 0) Then
            If Not (IsNumeric(e.FormattedValue)) And Not String.IsNullOrEmpty(e.FormattedValue) Then
                e.Cancel = True
                grdSegments.Rows(e.RowIndex).ErrorText =
                   "It must be a numeric value."
            End If
        End If

    End Sub

    Private Sub grdSegments_ClickEvent(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles grdSegments.CellClick
        If e.ColumnIndex = 0 Then

            Dim cell As New DataGridViewComboBoxCell()
            With cell
                .FlatStyle = FlatStyle.Flat
                .MaxDropDownItems = 7
                .Items.Add("WIRE")
                .Items.Add("CHAIN")
                .Items.Add("WIRE-P")
                .Items.Add("CHAIN-P")
                .Items.Add("POLY-P")
                .Items.Add("BUOY-P")
                .Items.Add("SPRING") '2.2.1
            End With

            cell.Value = "WIRE"
            grdSegments(e.ColumnIndex, e.RowIndex) = cell

        End If
    End Sub


    Private Sub grdSegments_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'cboSegType.Visible = False
        'txtSegEdit.Visible = False
    End Sub

    ' menus

    Public Sub mnuClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuClose.Click

        btnCancel_Click(btnCancel, New System.EventArgs())

    End Sub

    Public Sub mnuOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOpen.Click

        Dim ReadFile, isFileOpen As Boolean
        Dim msg As String

        isFileOpen = False

        '   should the user cancel the dialog box, exit
        On Error GoTo Errhandler

        dlgFileOpen.Filter = "All Files (*.*)|*.*|Mln Files (*.mln)|*.mln"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Mln Files (*.mln)|*.mln"
        dlgFileOpen.FilterIndex = 2
        dlgFileSave.FilterIndex = 2
        dlgFileOpen.InitialDirectory = CurProj.Directory
        dlgFileSave.InitialDirectory = CurProj.Directory
        dlgFileOpen.CheckFileExists = True
        dlgFileOpen.CheckPathExists = True
        dlgFileSave.CheckPathExists = True
        dlgFileOpen.ShowReadOnly = False
        dlgFileOpen.ShowDialog()
        dlgFileSave.FileName = dlgFileOpen.FileName

        '   open the file
        FileOpen(InFile, dlgFileOpen.FileName, OpenMode.Input, OpenAccess.Read)
        isFileOpen = True

        On Error GoTo Errhandler
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        ReadFile = MoorLines.InputML(InFile)
        '   close the input file and return
        FileClose(InFile)
        isFileOpen = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        If Not ReadFile Then
            msg = dlgFileOpen.FileName & " appears to be corrupted. " & "The original mooring system has been reloaded."
            MsgBox(msg, MsgBoxStyle.OkOnly, Me.Text)
            LoadMoorLines()
        End If
        LoadMoorData()

        Exit Sub

Errhandler:
        '   User pressed Cancel button, or some tragedy occurred
        If isFileOpen Then FileClose(InFile)
        Exit Sub

    End Sub

    Public Sub mnuSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSave.Click

        Dim WriteFile, FileOpen_Renamed As Boolean

        FileOpen_Renamed = False

        On Error GoTo Errhandler

        dlgFileOpen.Filter = "All Files (*.*)|*.*|Mln Files (*.mln)|*.mln"
        dlgFileSave.Filter = "All Files (*.*)|*.*|Mln Files (*.mln)|*.mln"
        dlgFileOpen.FilterIndex = 2
        dlgFileSave.FilterIndex = 2
        dlgFileOpen.FileName = "*.mln"
        dlgFileSave.FileName = "*.mln"
        dlgFileOpen.InitialDirectory = CurProj.Directory
        dlgFileSave.InitialDirectory = CurProj.Directory
        dlgFileSave.OverwritePrompt = True
        dlgFileOpen.ShowReadOnly = False
        dlgFileSave.ShowDialog()
        dlgFileOpen.FileName = dlgFileSave.FileName

        '   open the file for output
        FileOpen(OutFile, dlgFileOpen.FileName, OpenMode.Output)
        FileOpen_Renamed = True

        On Error GoTo 0
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        WriteFile = MoorLines.OutputML(OutFile)
        '   Close the output file
        FileClose(OutFile)
        FileOpen_Renamed = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

Errhandler:
        '   User pressed Cancel button, or some tragedy occurred
        If FileOpen_Renamed Then FileClose(OutFile)
        MsgBox("Error saving: " & Err.Description, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Error")
    End Sub

    Private Sub tabMoorLines_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) _
        Handles tabMoorLines.SelectedIndexChanged

        Dim SelLine As Short

        If Not Updated Then UpdateMoorData()

        SelLine = tabMoorLines.SelectedIndex

        If SelLine + 1 <> CurLine Then
            CurLine = SelLine + 1

            LoadMoorData()
            tabMoorLines.TabPages(SelLine).Controls.Add(Me.Panel1)
        End If

    End Sub

    ' command buttons
    ' tab level

    Private Sub btnAnchor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAnchor.Click
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        Dim SprdAngle, Scope, Alpha As Single
        Dim FLXs, FLYs As Single
        Dim AnchXg, AnchXs, AnchYs, AnchYg As Single

        Dim FL As New FairLead

        If LocChanged Then
            With ShipDesLoc
                .Xg = CSng(txtVslSt(0).Text) / LFactor
                .Yg = CSng(txtVslSt(1).Text) / LFactor
                .Heading = CSng(txtVslSt(2).Text) * Degrees2Radians
            End With
        End If

        SprdAngle = CSng(-txtGeneral(0).Text) * Degrees2Radians
        Scope = CSng(txtGeneral(1).Text) / LFactor

        FLXs = CSng(txtFL(0).Text) / LFactor
        FLYs = CSng(txtFL(1).Text) / LFactor

        FL.Node = _txtFLNode.Text
        FL.Xs = FLXs
        FL.Ys = FLYs

        AnchXs = Scope * System.Math.Cos(SprdAngle) + FLXs
        AnchYs = Scope * System.Math.Sin(SprdAngle) + FLYs

        Alpha = PI / 2.0# - ShipDesLoc.Heading
        AnchXg = System.Math.Cos(Alpha) * AnchXs - System.Math.Sin(Alpha) * AnchYs + ShipDesLoc.Xg
        AnchYg = System.Math.Sin(Alpha) * AnchXs + System.Math.Cos(Alpha) * AnchYs + ShipDesLoc.Yg

        txtAnchor(0).Text = Format(AnchXg * LFactor, "0.00")
        txtAnchor(1).Text = Format(AnchYg * LFactor, "0.00")

        btnAnchor.Enabled = False
        btnScope.Enabled = False

    End Sub

    Private Sub btnCatenary_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCatenary.Click

        Dim i, NumSeg As Short

        Dim Length, LenStr As Single

        Dim SegLength(MaxNumSeg) As Single
        Dim SegTension(MaxNumSeg) As Single
        Dim SegAngle(MaxNumSeg) As Single
        Dim SegPosition(MaxNumSeg) As Single
        Dim CatX(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim CatY(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim Connector(MaxNumSeg + 1) As Short

        If Not TenCalc Then btnTopTension_Click(btnTopTension, New System.EventArgs())

        With MoorLines.MoorLines(CurLine)
            NumSeg = .SegmentCount
            If NumSeg > MaxNumSeg Then NumSeg = MaxNumSeg
            For i = 1 To NumSeg Step 1
                With .Segments(i)
                    SegLength(i) = .Length
                    SegTension(i) = .TenUpp
                    SegAngle(i) = .AngUpp
                    If i < NumSeg Then SegPosition(i + 1) = .YLow
                End With
            Next i
            SegLength(1) = .Payout
            SegPosition(1) = .Draft - .FairLead.z

            Length = .LineLen
            LenStr = .LineLenStr

            If Not .CatenaryPoints(CatX, CatY, Connector) Then Exit Sub

            frmCatenary.UpdateForm(NumLines, CurLine, NumSeg, .DesScope, Length, LenStr - .GrdLen, .GrdLen, LenStr - Length, .Payout, .TopTension, .AnchPull, .FLAngle, .BtmAngle, SegPosition(1), SegLength, SegTension, SegAngle, SegPosition, CatX, CatY, Connector)
        End With

        Call VB6.ShowForm(frmCatenary, 1, Me)

    End Sub

    Private Sub btnNew_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnNew.Click

        Dim Segment_Renamed As Segment
        Dim MoorLine_Renamed As MoorLine
        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single

        MoorLines.MoorLineAdd()
        NumLines = NumLines + 1
        InitiateTabs()

        MoorLine_Renamed = MoorLines.MoorLines(CurLine)

        With MoorLines.MoorLines(NumLines)
            .BottomSlope = MoorLine_Renamed.BottomSlope
            .Connected = MoorLine_Renamed.Connected
            .DesScope = MoorLine_Renamed.DesScope

            .Payout = MoorLine_Renamed.Payout
            .PayoutSur = MoorLine_Renamed.PayoutSur
            .PayoutOpr = MoorLine_Renamed.PayoutOpr
            .PretensionSur = MoorLine_Renamed.PretensionSur
            .PretensionOpr = MoorLine_Renamed.PretensionOpr

            .FairLead.Xs = MoorLine_Renamed.FairLead.Xs
            .FairLead.Ys = MoorLine_Renamed.FairLead.Ys
            .FairLead.z = MoorLine_Renamed.FairLead.z
            .FairLead.Node = MoorLine_Renamed.FairLead.Node

            .Anchor.Xg = MoorLine_Renamed.Anchor.Xg
            .Anchor.Yg = MoorLine_Renamed.Anchor.Yg
            .Anchor.Node = MoorLine_Renamed.Anchor.Node
            .WaterDepth = MoorLine_Renamed.WaterDepth

            .Anchor.HoldCap = MoorLine_Renamed.Anchor.HoldCap
            .Anchor.Model = MoorLine_Renamed.Anchor.Model
            .Anchor.Remark = MoorLine_Renamed.Anchor.Remark
            .WinchCap = MoorLine_Renamed.WinchCap

            .SegmentClear()
        End With

        For Each Segment_Renamed In MoorLine_Renamed
            With Segment_Renamed
                SegTp = .SegType
                Lg = .Length
                TLg = .TotalLength
                dia = .Diameter
                BS = .BS
                E1 = .E1
                E2 = .E2
                DryWt = .UnitDryWeight
                WetWt = .UnitWetWeight
                Buoy = .Buoy
                BuoyL = .BuoyLength
                FrCoef = .FrictionCoef
            End With
            Call MoorLines.MoorLines(NumLines).SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
        Next Segment_Renamed
    End Sub

    Private Sub btnPayout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPayout.Click

        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        Dim TopTension As Single
        Dim Scope, POL As Single
        '
        '    If btnAnchor.Enabled Then btnAnchor_Click
        '    If btnScope.Enabled Then btnScope_Click
        '
        If Not Updated Then
            If Not UpdateMoorData() Then Exit Sub
        End If

        Scope = CSng(txtGeneral(1).Text) / LFactor
        TopTension = CSng(txtTopTen.Text) * 1000.0# / FrcFactor
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        POL = MoorLines.MoorLines(CurLine).PayoutByScopeTopTension(Scope, TopTension)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If POL <= 0# Then
            MsgBox("Program may not find wire payout for the given top tension and scope!", MsgBoxStyle.OkOnly, "MARS - Warning")
            If POL = -1.0# Then
                txtGeneral(2).Text = "0.000"
            Else
                txtGeneral(2).Text = Format(-POL * LFactor, "0.000")
            End If
            TenCalc = False
        Else
            txtGeneral(2).Text = Format(POL * LFactor, "0.000")
            TenCalc = True
        End If

        btnPayout.Enabled = False

    End Sub

    Private Sub btnPreten_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPreten.Click
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        Dim MoorLine_Renamed As MoorLine
        Dim i As Short

        If Not TenCalc Then btnTopTension_Click(btnTopTension, New System.EventArgs())

        With MoorLines.MoorLines(CurLine)
            Select Case cboDraft.SelectedIndex
                Case 1
                    .PretensionSur = CSng(txtTopTen.Text) * 1000.0# / FrcFactor
                    .PayoutSur = CSng(txtGeneral(2).Text) / LFactor
                    txtPreTen(0).Text = txtTopTen.Text

                    If chkBatch.CheckState Then
                        For i = 1 To NumLines
                            If i <> CurLine Then
                                MoorLine_Renamed = MoorLines.MoorLines(i)

                                MoorLine_Renamed.PretensionSur = .PretensionSur
                                MoorLine_Renamed.PayoutSur = .PayoutSur
                            End If
                        Next
                    End If

                Case 2
                    .PretensionOpr = CSng(txtTopTen.Text) * 1000.0# / FrcFactor
                    .PayoutOpr = CSng(txtGeneral(2).Text) / LFactor
                    txtPreTen(1).Text = txtTopTen.Text

                    If chkBatch.CheckState Then
                        For i = 1 To NumLines
                            If i <> CurLine Then
                                MoorLine_Renamed = MoorLines.MoorLines(i)

                                MoorLine_Renamed.PretensionOpr = .PretensionOpr
                                MoorLine_Renamed.PayoutOpr = .PayoutOpr
                            End If
                        Next
                    End If
            End Select

        End With

    End Sub

    Private Sub btnRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRemove.Click

        MoorLines.MoorLineDelete((CurLine))
        RemoveTabs()

        If NumLines = 1 Then
            MoorLines.MoorLineAdd()
        Else
            NumLines = NumLines - 1
            If CurLine > NumLines Then CurLine = 1
        End If
        LoadMoorData()
    End Sub

    Private Sub btnScope_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnScope.Click
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        With ShipDesLoc
            .Xg = CSng(txtVslSt(0).Text) / LFactor
            .Yg = CSng(txtVslSt(1).Text) / LFactor
            .Heading = CSng(txtVslSt(2).Text) * Degrees2Radians
        End With

        With MoorLines.MoorLines(CurLine)
            .Anchor.Xg = CSng(txtAnchor(0).Text) / LFactor
            .Anchor.Yg = CSng(txtAnchor(1).Text) / LFactor
            .FairLead.Xs = CSng(txtFL(0).Text) / LFactor
            .FairLead.Ys = CSng(txtFL(1).Text) / LFactor
            .FairLead.Node = CInt(_txtFLNode.Text)
            .Anchor.Node = CInt(_txtAnchorNode.Text)
            txtGeneral(1).Text = Format(.ScopeByVesselLocation(ShipDesLoc, True) * LFactor, "0.000")
            txtGeneral(0).Text = Format(.SprdAngle(ShipDesLoc, True) * Radians2Degrees, "0.000")
        End With

        btnAnchor.Enabled = False
        btnScope.Enabled = False

    End Sub

    Private Sub btnTopTension_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnTopTension.Click

        Dim TopTension, HorForce As Single
        Dim Scope, POL As Single

        If btnAnchor.Enabled Then btnAnchor_Click(btnAnchor, New System.EventArgs())
        If btnScope.Enabled Then btnScope_Click(btnScope, New System.EventArgs())

        If Not Updated Then
            If Not UpdateMoorData() Then Exit Sub
        End If

        With MoorLines.MoorLines(CurLine)
            Scope = .DesScope
            POL = .Payout
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            TopTension = .TensionByScopePOL(Scope, HorForce, POL)
            Cursor = System.Windows.Forms.Cursors.Default
            txtTopTen.Text = Format(TopTension * 0.001 * FrcFactor, "0.000")
        End With

        TenCalc = True

    End Sub

    ' combobox operation

    Private Sub cboDraft_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDraft.SelectedIndexChanged

        Dim i As Short
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If


        With MoorLines
            Select Case cboDraft.SelectedIndex
                Case 0
                    .ShipDraft = CurDraft
                Case 1
                    .ShipDraft = SurDraft
                Case 2
                    .ShipDraft = OprDraft
            End Select
            txtVslSt(3).Text = Format(.ShipDraft * LFactor, "0.000")
            Select Case cboDraft.SelectedIndex
                Case 0
                    txtTopTen.Text = Format(.MoorLines(CurLine).TopTension * 0.001 * FrcFactor, "0.000")
                Case 1
                    For i = 1 To NumLines
                        If .MoorLines(i).PayoutSur > 0# Then .MoorLines(i).Payout = .MoorLines(i).PayoutSur
                    Next
                    txtTopTen.Text = Format(.MoorLines(CurLine).PretensionSur * 0.001 * FrcFactor, "0.000")
                Case 2
                    For i = 1 To NumLines
                        If .MoorLines(i).PayoutOpr > 0# Then .MoorLines(i).Payout = .MoorLines(i).PayoutOpr
                    Next
                    txtTopTen.Text = Format(.MoorLines(CurLine).PretensionOpr * 0.001 * FrcFactor, "0.000")
            End Select
            txtGeneral(2).Text = Format(.MoorLines(CurLine).Payout * LFactor, "0.000")
        End With

    End Sub

    'grid operation

    ' Private Sub grdSegments_DblClick(ByVal eventSender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdSegments.DoubleClick
    'If e.ColumnIndex = 0 Then '2.1.4
    'If cboSegType.Visible = False Then '2.1.4
    '           ExistingTxt = grdSegments.Text
    '          SetSegTypeCbo((grdSegments.Text))
    'GridCombo(grdSegments, cboSegType, True)
    'End If
    'Else
    '' MSFlexGridEdit(grdSegments, txtSegEdit, System.Windows.Forms.Keys.F2)
    'End If
    'End Sub

    'Private Sub grdSegments_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grdSegments.EnterCell

    '   ExistingTxt = grdSegments.Text
    '  JEC = True

    'End Sub

    ' Private Sub grdSegments_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyDownEvent) Handles grdSegments.KeyDownEvent

    'KeyHandler(grdSegments, txtSegEdit, eventArgs.KeyCode, eventArgs.Shift, JEC, ExistingTxt)

    'End Sub

    Private Sub grdSegments_LeaveCell(ByVal eventSender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles grdSegments.CellLeave

        If e.ColumnIndex = 0 Then
            Dim cell As New DataGridViewTextBoxCell()

            cell.Value = grdSegments(e.ColumnIndex, e.RowIndex).Value
        End If

        Updated = False
        TenCalc = False
        btnPayout.Enabled = True
        Changed = True

    End Sub


    Public Sub mnuInsert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInsert.Click

        AddRow(grdSegments, CheckingGrid)
        Updated = False
        TenCalc = False
        Changed = True

    End Sub

    Public Sub mnuDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDelete.Click


        DeleteRow(grdSegments, CheckingGrid)
        Updated = False
        TenCalc = False
        Changed = True

    End Sub

    Private Sub frm_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown

        If eventArgs.Button = System.Windows.Forms.MouseButtons.Right Then

            ContextMenuStripGridEdit.Show(MousePosition)

        End If

    End Sub

    Private Function CheckEntries(ByVal Entry As String, ByVal Column As Short) As String

        Dim StrLen As Short
        Dim msg As String

        StrLen = Len(Entry)
        Do While StrLen > 0
            If VB.Left(Entry, 1) = " " Then
                StrLen = StrLen - 1
                Entry = VB.Right(Entry, StrLen)
            ElseIf VB.Right(Entry, 1) = " " Then
                StrLen = StrLen - 1
                Entry = VB.Left(Entry, StrLen)
            Else
                Exit Do
            End If
        Loop

        If StrLen > 0 And Column > 1 Then
            Do While Not IsNumeric(Entry) And Entry <> ""
                msg = "'" & Entry & "' is not a valid numberic input. Please re-enter here:"
                Entry = InputBox(msg)
            Loop
        End If

        CheckEntries = Entry

    End Function

    ' text boxes
    Private Sub txtVslSt_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVslSt.TextChanged
        Dim Index As Short = txtVslSt.GetIndex(eventSender)

        Select Case Index
            Case 0, 1, 2
                LocChanged = True
                btnAnchor.Enabled = True
                btnScope.Enabled = True
                btnPayout.Enabled = True
        End Select

    End Sub

    Private Sub txtVslSt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVslSt.Leave
        Dim Index As Short = txtVslSt.GetIndex(eventSender)

        txtVslSt(Index).Text = CheckData(txtVslSt(Index).Text)

    End Sub

    Private Sub txtAnchor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnchor.TextChanged
        Dim Index As Short = txtAnchor.GetIndex(eventSender)
        On Error GoTo Errhandler
        Select Case Index
            Case 0, 1 ' Anchor X, Y
                btnAnchor.Enabled = False
                btnScope.Enabled = True
                btnPayout.Enabled = True
            Case 2 ' Anchor WD
                If Len(Trim(txtWD.Text)) > 0 And Val(txtWD.Text) <> 0 Then
                    ' calc average slope as default
                    If Val(txtWD.Text) > CSng(txtAnchor(2).Text) Then
                        txtAnchor(3).Text = -Format(Atan(Val(txtGeneral(1).Text), Val(txtWD.Text) - Val(txtAnchor(2).Text), True) * 180 / PI, "#0.00")
                    Else
                        txtAnchor(3).Text = Format(Atan(Val(txtGeneral(1).Text), Val(txtAnchor(2).Text) - Val(txtWD.Text), True) * 180 / PI, "#0.00")
                    End If
                End If
        End Select
        Updated = False
        TenCalc = False
        Exit Sub
Errhandler:
        MsgBox("Error " & Err.Description, MsgBoxStyle.Critical, "Error")

    End Sub

    Private Sub txtAnchor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAnchor.Leave
        Dim Index As Short = txtAnchor.GetIndex(eventSender)

        txtAnchor(Index).Text = CheckData(txtAnchor(Index).Text)

    End Sub

    Private Sub txtFL_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFL.TextChanged
        Dim Index As Short = txtFL.GetIndex(eventSender)

        Select Case Index
            Case 0, 1
                btnAnchor.Enabled = True
                btnScope.Enabled = True
                btnPayout.Enabled = True
        End Select
        Updated = False
        TenCalc = False

    End Sub

    Private Sub txtFL_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFL.Leave
        Dim Index As Short = txtFL.GetIndex(eventSender)

        txtFL(Index).Text = CheckData(txtFL(Index).Text)

    End Sub

    Private Sub txtGeneral_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGeneral.TextChanged
        Dim Index As Short = txtGeneral.GetIndex(eventSender)

        Select Case Index
            Case 0, 1
                btnAnchor.Enabled = True
                btnScope.Enabled = False
                btnPayout.Enabled = True
        End Select
        Updated = False
        TenCalc = False

    End Sub

    Private Sub txtGeneral_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGeneral.Leave
        Dim Index As Short = txtGeneral.GetIndex(eventSender)

        txtGeneral(Index).Text = CheckData(txtGeneral(Index).Text)

    End Sub

    Private Sub txtTopTen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopTen.TextChanged

        btnPayout.Enabled = True
        Updated = False
        TenCalc = False

        If TenCalc = False Then
            If cboDraft.Text = VB6.GetItemString(cboDraft, 1) Or cboDraft.Text = VB6.GetItemString(cboDraft, 2) Then
                btnPreten.Enabled = True
            End If
        End If
    End Sub

    Private Sub txtTopTen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTopTen.Leave

        txtTopTen.Text = CheckData(txtTopTen.Text)

    End Sub

    Private Sub IniSegments()

        Dim c As Short

        CheckingGrid = True

        Call SetLblSegments()

        With grdSegments
            .RowCount = MaxNumSeg
            For c = 1 To .RowCount
                .Rows(c - 1).HeaderCell.Value = c.ToString()

            Next
            .ColumnCount = 12
            For c = 0 To .ColumnCount - 1
                .Columns(c).FillWeight = 100 / .ColumnCount
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(c).HeaderText = SegLabels(c)
            Next c
            ' .Columns(0).FillWeight = 80 / .ColumnCount
            '.Columns(1).FillWeight = 120 / .ColumnCount
        End With

        With grdAnchor
            .RowCount = 1
            .ColumnCount = 8
            .Columns(0).FillWeight = 100 / 12
            .Columns(1).FillWeight = 100 / 12
            .Columns(2).FillWeight = 300 / 12
            .Columns(3).FillWeight = 300 / 12
            .Columns(4).FillWeight = 100 / 12
            .Columns(5).FillWeight = 100 / 12
            .Columns(6).FillWeight = 200 / 12
            .Columns(7).FillWeight = 100 / 12
            .Rows(0).Cells(0).Value = "Ends"
            .Rows(0).Cells(1).Value = "Anchor"
            .Rows(0).Cells(3).Value = "Holding Capacity ( " & FrcUnit & "):"
            .Rows(0).Cells(6).Value = "Winch Capacity (" & FrcUnit & "):"
            For c = 0 To .ColumnCount - 1
                .Columns(c).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Next c
        End With

        '   segment type
        'With cboSegType '2.1.4
        '.FlatStyle = FlatStyle.Flat
        '.MaxDropDownItems = 7
        ' .Items.Add("WIRE")
        '.Items.Add("CHAIN")
        '.Items.Add("WIRE-P")
        '.Items.Add("CHAIN-P")
        '.Items.Add("POLY-P")
        '.Items.Add("BUOY-P")
        '.Items.Add("SPRING") '2.2.1
        ' End With

        CheckingGrid = False

    End Sub
    ' operation subroutines
    ' initiating

    Private Sub SetLblSegments()

        ' SegLabels(0) = "No"
        SegLabels(0) = "Type"
        SegLabels(1) = "Deployed Length"
        SegLabels(2) = "Total Length"
        SegLabels(3) = "Diameter"
        SegLabels(4) = "B.S."
        SegLabels(5) = "E1"
        SegLabels(6) = "E2"
        SegLabels(7) = "Unit Dry Weight"
        SegLabels(8) = "Unit Wet Weight"
        SegLabels(9) = "Friction Coeff"
        SegLabels(10) = "Buoy Net" ' Buoyancy"
        SegLabels(11) = "Vertical Distance"

        If IsMetricUnit Then
            SegLabels(1) = SegLabels(1) & vbCrLf & "(m)"
            SegLabels(2) = SegLabels(2) & vbCrLf & "(m)"
            SegLabels(3) = SegLabels(3) & vbCrLf & "(cm)"
            SegLabels(4) = SegLabels(4) & vbCrLf & "(KN)"
            SegLabels(5) = SegLabels(5) & vbCrLf & "(MPa)"
            SegLabels(6) = SegLabels(6) & vbCrLf & "(MPa)"
            SegLabels(7) = SegLabels(7) & vbCrLf & "(N/m)"
            SegLabels(8) = SegLabels(8) & vbCrLf & "(N/m)"
            SegLabels(10) = SegLabels(10) & vbCrLf & "(KN)"
            SegLabels(11) = SegLabels(11) & vbCrLf & "(m)"
        Else
            SegLabels(1) = SegLabels(1) & vbCrLf & "(ft)"
            SegLabels(2) = SegLabels(2) & vbCrLf & "(ft)"
            SegLabels(3) = SegLabels(4) & vbCrLf & "(in)"
            SegLabels(4) = SegLabels(5) & vbCrLf & "(kip)"
            SegLabels(5) = SegLabels(6) & vbCrLf & "(ksi)"
            SegLabels(6) = SegLabels(7) & vbCrLf & "(ksi)"
            SegLabels(7) = SegLabels(8) & vbCrLf & "(lb/ft)"
            SegLabels(8) = SegLabels(9) & vbCrLf & "(lb/ft)"
            SegLabels(10) = SegLabels(10) & vbCrLf & "(kip)"
            SegLabels(11) = SegLabels(11) & vbCrLf & "(ft)"
        End If
    End Sub

    Private Sub InitiateTabs()

        Dim i, NumTab As Short

        NumTab = tabMoorLines.TabPages.Count

        If NumLines > NumTab Then
            For i = NumTab + 1 To NumLines

                tabMoorLines.TabPages.Add(i, "Line" & i)

            Next
        End If

        tabMoorLines.SelectedIndex = CurLine - 1

    End Sub

    Private Sub RemoveTabs()

        Dim i As Short

        If NumLines - 1 > 1 Then

            tabMoorLines.TabPages.RemoveAt(CurLine - 1)
            For i = CurLine - 1 To NumLines - 2
                tabMoorLines.TabPages(i).Text = "Line " & (i + 1)
            Next
        End If

    End Sub


    ' operation subroutines
    ' load to and update from form
    Private Sub LoadMoorData(Optional ByVal FirstTime As Boolean = False)
        On Error GoTo Errhandler
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        Dim r, c As Short
        Dim Scope, SprdAng As Single
        Dim NumSeg As Short

        '   set ship location at first time
        If FirstTime Then
            With CurVessel
                txtVslSt(0).Text = Format(.ShipDesGlob.Xg * LFactor, "0.000")
                txtVslSt(1).Text = Format(.ShipDesGlob.Yg * LFactor, "0.000")
                txtWD.Text = Format(.WaterDepth * LFactor, "0.000")
                txtVslSt(2).Text = Format(.ShipDesGlob.Heading * Radians2Degrees, "0.000")
                txtVslSt(3).Text = Format(.ShipDraft * LFactor, "0.000")
            End With

            '       initiate form appearance
            cboDraft.Text = VB6.GetItemString(cboDraft, 0)
            btnPreten.Enabled = False
            '    chkBatch.Value = False
        End If

        '   clear grid
        With grdSegments
            For r = 0 To .RowCount - 1
                For c = 0 To .ColumnCount - 1
                    .Rows(r).Cells(c).Value = ""
                Next
            Next
        End With

        With MoorLines.MoorLines(CurLine)
            '       anchor
            txtAnchor(0).Text = Format(.Anchor.Xg * LFactor, "0.000")
            txtAnchor(1).Text = Format(.Anchor.Yg * LFactor, "0.000")
            txtAnchor(2).Text = Format(.WaterDepth * LFactor, "0.000")
            txtAnchor(3).Text = Format(.BottomSlope * Radians2Degrees, "0.000")

            '       fairlead
            txtFL(0).Text = Format(.FairLead.Xs * LFactor, "0.000")
            txtFL(1).Text = Format(.FairLead.Ys * LFactor, "0.000")
            txtFL(2).Text = Format(.FairLead.z * LFactor, "0.000")
            _txtFLNode.Text = Format(.FairLead.Node, "0")
            _txtAnchorNode.Text = Format(.Anchor.Node, "0")
            '       general
            ' ????? .FairLead.SprdAngle = 0 will cause actual 0 deg angle to be changed to 180 deg ??????
            If .FairLead.SprdAngle = 0# Then
                SprdAng = .SprdAngle(ShipDesLoc) ' calc SpreadAngle from Anchor and FL pts
            Else
                SprdAng = .FairLead.SprdAngle
            End If
            If .DesScope = 0# Then
                Scope = .ScopeByVesselLocation(ShipDesLoc)
            Else
                Scope = .DesScope
            End If

            ' This assumes HD from vessel center, both HD & angle will be affected
            '        GetHorDist HD, Dir, Scope, SprdAng, .FairLead.Xs, .FairLead.Ys
            '        txtGeneral(0) = Format(Dir * Radians2Degrees, "0.00")
            '        txtGeneral(1) = Format(HD * LFactor, "0.0")

            ' This assumes scope from fairlead
            txtGeneral(0).Text = Format(SprdAng * Radians2Degrees, "0.000")
            txtGeneral(1).Text = Format(Scope * LFactor, "0.000")
            If .FairLead.SprdAngle = 0# And .DesScope = 0# Then
                btnAnchor.Enabled = False
                btnScope.Enabled = False

                .FairLead.SprdAngle = SprdAng
                .DesScope = Scope
            Else
                btnAnchor.Enabled = True
                btnScope.Enabled = False
            End If

            txtGeneral(2).Text = Format(.Payout * LFactor, "0.000")

            '       top tensions
            txtTopTen.Text = Format(.TopTension * 0.001 * FrcFactor, "0.000")
            txtPreTen(0).Text = Format(.PretensionSur * 0.001 * FrcFactor, "0.000")
            txtPreTen(1).Text = Format(.PretensionOpr * 0.001 * FrcFactor, "0.000")
            btnPayout.Enabled = False

            '       segments
            CheckingGrid = True

            NumSeg = .SegmentCount

            For r = 1 To NumSeg
                grdSegments.Rows(r - 1).HeaderCell.Value = r.ToString()
                grdSegments.Rows(r - 1).Cells(0).Value = .Segments(r).SegType
                grdSegments.Rows(r - 1).Cells(1).Value = Format(.Segments(r).Length * LFactor, "0.0")
                grdSegments.Rows(r - 1).Cells(2).Value = Format(.Segments(r).TotalLength * LFactor, "0.0")
                grdSegments.Rows(r - 1).Cells(3).Value = Format(.Segments(r).Diameter * DiameterFactor, "0.000")
                grdSegments.Rows(r - 1).Cells(4).Value = Format(.Segments(r).BS * 0.001 * FrcFactor, "0.0")
                grdSegments.Rows(r - 1).Cells(5).Value = Format(.Segments(r).E1 * 0.001 * StressFactor, "0.0")

                grdSegments.Rows(r - 1).Cells(6).Value = Format(.Segments(r).E2 * 0.001 * StressFactor, "0.0")
                grdSegments.Rows(r - 1).Cells(7).Value = Format(.Segments(r).UnitDryWeight * (FrcFactor / LFactor), "0.00")
                grdSegments.Rows(r - 1).Cells(8).Value = Format(.Segments(r).UnitWetWeight * (FrcFactor / LFactor), "0.00")

                grdSegments.Rows(r - 1).Cells(9).Value = Format(.Segments(r).FrictionCoef, "0.00")
                grdSegments.Rows(r - 1).Cells(10).Value = Format(.Segments(r).Buoy / 1000 * FrcFactor, "0.0")
                grdSegments.Rows(r - 1).Cells(11).Value = Format(.Segments(r).BuoyLength * LFactor, "0.0")
            Next
            'grdSegments.RowCount = NumSeg 'JLIU TODO 
            grdAnchor.Rows(0).Cells(2).Value = .Anchor.Model
            grdAnchor.Rows(0).Cells(4).Value = Format(.Anchor.HoldCap * 0.001 * FrcFactor, "0.0")
            grdAnchor.Rows(0).Cells(5).Value = .Anchor.Remark
            grdAnchor.Rows(0).Cells(7).Value = Format(.WinchCap * 0.001 * FrcFactor, "0.0")
            CheckingGrid = False
        End With

        Updated = True
        TenCalc = False
        Exit Sub
Errhandler:
        MsgBox("Error " & Err.Description, MsgBoxStyle.Critical, "Error")


        '   If Val(txtTopTen.Text) < 0.9 Then btnTopTension_Click

    End Sub


    Private Sub UpdateBatchLines()

        Dim Segment_Renamed As Segment
        Dim MoorLine_Renamed As MoorLine
        Dim j, i, NumSeg As Short

        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single

        MoorLine_Renamed = MoorLines.MoorLines(CurLine)
        '    With MoorLine
        '        GetHorDist HD, Dir, .DesScope, .FairLead.SprdAngle, .FairLead.Xs, .FairLead.Ys
        '    End With

        For i = 1 To NumLines
            If i <> CurLine Then
                With MoorLines.MoorLines(i)
                    .BottomSlope = MoorLine_Renamed.BottomSlope
                    .DesScope = MoorLine_Renamed.DesScope
                    '            .DesScope = ScopeByHD(HD, .FairLead.SprdAngle, .FairLead.Xs, .FairLead.Ys)
                    .Payout = MoorLine_Renamed.Payout
                    .WaterDepth = MoorLine_Renamed.WaterDepth
                    .FairLead.z = MoorLine_Renamed.FairLead.z

                    .Anchor.HoldCap = MoorLine_Renamed.Anchor.HoldCap
                    .Anchor.Model = MoorLine_Renamed.Anchor.Model
                    .Anchor.Remark = MoorLine_Renamed.Anchor.Remark
                    .WinchCap = MoorLine_Renamed.WinchCap

                    If Changed Then
                        NumSeg = .SegmentCount
                        For j = 1 To NumSeg
                            .SegmentDelete((1))
                        Next j
                    End If
                End With

                If Changed Then
                    For Each Segment_Renamed In MoorLine_Renamed
                        With Segment_Renamed
                            SegTp = .SegType
                            Lg = .Length
                            TLg = .TotalLength
                            dia = .Diameter
                            BS = .BS
                            E1 = .E1
                            E2 = .E2
                            DryWt = .UnitDryWeight
                            WetWt = .UnitWetWeight
                            Buoy = .Buoy
                            BuoyL = .BuoyLength
                            FrCoef = .FrictionCoef
                        End With
                        Call MoorLines.MoorLines(i).SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
                    Next Segment_Renamed
                End If
            End If
        Next i

    End Sub

    Private Function UpdateMoorData() As Boolean
        If IsMetricUnit Then
            LFactor = 0.3048 ' ft -> m
            FrcFactor = 4.448222 ' kips -> KN
            StressFactor = 6.894757 ' ksi -> MPa
            DiameterFactor = 2.54 ' in -> cm
            LUnit = "m"
            FrcUnit = "KN"
        Else
            LFactor = 1
            FrcFactor = 1
            StressFactor = 1
            DiameterFactor = 1
            LUnit = "ft"
            FrcUnit = "kips"
        End If

        Dim r As Short
        Dim NumRows, NumSegments As Short

        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single
        Dim SprdAng As Single

        UpdateMoorData = True

        If LocChanged Then
            With ShipDesLoc
                .Xg = CSng(txtVslSt(0).Text) / LFactor
                .Yg = CSng(txtVslSt(1).Text) / LFactor
                .Heading = CSng(txtVslSt(2).Text) * Degrees2Radians
            End With
        End If


        With MoorLines.MoorLines(CurLine)

            '       general

            '        HD, Dir are in reference to vessel center for scope and Spread Angle
            '        Dir = txtGeneral(0) * Degrees2Radians
            '        HD = txtGeneral(1)
            SprdAng = CSng(txtGeneral(0).Text) * Degrees2Radians
            .DesScope = CSng(txtGeneral(1).Text) / LFactor ' This assume scope is referenced from fairlead
            .Payout = CSng(txtGeneral(2).Text) / LFactor '       top tension
            .TopTension = CSng(txtTopTen.Text) * 1000.0# / FrcFactor

            '       fairlead
            .FairLead.Xs = CSng(txtFL(0).Text) / LFactor
            .FairLead.Ys = CSng(txtFL(1).Text) / LFactor
            .FairLead.z = CSng(txtFL(2).Text) / LFactor
            .FairLead.Node = CInt(_txtFLNode.Text)
            .Anchor.Node = CInt(_txtAnchorNode.Text)
            '     GetScopeSprd Scope, SprdAng, HD, Dir, .FairLead.Xs, .FairLead.Ys
            .FairLead.SprdAngle = SprdAng

            '       anchor
            .Anchor.Xg = CSng(txtAnchor(0).Text) / LFactor
            .Anchor.Yg = CSng(txtAnchor(1).Text) / LFactor
            .WaterDepth = CSng(txtAnchor(2).Text) / LFactor
            .BottomSlope = CSng(txtAnchor(3).Text) * Degrees2Radians

            .Anchor.Model = grdAnchor.Rows(0).Cells(1).Value

            .Anchor.HoldCap = CSng(grdAnchor.Rows(0).Cells(4).Value) * 1000.0# / FrcFactor

            .Anchor.Remark = grdAnchor.Rows(0).Cells(5).Value

            .WinchCap = CSng(grdAnchor.Rows(0).Cells(7).Value) * 1000.0# / FrcFactor
        End With

        '   segments
        '   avoid triggering the "Leave Cell" event
        If Not Changed Then
            Updated = True
            If chkBatch.CheckState Then UpdateBatchLines()

            Exit Function
        End If

        CheckingGrid = True

        With grdSegments

            '       scan for a blank cell in this row; if one is found, bail out
            NumRows = .RowCount

            '       delete the old input
            NumSegments = MoorLines.MoorLines(CurLine).SegmentCount
            For r = 1 To NumSegments
                MoorLines.MoorLines(CurLine).SegmentDelete((1))
            Next r

            '       insert the updated input
            For r = 0 To NumRows - 1
                If Len(.Rows(r).Cells(0).Value) > 0 Then
                    SegTp = .Rows(r).Cells(0).Value
                    '           store the data from this row in the new segment

                    Lg = CSng(.Rows(r).Cells(1).Value) / LFactor

                    TLg = CSng(.Rows(r).Cells(2).Value) / LFactor

                    'UPGRADE_WARNING: Couldn't resolve default property of object DiameterFactor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    dia = CSng(.Rows(r).Cells(3).Value) / DiameterFactor

                    BS = CSng(.Rows(r).Cells(4).Value) * 1000.0# / FrcFactor

                    E1 = CSng(.Rows(r).Cells(5).Value) * 1000.0# / StressFactor

                    E2 = CSng(.Rows(r).Cells(6).Value) * 1000.0# / StressFactor

                    DryWt = CSng(.Rows(r).Cells(7).Value) / (FrcFactor / LFactor)

                    WetWt = CSng(.Rows(r).Cells(8).Value) / (FrcFactor / LFactor)

                    FrCoef = CSng(.Rows(r).Cells(9).Value)

                    Buoy = CSng(.Rows(r).Cells(10).Value) * 1000 / FrcFactor

                    BuoyL = CSng(.Rows(r).Cells(11).Value) / LFactor


                    Call MoorLines.MoorLines(CurLine).SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
                End If
            Next r
        End With

        '   reset the "Leaving Cell" trigger
        CheckingGrid = False

        If chkBatch.CheckState Then UpdateBatchLines()

        Updated = True

    End Function


    ' operation subroutines
    ' load form and save to project

    Private Sub LoadMoorLines()
        ' This routine makes a copy of the mooring data into MoorLines object

        Dim Segment_Renamed As Segment
        Dim MoorLine_Renamed As MoorLine
        Dim i As Short

        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single

        MoorLines = New MoorSystem

        With CurVessel
            NumLines = .MoorSystem.MoorLineCount
            CurDraft = .ShipDraft
            SurDraft = .ShipDraftSur
            OprDraft = .ShipDraftOpr
        End With

        MoorLines.ShipDraft = CurDraft
        ShipDesLoc = MoorLines.ShipGlob

        With CurVessel.ShipDesGlob
            ShipDesLoc.Xg = .Xg
            ShipDesLoc.Yg = .Yg
            ShipDesLoc.Heading = .Heading
        End With

        For i = 1 To NumLines
            MoorLines.MoorLineAdd()
            MoorLine_Renamed = CurVessel.MoorSystem.MoorLines(i)

            With MoorLines.MoorLines(i)
                .BottomSlope = MoorLine_Renamed.BottomSlope
                .Connected = True
                .DesScope = MoorLine_Renamed.DesScope

                .Payout = MoorLine_Renamed.Payout
                .PayoutSur = MoorLine_Renamed.PayoutSur
                .PayoutOpr = MoorLine_Renamed.PayoutOpr
                .PretensionSur = MoorLine_Renamed.PretensionSur
                .PretensionOpr = MoorLine_Renamed.PretensionOpr

                .FairLead.SprdAngle = MoorLine_Renamed.FairLead.SprdAngle
                .FairLead.Xs = MoorLine_Renamed.FairLead.Xs
                .FairLead.Ys = MoorLine_Renamed.FairLead.Ys
                .FairLead.z = MoorLine_Renamed.FairLead.z
                .FairLead.Node = MoorLine_Renamed.FairLead.Node
                .Anchor.Xg = MoorLine_Renamed.Anchor.Xg
                .Anchor.Yg = MoorLine_Renamed.Anchor.Yg
                .Anchor.Node = MoorLine_Renamed.Anchor.Node

                .WaterDepth = MoorLine_Renamed.WaterDepth

                .Anchor.HoldCap = MoorLine_Renamed.Anchor.HoldCap
                .Anchor.Model = MoorLine_Renamed.Anchor.Model
                .Anchor.Remark = MoorLine_Renamed.Anchor.Remark
                .WinchCap = MoorLine_Renamed.WinchCap

            End With

            For Each Segment_Renamed In MoorLine_Renamed
                With Segment_Renamed
                    SegTp = .SegType
                    Lg = .Length
                    TLg = .TotalLength
                    dia = .Diameter
                    BS = .BS
                    E1 = .E1
                    E2 = .E2
                    DryWt = .UnitDryWeight
                    WetWt = .UnitWetWeight
                    Buoy = .Buoy
                    BuoyL = .BuoyLength
                    FrCoef = .FrictionCoef
                End With
                Call MoorLines.MoorLines(i).SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
            Next Segment_Renamed
        Next

        If NumLines < 1 Then
            NumLines = 1
            MoorLines.MoorLineAdd()
        End If

    End Sub


    Private Sub SaveMoorLines()

        Dim Segment_Renamed As Segment
        Dim MoorLine_Renamed As MoorLine
        Dim DLine, i, j, NumSeg As Short

        Dim SegTp As String
        Dim TLg, Lg, dia As Single
        Dim E1, BS, E2 As Single
        Dim Buoy, DryWt, WetWt, BuoyL As Single
        Dim FrCoef As Single

        With CurVessel.MoorSystem
            DLine = NumLines - .MoorLineCount
            If DLine > 0 Then
                For i = 1 To DLine
                    .MoorLineAdd()
                Next i
            ElseIf DLine < 0 Then
                For i = NumLines - DLine To NumLines + 1 Step -1
                    .MoorLineDelete((i))
                Next i
            End If
        End With

        Dim outTh As Single

        For i = 1 To NumLines Step 1
            MoorLine_Renamed = MoorLines.MoorLines(i)

            With CurVessel.MoorSystem.MoorLines(i)

                NumSeg = .SegmentCount
                For j = 1 To NumSeg
                    .SegmentDelete((1))
                Next j
                For Each Segment_Renamed In MoorLine_Renamed
                    With Segment_Renamed
                        SegTp = .SegType
                        Lg = .Length
                        TLg = .TotalLength
                        dia = .Diameter
                        BS = .BS
                        E1 = .E1
                        E2 = .E2
                        DryWt = .UnitDryWeight
                        WetWt = .UnitWetWeight
                        Buoy = .Buoy
                        BuoyL = .BuoyLength
                        FrCoef = .FrictionCoef
                    End With
                    Call CurVessel.MoorSystem.MoorLines(i).SegmentAdd(SegTp, Lg, TLg, dia, BS, E1, E2, DryWt, WetWt, Buoy, BuoyL, FrCoef)
                Next Segment_Renamed

                .BottomSlope = MoorLine_Renamed.BottomSlope
                .DesScope = MoorLine_Renamed.DesScope

                .Payout = MoorLine_Renamed.Payout
                .PayoutSur = MoorLine_Renamed.PayoutSur
                .PayoutOpr = MoorLine_Renamed.PayoutOpr
                .PretensionSur = MoorLine_Renamed.PretensionSur
                .PretensionOpr = MoorLine_Renamed.PretensionOpr

                .FairLead.SprdAngle = MoorLine_Renamed.FairLead.SprdAngle
                .FairLead.Xs = MoorLine_Renamed.FairLead.Xs
                .FairLead.Ys = MoorLine_Renamed.FairLead.Ys
                .FairLead.z = MoorLine_Renamed.FairLead.z
                .FairLead.node = MoorLine_Renamed.FairLead.node

                .Anchor.Xg = MoorLine_Renamed.Anchor.Xg
                .Anchor.Yg = MoorLine_Renamed.Anchor.Yg
                .Anchor.Node = MoorLine_Renamed.Anchor.Node
                .WaterDepth = MoorLine_Renamed.WaterDepth

                .Anchor.HoldCap = MoorLine_Renamed.Anchor.HoldCap
                .Anchor.Model = MoorLine_Renamed.Anchor.Model
                .Anchor.Remark = MoorLine_Renamed.Anchor.Remark
                .WinchCap = MoorLine_Renamed.WinchCap
            End With

            With CurVessel.MoorSystem.MoorLines(i)
                If MoorLine_Renamed.TopTension = 0 Then
                    .TopTension = MoorLine_Renamed.TensionByScopePOL(.DesScope, outTh, .Payout)
                Else
                    .TopTension = MoorLine_Renamed.TopTension
                End If
            End With
        Next i

    End Sub

    ' programs involves catenary forms

    Friend Sub UpdateCat(Optional ByRef TopTension As Single = 0#, Optional ByRef HorForce As Single = 0#)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        Dim i, NumSeg As Short

        Dim Scope, FrcHor As Single

        Dim Length, LenStr As Single

        Dim SegLength(MaxNumSeg) As Single
        Dim SegTension(MaxNumSeg) As Single
        Dim SegAngle(MaxNumSeg) As Single
        Dim SegPosition(MaxNumSeg) As Single
        Dim CatX(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim CatY(MaxNumSubSeg * MaxNumSeg + 1) As Single
        Dim Connector(MaxNumSeg + 1) As Short
		
		With MoorLines.MoorLines(CurLine)
			If TopTension = 0# Then
				If HorForce = 0# Then
					If Not TenCalc Then btnTopTension_Click(btnTopTension, New System.EventArgs())
					Scope = .DesScope
				Else
					Scope = .ScopeByFrcHorPOL(HorForce, TopTension, .Payout)
				End If
			Else
				Scope = .ScopeByTopTensionPOL(TopTension, FrcHor, .Payout)
			End If
			
			Length = .LineLen
			LenStr = .LineLenStr
			
			NumSeg = .SegmentCount
			If NumSeg > MaxNumSeg Then NumSeg = MaxNumSeg
			For i = 1 To NumSeg Step 1
				With .Segments(i)
					SegLength(i) = .Length
					SegTension(i) = .TenUpp
					SegAngle(i) = .AngUpp
					If i < NumSeg Then SegPosition(i + 1) = .YLow
				End With
			Next i
			SegLength(1) = .Payout
			SegPosition(1) = .Draft - .FairLead.z
			
			If Not .CatenaryPoints(CatX, CatY, Connector) Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                If frmCatenary.Visible Then
					MsgBox("Catenary Calculation failed.")
				End If
				Exit Sub
			End If
			
			frmCatenary.UpdateForm(NumLines, CurLine, NumSeg, Scope, Length, LenStr - .GrdLen, .GrdLen, LenStr - Length, .Payout, .TopTension, .AnchPull, .FLAngle, .BtmAngle, SegPosition(1), SegLength, SegTension, SegAngle, SegPosition, CatX, CatY, Connector)
		End With
		
		TenCalc = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub
End Class