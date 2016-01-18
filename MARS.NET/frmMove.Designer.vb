<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMove
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents btnReport As System.Windows.Forms.Button
	Public WithEvents txtClientName As System.Windows.Forms.TextBox
	Public WithEvents txtLocationName As System.Windows.Forms.TextBox
	Public WithEvents lblClientName As System.Windows.Forms.Label
	Public WithEvents lblLocation As System.Windows.Forms.Label
	Public WithEvents fraReport As System.Windows.Forms.GroupBox
	Public WithEvents _txtExtLoad_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_3 As System.Windows.Forms.TextBox
	Public WithEvents _lblForceUnit_6 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_8 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_5 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_4 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_3 As System.Windows.Forms.Label
	Public WithEvents _lblExtUnits_3 As System.Windows.Forms.Label
	Public WithEvents _fraExternalLoad_1 As System.Windows.Forms.GroupBox
	Public WithEvents btnSave As System.Windows.Forms.Button
    'Public WithEvents SysInfo1 As AxSysInfoLib.AxSysInfo
    Public WithEvents btnCancel As System.Windows.Forms.Button
	Public WithEvents _txtExtLoad_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtExtLoad_2 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_7 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblExtUnits_2 As System.Windows.Forms.Label
	Public WithEvents _lblForceUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_0 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_1 As System.Windows.Forms.Label
	Public WithEvents _lblExtLoad_2 As System.Windows.Forms.Label
	Public WithEvents _fraExternalLoad_0 As System.Windows.Forms.GroupBox
	Public WithEvents _btnPayout_0 As System.Windows.Forms.Button
	Public WithEvents _btnPayout_1 As System.Windows.Forms.Button
    'Public WithEvents grdPayout As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents fraPayout As System.Windows.Forms.GroupBox
	Public WithEvents txtConditions As System.Windows.Forms.TextBox
    'Public WithEvents grdLC As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents _lblForceUnit_5 As System.Windows.Forms.Label
	Public WithEvents lblConditions As System.Windows.Forms.Label
	Public WithEvents fraConditions As System.Windows.Forms.GroupBox
	Public WithEvents btnPosition As System.Windows.Forms.Button
	Public WithEvents _txtVslSt_9 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_8 As System.Windows.Forms.TextBox
	Public WithEvents _cboWells_1 As System.Windows.Forms.ComboBox
	Public WithEvents _optInputSystem_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optInputSystem_2 As System.Windows.Forms.RadioButton
	Public WithEvents _txtVslSt_10 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_7 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_6 As System.Windows.Forms.TextBox
	Public WithEvents _lblLengthUnit_6 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_5 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_4 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_9 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_7 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_10 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_8 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_6 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_5 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc1 As System.Windows.Forms.GroupBox
	Public WithEvents _optInputSystem_0 As System.Windows.Forms.RadioButton
	Public WithEvents _cboWells_0 As System.Windows.Forms.ComboBox
	Public WithEvents _txtVslSt_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtVslSt_5 As System.Windows.Forms.TextBox
	Public WithEvents _optInputSystem_1 As System.Windows.Forms.RadioButton
	Public WithEvents _lblLengthUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_2 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_2 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_3 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_4 As System.Windows.Forms.Label
	Public WithEvents _lblLengthUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_0 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_1 As System.Windows.Forms.Label
	Public WithEvents _lblVslSt_3 As System.Windows.Forms.Label
	Public WithEvents _lblVslStUnit_4 As System.Windows.Forms.Label
	Public WithEvents fraVesselLoc As System.Windows.Forms.GroupBox
	Public WithEvents _lblForceUnit_7 As System.Windows.Forms.Label
	Public WithEvents _lblVelUnit_0 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_3 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_2 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_1 As System.Windows.Forms.Label
	Public WithEvents _lblGenCmt_0 As System.Windows.Forms.Label
	Public WithEvents btnPayout As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents cboWells As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents fraExternalLoad As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents lblExtLoad As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblExtUnits As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblForceUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblGenCmt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblLengthUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVelUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslSt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblVslStUnit As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optInputSystem As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents txtExtLoad As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtVslSt As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'frmMove
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "frmMove"
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class