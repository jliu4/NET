Option Strict Off
Option Explicit On
Friend Class frmBrowseDir
	Inherits System.Windows.Forms.Form
	Private mCtl As System.Windows.Forms.Control
	
	Public WriteOnly Property ControlToUpdate() As System.Windows.Forms.Control
		Set(ByVal Value As System.Windows.Forms.Control)
			mCtl = Value
		End Set
	End Property
	
	Public WriteOnly Property InitPath() As String
		Set(ByVal Value As String)
			On Error Resume Next
			Dir1.Path = Value
		End Set
	End Property
	
	Private Sub btnCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub
	
	Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOK.Click
        'UPGRADE_WARNING: Couldn't resolve default property of object mCtl. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mCtl.Text = Dir1.Path
        Me.Close()
	End Sub
	
	Private Sub Dir1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Change
		File1.Path = Dir1.Path
		File1.Refresh()
	End Sub
	
	Private Sub Drive1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Drive1.SelectedIndexChanged
		Dir1.Path = Drive1.Drive
	End Sub
	
	'UPGRADE_WARNING: Form event frmBrowseDir.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmBrowseDir_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		File1.Refresh()
	End Sub
End Class