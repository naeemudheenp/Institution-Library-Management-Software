Option Strict Off
Option Explicit On
Friend Class Form7
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Me.Hide()
		Form2.Show()
	End Sub
	
	Private Sub Form7_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Object
		
		Text1.Text = Form5.Calendar1.Value
		Adodc1.RecordSource = "Select * FROM sales where time < '" & Text1.Text & "' "
		Adodc1.Refresh()
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid1.CtlRefresh()
		Dim psum As Short
		Dim pqty As Short
		'UPGRADE_NOTE: val was upgraded to val_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim val_Renamed As String
		Dim qval As String
		psum = 0
		pqty = 0
		For i = 0 To DataGrid1.ApproxCount - 1
			val_Renamed = Adodc1.Recordset.Fields("price").Value
			qval = Adodc1.Recordset.Fields("qty").Value
			psum = CDbl(val_Renamed) + psum
			pqty = CDbl(qval) + pqty
			If Adodc1.Recordset.EOF = True Then
				GoTo fr
			End If
			Adodc1.Recordset.MoveNext()
		Next 
fr: 
		Text2.Text = CStr(psum)
		Text3.Text = CStr(pqty)
	End Sub
End Class