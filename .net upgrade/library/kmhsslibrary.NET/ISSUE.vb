Option Strict Off
Option Explicit On
Friend Class Form4
	Inherits System.Windows.Forms.Form
	Dim c As Short
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Adodc1.Refresh()
		Adodc2.Refresh()
		
		Text4.Text = Text2.Text
		
		
		Dim a As String
		Dim k As Short
		Dim tb2 As Short
		Dim lm As Short
		
		Dim i As Short
		
		a = Text1.Text
		Dim h As String
		Dim n As Short
		Dim o As Short
		Dim j As Short
		
		Dim f As String
		If a = "" Then
			MsgBox("You Cannot Leave The Box Empty")
		Else
			
			Adodc2.Recordset.FIND("bookid = '" & Text4.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			If Adodc2.Recordset.EOF Then
				MsgBox("BOOK NOT FOUND WITH THIS ID")
			Else
				f = Adodc2.Recordset.Fields("booksleft").Value
				f = CStr(CDbl(f) - 1)
				Adodc2.Recordset.Fields("booksleft").Value = f
				Adodc2.Recordset.Update()
				
				
				Adodc1.Recordset.AddNew()
				Adodc1.Recordset.Fields("studentname").Value = Text1.Text
				Adodc1.Recordset.Fields("bookid").Value = Text2.Text
				Adodc1.Recordset.Fields("Time").Value = Text3.Text
				Adodc1.Recordset.Update()
				MsgBox("BOOK ISSUED")
				Text1.Text = ""
				Text2.Text = ""
				Adodc2.Recordset.MoveFirst()
				
				
				
			End If
		End If
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		
		
		
		
		
		Form1.Show()
	End Sub
	
	
	
	Private Sub Form4_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Adodc1.Refresh()
		
		Dim dt As Date
		Dim tm As Date
		
		dt = Today
		
		dt = dt
		Text3.Text = CStr(dt)
	End Sub
	
	Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
		Dim too As String
		too = Text1.Text
		Text1.Text = ""
		Text1.Text = too
		
	End Sub
End Class