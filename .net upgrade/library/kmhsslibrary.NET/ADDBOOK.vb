Option Strict Off
Option Explicit On
Friend Class Form2
	Inherits System.Windows.Forms.Form
	Dim counter As Short
	
	Dim c As Short
	
	
	
	Public Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim i As Object
		
		Dim n As String
		
		Dim value As String
		Dim value1 As String
		Dim value2 As String
		Dim value3 As String
		Dim value4 As String
		Dim value5 As String
		Dim value6 As String
		Dim value7 As String
		Dim limit As Short
		Dim h As String
		value = Text5.Text
		value1 = Text7.Text
		value2 = Text6.Text
		value3 = Text8.Text
		value4 = Text4.Text
		value5 = Text3.Text
		value6 = Text2.Text
		value7 = Text1.Text
		Dim gl As Short
		Dim j As Short
		limit = DataGrid1.ApproxCount
		j = 0
		n = Text4.Text
		
		
		
		
		
		
		If value1 = "" Or value2 = "" Or value3 = "" Or value4 = "" Or value5 = "" Or value6 = "" Or value7 = "" Or value = "" Then
			MsgBox("You Cannot Leave Textbox empty")
		Else
			For i = 0 To limit - 1
				h = Adodc1.Recordset.Fields("bookid").Value
				
				If h = value Then
					MsgBox("Book With This Id Has Been Entered Already")
					GoTo lst
				Else
					Adodc1.Recordset.MoveNext()
				End If
			Next 
			Adodc1.Recordset.AddNew()
			Adodc1.Recordset.Fields("bookname").Value = Text1.Text
			Adodc1.Recordset.Fields("author").Value = Text2.Text
			Adodc1.Recordset.Fields("price").Value = Text3.Text
			Adodc1.Recordset.Fields("qty").Value = Text4.Text
			Adodc1.Recordset.Fields("bookid").Value = Text5.Text
			Adodc1.Recordset.Fields("publication").Value = Text7.Text
			Adodc1.Recordset.Fields("shelf").Value = Text6.Text
			Adodc1.Recordset.Fields("rack").Value = Text8.Text
			Adodc1.Recordset.Fields("booksleft").Value = Text4.Text
			
			
			
			Adodc1.Recordset.Update()
			MsgBox("BOOK ADDED")
			Adodc1.Refresh()
			
			
			Text1.Text = ""
			Text2.Text = ""
			Text3.Text = ""
			Text4.Text = ""
			Text5.Text = ""
			Text6.Text = ""
			Text7.Text = ""
			Text8.Text = ""
		End If
lst: 
		Adodc1.Refresh()
		
	End Sub
	
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		
		
		
		
		Form1.Show()
		
	End Sub
	
	
	Private Sub Form2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Adodc1.Refresh()
		
		
	End Sub
	
	Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
		Dim too As String
		too = Text1.Text
		Text1.Text = ""
		Text1.Text = too
		
	End Sub
	
	Private Sub Text2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.Leave
		Dim too2 As String
		
		too2 = Text2.Text
		Text2.Text = ""
		Text2.Text = too2
	End Sub
	
	Private Sub Text7_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text7.Leave
		Dim too1 As String
		
		too1 = Text7.Text
		Text7.Text = ""
		Text7.Text = too1
	End Sub
End Class