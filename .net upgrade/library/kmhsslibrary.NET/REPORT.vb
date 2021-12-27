Option Strict Off
Option Explicit On
Friend Class Form8
	Inherits System.Windows.Forms.Form
	Dim c As Short
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Command1.Visible = False
		
		
		
		
		
		
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
		
		Command1.Visible = True
		
		
		
		
	End Sub
	
	Private Sub Command3_Click()
		
		Command1.Visible = False
		
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
		Command1.Visible = True
		
	End Sub
	
	Private Sub Command4_Click()
		Form3.Command4_Click(Nothing, New System.EventArgs())
		
		
		
		
		
		
	End Sub
	
	Private Sub Command5_Click()
		Form5.Command1.Visible = False
		Form5.Command2.Visible = False
		
		
		
		'UPGRADE_ISSUE: PrintForm Component might require to be declared. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6D42BDC5-2EFB-44F3-942C-EAB9A2DB05F0"'
		Form5.PrintForm1.Print(Form5, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
		Form5.Command1.Visible = True
		Form5.Command2.Visible = True
		
		
		
		
		
	End Sub
	
	Private Sub Command6_Click()
		Dim dd As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object dd.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dd.Visible = True
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object dd.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dd.Height = 10000000
	End Sub
	
	Private Sub Command2_Click()
		Me.Hide()
		
		
		
		
		Form1.Show()
		
	End Sub
	
	Private Sub Form8_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim sum As Short
		Dim i As Short
		Dim j As Short
		Dim lft As Short
		Dim t As Short
		Dim price As Short
		Dim p As Short
		Dim iss As Short
		
		
		Adodc1.Refresh()
		iss = DataGrid2.ApproxCount
		
		Dim dt As Date
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			
			
			
			sum = 0
			lft = 0
			price = 0
			GoTo l
		Else
			For i = 0 To DataGrid1.ApproxCount - 1
				t = Adodc1.Recordset.Fields("booksleft").Value
				j = Adodc1.Recordset.Fields("qty").Value
				p = Adodc1.Recordset.Fields("price").Value
				lft = lft + t
				sum = sum + j
				price = price + p
				If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
					GoTo l
					If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
					End If
				End If
				Adodc1.Recordset.MoveNext()
			Next 
			
l: 
			Text1.Text = CStr(sum)
			Text2.Text = CStr(lft)
			Text3.Text = CStr(price)
			Text4.Text = CStr(iss)
			dt = Now
			
			
			
			Text5.Text = CStr(dt)
			
			
			
			
			
		End If
		
	End Sub
End Class