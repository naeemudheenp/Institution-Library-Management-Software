Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class Form5
	Inherits System.Windows.Forms.Form
	Dim counter As Short
	
	Dim c As Short
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim n As Object
		
		Dim a As String
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("DATABASE EMPTY")
			
		Else
			Text1.Text = Adodc1.Recordset.Fields("bookid").Value
			Text2.Text = Text1.Text
			If Adodc2.Recordset.EOF = True And Adodc2.Recordset.BOF = True Then
				GoTo ftr
			Else
				
				
				
				
				Adodc2.Recordset.MoveFirst()
				
				
				Adodc2.Recordset.FIND("bookid = '" & Text2.Text & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
				If Adodc2.Recordset.EOF = True Then
					GoTo ftr
				Else
					
					
					
					'UPGRADE_WARNING: Couldn't resolve default property of object n. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					n = Adodc2.Recordset.Fields("booksleft").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object n. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					n = n + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object n. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Adodc2.Recordset.Fields("booksleft").Value = n
					Adodc2.Recordset.Update()
					If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
						MsgBox("DATABASE EMPTY")
						GoTo lst
					End If
ftr: 
					Adodc1.Recordset.Delete()
					MsgBox("ISSUE DELETED")
lst: 
				End If
			End If
			Adodc1.Refresh()
			Adodc2.Refresh()
		End If
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Hide()
		
		
		
		
		Form1.Show()
		
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Dim Printer As New Printer
		Dim i As Short
		Dim n As Short
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("UNABLE TO PRINT DATABASE EMPTY")
			GoTo SD
		Else
kt: 
			
			
			If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = False Then
				
				GoTo fr
			Else
				Printer.Print("ISSUE LIST")
				Printer.Print("BOOKNAME            STUDENTNAME            DATE\TIME")
				n = 0
st: 
				If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = False Then
					GoTo fr
				Else
					str_Renamed = ""
					For i = 0 To DataGrid1.Columns.Count - 1
						
						DataGrid1.Col = i
						str_Renamed = str_Renamed & DataGrid1.Text & "         "
						
						
					Next 
					
					Printer.Print(str_Renamed)
					n = n + 1
					
					Adodc1.Recordset.MoveNext()
					If n = 50 Then
						Printer.NewPage()
						GoTo kt
					Else
						GoTo st
					End If
				End If
			End If
		End If
fr: 
		Printer.EndDoc()
SD: 
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		Adodc1.Refresh()
		Adodc2.Refresh()
		
	End Sub
	
	Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
		DataGrid2.Visible = False
		Command9.Visible = False
	End Sub
	
	Private Sub Form5_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		If Adodc1.Recordset.EOF = True And Adodc1.Recordset.BOF = True Then
			MsgBox("ISSUE LIST EMPTY")
			
		End If
		
		
		counter = 1
		Picture1.Image = System.Drawing.Image.FromFile("E:\software\images\image" & counter & ".jpg")
		Timer1.Interval = 5000 '5 seconds
		Timer1.Enabled = True
	End Sub
	
	'UPGRADE_WARNING: Event Text3.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text3.TextChanged
		Adodc1.Refresh()
		
		Adodc3.Refresh()
		Adodc3.RecordSource = "Select * from issue where studentname Like '" & Text3.Text & "%' OR Bookid Like '" & Text3.Text & "%' "
		'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		DataGrid2.CtlRefresh()
		Dim fnd As String
		If Adodc3.Recordset.EOF = True Then
			GoTo fr
		Else
			DataGrid2.Visible = True
			
			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			DataGrid2.CtlRefresh()
			fnd = Adodc3.Recordset.Fields("bookid").Value
			
			Adodc1.Recordset.FIND("bookid = '" & fnd & "'", 0, ADODB.SearchDirectionEnum.adSearchForward)
			
			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			DataGrid1.CtlRefresh()
		End If
fr: 
		If DataGrid2.Visible = True Then
			Command9.Visible = True
		End If
	End Sub
	
	Private Sub Text3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text3.Click
		Text3.Text = ""
		
	End Sub
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		counter = counter + 1
		If counter > 6 Then counter = 1
		Picture1.Image = System.Drawing.Image.FromFile("E:\software\images\image" & counter & ".jpg")
	End Sub
End Class