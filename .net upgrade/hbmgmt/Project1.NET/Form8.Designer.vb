<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form8
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		VB6_AddADODataBinding()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			VB6_RemoveADODataBinding()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text3 As System.Windows.Forms.TextBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form8))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.Text3 = New System.Windows.Forms.TextBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc1 = New VB6.ADODC
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Form8"
		Me.ClientSize = New System.Drawing.Size(937, 513)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form8"
		Me.Text1.AutoSize = False
		Me.Text1.Enabled = False
		Me.Text1.Size = New System.Drawing.Size(273, 25)
		Me.Text1.Location = New System.Drawing.Point(272, 32)
		Me.Text1.TabIndex = 4
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(161, 41)
		Me.Text2.Location = New System.Drawing.Point(776, 120)
		Me.Text2.TabIndex = 2
		Me.Text2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.BackColor = System.Drawing.SystemColors.Window
		Me.Text2.CausesValidation = True
		Me.Text2.Enabled = True
		Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text2.HideSelection = True
		Me.Text2.ReadOnly = False
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text2.Name = "Text2"
		Me.Text3.AutoSize = False
		Me.Text3.Size = New System.Drawing.Size(161, 41)
		Me.Text3.Location = New System.Drawing.Point(776, 208)
		Me.Text3.TabIndex = 1
		Me.Text3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text3.AcceptsReturn = True
		Me.Text3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text3.BackColor = System.Drawing.SystemColors.Window
		Me.Text3.CausesValidation = True
		Me.Text3.Enabled = True
		Me.Text3.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text3.HideSelection = True
		Me.Text3.ReadOnly = False
		Me.Text3.Maxlength = 0
		Me.Text3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text3.MultiLine = False
		Me.Text3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text3.TabStop = True
		Me.Text3.Visible = True
		Me.Text3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text3.Name = "Text3"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "MAINMENU"
		Me.Command1.Size = New System.Drawing.Size(161, 49)
		Me.Command1.Location = New System.Drawing.Point(776, 248)
		Me.Command1.TabIndex = 0
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(777, 457)
		Me.DataGrid1.Location = New System.Drawing.Point(0, 56)
		Me.DataGrid1.TabIndex = 3
		Me.DataGrid1.Name = "DataGrid1"
		Me.Adodc1.Size = New System.Drawing.Size(129, 22)
		Me.Adodc1.Location = New System.Drawing.Point(40, 288)
		Me.Adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc1.ConnectionTimeout = 15
		Me.Adodc1.CommandTimeout = 30
		Me.Adodc1.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc1.CommandType = ADODB.CommandTypeEnum.adCmdUnknown
		Me.Adodc1.CacheSize = 50
		Me.Adodc1.MaxRecords = 0
		Me.Adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc1.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc1.Enabled = True
		Me.Adodc1.UserName = ""
		Me.Adodc1.RecordSource = "select * from sales;"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Label1.Text = "                                                                     REPORT ON SPECIFIED DATE"
		Me.Label1.Size = New System.Drawing.Size(929, 25)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 9
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label2.Text = "SELECTED DATE:"
		Me.Label2.Size = New System.Drawing.Size(265, 17)
		Me.Label2.Location = New System.Drawing.Point(0, 32)
		Me.Label2.TabIndex = 8
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label3.Text = "           TOTAL AMOUNT "
		Me.Label3.Size = New System.Drawing.Size(153, 41)
		Me.Label3.Location = New System.Drawing.Point(784, 72)
		Me.Label3.TabIndex = 7
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label4.Text = "          TOTAL QUANTITY"
		Me.Label4.Size = New System.Drawing.Size(153, 41)
		Me.Label4.Location = New System.Drawing.Point(784, 160)
		Me.Label4.TabIndex = 6
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label5.Text = "            FINAL REPORT"
		Me.Label5.Size = New System.Drawing.Size(145, 33)
		Me.Label5.Location = New System.Drawing.Point(784, 32)
		Me.Label5.TabIndex = 5
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Controls.Add(Text1)
		Me.Controls.Add(Text2)
		Me.Controls.Add(Text3)
		Me.Controls.Add(Command1)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label5)
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
#Region "Upgrade Support"
	Public Sub VB6_AddADODataBinding()
		DataGrid1.DataSource = CType(Adodc1, MSDATASRC.DataSource)
	End Sub
	Public Sub VB6_RemoveADODataBinding()
		DataGrid1.DataSource = Nothing
	End Sub
#End Region 
End Class