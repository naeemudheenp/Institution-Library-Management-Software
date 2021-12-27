<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
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
	Public WithEvents Command12 As System.Windows.Forms.Button
	Public WithEvents Command9 As System.Windows.Forms.Button
	Public WithEvents Command6 As System.Windows.Forms.Button
	Public WithEvents Command11 As System.Windows.Forms.Button
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Command8 As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Command7 As System.Windows.Forms.Button
	Public WithEvents Command5 As System.Windows.Forms.Button
	Public WithEvents Command10 As System.Windows.Forms.Button
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command12 = New System.Windows.Forms.Button
		Me.Command9 = New System.Windows.Forms.Button
		Me.Command6 = New System.Windows.Forms.Button
		Me.Command11 = New System.Windows.Forms.Button
		Me.Adodc1 = New VB6.ADODC
		Me.Command8 = New System.Windows.Forms.Button
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Command7 = New System.Windows.Forms.Button
		Me.Command5 = New System.Windows.Forms.Button
		Me.Command10 = New System.Windows.Forms.Button
		Me.Command4 = New System.Windows.Forms.Button
		Me.Command3 = New System.Windows.Forms.Button
		Me.Command2 = New System.Windows.Forms.Button
		Me.Command1 = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "MAIN SCREEN"
		Me.ClientSize = New System.Drawing.Size(1350, 633)
		Me.Location = New System.Drawing.Point(13, 36)
		Me.Icon = CType(resources.GetObject("Form1.Icon"), System.Drawing.Icon)
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("Form1.BackgroundImage"), System.Drawing.Image)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
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
		Me.Name = "Form1"
		Me.Command12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command12.Text = "SLEEP"
		Me.Command12.Size = New System.Drawing.Size(353, 49)
		Me.Command12.Location = New System.Drawing.Point(0, 632)
		Me.Command12.TabIndex = 14
		Me.Command12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command12.BackColor = System.Drawing.SystemColors.Control
		Me.Command12.CausesValidation = True
		Me.Command12.Enabled = True
		Me.Command12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command12.TabStop = True
		Me.Command12.Name = "Command12"
		Me.Command9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command9.Text = "CLOSE GRID"
		Me.Command9.Size = New System.Drawing.Size(89, 33)
		Me.Command9.Location = New System.Drawing.Point(1000, 0)
		Me.Command9.TabIndex = 13
		Me.Command9.Visible = False
		Me.Command9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command9.BackColor = System.Drawing.SystemColors.Control
		Me.Command9.CausesValidation = True
		Me.Command9.Enabled = True
		Me.Command9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command9.TabStop = True
		Me.Command9.Name = "Command9"
		Me.Command6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command6.Text = "ISSUE LIST"
		Me.Command6.Size = New System.Drawing.Size(137, 177)
		Me.Command6.Location = New System.Drawing.Point(216, 160)
		Me.Command6.TabIndex = 5
		Me.Command6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command6.BackColor = System.Drawing.SystemColors.Control
		Me.Command6.CausesValidation = True
		Me.Command6.Enabled = True
		Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command6.TabStop = True
		Me.Command6.Name = "Command6"
		Me.Command11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command11.Text = "REPORT"
		Me.Command11.Size = New System.Drawing.Size(137, 137)
		Me.Command11.Location = New System.Drawing.Point(216, 496)
		Me.Command11.TabIndex = 9
		Me.Command11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command11.BackColor = System.Drawing.SystemColors.Control
		Me.Command11.CausesValidation = True
		Me.Command11.Enabled = True
		Me.Command11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command11.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command11.TabStop = True
		Me.Command11.Name = "Command11"
		Me.Adodc1.Size = New System.Drawing.Size(105, 57)
		Me.Adodc1.Location = New System.Drawing.Point(640, 200)
		Me.Adodc1.Visible = 0
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
		Me.Adodc1.RecordSource = "SELECT* FROM BOOKLIST;"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Command8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command8.Text = " SEARCH"
		Me.Command8.Size = New System.Drawing.Size(129, 33)
		Me.Command8.Location = New System.Drawing.Point(872, 0)
		Me.Command8.TabIndex = 12
		Me.Command8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command8.BackColor = System.Drawing.SystemColors.Control
		Me.Command8.CausesValidation = True
		Me.Command8.Enabled = True
		Me.Command8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command8.TabStop = True
		Me.Command8.Name = "Command8"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(521, 33)
		Me.Text1.Location = New System.Drawing.Point(352, 0)
		Me.Text1.TabIndex = 0
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
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
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(1089, 697)
		Me.DataGrid1.Location = New System.Drawing.Point(352, 32)
		Me.DataGrid1.TabIndex = 11
		Me.DataGrid1.Visible = False
		Me.DataGrid1.Name = "DataGrid1"
		Me.Command7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command7.Text = "NOT ISSUED BOOKS"
		Me.Command7.Size = New System.Drawing.Size(137, 161)
		Me.Command7.Location = New System.Drawing.Point(216, 336)
		Me.Command7.TabIndex = 7
		Me.Command7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command7.BackColor = System.Drawing.SystemColors.Control
		Me.Command7.CausesValidation = True
		Me.Command7.Enabled = True
		Me.Command7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command7.TabStop = True
		Me.Command7.Name = "Command7"
		Me.Command5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command5.Text = "SETTINGS"
		Me.Command5.Size = New System.Drawing.Size(225, 161)
		Me.Command5.Location = New System.Drawing.Point(-8, 336)
		Me.Command5.TabIndex = 6
		Me.Command5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command5.BackColor = System.Drawing.SystemColors.Control
		Me.Command5.CausesValidation = True
		Me.Command5.Enabled = True
		Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command5.TabStop = True
		Me.Command5.Name = "Command5"
		Me.Command10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command10.Text = "ISSUE"
		Me.Command10.Size = New System.Drawing.Size(137, 129)
		Me.Command10.Location = New System.Drawing.Point(216, 32)
		Me.Command10.TabIndex = 3
		Me.Command10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command10.BackColor = System.Drawing.SystemColors.Control
		Me.Command10.CausesValidation = True
		Me.Command10.Enabled = True
		Me.Command10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command10.TabStop = True
		Me.Command10.Name = "Command10"
		Me.Command4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command4.Text = "ABOUT"
		Me.Command4.Size = New System.Drawing.Size(225, 137)
		Me.Command4.Location = New System.Drawing.Point(-8, 496)
		Me.Command4.TabIndex = 8
		Me.Command4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command4.BackColor = System.Drawing.SystemColors.Control
		Me.Command4.CausesValidation = True
		Me.Command4.Enabled = True
		Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command4.TabStop = True
		Me.Command4.Name = "Command4"
		Me.Command3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command3.BackColor = System.Drawing.SystemColors.Control
		Me.Command3.Text = "EXIT"
		Me.Command3.Size = New System.Drawing.Size(369, 49)
		Me.Command3.Location = New System.Drawing.Point(-16, 680)
		Me.Command3.TabIndex = 10
		Me.Command3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command3.CausesValidation = True
		Me.Command3.Enabled = True
		Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command3.TabStop = True
		Me.Command3.Name = "Command3"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "BOOK LIST "
		Me.Command2.Size = New System.Drawing.Size(217, 177)
		Me.Command2.Location = New System.Drawing.Point(0, 160)
		Me.Command2.TabIndex = 4
		Me.Command2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.Text = "ADD BOOK"
		Me.Command1.Size = New System.Drawing.Size(217, 129)
		Me.Command1.Location = New System.Drawing.Point(0, 32)
		Me.Command1.TabIndex = 2
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Label1.Text = "                    CHOOSE THE SERVICE"
		Me.Label1.Size = New System.Drawing.Size(353, 33)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 1
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
		Me.Controls.Add(Command12)
		Me.Controls.Add(Command9)
		Me.Controls.Add(Command6)
		Me.Controls.Add(Command11)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Command8)
		Me.Controls.Add(Text1)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Command7)
		Me.Controls.Add(Command5)
		Me.Controls.Add(Command10)
		Me.Controls.Add(Command4)
		Me.Controls.Add(Command3)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Label1)
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