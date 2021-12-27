<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form3
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
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents DataGrid1 As AxMSDataGridLib.AxDataGrid
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents tprice As System.Windows.Forms.TextBox
	Public WithEvents tpname As System.Windows.Forms.TextBox
	Public WithEvents tid As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form3))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Command3 = New System.Windows.Forms.Button
		Me.DataGrid1 = New AxMSDataGridLib.AxDataGrid
		Me.Adodc1 = New VB6.ADODC
		Me.Command2 = New System.Windows.Forms.Button
		Me.Command1 = New System.Windows.Forms.Button
		Me.tprice = New System.Windows.Forms.TextBox
		Me.tpname = New System.Windows.Forms.TextBox
		Me.tid = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.ActiveBorder
		Me.Text = "Form3"
		Me.ClientSize = New System.Drawing.Size(387, 184)
		Me.Location = New System.Drawing.Point(8, 31)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
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
		Me.Name = "Form3"
		Me.Command3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command3.Text = "BACK"
		Me.Command3.Size = New System.Drawing.Size(49, 25)
		Me.Command3.Location = New System.Drawing.Point(200, 160)
		Me.Command3.TabIndex = 10
		Me.Command3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command3.BackColor = System.Drawing.SystemColors.Control
		Me.Command3.CausesValidation = True
		Me.Command3.Enabled = True
		Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command3.TabStop = True
		Me.Command3.Name = "Command3"
		DataGrid1.OcxState = CType(resources.GetObject("DataGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DataGrid1.Size = New System.Drawing.Size(297, 65)
		Me.DataGrid1.Location = New System.Drawing.Point(128, 208)
		Me.DataGrid1.TabIndex = 9
		Me.DataGrid1.Name = "DataGrid1"
		Me.Adodc1.Size = New System.Drawing.Size(145, 25)
		Me.Adodc1.Location = New System.Drawing.Point(0, 160)
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
		Me.Adodc1.RecordSource = "select * from products;"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command2.Text = "RESET"
		Me.Command2.Size = New System.Drawing.Size(73, 25)
		Me.Command2.Location = New System.Drawing.Point(248, 160)
		Me.Command2.TabIndex = 8
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
		Me.Command1.Text = "ADD "
		Me.Command1.Size = New System.Drawing.Size(73, 25)
		Me.Command1.Location = New System.Drawing.Point(320, 160)
		Me.Command1.TabIndex = 7
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.tprice.AutoSize = False
		Me.tprice.Size = New System.Drawing.Size(193, 25)
		Me.tprice.Location = New System.Drawing.Point(200, 128)
		Me.tprice.TabIndex = 6
		Me.tprice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tprice.AcceptsReturn = True
		Me.tprice.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.tprice.BackColor = System.Drawing.SystemColors.Window
		Me.tprice.CausesValidation = True
		Me.tprice.Enabled = True
		Me.tprice.ForeColor = System.Drawing.SystemColors.WindowText
		Me.tprice.HideSelection = True
		Me.tprice.ReadOnly = False
		Me.tprice.Maxlength = 0
		Me.tprice.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.tprice.MultiLine = False
		Me.tprice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.tprice.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.tprice.TabStop = True
		Me.tprice.Visible = True
		Me.tprice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.tprice.Name = "tprice"
		Me.tpname.AutoSize = False
		Me.tpname.Size = New System.Drawing.Size(193, 25)
		Me.tpname.Location = New System.Drawing.Point(200, 96)
		Me.tpname.TabIndex = 5
		Me.tpname.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tpname.AcceptsReturn = True
		Me.tpname.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.tpname.BackColor = System.Drawing.SystemColors.Window
		Me.tpname.CausesValidation = True
		Me.tpname.Enabled = True
		Me.tpname.ForeColor = System.Drawing.SystemColors.WindowText
		Me.tpname.HideSelection = True
		Me.tpname.ReadOnly = False
		Me.tpname.Maxlength = 0
		Me.tpname.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.tpname.MultiLine = False
		Me.tpname.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.tpname.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.tpname.TabStop = True
		Me.tpname.Visible = True
		Me.tpname.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.tpname.Name = "tpname"
		Me.tid.AutoSize = False
		Me.tid.Size = New System.Drawing.Size(193, 25)
		Me.tid.Location = New System.Drawing.Point(200, 64)
		Me.tid.TabIndex = 4
		Me.tid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tid.AcceptsReturn = True
		Me.tid.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.tid.BackColor = System.Drawing.SystemColors.Window
		Me.tid.CausesValidation = True
		Me.tid.Enabled = True
		Me.tid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.tid.HideSelection = True
		Me.tid.ReadOnly = False
		Me.tid.Maxlength = 0
		Me.tid.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.tid.MultiLine = False
		Me.tid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.tid.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.tid.TabStop = True
		Me.tid.Visible = True
		Me.tid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.tid.Name = "tid"
		Me.Label4.Text = "PRICE"
		Me.Label4.Size = New System.Drawing.Size(185, 25)
		Me.Label4.Location = New System.Drawing.Point(0, 128)
		Me.Label4.TabIndex = 3
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
		Me.Label3.Text = "PRODUCT NAME"
		Me.Label3.Size = New System.Drawing.Size(185, 25)
		Me.Label3.Location = New System.Drawing.Point(0, 96)
		Me.Label3.TabIndex = 2
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
		Me.Label2.Text = "UNIQUE ID"
		Me.Label2.Size = New System.Drawing.Size(185, 25)
		Me.Label2.Location = New System.Drawing.Point(0, 64)
		Me.Label2.TabIndex = 1
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
		Me.Label1.Text = "                                                           STOCK"
		Me.Label1.Size = New System.Drawing.Size(393, 33)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(Command3)
		Me.Controls.Add(DataGrid1)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Command1)
		Me.Controls.Add(tprice)
		Me.Controls.Add(tpname)
		Me.Controls.Add(tid)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
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