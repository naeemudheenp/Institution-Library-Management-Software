<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLogin
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
	Public WithEvents Adodc2 As VB6.ADODC
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents u2 As System.Windows.Forms.TextBox
	Public WithEvents u1 As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Adodc2 = New VB6.ADODC
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.Adodc1 = New VB6.ADODC
		Me.Command1 = New System.Windows.Forms.Button
		Me.u2 = New System.Windows.Forms.TextBox
		Me.u1 = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "Login"
		Me.ClientSize = New System.Drawing.Size(446, 154)
		Me.Location = New System.Drawing.Point(374, 182)
		Me.Icon = CType(resources.GetObject("frmLogin.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		Me.BackgroundImage = CType(resources.GetObject("frmLogin.BackgroundImage"), System.Drawing.Image)
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmLogin"
		Me.Adodc2.Size = New System.Drawing.Size(97, 49)
		Me.Adodc2.Location = New System.Drawing.Point(128, 104)
		Me.Adodc2.Visible = 0
		Me.Adodc2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc2.ConnectionTimeout = 15
		Me.Adodc2.CommandTimeout = 30
		Me.Adodc2.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc2.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc2.CommandType = ADODB.CommandTypeEnum.adCmdTable
		Me.Adodc2.CacheSize = 50
		Me.Adodc2.MaxRecords = 0
		Me.Adodc2.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc2.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc2.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc2.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc2.Enabled = True
		Me.Adodc2.UserName = ""
		Me.Adodc2.RecordSource = "datagrid"
		Me.Adodc2.Text = "Adodc2"
		Me.Adodc2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc2.Name = "Adodc2"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(81, 33)
		Me.Text1.Location = New System.Drawing.Point(144, 88)
		Me.Text1.TabIndex = 6
		Me.Text1.Text = "Text1"
		Me.Text1.Visible = False
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
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.Picture1.Size = New System.Drawing.Size(145, 97)
		Me.Picture1.Location = New System.Drawing.Point(264, 0)
		Me.Picture1.TabIndex = 5
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.BackColor = System.Drawing.SystemColors.Control
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture1.Name = "Picture1"
		Me.Adodc1.Size = New System.Drawing.Size(97, 49)
		Me.Adodc1.Location = New System.Drawing.Point(32, 72)
		Me.Adodc1.Visible = 0
		Me.Adodc1.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		Me.Adodc1.ConnectionTimeout = 15
		Me.Adodc1.CommandTimeout = 30
		Me.Adodc1.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		Me.Adodc1.LockType = ADODB.LockTypeEnum.adLockOptimistic
		Me.Adodc1.CommandType = ADODB.CommandTypeEnum.adCmdTable
		Me.Adodc1.CacheSize = 50
		Me.Adodc1.MaxRecords = 0
		Me.Adodc1.BOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.BOFActionEnum.adDoMoveFirst
		Me.Adodc1.EOFAction = Microsoft.VisualBasic.Compatibility.VB6.ADODC.EOFActionEnum.adDoMoveLast
		Me.Adodc1.BackColor = System.Drawing.SystemColors.Window
		Me.Adodc1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Adodc1.Orientation = Microsoft.VisualBasic.Compatibility.VB6.ADODC.OrientationEnum.adHorizontal
		Me.Adodc1.Enabled = True
		Me.Adodc1.UserName = ""
		Me.Adodc1.RecordSource = "login"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\software\kms.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "LOGIN"
		Me.Command1.Size = New System.Drawing.Size(113, 33)
		Me.Command1.Location = New System.Drawing.Point(144, 64)
		Me.Command1.TabIndex = 4
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.u2.AutoSize = False
		Me.u2.Size = New System.Drawing.Size(113, 28)
		Me.u2.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.u2.Location = New System.Drawing.Point(144, 40)
		Me.u2.PasswordChar = ChrW(42)
		Me.u2.TabIndex = 1
		Me.u2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.u2.AcceptsReturn = True
		Me.u2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.u2.BackColor = System.Drawing.SystemColors.Window
		Me.u2.CausesValidation = True
		Me.u2.Enabled = True
		Me.u2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.u2.HideSelection = True
		Me.u2.ReadOnly = False
		Me.u2.Maxlength = 0
		Me.u2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.u2.MultiLine = False
		Me.u2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.u2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.u2.TabStop = True
		Me.u2.Visible = True
		Me.u2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.u2.Name = "u2"
		Me.u1.AutoSize = False
		Me.u1.Size = New System.Drawing.Size(113, 25)
		Me.u1.Location = New System.Drawing.Point(144, 0)
		Me.u1.TabIndex = 0
		Me.u1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.u1.AcceptsReturn = True
		Me.u1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.u1.BackColor = System.Drawing.SystemColors.Window
		Me.u1.CausesValidation = True
		Me.u1.Enabled = True
		Me.u1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.u1.HideSelection = True
		Me.u1.ReadOnly = False
		Me.u1.Maxlength = 0
		Me.u1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.u1.MultiLine = False
		Me.u1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.u1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.u1.TabStop = True
		Me.u1.Visible = True
		Me.u1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.u1.Name = "u1"
		Me.Label2.Text = "PASSWORD"
		Me.Label2.Font = New System.Drawing.Font("Arial", 15!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(169, 25)
		Me.Label2.Location = New System.Drawing.Point(8, 40)
		Me.Label2.TabIndex = 3
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.Text = "USERNAME"
		Me.Label1.Font = New System.Drawing.Font("Arial", 15!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(169, 25)
		Me.Label1.Location = New System.Drawing.Point(8, 0)
		Me.Label1.TabIndex = 2
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Adodc2)
		Me.Controls.Add(Text1)
		Me.Controls.Add(Picture1)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(Command1)
		Me.Controls.Add(u2)
		Me.Controls.Add(u1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class