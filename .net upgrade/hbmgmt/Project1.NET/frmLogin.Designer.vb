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
	Public WithEvents Picture2 As System.Windows.Forms.PictureBox
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents txtusername As System.Windows.Forms.TextBox
	Public WithEvents Adodc1 As VB6.ADODC
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents Shape1 As Microsoft.VisualBasic.PowerPacks.RectangleShape
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.Picture2 = New System.Windows.Forms.PictureBox
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.txtusername = New System.Windows.Forms.TextBox
		Me.Adodc1 = New VB6.ADODC
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.Shape1 = New Microsoft.VisualBasic.PowerPacks.RectangleShape
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.SystemColors.Window
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Text = "Login"
		Me.ClientSize = New System.Drawing.Size(207, 205)
		Me.Location = New System.Drawing.Point(186, 206)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmLogin"
		Me.Picture2.BackColor = System.Drawing.SystemColors.Window
		Me.Picture2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture2.Size = New System.Drawing.Size(11, 12)
		Me.Picture2.Location = New System.Drawing.Point(184, 8)
		Me.Picture2.Image = CType(resources.GetObject("Picture2.Image"), System.Drawing.Image)
		Me.Picture2.TabIndex = 4
		Me.Picture2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture2.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture2.CausesValidation = True
		Me.Picture2.Enabled = True
		Me.Picture2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture2.TabStop = True
		Me.Picture2.Visible = True
		Me.Picture2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture2.Name = "Picture2"
		Me.Picture1.BackColor = System.Drawing.SystemColors.Window
		Me.Picture1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Picture1.Size = New System.Drawing.Size(73, 25)
		Me.Picture1.Location = New System.Drawing.Point(64, 128)
		Me.Picture1.Image = CType(resources.GetObject("Picture1.Image"), System.Drawing.Image)
		Me.Picture1.TabIndex = 2
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Picture1.Name = "Picture1"
		Me.txtusername.AutoSize = False
		Me.txtusername.BackColor = System.Drawing.Color.White
		Me.txtusername.ForeColor = System.Drawing.SystemColors.ActiveBorder
		Me.txtusername.Size = New System.Drawing.Size(153, 19)
		Me.txtusername.Location = New System.Drawing.Point(24, 56)
		Me.txtusername.TabIndex = 0
		Me.txtusername.Text = "                 username"
		Me.txtusername.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtusername.AcceptsReturn = True
		Me.txtusername.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtusername.CausesValidation = True
		Me.txtusername.Enabled = True
		Me.txtusername.HideSelection = True
		Me.txtusername.ReadOnly = False
		Me.txtusername.Maxlength = 0
		Me.txtusername.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtusername.MultiLine = False
		Me.txtusername.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtusername.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtusername.TabStop = True
		Me.txtusername.Visible = True
		Me.txtusername.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtusername.Name = "txtusername"
		Me.Adodc1.Size = New System.Drawing.Size(80, 22)
		Me.Adodc1.Location = New System.Drawing.Point(224, 80)
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
		Me.Adodc1.RecordSource = "select * from security;"
		Me.Adodc1.Text = "Adodc1"
		Me.Adodc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\MACBETH\HMGMT\product.mdb;Persist Security Info=False"
		Me.Adodc1.Name = "Adodc1"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(153, 19)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(24, 99)
		Me.txtPassword.PasswordChar = ChrW(42)
		Me.txtPassword.TabIndex = 1
		Me.txtPassword.Text = "password"
		Me.txtPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtPassword.CausesValidation = True
		Me.txtPassword.Enabled = True
		Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPassword.HideSelection = True
		Me.txtPassword.ReadOnly = False
		Me.txtPassword.Maxlength = 0
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.MultiLine = False
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPassword.TabStop = True
		Me.txtPassword.Visible = True
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.Shape1.Size = New System.Drawing.Size(201, 201)
		Me.Shape1.Location = New System.Drawing.Point(0, 0)
		Me.Shape1.BackColor = System.Drawing.SystemColors.Window
		Me.Shape1.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Transparent
		Me.Shape1.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Shape1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Shape1.BorderWidth = 1
		Me.Shape1.FillColor = System.Drawing.Color.Black
		Me.Shape1.FillStyle = Microsoft.VisualBasic.PowerPacks.FillStyle.Transparent
		Me.Shape1.Visible = True
		Me.Shape1.Name = "Shape1"
		Me.Label1.BackColor = System.Drawing.Color.White
		Me.Label1.Text = "                SIGN IN"
		Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveBorder
		Me.Label1.Size = New System.Drawing.Size(145, 25)
		Me.Label1.Location = New System.Drawing.Point(32, 16)
		Me.Label1.TabIndex = 3
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(Picture2)
		Me.Controls.Add(Picture1)
		Me.Controls.Add(txtusername)
		Me.Controls.Add(Adodc1)
		Me.Controls.Add(txtPassword)
		Me.ShapeContainer1.Shapes.Add(Shape1)
		Me.Controls.Add(Label1)
		Me.Controls.Add(ShapeContainer1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class