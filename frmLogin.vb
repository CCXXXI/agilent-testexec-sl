Option Strict Off
Option Explicit On
Friend Class frmLogin
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
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
	Public WithEvents txtPassWord As System.Windows.Forms.TextBox
    Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblPassword As System.Windows.Forms.Label
	Public WithEvents lblName As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents rtbName As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtPassWord = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.rtbName = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'txtPassWord
        '
        Me.txtPassWord.AcceptsReturn = True
        Me.txtPassWord.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassWord.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPassWord.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassWord.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassWord.Location = New System.Drawing.Point(37, 98)
        Me.txtPassWord.MaxLength = 0
        Me.txtPassWord.Name = "txtPassWord"
        Me.txtPassWord.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassWord.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPassWord.Size = New System.Drawing.Size(219, 30)
        Me.txtPassWord.TabIndex = 3
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(175, 158)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(97, 40)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(24, 158)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(97, 40)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'lblPassword
        '
        Me.lblPassword.BackColor = System.Drawing.SystemColors.Control
        Me.lblPassword.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPassword.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassword.Location = New System.Drawing.Point(38, 75)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPassword.Size = New System.Drawing.Size(219, 21)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "Password"
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.SystemColors.Control
        Me.lblName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblName.Location = New System.Drawing.Point(38, 9)
        Me.lblName.Name = "lblName"
        Me.lblName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblName.Size = New System.Drawing.Size(219, 21)
        Me.lblName.TabIndex = 0
        Me.lblName.Text = "Name"
        '
        'rtbName
        '
        Me.rtbName.Location = New System.Drawing.Point(38, 39)
        Me.rtbName.Name = "rtbName"
        Me.rtbName.Size = New System.Drawing.Size(221, 30)
        Me.rtbName.TabIndex = 1
        Me.rtbName.Text = ""
        '
        'frmLogin
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(492, 299)
        Me.Controls.Add(Me.rtbName)
        Me.Controls.Add(Me.txtPassWord)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.lblName)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(306, 309)
        Me.Name = "frmLogin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmLogin
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmLogin
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmLogin()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	'*****************************************************************************
	'frmLogin.frm
	'frmLogin.frx
	'
	'These files support the logging in of operators.
	'*****************************************************************************
	
	Private LocalSource As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click 'TxSLErrorTrapSite
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":cmdOK_Click"
		
		'Toggle the buttons, so that they know something happening
		cmdOK.Enabled = False
		cmdCancel.Enabled = False
        'If frmMain.TestExecSL1.Security.Login(rtbName.Text, txtPassWord.Text) Then
        frmMain.AfterSuccessfulLogin()
        'reset all the fields
        rtbName.Text = ""
        txtPassWord.Text = ""
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        'now hide the form
        Hide()
        'Else
        'Beep()
        'reenable the buttons back, so that they can be used again
        'cmdOK.Enabled = True
        'cmdCancel.Enabled = True
        'rtbName.Focus()
        'End If

        Exit Sub
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
	End Sub
	
    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim assemblyLocation As String = System.Reflection.Assembly.GetExecutingAssembly.Location()
        Dim fileVersionInfo As System.Diagnostics.FileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation)

        'Me.Text = Me.Text & "  Version " & CStr(fileVersionInfo.FileMajorPart) & "." & CStr(fileVersionInfo.FileMinorPart) & "." & CStr(fileVersionInfo.FileVersion) 'NoLocalize
        Me.Text = Me.Text & "  Version " & fileVersionInfo.FileVersion
    End Sub

    Private Sub rtbName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles rtbName.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            'this will allow an enter to be used to actually do the login
            cmdOK_Click(cmdOK, New System.EventArgs)

            'These lines will cause it to go to the next field, in the
            'tab order.  This should be the password text box
            'This is NOT how the TxSL development environment dialog works, so we didn't do it this way
            'SendKeys "{tab}"
            'KeyAscii = 0
        End If
    End Sub

    Private Sub txtPassword_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPassWord.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            System.Windows.Forms.SendKeys.Send("{tab}")
            KeyAscii = 0
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

End Class