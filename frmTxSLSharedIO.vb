Option Strict Off
Option Explicit On
Friend Class frmTxSLSharedIO
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
    'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 24)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(424, 88)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'frmTxSLSharedIO
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(639, 459)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(43, 75)
        Me.Name = "frmTxSLSharedIO"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TxSL Shared IO"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmTxSLSharedIO
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmTxSLSharedIO
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmTxSLSharedIO()
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
	'frmTxSLSharedIO.frm
	'frmTxSLSharedIO.frx
	'
	'The frmTxSLSharedIO.frm is provided as a container to hold
	'code that will be used to control IO that is normally controlled
	'by the test system.  Since we don't know what IO is available, we
	'have only provided some structure here, but don't actually do
	'any control.  A typical application would be to set a pass/fail
	'indicator upon pass or fail of a UUT, or to light a red indicator,
	'should operator assistance be required.
	
	'Also note that the subs here mirror the event names of the TxSL control, and of
	'the main form.  This is to provide some consistency in relating subs here with
	'the execution of the rest of the program.  This is a recommended approach, but not
	'required.
	'*****************************************************************************
	
	Private LocalSource As String
	Public Sub AfterLogin()
		
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":AfterLogin " 'Used to track where the error happened
		'modify to send user messages
		Exit Sub
		
LocalErrorHandler: 
		frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
		
	End Sub
	
	Public Sub AfterTestplanStop(ByVal Reason As HPTestExecSL.TestplanState)
		'It is most likely that automation would be necessary following the
		'AfterTestplanStop event from TxSL
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":AfterTestplanStop " 'Used to track where the error happened
		
		Select Case Reason
			Case HPTestExecSL.TestplanState.TestplanPassed
				'Set bits appropriate for a passed UUT
			Case HPTestExecSL.TestplanState.TestplanFailed
				'Set bits appropriate for a failed UUT
			Case Else
				'If we get anything other than a pass or fail, we will probably want to fail the UUT
		End Select
		
		Exit Sub
		
LocalErrorHandler: 
		frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
		
	End Sub
	
	Private Sub frmTxSLSharedIO_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Configure the flags to whatever they should be when this form is loaded.
		'This may be nothing
		
	End Sub
	
	Public Sub AfterTestplanUnload()
		'If you want to do something After the Testplan has been unloaded, put the
		'statements here.
		
	End Sub
	
	Private Sub frmTxSLSharedIO_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		
		'Configure the flags to whatever they should be when this form is loaded
		'Since this is what would likely happen after an orderly shutdown, this may
		'be a good time to ask for assistance.
		'
		
	End Sub
End Class